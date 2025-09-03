//  GraphAPIManager.swift
//  Intune Manager
//
//  Created by Eddie Jimenez on 8/13/25.
//

import Foundation
import AppKit
import UniformTypeIdentifiers


// MARK: - Paged response (reduces type-checking complexity)
private struct GraphPage<T: Decodable>: Decodable {
    let value: [T]
    let nextLink: String?

    enum CodingKeys: String, CodingKey {
        case value
        case nextLink = "@odata.nextLink"
    }
}

// MARK: - Managed Apps (Graph beta: users/{id}/mobileAppIntentAndStates)
private struct GraphManagedIntentPage: Decodable {
    let value: [GraphManagedIntentItem]?
}
private struct GraphManagedIntentItem: Decodable {
    let id: String?
    let deviceId: String?
    let mobileAppList: [GraphManagedMobileApp]?
}
private struct GraphManagedMobileApp: Decodable {
    let id: String?
    let displayName: String?
    let publisher: String?
    let version: String?
}


final class GraphAPIManager: ObservableObject {
    // Public state
    @Published var devices: [IntuneDevice] = []
    @Published var users: [EntraUser] = []
    @Published var apps: [DetectedApp] = []  // All apps cache
    @Published var discoveredApps: [DetectedApp] = []  // Discovered apps cache
    @Published var isLoading = false
    @Published var errorMessage: String?
    @Published var loadingProgress = 0.0
    @Published var currentOperation = ""
    
    // Pagination Support
    @Published var hasMoreDevices = false
    @Published var isLoadingMore = false
    private var nextDeviceLink: String?
    private let pageSize = 500  // Match Intune portal behavior
    private var totalDeviceCount: Int?  // Track total count if available

    // Endpoints
    private let baseURL = "https://graph.microsoft.com/v1.0"
    let betaURL = "https://graph.microsoft.com/beta"

    // Simple user cache
    private var userCache: [String: EntraUser] = [:]
    
    // App assignments cache: appId -> [userId]
    private var appAssignmentsCache: [String: [String]] = [:]

    private let jsonDecoder: JSONDecoder = {
        let d = JSONDecoder()
        return d
    }()
    
    // Rate limiting
    private let deviceSemaphore = DispatchSemaphore(value: 5)  // Limit concurrent device requests
    private let userSemaphore = DispatchSemaphore(value: 3)    // Limit concurrent user requests

    // MARK: - Token Validation
    
    /// Check if the access token contains the required scope for Apps operations
    func tokenHasAppsScope(_ accessToken: String) -> Bool {
        // Decode JWT payload to check scopes
        guard let payload = accessToken.split(separator: ".").dropFirst().first else { return false }
        
        let paddedPayload = String(payload)
            .replacingOccurrences(of: "-", with: "+")
            .replacingOccurrences(of: "_", with: "/")
            .padding(toLength: ((String(payload).count+3)/4)*4, withPad: "=", startingAt: 0)
        
        guard let data = Data(base64Encoded: paddedPayload),
              let json = (try? JSONSerialization.jsonObject(with: data)) as? [String: Any],
              let scp = json["scp"] as? String else { return false }
        
        return scp.split(separator: " ").contains { scope in
            scope == "DeviceManagementApps.Read.All" || scope == "DeviceManagementApps.ReadWrite.All"
        }
    }

    // MARK: - Paginated Device Loading

    /// Load initial page of devices
    func fetchInitialDevices(accessToken: String) {
        print("📍 fetchInitialDevices actually called in GraphAPIManager")
        // Reset state
        devices = []
        userCache = [:]
        nextDeviceLink = nil
        hasMoreDevices = false
        totalDeviceCount = nil
        isLoading = true
        errorMessage = nil
        loadingProgress = 0
        currentOperation = "Loading devices..."
        
        // Build initial URL - simplified to avoid filtering issues
        let url = "\(betaURL)/deviceManagement/managedDevices?$top=\(pageSize)"
        fetchDevicesPageWithRateLimit(accessToken: accessToken, url: url, isInitial: true)
    }
    
    /// Load next page of devices
    func fetchMoreDevices(accessToken: String) {
        guard let nextLink = nextDeviceLink, !isLoadingMore else { return }
        
        isLoadingMore = true
        currentOperation = "Loading more devices..."
        fetchDevicesPageWithRateLimit(accessToken: accessToken, url: nextLink, isInitial: false)
    }

    // MARK: - Old Device Fetching (for compatibility)
    
    func fetchAllDevices(accessToken: String) {
        // For backward compatibility, load all devices but with pagination
        fetchInitialDevices(accessToken: accessToken)
        
        // Automatically load all remaining pages
        if hasMoreDevices {
            loadAllRemainingDevices(accessToken: accessToken)
        }
    }
    
    private func loadAllRemainingDevices(accessToken: String) {
        guard hasMoreDevices, let nextLink = nextDeviceLink else {
            isLoading = false
            currentOperation = ""
            return
        }
        
        fetchDevicesPageWithRateLimit(accessToken: accessToken, url: nextLink, isInitial: false) { [weak self] in
                    // After each page, continue loading if there are more
                    if self?.hasMoreDevices == true {
                        self?.loadAllRemainingDevices(accessToken: accessToken)
                    } else {
                        self?.isLoading = false
                        self?.currentOperation = ""
                    }
                }
            }

            /// Fetch device page with rate limit handling
            private func fetchDevicesPageWithRateLimit(accessToken: String, url: String, isInitial: Bool, completion: (() -> Void)? = nil) {
                print("📍 fetchDevicesPageWithRateLimit called with URL: \(url)")
                guard let requestURL = URL(string: url) else {
                    print("❌ Failed to create URL from string: \(url)")
                    finishWithError("Invalid devices URL")
                    completion?()
                    return
                }
                
                var req = URLRequest(url: requestURL)
                req.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
                req.setValue("application/json", forHTTPHeaderField: "Accept")
                
                print("📍 Starting network request...")
                URLSession.shared.dataTask(with: req) { [weak self] data, resp, err in
                    print("📍 Network response received - Error: \(err?.localizedDescription ?? "none")")
                    if let http = resp as? HTTPURLResponse {
                        print("📍 HTTP Status: \(http.statusCode)")
                    }
                    DispatchQueue.main.async {
                        self?.handleDevicesPageWithRetry(
                            data: data,
                            response: resp,
                            error: err,
                            accessToken: accessToken,
                            url: url,
                            isInitial: isInitial,
                            retryCount: 0,
                            completion: completion
                        )
                    }
                }.resume()
            }

            /// Handle device page response with retry logic for rate limiting
            private func handleDevicesPageWithRetry(
                data: Data?,
                response: URLResponse?,
                error: Error?,
                accessToken: String,
                url: String,
                isInitial: Bool,
                retryCount: Int,
                completion: (() -> Void)? = nil
            ) {
                if let error = error {
                    if isInitial {
                        finishWithError("Devices request failed: \(error.localizedDescription)")
                    } else {
                        isLoadingMore = false
                        errorMessage = "Failed to load more devices: \(error.localizedDescription)"
                    }
                    completion?()
                    return
                }
                
                guard let http = response as? HTTPURLResponse else {
                    if isInitial {
                        finishWithError("Invalid devices response")
                    } else {
                        isLoadingMore = false
                        errorMessage = "Invalid response when loading more devices"
                    }
                    completion?()
                    return
                }

                
                // Handle rate limiting (429) with exponential backoff
                if http.statusCode == 429 {
                    let retryAfter = http.value(forHTTPHeaderField: "Retry-After") ?? "10"
                    let baseDelay = Int(retryAfter) ?? 10
                    let delay = min(baseDelay * (retryCount + 1), 60)  // Exponential backoff, max 60 seconds
                    
                    if retryCount < 3 {  // Max 3 retries
                        currentOperation = "Rate limited. Retrying in \(delay) seconds..."
                        
                        DispatchQueue.main.asyncAfter(deadline: .now() + Double(delay)) { [weak self] in
                            self?.fetchDevicesPageWithRateLimit(
                                accessToken: accessToken,
                                url: url,
                                isInitial: isInitial
                            )
                        }
                        return
                    } else {
                        if isInitial {
                            finishWithError("Rate limit exceeded. Please try again later.")
                        } else {
                            isLoadingMore = false
                            errorMessage = "Rate limit exceeded while loading more devices"
                        }
                        completion?()
                        return
                    }
                }
                
                // Handle throttling (503)
                if http.statusCode == 503 {
                    if retryCount < 3 {
                        let delay = 5 * (retryCount + 1)  // 5, 10, 15 seconds
                        currentOperation = "Service unavailable. Retrying in \(delay) seconds..."
                        
                        DispatchQueue.main.asyncAfter(deadline: .now() + Double(delay)) { [weak self] in
                            self?.handleDevicesPageWithRetry(
                                data: data,
                                response: response,
                                error: error,
                                accessToken: accessToken,
                                url: url,
                                isInitial: isInitial,
                                retryCount: retryCount + 1,
                                completion: completion
                            )
                        }
                        return
                    }
                }
                
                guard http.statusCode == 200 else {
                    let message = "Devices API returned \(http.statusCode)"
                    if isInitial {
                        finishWithError(message)
                    } else {
                        isLoadingMore = false
                        errorMessage = message
                    }
                    completion?()
                    return
                }
                
                guard let data = data else {
                    if isInitial {
                        finishWithError("Devices API returned no data")
                    } else {
                        isLoadingMore = false
                        errorMessage = "No data received"
                    }
                    completion?()
                    return
                }
                
                do {
                    print("📍 Attempting to decode response data of size: \(data.count) bytes")
                    
                    // Try to extract total count from response
                    if let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any] {
                        print("📍 JSON parsed successfully")
                        if let count = json["@odata.count"] as? Int {
                            totalDeviceCount = count
                            print("📍 Total device count: \(count)")
                        }
                        
                        // Debug: print keys to see what we got
                        print("📍 Response keys: \(json.keys.joined(separator: ", "))")
                        
                        // Debug: check if value array exists
                        if let valueArray = json["value"] as? [[String: Any]] {
                            print("📍 Found \(valueArray.count) devices in value array")
                        }
                    }
                    
                    let page = try jsonDecoder.decode(GraphPage<IntuneDevice>.self, from: data)
                    print("📍 Successfully decoded GraphPage with \(page.value.count) devices")
                    
                    // Append new devices
                    devices.append(contentsOf: page.value)
                    
                    // Update pagination state
                    nextDeviceLink = page.nextLink
                    hasMoreDevices = page.nextLink != nil
                    
                    // Update UI
                    if let total = totalDeviceCount {
                        currentOperation = "Loaded \(devices.count) of \(total) devices"
                        loadingProgress = Double(devices.count) / Double(total)
                    } else {
                        currentOperation = "Loaded \(devices.count) devices"
                        loadingProgress = min(0.95, Double(devices.count) / 1000.0)
                    }
                    
                    // For initial load, mark as complete
                    if isInitial {
                        isLoading = false
                    } else {
                        isLoadingMore = false
                    }
                    
                    // Enrich batch with user data in background (non-blocking)
                    enrichDevicesBatch(page.value, accessToken: accessToken)
                    
                    completion?()
                    
                } catch {
                    print("📍 ❌ JSON Decode Error: \(error)")
                    print("📍 Error type: \(type(of: error))")
                    
                    // Try to print raw response for debugging
                    if let rawString = String(data: data, encoding: .utf8) {
                        let preview = String(rawString.prefix(500))
                        print("📍 Response preview: \(preview)")
                    }
                    
                    let message = "Failed to decode devices: \(error.localizedDescription)"
                    if isInitial {
                        finishWithError(message)
                    } else {
                        isLoadingMore = false
                        errorMessage = message
                    }
                    completion?()
                }
            }

    /// Enrich a batch of devices with user data (background, non-blocking)
    private func enrichDevicesBatch(_ batch: [IntuneDevice], accessToken: String) {
        let userIds = Set(batch.compactMap { $0.userId?.trimmingCharacters(in: .whitespacesAndNewlines) }
            .filter { !$0.isEmpty && userCache[$0] == nil })
        
        guard !userIds.isEmpty else { return }
        
        for userId in userIds {
            DispatchQueue.global(qos: .background).async { [weak self] in
                self?.userSemaphore.wait()
                
                self?.fetchUserDataWithRetry(
                    userId: userId,
                    accessToken: accessToken,
                    retryCount: 0
                ) { user in
                    if let user = user {
                        DispatchQueue.main.async {
                            self?.userCache[userId] = user
                            self?.updateDevicesWithUserData(userId: userId, user: user)
                        }
                    }
                    self?.userSemaphore.signal()
                }
            }
        }
    }

    /// Fetch user data with retry logic for rate limiting
    private func fetchUserDataWithRetry(
        userId: String,
        accessToken: String,
        retryCount: Int,
        completion: @escaping (EntraUser?) -> Void
    ) {
        let select = "id,displayName,userPrincipalName,department,jobTitle,officeLocation,companyName,country,city,mail,mobilePhone"
        let expand = "manager($select=displayName,id)"
        guard let url = URL(string: "\(baseURL)/users/\(userId)?$select=\(select)&$expand=\(expand)") else {
            completion(nil)
            return
        }
        
        var req = URLRequest(url: url)
        req.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        req.setValue("application/json", forHTTPHeaderField: "Accept")
        
        URLSession.shared.dataTask(with: req) { [weak self] data, resp, err in
            guard err == nil,
                  let data = data,
                  let http = resp as? HTTPURLResponse else {
                completion(nil)
                return
            }

            
            
            // Handle rate limiting with exponential backoff
            if http.statusCode == 429 && retryCount < 2 {
                let retryAfter = http.value(forHTTPHeaderField: "Retry-After") ?? "5"
                let baseDelay = Int(retryAfter) ?? 5
                let delay = min(baseDelay * (retryCount + 1), 30)  // Cap at 30 seconds
                
                DispatchQueue.global(qos: .background).asyncAfter(deadline: .now() + Double(delay)) {
                    self?.fetchUserDataWithRetry(
                        userId: userId,
                        accessToken: accessToken,
                        retryCount: retryCount + 1,
                        completion: completion
                    )
                }
                return
            }
            
            guard http.statusCode == 200,
                  let user = try? self?.jsonDecoder.decode(EntraUser.self, from: data) else {
                completion(nil)
                return
            }
            
            completion(user)
        }.resume()
    }

    /// Update devices with cached user data
    private func updateDevicesWithUserData(userId: String, user: EntraUser) {
        for i in devices.indices {
            if devices[i].userId == userId {
                devices[i].userDepartment = user.department
                devices[i].userJobTitle = user.jobTitle
                devices[i].userManager = user.manager?.displayName
                devices[i].userOfficeLocation = user.officeLocation
                devices[i].userCompanyName = user.companyName
                devices[i].userCountry = user.country
                devices[i].userCity = user.city
            }
        }
    }

    // MARK: - Batch User Enrichment (for export operations)

    /// Enrich all devices with user data before export (with rate limit handling)
    func enrichAllDevicesForExport(accessToken: String, completion: @escaping () -> Void) {
        let userIds = Set(devices.compactMap { $0.userId?.trimmingCharacters(in: .whitespacesAndNewlines) }
            .filter { !$0.isEmpty && userCache[$0] == nil })
        
        guard !userIds.isEmpty else {
            completion()
            return
        }
        
        currentOperation = "Enriching device data for export..."
        
        let group = DispatchGroup()
        let semaphore = DispatchSemaphore(value: 2)  // Very conservative for export
        var processed = 0
        let total = userIds.count
        
        for userId in userIds {
            group.enter()
            
            DispatchQueue.global(qos: .userInitiated).async { [weak self] in
                semaphore.wait()
                
                self?.fetchUserDataWithRetry(
                    userId: userId,
                    accessToken: accessToken,
                    retryCount: 0
                ) { user in
                    if let user = user {
                        DispatchQueue.main.async {
                            self?.userCache[userId] = user
                            self?.updateDevicesWithUserData(userId: userId, user: user)
                            processed += 1
                            self?.currentOperation = "Enriching data... (\(processed)/\(total))"
                            self?.loadingProgress = Double(processed) / Double(total)
                        }
                    }
                    semaphore.signal()
                    group.leave()
                }
            }
        }
        
        group.notify(queue: .main) { [weak self] in
            self?.currentOperation = ""
            self?.loadingProgress = 1.0
            completion()
        }
    }

    // MARK: - Legacy Enrichment (kept for compatibility)

    private func enrichDevicesWithUserData(accessToken: String) {
        let ids = Set(devices.compactMap { $0.userId?.trimmingCharacters(in: .whitespacesAndNewlines) }.filter { !$0.isEmpty })
        guard !ids.isEmpty else {
            isLoading = false
            currentOperation = ""
            return
        }

        let total = ids.count
        var done = 0
        let group = DispatchGroup()

        for userId in ids {
            group.enter()
            fetchUserData(userId: userId, accessToken: accessToken) { [weak self] user in
                if let user = user { self?.userCache[userId] = user }
                done += 1
                self?.loadingProgress = 0.95 + (0.05 * Double(done) / Double(total))
                self?.currentOperation = "Loading user data… (\(done)/\(total))"
                group.leave()
            }
        }

        group.notify(queue: .main) { [weak self] in
            self?.applyUserDataToDevices()
            self?.isLoading = false
            self?.currentOperation = ""
            self?.loadingProgress = 1.0
        }
    }

    private func fetchUserData(userId: String, accessToken: String, completion: @escaping (EntraUser?) -> Void) {
        fetchUserDataWithRetry(userId: userId, accessToken: accessToken, retryCount: 0, completion: completion)
    }

    private func applyUserDataToDevices() {
        guard !userCache.isEmpty else { return }
        for i in devices.indices {
            guard let uid = devices[i].userId, let u = userCache[uid] else { continue }
            devices[i].userDepartment     = u.department
            devices[i].userJobTitle       = u.jobTitle
            devices[i].userManager        = u.manager?.displayName
            devices[i].userOfficeLocation = u.officeLocation
            devices[i].userCompanyName    = u.companyName
            devices[i].userCountry        = u.country
            devices[i].userCity           = u.city
        }
    }

    // MARK: - Users

    func fetchAllUsers(accessToken: String) {
        isLoading = true
        errorMessage = nil
        users = []
        loadingProgress = 0
        currentOperation = "Fetching users…"

        let select = "id,displayName,userPrincipalName,department,jobTitle,officeLocation,companyName,country,city,mail,mobilePhone"
        let expand = "manager($select=id,displayName)"
        let first = "\(baseURL)/users?$select=\(select)&$expand=\(expand)&$top=999"
        fetchUsersPage(accessToken: accessToken, url: first)
    }

    private func fetchUsersPage(accessToken: String, url: String) {
        guard let requestURL = URL(string: url) else {
            finishWithError("Invalid users URL")
            return
        }

        var req = URLRequest(url: requestURL)
        req.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        req.setValue("application/json", forHTTPHeaderField: "Accept")

        URLSession.shared.dataTask(with: req) { [weak self] data, resp, err in
            DispatchQueue.main.async {
                self?.handleUsersPage(data: data, response: resp, error: err, accessToken: accessToken)
            }
        }.resume()
    }

    private func handleUsersPage(data: Data?, response: URLResponse?, error: Error?, accessToken: String) {
        if let error = error { return finishWithError("Users request failed: \(error.localizedDescription)") }
        guard let http = response as? HTTPURLResponse else { return finishWithError("Invalid users response") }
        
        // Handle rate limiting
        if http.statusCode == 429 {
            let retryAfter = http.value(forHTTPHeaderField: "Retry-After") ?? "10"
            let delay = Int(retryAfter) ?? 10
            currentOperation = "Rate limited. Retrying in \(delay) seconds..."
            
            DispatchQueue.main.asyncAfter(deadline: .now() + Double(delay)) { [weak self] in
                guard let url = response?.url?.absoluteString else { return }
                self?.fetchUsersPage(accessToken: accessToken, url: url)
            }
            return
        }
        
        guard http.statusCode == 200 else { return finishWithError("Users API returned \(http.statusCode)") }
        guard let data = data else { return finishWithError("Users API returned no data") }

        do {
            let page = try jsonDecoder.decode(GraphPage<EntraUser>.self, from: data)
            users.append(contentsOf: page.value)
            loadingProgress = min(0.99, loadingProgress + 0.05)
            currentOperation = "Loaded \(users.count) users…"

            if let next = page.nextLink {
                fetchUsersPage(accessToken: accessToken, url: next)
            } else {
                loadingProgress = 1.0
                isLoading = false
                currentOperation = ""
            }
        } catch {
            finishWithError("Failed to decode users: \(error.localizedDescription)")
        }
    }

    // MARK: - Apps (Updated for proper All apps / Discovered apps separation)

    /// Fetch all apps from both endpoints and combine them
    func fetchAllApps(accessToken: String, completion: @escaping ([DetectedApp]) -> Void) {
        // Check token has required scope first
        guard tokenHasAppsScope(accessToken) else {
            DispatchQueue.main.async {
                self.errorMessage = "Authentication is missing Apps scope. Please sign out and sign in again to grant DeviceManagementApps.Read.All permission."
                self.isLoading = false
                self.currentOperation = ""
            }
            completion([])
            return
        }
        
        isLoading = true
        errorMessage = nil
        currentOperation = "Fetching apps…"
        apps = []
        discoveredApps = []
        
        let group = DispatchGroup()
        var mobileApps: [DetectedApp] = []
        var detectedApps: [DetectedApp] = []
        var fetchError: String?
        
        // Fetch mobile apps (managed apps)
        group.enter()
        fetchMobileApps(accessToken: accessToken) { result in
            switch result {
            case .success(let apps):
                mobileApps = apps
            case .failure(let error):
                fetchError = error.localizedDescription
            }
            group.leave()
        }
        
        // Fetch discovered apps
        group.enter()
        fetchDiscoveredApps(accessToken: accessToken) { result in
            switch result {
            case .success(let apps):
                detectedApps = apps
            case .failure(let error):
                // Don't overwrite if we already have an error
                if fetchError == nil {
                    fetchError = error.localizedDescription
                }
            }
            group.leave()
        }
        
        group.notify(queue: .main) {
            if let error = fetchError {
                self.errorMessage = error
                self.isLoading = false
                self.currentOperation = ""
                completion([])
                return
            }
            
            // Store apps separately
            self.apps = mobileApps
            self.discoveredApps = detectedApps
            
            // Combine all apps for the view
            let allApps = mobileApps + detectedApps
            
            self.isLoading = false
            self.currentOperation = ""
            completion(allApps)
        }
    }

    /// Fetch mobile apps (managed apps) from deviceAppManagement endpoint
    private func fetchMobileApps(accessToken: String, completion: @escaping (Result<[DetectedApp], Error>) -> Void) {
        currentOperation = "Fetching managed apps…"
        
        // Simplify - just get all mobile apps without the complex filter
        let url = "\(betaURL)/deviceAppManagement/mobileApps?$orderby=displayName asc&$top=100"
        
        func fetchPage(_ urlStr: String, _ accumulated: [DetectedApp]) {
            guard let url = URL(string: urlStr) else {
                completion(.success(accumulated))
                return
            }
            
            var request = URLRequest(url: url)
            request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            request.setValue("application/json", forHTTPHeaderField: "Accept")
            
            URLSession.shared.dataTask(with: request) { data, response, error in
                if let error = error {
                    DispatchQueue.main.async {
                        completion(.failure(error))
                    }
                    return
                }
                
                guard let data = data,
                      let http = response as? HTTPURLResponse else {
                    DispatchQueue.main.async {
                        completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "No response from server"])))
                    }
                    return
                }
                
                // Handle rate limiting
                if http.statusCode == 429 {
                    let retryAfter = http.value(forHTTPHeaderField: "Retry-After") ?? "10"
                    let delay = Int(retryAfter) ?? 10
                    
                    DispatchQueue.global().asyncAfter(deadline: .now() + Double(delay)) {
                        fetchPage(urlStr, accumulated)
                    }
                    return
                }
                
                guard (200..<300).contains(http.statusCode) else {
                    DispatchQueue.main.async {
                        let message = self.getErrorMessage(for: http.statusCode, endpoint: "mobileApps")
                        completion(.failure(NSError(domain: "GraphAPI", code: http.statusCode, userInfo: [NSLocalizedDescriptionKey: message])))
                    }
                    return
                }
                
                guard let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let items = json["value"] as? [[String: Any]] else {
                    DispatchQueue.main.async {
                        completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "Invalid response format"])))
                    }
                    return
                }
                
                var nextBatch: [DetectedApp] = []
                
                for obj in items {
                    guard let id = obj["id"] as? String else { continue }
                    
                    let name = self.sanitize(obj["displayName"] as? String)
                    guard !name.isEmpty else { continue }
                    
                    let version = self.sanitize(obj["version"] as? String, isVersion: true)
                    let publisher = self.sanitize(obj["publisher"] as? String)
                    
                    var app = DetectedApp(
                        id: id,
                        displayName: name,
                        version: version,
                        publisher: publisher
                    )
                    app.isManaged = true  // All apps from mobileApps endpoint are managed
                    
                    nextBatch.append(app)
                }
                
                let combined = accumulated + nextBatch
                
                DispatchQueue.main.async {
                    self.currentOperation = "Loaded \(combined.count) managed apps…"
                }
                
                if let next = json["@odata.nextLink"] as? String {
                    fetchPage(next, combined)
                } else {
                    DispatchQueue.main.async {
                        completion(.success(combined))
                    }
                }
            }.resume()
        }
        
        fetchPage(url, [])
    }

    /// Fetch discovered apps from detectedApps endpoint
    private func fetchDiscoveredApps(accessToken: String, completion: @escaping (Result<[DetectedApp], Error>) -> Void) {
        currentOperation = "Fetching discovered apps…"
        
        let url = "\(betaURL)/deviceManagement/detectedApps?$orderby=displayName&$top=100"
        
        func fetchPage(_ urlStr: String, _ accumulated: [DetectedApp]) {
            guard let url = URL(string: urlStr) else {
                completion(.success(accumulated))
                return
            }
            
            var request = URLRequest(url: url)
            request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            request.setValue("application/json", forHTTPHeaderField: "Accept")
            
            URLSession.shared.dataTask(with: request) { data, response, error in
                if let error = error {
                    DispatchQueue.main.async {
                        completion(.failure(error))
                    }
                    return
                }
                
                guard let data = data,
                      let http = response as? HTTPURLResponse else {
                    DispatchQueue.main.async {
                        completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "No response from server"])))
                    }
                    return
                }
                
                // Handle rate limiting
                if http.statusCode == 429 {
                    let retryAfter = http.value(forHTTPHeaderField: "Retry-After") ?? "10"
                    let delay = Int(retryAfter) ?? 10
                    
                    DispatchQueue.global().asyncAfter(deadline: .now() + Double(delay)) {
                        fetchPage(urlStr, accumulated)
                    }
                    return
                }
                
                guard (200..<300).contains(http.statusCode) else {
                    DispatchQueue.main.async {
                        let message = self.getErrorMessage(for: http.statusCode, endpoint: "detectedApps")
                        completion(.failure(NSError(domain: "GraphAPI", code: http.statusCode, userInfo: [NSLocalizedDescriptionKey: message])))
                    }
                    return
                }
                
                guard let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let items = json["value"] as? [[String: Any]] else {
                    DispatchQueue.main.async {
                        completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "Invalid response format"])))
                    }
                    return
                }
                
                var nextBatch: [DetectedApp] = []
                
                for obj in items {
                    guard let id = obj["id"] as? String else { continue }
                    
                    let name = self.sanitize(obj["displayName"] as? String)
                    guard !name.isEmpty else { continue }
                    
                    // Version might be in a different format for detected apps
                    var version: String? = nil
                    if let versionObj = obj["version"] as? [String: Any],
                       let versionValue = versionObj["value"] as? String {
                        version = self.sanitize(versionValue, isVersion: true)
                    } else if let versionString = obj["version"] as? String {
                        version = self.sanitize(versionString, isVersion: true)
                    }
                    
                    let publisher = self.sanitize(obj["publisher"] as? String)
                    
                    var app = DetectedApp(
                        id: id,
                        displayName: name,
                        version: version,
                        publisher: publisher
                    )
                    app.isManaged = false  // All apps from detectedApps endpoint are discovered
                    
                    nextBatch.append(app)
                }
                
                let combined = accumulated + nextBatch
                
                DispatchQueue.main.async {
                    self.currentOperation = "Loaded \(combined.count) discovered apps…"
                }
                
                if let next = json["@odata.nextLink"] as? String {
                    fetchPage(next, combined)
                } else {
                    DispatchQueue.main.async {
                        completion(.success(combined))
                    }
                }
            }.resume()
        }
        
        fetchPage(url, [])
    }

    /// Helper to get user-friendly error messages
    private func getErrorMessage(for statusCode: Int, endpoint: String) -> String {
        switch statusCode {
        case 403:
            return "Access denied to \(endpoint). Please ensure you have appropriate Intune role permissions."
        case 401:
            return "Authentication failed. Please sign in again."
        case 404:
            return "The \(endpoint) endpoint was not found. This might indicate missing Intune licensing."
        default:
            return "Failed to fetch \(endpoint) (status: \(statusCode))"
        }
    }

    /// Fetch detailed information about a specific app
    func fetchAppDetail(id: String, accessToken: String, completion: @escaping (Result<DetectedApp, Error>) -> Void) {
        // Check token scope first
        guard tokenHasAppsScope(accessToken) else {
            completion(.failure(NSError(domain: "GraphAPI", code: 403, userInfo: [NSLocalizedDescriptionKey: "Missing DeviceManagementApps.Read.All scope"])))
            return
        }
        
        // First check cache - try both managed and discovered apps
        if let cached = apps.first(where: { $0.id == id }) ?? discoveredApps.first(where: { $0.id == id }) {
            completion(.success(cached))
            return
        }
        
        // Try mobile apps endpoint first
        let mobileAppUrl = "\(betaURL)/deviceAppManagement/mobileApps/\(id)"
        guard let requestURL = URL(string: mobileAppUrl) else {
            completion(.failure(NSError(domain: "GraphAPI", code: 400, userInfo: [NSLocalizedDescriptionKey: "Invalid URL"])))
            return
        }
        
        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Accept")
        
        URLSession.shared.dataTask(with: request) { data, response, error in
            if let error = error {
                DispatchQueue.main.async {
                    completion(.failure(error))
                }
                return
            }
            
            guard let data = data,
                  let http = response as? HTTPURLResponse else {
                DispatchQueue.main.async {
                    completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "No response"])))
                }
                return
            }
            
            if (200..<300).contains(http.statusCode),
               let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
               let id = json["id"] as? String {
                
                let name = self.sanitize(json["displayName"] as? String) ?? ""
                let version = self.sanitize(json["version"] as? String, isVersion: true)
                let publisher = self.sanitize(json["publisher"] as? String)
                
                var app = DetectedApp(
                    id: id,
                    displayName: name,
                    version: version,
                    publisher: publisher
                )
                app.isManaged = true
                
                DispatchQueue.main.async {
                    completion(.success(app))
                }
                return
            }
            
            // If not found in mobile apps, try detected apps
            let detectedAppUrl = "\(self.betaURL)/deviceManagement/detectedApps/\(id)"
            guard let detectedURL = URL(string: detectedAppUrl) else {
                DispatchQueue.main.async {
                    completion(.failure(NSError(domain: "GraphAPI", code: 404, userInfo: [NSLocalizedDescriptionKey: "App not found"])))
                }
                return
            }
            
            var detectedRequest = URLRequest(url: detectedURL)
            detectedRequest.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            detectedRequest.setValue("application/json", forHTTPHeaderField: "Accept")
            
            URLSession.shared.dataTask(with: detectedRequest) { data, response, error in
                DispatchQueue.main.async {
                    if let error = error {
                        completion(.failure(error))
                        return
                    }
                    
                    guard let data = data,
                          let http = response as? HTTPURLResponse,
                          (200..<300).contains(http.statusCode),
                          let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any] else {
                        completion(.failure(NSError(domain: "GraphAPI", code: 404, userInfo: [NSLocalizedDescriptionKey: "App not found in either endpoint"])))
                        return
                    }
                    
                    guard let id = json["id"] as? String else {
                        completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "Invalid app data"])))
                        return
                    }
                    
                    let name = self.sanitize(json["displayName"] as? String) ?? ""
                    var version: String? = nil
                    if let versionObj = json["version"] as? [String: Any],
                       let versionValue = versionObj["value"] as? String {
                        version = self.sanitize(versionValue, isVersion: true)
                    } else if let versionString = json["version"] as? String {
                        version = self.sanitize(versionString, isVersion: true)
                    }
                    let publisher = self.sanitize(json["publisher"] as? String)
                    
                    var app = DetectedApp(
                        id: id,
                        displayName: name,
                        version: version,
                        publisher: publisher
                    )
                    app.isManaged = false
                    
                    completion(.success(app))
                }
            }.resume()
        }.resume()
    }

    /// Fetch users assigned to a specific app
    func fetchUsersForApp(appId: String, accessToken: String, completion: @escaping ([EntraUser]) -> Void) {
        // Check token scope first
        guard tokenHasAppsScope(accessToken) else {
            completion([])
            return
        }
        
        // First, fetch the app assignments
        let url = "\(betaURL)/deviceAppManagement/mobileApps/\(appId)/assignments"
        guard let requestURL = URL(string: url) else {
            completion([])
            return
        }
        
        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Accept")
        
        URLSession.shared.dataTask(with: request) { data, response, _ in
            guard let data = data,
                  let http = response as? HTTPURLResponse,
                  (200..<300).contains(http.statusCode),
                  let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                  let assignments = json["value"] as? [[String: Any]] else {
                DispatchQueue.main.async { completion([]) }
                return
            }
            
            // Extract group IDs from assignments
            var groupIds: Set<String> = []
            for assignment in assignments {
                if let target = assignment["target"] as? [String: Any],
                   let targetType = target["@odata.type"] as? String,
                   targetType.contains("groupAssignmentTarget"),
                   let groupId = target["groupId"] as? String {
                    groupIds.insert(groupId)
                }
            }
            
            guard !groupIds.isEmpty else {
                DispatchQueue.main.async { completion([]) }
                return
            }
            
            // Now fetch members of these groups
            self.fetchUsersFromGroups(Array(groupIds), accessToken: accessToken, completion: completion)
        }.resume()
    }

    /// Fetch devices that have a specific app installed
    func fetchDevicesForApp(appId: String, accessToken: String, completion: @escaping ([IntuneDevice]) -> Void) {
        // First, determine if this is a managed app or discovered app
        if let app = apps.first(where: { $0.id == appId }) ?? discoveredApps.first(where: { $0.id == appId }) {
            if app.isManaged == false {
                // This is a discovered app, use the detectedApps endpoint
                fetchDevicesForDiscoveredApp(appId: appId, accessToken: accessToken, completion: completion)
            } else {
                // This is a managed app, use the Reports API
                fetchDevicesForManagedApp(appId: appId, accessToken: accessToken, completion: completion)
            }
        } else {
            // App not in cache, try to determine type by checking if it exists in detectedApps
            checkIfDiscoveredApp(appId: appId, accessToken: accessToken) { isDiscovered in
                if isDiscovered {
                    self.fetchDevicesForDiscoveredApp(appId: appId, accessToken: accessToken, completion: completion)
                } else {
                    self.fetchDevicesForManagedApp(appId: appId, accessToken: accessToken, completion: completion)
                }
            }
        }
    }
    
    /// Check if an app ID corresponds to a discovered app
    private func checkIfDiscoveredApp(appId: String, accessToken: String, completion: @escaping (Bool) -> Void) {
        let url = "\(betaURL)/deviceManagement/detectedApps/\(appId)"
        guard let requestURL = URL(string: url) else {
            completion(false)
            return
        }
        
        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Accept")
        
        URLSession.shared.dataTask(with: request) { data, response, error in
            if let http = response as? HTTPURLResponse {
                // If we get a 200, it's a discovered app
                DispatchQueue.main.async {
                    completion((200..<300).contains(http.statusCode))
                }
            } else {
                DispatchQueue.main.async {
                    completion(false)
                }
            }
        }.resume()
    }
    
    /// Fetch devices for a discovered app using the detectedApps endpoint
    private func fetchDevicesForDiscoveredApp(appId: String, accessToken: String, completion: @escaping ([IntuneDevice]) -> Void) {
        let url = "\(betaURL)/deviceManagement/detectedApps('\(appId)')/managedDevices?$orderby=deviceName asc&$top=100"
        
        var allDevices: [IntuneDevice] = []
        
        func fetchPage(_ pageURL: String) {
            guard let url = URL(string: pageURL) else {
                DispatchQueue.main.async {
                    completion(allDevices)
                }
                return
            }
            
            var request = URLRequest(url: url)
            request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            request.setValue("application/json", forHTTPHeaderField: "Accept")
            
            URLSession.shared.dataTask(with: request) { [weak self] data, response, error in
                guard let self = self,
                      let data = data,
                      let http = response as? HTTPURLResponse,
                      (200..<300).contains(http.statusCode) else {
                    DispatchQueue.main.async {
                        completion(allDevices)
                    }
                    return
                }
                
                do {
                    let page = try self.jsonDecoder.decode(GraphPage<IntuneDevice>.self, from: data)
                    allDevices.append(contentsOf: page.value)
                    
                    if let next = page.nextLink {
                        fetchPage(next)
                    } else {
                        DispatchQueue.main.async {
                            // Get the app details to determine platform filtering
                            let app = self.discoveredApps.first(where: { $0.id == appId })
                            let appName = app?.displayName ?? ""
                            
                            // Filter devices based on platform compatibility
                            var filteredDevices = allDevices
                            
                            // Platform-specific app detection
                            if self.isWindowsOnlyApp(appName) {
                                filteredDevices = allDevices.filter { device in
                                    let os = device.operatingSystem?.lowercased() ?? ""
                                    return os.contains("windows")
                                }
                            } else if self.isMacOnlyApp(appName) {
                                filteredDevices = allDevices.filter { device in
                                    let os = device.operatingSystem?.lowercased() ?? ""
                                    return os.contains("macos") || os.contains("mac os")
                                }
                            } else if self.isiOSOnlyApp(appName) {
                                filteredDevices = allDevices.filter { device in
                                    let os = device.operatingSystem?.lowercased() ?? ""
                                    return os.contains("ios")
                                }
                            } else if self.isAndroidOnlyApp(appName) {
                                filteredDevices = allDevices.filter { device in
                                    let os = device.operatingSystem?.lowercased() ?? ""
                                    return os.contains("android")
                                }
                            }
                            // For cross-platform apps, return all devices
                            
                            // Enrich with cached user data if available
                            if !self.devices.isEmpty {
                                var enrichedDevices = filteredDevices
                                for i in enrichedDevices.indices {
                                    if let cachedDevice = self.devices.first(where: { $0.id == enrichedDevices[i].id }) {
                                        enrichedDevices[i].userDepartment = cachedDevice.userDepartment
                                        enrichedDevices[i].userJobTitle = cachedDevice.userJobTitle
                                        enrichedDevices[i].userManager = cachedDevice.userManager
                                        enrichedDevices[i].userOfficeLocation = cachedDevice.userOfficeLocation
                                        enrichedDevices[i].userCompanyName = cachedDevice.userCompanyName
                                        enrichedDevices[i].userCountry = cachedDevice.userCountry
                                        enrichedDevices[i].userCity = cachedDevice.userCity
                                    }
                                }
                                completion(enrichedDevices)
                            } else {
                                completion(filteredDevices)
                            }
                        }
                    }
                } catch {
                    DispatchQueue.main.async {
                        completion(allDevices)
                    }
                }
            }.resume()
        }
        
        fetchPage(url)
    }
    
    // MARK: - Platform Detection Helpers
    
    private func isWindowsOnlyApp(_ appName: String) -> Bool {
        let windowsApps = [
            "7-zip", "7zip", "winrar", "notepad++", "putty", "winscp",
            "registry editor", "task manager", "device manager", "disk management",
            "windows defender", "windows terminal", "powershell", "paint",
            "snipping tool", "windows media player", "internet explorer"
        ]
        let lowercaseName = appName.lowercased()
        return windowsApps.contains { lowercaseName.contains($0) }
    }
    
    private func isMacOnlyApp(_ appName: String) -> Bool {
        let macApps = [
            "xcode", "final cut", "logic pro", "keynote", "pages", "numbers",
            "about this mac", "activity monitor", "disk utility", "finder",
            "preview", "textedit", "terminal", "safari", "photos", "imovie",
            "garageband", "automator", "grapher", "console", "keychain access",
            "applescript", "script editor", "migration assistant", "time machine",
            "airdrop", "airplay", "system preferences", "system settings",
            "mission control", "launchpad", "spotlight", "siri", "archive utility",
            "bluetooth file exchange", "boot camp", "color sync", "directory utility",
            "dvd player", "font book", "image capture", "screenshot", "stickies",
            "voiceover utility", "audio midi setup", "digital color meter"
        ]
        let lowercaseName = appName.lowercased()
        
        // Also check for apps that contain "mac" or "macos" in the name
        if lowercaseName.contains("mac") || lowercaseName.contains("macos") {
            // But exclude some cross-platform apps that might have "mac" in the name
            let excludedApps = ["teamviewer", "anydesk", "chrome remote"]
            if !excludedApps.contains(where: { lowercaseName.contains($0) }) {
                return true
            }
        }
        
        return macApps.contains { lowercaseName.contains($0) }
    }
    
    private func isiOSOnlyApp(_ appName: String) -> Bool {
        let iosApps = [
            "procreate", "goodnotes", "notability", "shortcuts", "testflight",
            "clips", "measure", "compass", "health", "wallet", "find my",
            "home", "news", "stocks", "voice memos", "facetime", "books"
        ]
        let lowercaseName = appName.lowercased()
        return iosApps.contains { lowercaseName.contains($0) }
    }
    
    private func isAndroidOnlyApp(_ appName: String) -> Bool {
        let androidApps = [
            "google play", "play store", "android system", "android auto",
            "google services", "samsung", "galaxy", "oneplus"
        ]
        let lowercaseName = appName.lowercased()
        return androidApps.contains { lowercaseName.contains($0) }
    }
    
    /// Fetch devices for a managed app using the Reports API
    private func fetchDevicesForManagedApp(appId: String, accessToken: String, completion: @escaping ([IntuneDevice]) -> Void) {
        // Use the Reports API for managed apps
        fetchDeviceAppInstallationStatusReport(appId: appId, accessToken: accessToken) { [weak self] deviceIds in
            guard let self = self else {
                completion([])
                return
            }
            
            // If we have cached devices, filter them
            if !self.devices.isEmpty {
                let devicesWithApp = self.devices.filter { device in
                    deviceIds.contains(device.id)
                }
                completion(devicesWithApp)
            } else {
                // Need to fetch device details for the device IDs we found
                self.fetchDevicesByIds(deviceIds, accessToken: accessToken, completion: completion)
            }
        }
    }

    /// Fetch device app installation status using the Reports API
    private func fetchDeviceAppInstallationStatusReport(appId: String, accessToken: String, completion: @escaping ([String]) -> Void) {
        let url = "\(betaURL)/deviceManagement/reports/retrieveDeviceAppInstallationStatusReport"
        
        guard let requestURL = URL(string: url) else {
            completion([])
            return
        }
        
        // Build the request body matching what Intune uses
        let requestBody: [String: Any] = [
            "select": [
                "DeviceName",
                "UserPrincipalName",
                "Platform",
                "AppVersion",
                "InstallState",
                "InstallStateDetail",
                "AssignmentFilterIdsExist",
                "LastModifiedDateTime",
                "DeviceId",
                "ErrorCode",
                "UserName",
                "UserId",
                "ApplicationId",
                "AssignmentFilterIdsList",
                "AppInstallState",
                "AppInstallStateDetails",
                "HexErrorCode"
            ],
            "skip": 0,
            "top": 999,
            "filter": "(ApplicationId eq '\(appId)')",
            "orderBy": []
        ]
        
        guard let jsonData = try? JSONSerialization.data(withJSONObject: requestBody) else {
            completion([])
            return
        }
        
        var request = URLRequest(url: requestURL)
        request.httpMethod = "POST"
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        request.setValue("application/json", forHTTPHeaderField: "Accept")
        request.httpBody = jsonData
        
        URLSession.shared.dataTask(with: request) { data, response, error in
            guard let data = data,
                  let http = response as? HTTPURLResponse,
                  (200..<300).contains(http.statusCode) else {
                DispatchQueue.main.async {
                    completion([])
                }
                return
            }
            
            // Parse the report response
            var deviceIds: [String] = []
            
            do {
                if let json = try JSONSerialization.jsonObject(with: data) as? [String: Any],
                   let values = json["values"] as? [[Any]] {
                    
                    // The response contains an array of arrays
                    // We need to find the DeviceId column index from the schema
                    if let schema = json["schema"] as? [String: Any],
                       let columns = schema["columns"] as? [[String: Any]] {
                        
                        // Find the index of the DeviceId column
                        var deviceIdIndex = -1
                        var installStateIndex = -1
                        
                        for (index, column) in columns.enumerated() {
                            if let name = column["name"] as? String {
                                if name == "DeviceId" {
                                    deviceIdIndex = index
                                } else if name == "AppInstallState" || name == "InstallState" {
                                    installStateIndex = index
                                }
                            }
                        }
                        
                        // Extract device IDs where the app is installed
                        if deviceIdIndex >= 0 {
                            for row in values {
                                if row.count > deviceIdIndex,
                                   let deviceId = row[deviceIdIndex] as? String {
                                    
                                    // Check install state if we have it
                                    if installStateIndex >= 0 && row.count > installStateIndex {
                                        if let state = row[installStateIndex] as? String,
                                           state.lowercased() == "installed" {
                                            deviceIds.append(deviceId)
                                        }
                                    } else {
                                        // If no install state column, include all devices
                                        deviceIds.append(deviceId)
                                    }
                                }
                            }
                        }
                    }
                }
            } catch {
                print("Error parsing report response: \(error)")
            }
            
            DispatchQueue.main.async {
                completion(deviceIds)
            }
        }.resume()
    }

    /// Fetch devices by their IDs
    private func fetchDevicesByIds(_ deviceIds: [String], accessToken: String, completion: @escaping ([IntuneDevice]) -> Void) {
        guard !deviceIds.isEmpty else {
            completion([])
            return
        }
        
        var devices: [IntuneDevice] = []
        let group = DispatchGroup()
        let semaphore = DispatchSemaphore(value: 5) // Limit concurrent requests
        
        for deviceId in deviceIds {
            group.enter()
            
            DispatchQueue.global(qos: .userInitiated).async {
                semaphore.wait()
                
                let url = "\(self.betaURL)/deviceManagement/managedDevices/\(deviceId)"
                guard let requestURL = URL(string: url) else {
                    semaphore.signal()
                    group.leave()
                    return
                }
                
                var request = URLRequest(url: requestURL)
                request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
                request.setValue("application/json", forHTTPHeaderField: "Accept")
                
                URLSession.shared.dataTask(with: request) { data, response, error in
                    defer {
                        semaphore.signal()
                        group.leave()
                    }
                    
                    guard let data = data,
                          let http = response as? HTTPURLResponse,
                          (200..<300).contains(http.statusCode) else {
                        return
                    }
                    
                    do {
                        let device = try self.jsonDecoder.decode(IntuneDevice.self, from: data)
                        DispatchQueue.main.async {
                            devices.append(device)
                        }
                    } catch {
                        print("Error decoding device: \(error)")
                    }
                }.resume()
            }
        }
        
        group.notify(queue: .main) {
            completion(devices)
        }
    }

    private func fetchUsersFromGroups(_ groupIds: [String], accessToken: String, completion: @escaping ([EntraUser]) -> Void) {
        var allUsers: [EntraUser] = []
        let group = DispatchGroup()
        
        for groupId in groupIds {
            group.enter()
            
            let url = "\(baseURL)/groups/\(groupId)/members"
            guard let requestURL = URL(string: url) else {
                group.leave()
                continue
            }
            
            var request = URLRequest(url: requestURL)
            request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            request.setValue("application/json", forHTTPHeaderField: "Accept")
            
            URLSession.shared.dataTask(with: request) { data, _, _ in
                defer { group.leave() }
                
                guard let data = data,
                      let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let members = json["value"] as? [[String: Any]] else {
                    return
                }
                
                for member in members {
                    if let odataType = member["@odata.type"] as? String,
                       odataType.contains("user"),
                       let id = member["id"] as? String {
                        
                        // Try to find user in cache first
                        if let cachedUser = self.userCache[id] ?? self.users.first(where: { $0.id == id }) {
                            allUsers.append(cachedUser)
                        } else {
                            // Create a minimal user object
                            let user = EntraUser(
                                id: id,
                                displayName: member["displayName"] as? String,
                                userPrincipalName: member["userPrincipalName"] as? String,
                                department: nil,
                                jobTitle: nil,
                                manager: nil,
                                officeLocation: nil,
                                companyName: nil,
                                country: nil,
                                city: nil,
                                mail: member["mail"] as? String,
                                mobilePhone: nil
                            )
                            allUsers.append(user)
                        }
                    }
                }
            }.resume()
        }
        
        group.notify(queue: .main) {
            // Remove duplicates
            var seen = Set<String>()
            let uniqueUsers = allUsers.filter { user in
                guard !seen.contains(user.id) else { return false }
                seen.insert(user.id)
                return true
            }
            completion(uniqueUsers)
        }
    }

    // MARK: - Detail fetch (cache-first)

    func fetchDeviceDetail(id: String, completion: @escaping (Result<IntuneDevice, Error>) -> Void) {
        if let d = devices.first(where: { $0.id == id }) {
            completion(.success(d))
        } else {
            completion(.failure(NSError(domain: "GraphAPI", code: 404,
                                        userInfo: [NSLocalizedDescriptionKey: "Device not found in cache"])))
        }
    }

    func fetchUsersForDevice(id: String, completion: @escaping ([EntraUser]) -> Void) {
        guard let d = devices.first(where: { $0.id == id }),
              let uid = d.userId, !uid.isEmpty
        else { completion([]); return }

        if let cached = userCache[uid] {
            completion([cached])
        } else if let u = users.first(where: { $0.id == uid }) {
            completion([u])
        } else {
            completion([])
        }
    }

    func fetchUserDetail(id: String, completion: @escaping (Result<EntraUser, Error>) -> Void) {
        if let u = userCache[id] ?? users.first(where: { $0.id == id }) {
            completion(.success(u))
        } else {
            completion(.failure(NSError(domain: "GraphAPI", code: 404,
                                        userInfo: [NSLocalizedDescriptionKey: "User not found in cache"])))
        }
    }

    func fetchDevicesForUser(id: String, completion: @escaping ([IntuneDevice]) -> Void) {
        let owned = devices.filter { $0.userId == id }
        completion(owned)
    }

    // MARK: - Detail fetch with API fallback (for detail window)

    func fetchDeviceDetail(id: String, accessToken: String, completion: @escaping (Result<IntuneDevice, Error>) -> Void) {
        // First check cache
        if let cached = devices.first(where: { $0.id == id }) {
            completion(.success(cached))
            return
        }

        // Not in cache, fetch from API
        let url = "\(betaURL)/deviceManagement/managedDevices/\(id)"
        guard let requestURL = URL(string: url) else {
            completion(.failure(NSError(domain: "GraphAPI", code: 400, userInfo: [NSLocalizedDescriptionKey: "Invalid URL"])))
            return
        }

        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Accept")

        URLSession.shared.dataTask(with: request) { [weak self] data, response, error in
            DispatchQueue.main.async {
                if let error = error {
                    completion(.failure(error)); return
                }
                guard let data = data else {
                    completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "No data received"]))); return
                }
                do {
                    let device = try (self?.jsonDecoder ?? JSONDecoder()).decode(IntuneDevice.self, from: data)
                    completion(.success(device))
                } catch {
                    completion(.failure(error))
                }
            }
        }.resume()
    }

    func fetchUsersForDevice(id: String, accessToken: String, completion: @escaping ([EntraUser]) -> Void) {
        // First try cache
        if let device = devices.first(where: { $0.id == id }),
           let userId = device.userId,
           !userId.isEmpty {

            // Try to get from cache or fetch
            fetchUserDetail(id: userId, accessToken: accessToken) { result in
                switch result {
                case .success(let user): completion([user])
                case .failure:           completion([])
                }
            }
            return
        }

        // If not in cache, fetch the device first
        fetchDeviceDetail(id: id, accessToken: accessToken) { [weak self] result in
            switch result {
            case .success(let device):
                if let userId = device.userId, !userId.isEmpty {
                    self?.fetchUserDetail(id: userId, accessToken: accessToken) { userResult in
                        switch userResult {
                        case .success(let user): completion([user])
                        case .failure:           completion([])
                        }
                    }
                } else {
                    completion([])
                }
            case .failure:
                completion([])
            }
        }
    }

    func fetchUserDetail(id: String, accessToken: String, completion: @escaping (Result<EntraUser, Error>) -> Void) {
        // Check cache first
        if let cached = userCache[id] ?? users.first(where: { $0.id == id }) {
            completion(.success(cached))
            return
        }

        // Not in cache, fetch from API
        let select = "id,displayName,userPrincipalName,department,jobTitle,officeLocation,companyName,country,city,mail,mobilePhone"
        let expand = "manager($select=displayName,id)"
        guard let url = URL(string: "\(baseURL)/users/\(id)?$select=\(select)&$expand=\(expand)") else {
            completion(.failure(NSError(domain: "GraphAPI", code: 400, userInfo: [NSLocalizedDescriptionKey: "Invalid URL"])))
            return
        }

        var request = URLRequest(url: url)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Accept")

        URLSession.shared.dataTask(with: request) { [weak self] data, response, error in
            DispatchQueue.main.async {
                if let error = error {
                    completion(.failure(error)); return
                }
                guard let data = data else {
                    completion(.failure(NSError(domain: "GraphAPI", code: 0, userInfo: [NSLocalizedDescriptionKey: "No data received"]))); return
                }
                do {
                    let user = try (self?.jsonDecoder ?? JSONDecoder()).decode(EntraUser.self, from: data)
                    self?.userCache[id] = user
                    completion(.success(user))
                } catch {
                    completion(.failure(error))
                }
            }
        }.resume()
    }
    
    // MARK: - User Groups and Assignments

    func fetchUserGroups(userId: String, accessToken: String, completion: @escaping ([String]) -> Void) {
        let url = "\(baseURL)/users/\(userId)/memberOf"
        guard let requestURL = URL(string: url) else {
            completion([]); return
        }

        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Accept")

        URLSession.shared.dataTask(with: request) { data, _, _ in
            DispatchQueue.main.async {
                guard let data = data,
                      let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let groups = json["value"] as? [[String: Any]] else {
                    completion([]); return
                }
                let groupNames = groups.compactMap { $0["displayName"] as? String }
                completion(groupNames)
            }
        }.resume()
    }

    func fetchUserAssignments(userId: String, accessToken: String, completion: @escaping ([String]) -> Void) {
        var allAssignments: [String] = []
        let group = DispatchGroup()

        // 1. Enterprise app role assignments
        group.enter()
        fetchAppRoleAssignments(userId: userId, accessToken: accessToken) { assignments in
            allAssignments.append(contentsOf: assignments)
            group.leave()
        }

        // 2. Intune mobile app assignments (via groups) - only if token has apps scope
        if tokenHasAppsScope(accessToken) {
            group.enter()
            fetchIntuneAppAssignments(userId: userId, accessToken: accessToken) { assignments in
                allAssignments.append(contentsOf: assignments)
                group.leave()
            }
        }

        // 3. Compliance policies
        group.enter()
        fetchCompliancePolicies(userId: userId, accessToken: accessToken) { policies in
            allAssignments.append(contentsOf: policies)
            group.leave()
        }

        group.notify(queue: .main) {
            completion(Array(Set(allAssignments))) // de-dupe
        }
    }

    private func fetchAppRoleAssignments(userId: String, accessToken: String, completion: @escaping ([String]) -> Void) {
        let url = "\(baseURL)/users/\(userId)/appRoleAssignments"
        guard let requestURL = URL(string: url) else {
            completion([]); return
        }

        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")

        URLSession.shared.dataTask(with: request) { data, _, _ in
            DispatchQueue.main.async {
                guard let data = data,
                      let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let assignments = json["value"] as? [[String: Any]] else {
                    completion([]); return
                }
                let appNames = assignments.compactMap { $0["resourceDisplayName"] as? String }.map { "📱 \($0)" }
                completion(appNames)
            }
        }.resume()
    }

    private func fetchIntuneAppAssignments(userId: String, accessToken: String, completion: @escaping ([String]) -> Void) {
        fetchUserGroupIds(userId: userId, accessToken: accessToken) { groupIds in
            guard !groupIds.isEmpty else { completion([]); return }

            let url = "\(self.betaURL)/deviceAppManagement/mobileApps?$expand=assignments"
            guard let requestURL = URL(string: url) else {
                completion([]); return
            }

            var request = URLRequest(url: requestURL)
            request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")

            URLSession.shared.dataTask(with: request) { data, _, _ in
                DispatchQueue.main.async {
                    guard let data = data,
                          let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                          let apps = json["value"] as? [[String: Any]] else {
                        completion([]); return
                    }

                    var assignedApps: [String] = []
                    for app in apps {
                        guard let appName = app["displayName"] as? String,
                              let assignments = app["assignments"] as? [[String: Any]] else { continue }

                        for assignment in assignments {
                            if let target = assignment["target"] as? [String: Any],
                               let targetType = target["@odata.type"] as? String,
                               targetType.contains("groupAssignmentTarget"),
                               let targetGroupId = target["groupId"] as? String,
                               groupIds.contains(targetGroupId) {
                                assignedApps.append("📲 \(appName)")
                                break
                            }
                        }
                    }
                    completion(assignedApps)
                }
            }.resume()
        }
    }

    private func fetchUserGroupIds(userId: String, accessToken: String, completion: @escaping ([String]) -> Void) {
        let url = "\(baseURL)/users/\(userId)/memberOf"
        guard let requestURL = URL(string: url) else {
            completion([]); return
        }

        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")

        URLSession.shared.dataTask(with: request) { data, _, _ in
            DispatchQueue.main.async {
                guard let data = data,
                      let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let groups = json["value"] as? [[String: Any]] else {
                    completion([]); return
                }

                let groupIds = groups.compactMap { $0["id"] as? String }
                completion(groupIds)
            }
        }.resume()
    }

    private func fetchCompliancePolicies(userId: String, accessToken: String, completion: @escaping ([String]) -> Void) {
        let url = "\(betaURL)/deviceManagement/deviceCompliancePolicies"
        guard let requestURL = URL(string: url) else {
            completion([]); return
        }

        var request = URLRequest(url: requestURL)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")

        URLSession.shared.dataTask(with: request) { data, _, _ in
            DispatchQueue.main.async {
                guard let data = data,
                      let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let policies = json["value"] as? [[String: Any]] else {
                    completion([]); return
                }

                let policyNames = policies.compactMap { $0["displayName"] as? String }.map { "📋 \($0)" }
                completion(policyNames)
            }
        }.resume()
    }

    // MARK: - Device Assignments & Apps

    /// Remove emoji-like scalars that sometimes show up in Graph displayName fields.
    /// For version fields, we'll be less aggressive to preserve numbers and dots.
    private func sanitize(_ s: String?, isVersion: Bool = false) -> String {
        guard let s = s else { return "" }
        if isVersion {
            // For versions, just trim whitespace and keep all other characters
            return s.trimmingCharacters(in: .whitespacesAndNewlines)
        } else {
            // For display names, remove emoji
            let filtered = s.unicodeScalars.filter { !($0.properties.isEmoji || $0.properties.isEmojiPresentation) }
            return String(String.UnicodeScalarView(filtered)).trimmingCharacters(in: .whitespacesAndNewlines)
        }
    }

    /// Fetch the display names of Configuration and Compliance policies assigned to a device.
    /// Returns two arrays: (configurationPolicies, compliancePolicies).
    func fetchDeviceAssignments(deviceId: String,
                                accessToken: String,
                                completion: @escaping (_ configurationPolicies: [String], _ compliancePolicies: [String]) -> Void) {

        let configURL = "\(betaURL)/deviceManagement/managedDevices/\(deviceId)/deviceConfigurationStates?$select=displayName"
        let compURL   = "\(betaURL)/deviceManagement/managedDevices/\(deviceId)/deviceCompliancePolicyStates?$select=displayName"

        let group = DispatchGroup()
        var configs: [String] = []
        var comps:   [String] = []

        func fetchNames(_ urlStr: String, sink: @escaping ([String]) -> Void) {
            guard let url = URL(string: urlStr) else { sink([]); return }
            var req = URLRequest(url: url)
            req.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            req.setValue("application/json", forHTTPHeaderField: "Accept")

            URLSession.shared.dataTask(with: req) { data, resp, _ in
                var names: [String] = []
                defer { sink(names) }

                guard let data = data,
                      let http = resp as? HTTPURLResponse, (200..<300).contains(http.statusCode),
                      let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let items = json["value"] as? [[String: Any]] else { return }

                for obj in items {
                    if let raw = obj["displayName"] as? String {
                        let name = self.sanitize(raw)
                        if !name.isEmpty { names.append(name) }
                    }
                }
            }.resume()
        }

        group.enter()
        fetchNames(configURL) { list in configs = list; group.leave() }

        group.enter()
        fetchNames(compURL) { list in comps = list; group.leave() }

        group.notify(queue: .main) {
            completion(configs, comps)
        }
    }

    /// Discovered apps (device → detectedApps) with full pagination.
    func fetchDetectedApps(deviceId: String,
                           accessToken: String,
                           completion: @escaping ([DetectedApp]) -> Void) {

        let first = "\(betaURL)/deviceManagement/managedDevices/\(deviceId)/detectedApps?$orderBy=displayName asc"

        func page(_ urlStr: String, _ accum: [DetectedApp]) {
            guard let url = URL(string: urlStr) else {
                DispatchQueue.main.async { completion(accum) }; return
            }

            var req = URLRequest(url: url)
            req.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            req.setValue("application/json", forHTTPHeaderField: "Accept")

            URLSession.shared.dataTask(with: req) { data, resp, _ in
                guard let data = data,
                      let http = resp as? HTTPURLResponse,
                      (200..<300).contains(http.statusCode),
                      let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                      let items = json["value"] as? [[String: Any]] else {
                    DispatchQueue.main.async { completion(accum) }
                    return
                }

                var nextBatch: [DetectedApp] = []

                for obj in items {
                    guard let id = obj["id"] as? String else { continue }

                    // name
                    let name = self.sanitize(obj["displayName"] as? String)
                    guard !name.isEmpty else { continue }

                    // version (string or { value: "..." })
                    var version: String? = nil
                    if let dict = obj["version"] as? [String: Any],
                       let v = dict["value"] as? String {
                        let s = self.sanitize(v, isVersion: true); version = s.isEmpty ? nil : s
                    } else if let v = obj["version"] as? String {
                        let s = self.sanitize(v, isVersion: true); version = s.isEmpty ? nil : s
                    }

                    let pubRaw = obj["publisher"] as? String
                    let publisher = {
                        let s = self.sanitize(pubRaw)
                        return s.isEmpty ? nil : s
                    }()

                    nextBatch.append(DetectedApp(id: id, displayName: name, version: version, publisher: publisher))
                }

                let combined = accum + nextBatch
                if let next = json["@odata.nextLink"] as? String {
                    page(next, combined)
                } else {
                    DispatchQueue.main.async { completion(combined) }
                }
            }.resume()
        }

        page(first, [])
    }

    /// Managed apps for a (user, device). Uses Graph **beta**.
    /// Tries filter form first, then falls back to the path form.
    func fetchManagedApps(userId: String,
                          deviceId: String,
                          accessToken: String,
                          completion: @escaping ([DetectedApp]) -> Void) {

        func mapStates(_ states: [GraphManagedIntentItem]) -> [DetectedApp] {
            var out: [DetectedApp] = []
            for s in states {
                for m in (s.mobileAppList ?? []) {
                    let name = self.sanitize(m.displayName)
                    guard !name.isEmpty else { continue }
                    let version = {
                        let v = self.sanitize(m.version, isVersion: true)
                        return v.isEmpty ? nil : v
                    }()
                    let publisher = {
                        let p = self.sanitize(m.publisher)
                        return p.isEmpty ? nil : p
                    }()
                    var app = DetectedApp(
                        id: m.id ?? UUID().uuidString,
                        displayName: name,
                        version: version,
                        publisher: publisher
                    )
                    app.isManaged = true
                    out.append(app)
                }
            }
            // De-dupe by lowercased name
            var seen = Set<String>(), dedup: [DetectedApp] = []
            for a in out {
                let k = a.displayName.lowercased()
                if !seen.contains(k) { seen.insert(k); dedup.append(a) }
            }
            return dedup.sorted { $0.displayName.localizedCaseInsensitiveCompare($1.displayName) == .orderedAscending }
        }

        func call(_ urlStr: String, fallback: (() -> Void)? = nil) {
            guard let url = URL(string: urlStr) else { fallback?(); return }
            var req = URLRequest(url: url)
            req.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
            req.setValue("application/json", forHTTPHeaderField: "Accept")

            URLSession.shared.dataTask(with: req) { data, resp, _ in
                guard let data = data, let http = resp as? HTTPURLResponse else {
                    DispatchQueue.main.async { fallback?() ?? completion([]) }
                    return
                }

                // If 404/403/empty -> fallback (if any)
                if !(200..<300).contains(http.statusCode) {
                    DispatchQueue.main.async { fallback?() ?? completion([]) }
                    return
                }

                // Try page form first
                if let page = try? self.jsonDecoder.decode(GraphManagedIntentPage.self, from: data),
                   let states = page.value, !states.isEmpty {
                    let apps = mapStates(states)
                    DispatchQueue.main.async { completion(apps) }
                    return
                }

                // Try single-object form
                if let single = try? self.jsonDecoder.decode(GraphManagedIntentItem.self, from: data) {
                    let apps = mapStates([single])
                    DispatchQueue.main.async { completion(apps) }
                    return
                }

                // Nothing usable
                DispatchQueue.main.async { fallback?() ?? completion([]) }
            }.resume()
        }

        // 1) Filter form: .../mobileAppIntentAndStates?$filter=deviceId eq '...'
        let filterURL = "\(betaURL)/users/\(userId)/mobileAppIntentAndStates?$filter=deviceId eq '\(deviceId)'"
        // 2) Fallback path form: .../mobileAppIntentAndStates('{deviceId}')
        let pathURL   = "\(betaURL)/users/\(userId)/mobileAppIntentAndStates('\(deviceId)')"

        call(filterURL) { call(pathURL) { completion([]) } }
    }

    // MARK: - CSV

    func exportToCSV() -> String {
        var csv = """
        Device Name,User Display Name,Email,Department,Job Title,Manager,Office Location,Company,Country,City,OS,OS Version,Model,Serial Number,Compliance State,Management State,Ownership,Last Sync,Enrolled Date,Device Type,Azure AD Device ID,Encrypted,Supervised,Jailbroken
        """

        for d in devices {
            let fields: [String] = csvFields(for: d)
            let line = fields.map(escapeCSV).joined(separator: ",")
            csv.append("\n")
            csv.append(line)
        }
        return csv
    }

    private func csvFields(for d: IntuneDevice) -> [String] {
        let encrypted  = (d.isEncrypted ?? false) ? "true" : "false"
        let supervised = (d.supervisedStatus ?? false) ? "true" : "false"
        let jailbroken = d.jailBroken ?? ""

        let lastSync   = formatDate(d.lastSyncDateTime ?? "")
        let enrolled   = formatDate(d.enrolledDateTime ?? "")

        // Break out each piece to keep type checker happy
        let name       = d.deviceName ?? ""
        let userName   = d.userDisplayName ?? ""
        let email      = d.emailAddress ?? ""
        let dept       = d.userDepartment ?? ""
        let job        = d.userJobTitle ?? ""
        let manager    = d.userManager ?? ""
        let office     = d.userOfficeLocation ?? ""
        let company    = d.userCompanyName ?? ""
        let country    = d.userCountry ?? ""
        let city       = d.userCity ?? ""
        let os         = d.operatingSystem ?? ""
        let osVer      = d.osVersion ?? ""
        let model      = d.model ?? ""
        let serial     = d.serialNumber ?? ""
        let comp       = d.complianceState ?? ""
        let mgmt       = d.managementState ?? ""
        let owner      = d.managedDeviceOwnerType ?? ""
        let type       = d.deviceType ?? ""
        let aad        = d.azureADDeviceId ?? ""

        return [
            name, userName, email, dept, job, manager, office, company, country, city,
            os, osVer, model, serial, comp, mgmt, owner, lastSync, enrolled, type, aad,
            encrypted, supervised, jailbroken
        ]
    }

    private func escapeCSV(_ field: String) -> String {
        guard field.contains(",") || field.contains("\"") || field.contains("\n") else { return field }
        let escaped = field.replacingOccurrences(of: "\"", with: "\"\"")
        return "\"\(escaped)\""
    }

    func saveCSV(_ content: String, suggested filename: String) {
        let panel = NSSavePanel()
        panel.nameFieldStringValue = filename
        panel.allowedContentTypes = [.commaSeparatedText]
        panel.begin { result in
            guard result == .OK, let url = panel.url else { return }
            do { try content.write(to: url, atomically: true, encoding: .utf8) }
            catch { print("Failed to save CSV: \(error)") }
        }
    }

    /// Save a group of CSV files into a named subfolder chosen by the user.
    /// `files` is a map of filename -> CSV content.
    func saveCSVBundle(_ files: [String: String], suggestedFolderName: String) {
        let panel = NSOpenPanel()
        panel.canChooseFiles = false
        panel.canChooseDirectories = true
        panel.allowsMultipleSelection = false
        panel.prompt = "Choose Folder"
        panel.message = "Select a folder to save the export."

        panel.begin { result in
            guard result == .OK, let baseDir = panel.url else { return }
            let folder = baseDir.appendingPathComponent(suggestedFolderName, isDirectory: true)
            do {
                try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
                for (filename, content) in files {
                    let safeName = filename.isEmpty ? "export.csv" : filename
                    let fileURL = folder.appendingPathComponent(safeName)
                    try content.write(to: fileURL, atomically: true, encoding: .utf8)
                }
            } catch {
                print("Failed to save CSV bundle: \(error)")
            }
        }
    }

    // MARK: - Helpers

    private func formatDate(_ iso: String) -> String {
        guard let date = ISO8601DateFormatter().date(from: iso) else { return iso }
        let df = DateFormatter()
        df.dateStyle = .short
        df.timeStyle = .none
        return df.string(from: date)
    }

    private func finishWithError(_ message: String) {
        errorMessage = message
        isLoading = false
        currentOperation = ""
    }
}
