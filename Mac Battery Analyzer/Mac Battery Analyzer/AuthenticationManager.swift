//
//  AuthenticationManager.swift
//  Mac Battery Analyzer
//
//  Created by Eddie Jimenez on 8/13/25.
//

import Foundation
import AuthenticationServices

class AuthenticationManager: NSObject, ObservableObject {
    @Published var isAuthenticated = false
    @Published var accessToken: String?
    @Published var isLoading = false
    @Published var errorMessage: String?
    
    // Azure AD App Registration details
    private let clientId = "YOUR CLIENT ID HERE" // Enter your app registration clientId here
    private let tenantId = "YOUR TENANT ID HERE" // Enter your tenant ID here
    private let redirectUri = "msauth.com.battery.analyzer://auth" // For convenience leave this as-is and use this name as your redirect Uri in the app reg.
    
    private let scopes = [
        "https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All",
        "https://graph.microsoft.com/DeviceManagementManagedDevices.ReadWrite.All",
        "https://graph.microsoft.com/DeviceManagementApps.Read.All",
        "https://graph.microsoft.com/User.Read.All",
        "https://graph.microsoft.com/Directory.Read.All"
    ]
    
    func authenticate() {
        isLoading = true
        errorMessage = nil
        
        // Construct the Microsoft OAuth2 authorization URL
        var components = URLComponents(string: "https://login.microsoftonline.com/\(tenantId)/oauth2/v2.0/authorize")!
        components.queryItems = [
            URLQueryItem(name: "client_id", value: clientId),
            URLQueryItem(name: "response_type", value: "code"),
            URLQueryItem(name: "redirect_uri", value: redirectUri),
            URLQueryItem(name: "scope", value: scopes.joined(separator: " ")),
            URLQueryItem(name: "response_mode", value: "query"),
            URLQueryItem(name: "state", value: UUID().uuidString),
            URLQueryItem(name: "prompt", value: "select_account")
        ]
        
        guard let authURL = components.url else {
            DispatchQueue.main.async {
                self.errorMessage = "Failed to create authentication URL"
                self.isLoading = false
            }
            return
        }
        
        // Start the authentication session
        // FIXED: Using the correct callback scheme that matches your bundle ID
        let session = ASWebAuthenticationSession(url: authURL, callbackURLScheme: "msauth.com.battery.analyzer") { [weak self] url, error in
            DispatchQueue.main.async {
                self?.handleAuthCallback(url: url, error: error)
            }
        }
        
        session.presentationContextProvider = self
        session.prefersEphemeralWebBrowserSession = false
        session.start()
    }
    
    private func handleAuthCallback(url: URL?, error: Error?) {
        isLoading = false
        
        if let error = error {
            if (error as NSError).code == ASWebAuthenticationSessionError.canceledLogin.rawValue {
                errorMessage = "Authentication was cancelled"
            } else {
                errorMessage = "Authentication failed: \(error.localizedDescription)"
            }
            return
        }
        
        guard let url = url,
              let components = URLComponents(url: url, resolvingAgainstBaseURL: false),
              let queryItems = components.queryItems else {
            errorMessage = "Invalid callback URL"
            return
        }
        
        // Check for error in callback
        if (queryItems.first(where: { $0.name == "error" })?.value) != nil {
            let errorDescription = queryItems.first(where: { $0.name == "error_description" })?.value ?? "Unknown error"
            errorMessage = "Authentication error: \(errorDescription)"
            return
        }
        
        // Extract the authorization code
        guard let code = queryItems.first(where: { $0.name == "code" })?.value else {
            errorMessage = "No authorization code received"
            return
        }
        
        // Exchange authorization code for access token
        exchangeCodeForToken(code: code)
    }
    
    private func exchangeCodeForToken(code: String) {
        isLoading = true
        
        guard let tokenURL = URL(string: "https://login.microsoftonline.com/\(tenantId)/oauth2/v2.0/token") else {
            errorMessage = "Invalid token URL"
            isLoading = false
            return
        }
        
        var request = URLRequest(url: tokenURL)
        request.httpMethod = "POST"
        request.setValue("application/x-www-form-urlencoded", forHTTPHeaderField: "Content-Type")
        
        let bodyParameters = [
            "client_id": clientId,
            "scope": scopes.joined(separator: " "),
            "code": code,
            "redirect_uri": redirectUri,
            "grant_type": "authorization_code"
        ]
        
        let bodyString = bodyParameters.map { "\($0.key)=\($0.value.addingPercentEncoding(withAllowedCharacters: .urlQueryAllowed) ?? $0.value)" }
            .joined(separator: "&")
        
        request.httpBody = bodyString.data(using: .utf8)
        
        URLSession.shared.dataTask(with: request) { [weak self] data, response, error in
            DispatchQueue.main.async {
                self?.handleTokenResponse(data: data, response: response, error: error)
            }
        }.resume()
    }
    
    private func handleTokenResponse(data: Data?, response: URLResponse?, error: Error?) {
        isLoading = false
        
        if let error = error {
            errorMessage = "Token exchange failed: \(error.localizedDescription)"
            return
        }
        
        guard let data = data else {
            errorMessage = "No response data received"
            return
        }
        
        do {
            guard let json = try JSONSerialization.jsonObject(with: data) as? [String: Any] else {
                errorMessage = "Invalid response format"
                return
            }
            
            if json["error"] is String {
                let errorDescription = json["error_description"] as? String ?? "Unknown error"
                errorMessage = "Token error: \(errorDescription)"
                return
            }
            
            guard let token = json["access_token"] as? String else {
                errorMessage = "No access token in response"
                return
            }
            
            // Successfully received access token
            accessToken = token
            isAuthenticated = true
            errorMessage = nil
            
        } catch {
            errorMessage = "Failed to parse token response: \(error.localizedDescription)"
        }
    }
    
    func signOut() {
        accessToken = nil
        isAuthenticated = false
        errorMessage = nil
    }
}

// MARK: - ASWebAuthenticationPresentationContextProviding

extension AuthenticationManager: ASWebAuthenticationPresentationContextProviding {
    func presentationAnchor(for session: ASWebAuthenticationSession) -> ASPresentationAnchor {
        return NSApplication.shared.windows.first ?? ASPresentationAnchor()
    }
}
