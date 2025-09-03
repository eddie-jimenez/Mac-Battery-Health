//
//  Models.swift
//  Intune Manager
//
//  Created by Eddie Jimenez on 8/13/25.
//

import Foundation

// MARK: - Device Models

struct IntuneDevice: Identifiable, Codable {
    let id: String
    let deviceName: String?
    let managementState: String?
    let complianceState: String?
    let osVersion: String?
    let deviceType: String?
    let lastSyncDateTime: String?
    let enrolledDateTime: String?
    let model: String?
    let manufacturer: String?
    let serialNumber: String?
    let userPrincipalName: String?
    let userId: String?
    let azureADDeviceId: String?
    let managedDeviceOwnerType: String?
    let emailAddress: String?
    let azureADRegistered: Bool?
    let enrollmentType: String?
    let activationLockBypassCode: String?
    let phoneNumber: String?
    let operatingSystem: String?
    let jailBroken: String?
    let managementAgent: String?
    let deviceRegistrationState: String?
    let supervisedStatus: Bool?
    let exchangeAccessState: String?
    let exchangeAccessStateReason: String?
    let remoteAssistanceSessionUrl: String?
    let isEncrypted: Bool?
    let userDisplayName: String?
    let configurationManagerClientEnabledFeatures: ConfigManagerFeatures?
    let wiFiMacAddress: String?
    let primaryUser: String?
    let enrollmentProfileName: String?
    
    // Extended user properties from Entra
    var userDepartment: String?
    var userJobTitle: String?
    var userManager: String?
    var userOfficeLocation: String?
    var userCompanyName: String?
    var userCountry: String?
    var userCity: String?
    
    enum CodingKeys: String, CodingKey {
        case id
        case deviceName
        case managementState
        case complianceState
        case osVersion
        case deviceType
        case lastSyncDateTime
        case enrolledDateTime
        case model
        case manufacturer
        case serialNumber
        case userPrincipalName
        case userId
        case azureADDeviceId
        case managedDeviceOwnerType
        case emailAddress
        case azureADRegistered
        case enrollmentType
        case activationLockBypassCode
        case phoneNumber
        case operatingSystem
        case jailBroken
        case managementAgent
        case deviceRegistrationState
        case supervisedStatus
        case exchangeAccessState
        case exchangeAccessStateReason
        case remoteAssistanceSessionUrl
        case isEncrypted
        case userDisplayName
        case configurationManagerClientEnabledFeatures
        case wiFiMacAddress
        case primaryUser
        case enrollmentProfileName
    }
}

struct ConfigManagerFeatures: Codable {
    let isInventoryEnabled: Bool?
    let isModernAppManagementEnabled: Bool?
    let isOfficeAppsEnabled: Bool?
    let isResourceAccessEnabled: Bool?
    let isWindowsUpdateEnabled: Bool?
}

// MARK: - User Models

struct EntraUser: Codable {
    let id: String
    let displayName: String?
    let userPrincipalName: String?
    let department: String?
    let jobTitle: String?
    let manager: Manager?
    let officeLocation: String?
    let companyName: String?
    let country: String?
    let city: String?
    let mail: String?
    let mobilePhone: String?
    
    struct Manager: Codable {
        let displayName: String?
        let id: String?
    }
}

// MARK: - Filter Options

enum ComplianceFilter: String, CaseIterable {
    case all = "All"
    case compliant = "Compliant"
    case noncompliant = "Noncompliant"
    case unknown = "Unknown"
    case inGracePeriod = "InGracePeriod"
    case configManager = "ConfigManager"
}

enum PlatformFilter: String, CaseIterable {
    case all = "All"
    case iOS = "iOS"
    case android = "Android"
    case windows = "Windows"
    case macOS = "macOS"
    case other = "Other"
}

enum OwnershipFilter: String, CaseIterable {
    case all = "All"
    case personal = "Personal"
    case corporate = "Corporate"
    case unknown = "Unknown"
}

// MARK: - DetectedApp (discovered apps for a device)
struct DetectedApp: Identifiable, Codable, Hashable {
    let id: String
    let displayName: String
    let version: String?
    let publisher: String?
    var isManaged: Bool? = nil  // To distinguish between discovered and managed apps
    
    // Additional properties for app details
    var description: String?
    var largeIcon: String?
    var createdDateTime: String?
    var lastModifiedDateTime: String?
    var isFeatured: Bool?
    var privacyInformationUrl: String?
    var informationUrl: String?
    var owner: String?
    var developer: String?
    var notes: String?
}
