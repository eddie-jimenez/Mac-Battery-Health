//
//  Mac_Battery_AnalyzerApp.swift
//  Mac Battery Analyzer
//
//  Created by Eddie Jimenez on 8/28/25.
//

import SwiftUI

@main
struct Mac_Battery_AnalyzerApp: App {
    @StateObject private var authManager = AuthenticationManager()

    var body: some Scene {
        WindowGroup {
            ContentView()
                .environmentObject(authManager)
                .accentColor(.green)  // Add this line
        }
    }
}

