//
//  ContentView.swift
//  Mac Battery Analyzer
//
import SwiftUI

struct ContentView: View {
    @EnvironmentObject var authManager: AuthenticationManager
    @State private var animateIn = false
    
    var body: some View {
        Group {
            if authManager.isAuthenticated {
                BatteryHealthView()
                    .environmentObject(authManager)
                    .navigationTitle("Mac Battery Analyzer")
                    .toolbar {
                        ToolbarItem(placement: .navigation) {
                            Image(systemName: "bolt.batteryblock")
                                .foregroundStyle(.green)
                        }
                        ToolbarItem(placement: .automatic) {
                            Button {
                                authManager.accessToken = nil
                                authManager.isAuthenticated = false
                            } label: {
                                Label("Sign Out", systemImage: "rectangle.portrait.and.arrow.right")
                            }
                        }
                    }
            } else {
                // Centered sign-in prompt when logged out
                VStack(spacing: 30) {
                    // Power Icon - same as in LocalBatteryHealthView
                    ZStack {
                        Circle()
                            .fill(
                                LinearGradient(
                                    colors: [.green.opacity(0.3), .green.opacity(0.1)],
                                    startPoint: .topLeading,
                                    endPoint: .bottomTrailing
                                )
                            )
                            .frame(width: 100, height: 100)
                            .overlay(
                                Circle()
                                    .stroke(
                                        LinearGradient(
                                            colors: [.green, .green.opacity(0.6)],
                                            startPoint: .topLeading,
                                            endPoint: .bottomTrailing
                                        ),
                                        lineWidth: 3
                                    )
                            )
                        
                        Image(systemName: "bolt.fill")
                            .font(.system(size: 45, weight: .semibold))
                            .foregroundStyle(
                                LinearGradient(
                                    colors: [.green, .green.opacity(0.8)],
                                    startPoint: .top,
                                    endPoint: .bottom
                                )
                            )
                            .scaleEffect(animateIn ? 1.0 : 0.5)
                            .animation(.spring(response: 0.6, dampingFraction: 0.8), value: animateIn)
                    }
                    
                    VStack(spacing: 20) {
                        Text("Sign in to continue")
                            .font(.title2).bold()
                        
                        Button {
                            authManager.authenticate()
                        } label: {
                            HStack {
                                Image(systemName: "lock")
                                Text("Sign In with Microsoft")
                            }
                            .frame(minWidth: 200)
                        }
                        .buttonStyle(.borderedProminent)
                        .controlSize(.large)
                    }
                }
                .frame(maxWidth: .infinity, maxHeight: .infinity)
                .background(Color(NSColor.windowBackgroundColor))
                .navigationTitle("Mac Battery Analyzer")
                .toolbar {
                    ToolbarItem(placement: .navigation) {
                        Image(systemName: "bolt.batteryblock")
                            .foregroundStyle(.green)
                    }
                }
                .onAppear {
                    withAnimation {
                        animateIn = true
                    }
                }
            }
        }
        .frame(minWidth: 1200, minHeight: 700)
    }
}

#Preview {
    ContentView()
        .environmentObject(AuthenticationManager())
}
