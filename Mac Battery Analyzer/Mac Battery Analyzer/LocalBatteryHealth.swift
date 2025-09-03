//
//  LocalBatteryHealth.swift
//  Mac Battery Analyzer
//
//  Local device battery health visualization
//

import SwiftUI
import Combine

struct LocalBatteryHealthView: View {
    @State private var batteryData: LocalBatteryData?
    @State private var powerData: PowerData?
    @State private var isLoading = true
    @State private var timer: Timer?
    
    var body: some View {
        ZStack {
            if let data = batteryData {
                BatteryVisual(data: data)
            } else if let power = powerData {
                PowerVisual(data: power)
            } else if isLoading {
                ProgressView()
                    .scaleEffect(0.8)
            } else {
                VStack(spacing: 8) {
                    Image(systemName: "bolt.slash")
                        .font(.system(size: 40))
                        .foregroundStyle(.tertiary)
                    Text("No Power Data")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
            }
        }
        .frame(width: 320, height: 294)
        .background(.quaternary.opacity(0.08), in: RoundedRectangle(cornerRadius: 14, style: .continuous))
        .task {
            await loadBatteryData()
            // Refresh every 30 seconds
            timer = Timer.scheduledTimer(withTimeInterval: 30, repeats: true) { _ in
                Task { await loadBatteryData() }
            }
        }
        .onDisappear {
            timer?.invalidate()
        }
    }
    
    private func loadBatteryData() async {
        // First try the battery script
        let batteryScript = """
        #!/bin/bash
        PB="/usr/libexec/PlistBuddy"
        TMP="$(/usr/bin/mktemp -t batt).plist"
        trap 'rm -f "$TMP"' EXIT
        
        /usr/sbin/ioreg -a -r -c AppleSmartBattery > "$TMP" 2>/dev/null || true
        
        if ! /usr/bin/grep -q "<dict>" "$TMP" 2>/dev/null; then
          echo "None"
          exit 0
        fi
        
        pb() { "$PB" -c "Print :0:$1" "$TMP" 2>/dev/null || true; }
        intval() { case "$1" in (''|*[!0-9-]*) echo 0 ;; (*) echo "$1" ;; esac; }
        yn() {
          v="$(echo "$1" | /usr/bin/tr '[:upper:]' '[:lower:]')"
          case "$v" in (yes|true|1) echo "True" ;; (no|false|0) echo "False" ;; (*) echo "None" ;; esac
        }
        val_or_none() { case "$1" in (''|*[!0-9-]*) echo "None" ;; (*) echo "$1" ;; esac; }
        
        design_s="$(pb DesignCapacity)"
        rawmax_s="$(pb AppleRawMaxCapacity)"
        nominal_s="$(pb NominalChargeCapacity)"
        max_s="$(pb MaxCapacity)"
        cur_s="$(pb CurrentCapacity)"
        rawcur_s="$(pb AppleRawCurrentCapacity)"
        cycles_s="$(pb CycleCount)"
        ischg_s="$(pb IsCharging)"
        ext_s="$(pb ExternalConnected)"
        trem_s="$(pb TimeRemaining)"
        volt_s="$(pb Voltage)"
        cond_s="$(pb BatteryHealth)"; [ -z "$cond_s" ] && cond_s="$(pb Condition)"
        
        i_design=$(intval "$design_s")
        i_rawmax=$(intval "$rawmax_s")
        i_nominal=$(intval "$nominal_s")
        i_max=$(intval "$max_s")
        i_cur=$(intval "$cur_s")
        i_rawcur=$(intval "$rawcur_s")
        i_cycles=$(intval "$cycles_s")
        i_trem=$(intval "$trem_s")
        
        fcc=0
        if [ "$i_rawmax" -gt 0 ]; then
          fcc=$i_rawmax
        elif [ "$i_max" -gt 200 ]; then
          fcc=$i_max
        elif [ "$i_nominal" -gt 0 ] && [ "$i_max" -ge 0 ] && [ "$i_max" -le 110 ]; then
          fcc=$(( (i_nominal * i_max) / 100 ))
        fi
        
        curr=0
        if [ "$i_rawcur" -gt 0 ]; then
          curr=$i_rawcur
        elif [ "$i_cur" -gt 200 ]; then
          curr=$i_cur
        fi
        
        health="None"
        if [ "$i_design" -gt 0 ] && [ "$fcc" -gt 0 ]; then
          health=$(( (100 * fcc) / i_design ))
        elif [ "$i_max" -ge 0 ] && [ "$i_max" -le 110 ]; then
          health=$i_max
        fi
        
        trem_out="None"
        if [ "$i_trem" -ge 0 ] && [ "$i_trem" -lt 60000 ]; then trem_out="$i_trem"; fi
        
        [ -z "$cond_s" ] && cond_s="None"
        
        CYCLE_MAX=1000
        over="False"; [ "$i_cycles" -ge "$CYCLE_MAX" ] && over="True"
        
        echo "$( [ "$health" = "None" ] && echo "None" || echo "$health"),"\\
        "$(val_or_none "$i_cycles"),"\\
        "$(val_or_none "$fcc"),"\\
        "$(val_or_none "$i_design"),"\\
        "$(val_or_none "$curr"),"\\
        "$(yn "$ischg_s"),"\\
        "$(yn "$ext_s"),"\\
        "$trem_out,"\\
        "$(val_or_none "$volt_s"),"\\
        "$cond_s,"\\
        "$over"
        """
        
        do {
            let process = Process()
            process.executableURL = URL(fileURLWithPath: "/bin/bash")
            process.arguments = ["-c", batteryScript]
            
            let pipe = Pipe()
            process.standardOutput = pipe
            
            try process.run()
            process.waitUntilExit()
            
            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            if let output = String(data: data, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines) {
                if output == "None" {
                    // No battery found, try to get power info
                    await loadPowerData()
                } else {
                    await MainActor.run {
                        self.batteryData = parseBatteryData(output)
                        self.isLoading = false
                    }
                }
            }
        } catch {
            await loadPowerData()
        }
    }
    
    private func loadPowerData() async {
        // Get power adapter and system info for desktop Macs
        let powerScript = """
        #!/bin/bash
        
        # Get system model
        model=$(system_profiler SPHardwareDataType 2>/dev/null | grep "Model Name" | awk -F': ' '{print $2}')
        
        # Get chip info (Apple Silicon or Intel)
        chip=$(system_profiler SPHardwareDataType 2>/dev/null | grep "Chip:" | awk -F': ' '{print $2}')
        if [ -z "$chip" ]; then
            chip=$(sysctl -n machdep.cpu.brand_string | awk '{print $1, $2}')
        fi
        
        # Get number of CPU cores
        cores=$(sysctl -n hw.ncpu)
        
        # Get power adapter info if available
        adapter_info=$(system_profiler SPPowerDataType 2>/dev/null | grep -A 20 "AC Charger Information")
        
        # Extract wattage if available
        wattage=$(echo "$adapter_info" | grep "Wattage" | awk -F': ' '{print $2}' | awk '{print $1}')
        [ -z "$wattage" ] && wattage="Unknown"
        
        # Check if connected to power
        connected=$(echo "$adapter_info" | grep "Connected:" | awk -F': ' '{print $2}')
        [ -z "$connected" ] && connected="Yes"  # Desktop Macs are always "connected"
        
        # Get system load (1 minute average)
        load=$(uptime | awk -F'load averages:' '{print $2}' | awk '{print $1}')
        
        # Get memory info
        memory=$(system_profiler SPHardwareDataType | grep "Memory:" | awk -F': ' '{print $2}')
        
        # Get memory pressure
        mem_pressure=$(memory_pressure | grep "System-wide memory free percentage:" | awk -F': ' '{print $2}' | tr -d '%')
        
        echo "$model|$wattage|$connected|$load|$chip|$memory|$cores|$mem_pressure"
        """
        
        do {
            let process = Process()
            process.executableURL = URL(fileURLWithPath: "/bin/bash")
            process.arguments = ["-c", powerScript]
            
            let pipe = Pipe()
            process.standardOutput = pipe
            
            try process.run()
            process.waitUntilExit()
            
            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            if let output = String(data: data, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines) {
                await MainActor.run {
                    self.powerData = parsePowerData(output)
                    self.isLoading = false
                }
            }
        } catch {
            await MainActor.run {
                self.isLoading = false
            }
        }
    }
    
    private func parsePowerData(_ output: String) -> PowerData {
        let parts = output.split(separator: "|").map { String($0) }
        
        return PowerData(
            model: parts.count > 0 ? parts[0] : "Unknown",
            wattage: parts.count > 1 ? parts[1] : "Unknown",
            connected: parts.count > 2 ? parts[2] == "Yes" : true,
            load: parts.count > 3 ? Double(parts[3]) : nil,
            chip: parts.count > 4 ? parts[4] : "Unknown",
            memory: parts.count > 5 ? parts[5] : "Unknown",
            cores: parts.count > 6 ? Int(parts[6]) ?? 0 : 0,
            memoryPressure: parts.count > 7 ? Int(parts[7]) ?? 0 : 0
        )
    }
    
    private func parseBatteryData(_ csv: String) -> LocalBatteryData? {
        if csv == "None" { return nil }
        
        let parts = csv.split(separator: ",").map { String($0).trimmingCharacters(in: .whitespaces) }
        guard parts.count >= 11 else { return nil }
        
        return LocalBatteryData(
            healthPercent: Int(parts[0]) ?? 0,
            cycleCount: Int(parts[1]),
            fullCharge_mAh: Int(parts[2]),
            design_mAh: Int(parts[3]),
            current_mAh: Int(parts[4]),
            isCharging: parts[5] == "True",
            externalPower: parts[6] == "True",
            timeRemaining: Int(parts[7]),
            voltage_mV: Int(parts[8]),
            condition: parts[9] == "None" ? nil : parts[9],
            overThreshold: parts[10] == "True"
        )
    }
}

struct BatteryVisual: View {
    let data: LocalBatteryData
    @State private var animateIn = false
    
    private var healthColor: Color {
        switch data.healthPercent {
        case 90...100: return .green
        case 70..<90: return .yellow
        case 50..<70: return .orange
        default: return .red
        }
    }
    
    private var chargeLevel: Double {
        guard let current = data.current_mAh,
              let full = data.fullCharge_mAh,
              full > 0 else { return 0 }
        return Double(current) / Double(full)
    }
    
    var body: some View {
        VStack(spacing: 20) {
            // Battery Icon with Fill
            ZStack {
                // Battery outline
                BatteryShape()
                    .stroke(healthColor.opacity(0.3), lineWidth: 3)
                    .frame(width: 180, height: 80)
                
                // Battery fill based on current charge
                BatteryShape()
                    .fill(
                        LinearGradient(
                            colors: [healthColor, healthColor.opacity(0.7)],
                            startPoint: .leading,
                            endPoint: .trailing
                        )
                    )
                    .mask(
                        GeometryReader { geo in
                            Rectangle()
                                .frame(width: geo.size.width * chargeLevel)
                        }
                    )
                    .frame(width: 180, height: 80)
                    .animation(.easeInOut(duration: 1.5), value: animateIn)
                
                // Charging indicator
                if data.isCharging {
                    Image(systemName: "bolt.fill")
                        .font(.system(size: 30, weight: .semibold))
                        .foregroundStyle(.white)
                        .shadow(radius: 3)
                }
                
                // Percentage text
                if !data.isCharging {
                    Text("\(Int(chargeLevel * 100))%")
                        .font(.system(size: 24, weight: .semibold, design: .rounded))
                        .foregroundStyle(.white)
                        .shadow(radius: 3)
                }
            }
            
            // Health percentage with ring
            ZStack {
                // Background ring
                Circle()
                    .stroke(Color.gray.opacity(0.2), lineWidth: 12)
                    .frame(width: 100, height: 100)
                
                // Health ring
                Circle()
                    .trim(from: 0, to: animateIn ? Double(data.healthPercent) / 100 : 0)
                    .stroke(
                        healthColor,
                        style: StrokeStyle(lineWidth: 12, lineCap: .round)
                    )
                    .frame(width: 100, height: 100)
                    .rotationEffect(.degrees(-90))
                    .animation(.easeInOut(duration: 1.5), value: animateIn)
                
                VStack(spacing: 2) {
                    Text("\(data.healthPercent)%")
                        .font(.system(size: 28, weight: .bold, design: .rounded))
                    Text("Health")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
            }
            
            // Stats grid
            HStack(spacing: 20) {
                StatItem(
                    icon: "arrow.triangle.2.circlepath",
                    value: data.cycleCount.map(String.init) ?? "—",
                    label: "Cycles",
                    highlight: data.overThreshold
                )
                
                StatItem(
                    icon: data.externalPower ? "bolt.circle.fill" : "battery.25",
                    value: data.externalPower ? "Connected" : "Battery",
                    label: "Power",
                    highlight: data.externalPower
                )
                
                if let minutes = data.timeRemaining {
                    StatItem(
                        icon: "clock",
                        value: formatTime(minutes),
                        label: "Remaining",
                        highlight: false
                    )
                }
            }
            .padding(.top, 8)
        }
        .padding()
        .onAppear {
            withAnimation {
                animateIn = true
            }
        }
    }
    
    private func formatTime(_ minutes: Int) -> String {
        let hours = minutes / 60
        let mins = minutes % 60
        if hours > 0 {
            return "\(hours)h \(mins)m"
        } else {
            return "\(mins)m"
        }
    }
}

struct StatItem: View {
    let icon: String
    let value: String
    let label: String
    let highlight: Bool
    
    var body: some View {
        VStack(spacing: 4) {
            Image(systemName: icon)
                .font(.system(size: 18))
                .foregroundStyle(highlight ? .green : .secondary)
            Text(value)
                .font(.system(size: 14, weight: .semibold, design: .rounded))
            Text(label)
                .font(.caption2)
                .foregroundStyle(.secondary)
        }
        .frame(minWidth: 60)
    }
}

struct BatteryShape: Shape {
    func path(in rect: CGRect) -> Path {
        var path = Path()
        
        let bodyWidth = rect.width * 0.95
        let bodyHeight = rect.height
        let capWidth = rect.width * 0.05
        let capHeight = rect.height * 0.35
        let cornerRadius: CGFloat = 8
        
        // Main battery body
        path.addRoundedRect(
            in: CGRect(x: 0, y: 0, width: bodyWidth, height: bodyHeight),
            cornerSize: CGSize(width: cornerRadius, height: cornerRadius)
        )
        
        // Battery cap
        path.addRoundedRect(
            in: CGRect(
                x: bodyWidth,
                y: (bodyHeight - capHeight) / 2,
                width: capWidth,
                height: capHeight
            ),
            cornerSize: CGSize(width: 2, height: 2)
        )
        
        return path
    }
}

struct LocalBatteryData {
    let healthPercent: Int
    let cycleCount: Int?
    let fullCharge_mAh: Int?
    let design_mAh: Int?
    let current_mAh: Int?
    let isCharging: Bool
    let externalPower: Bool
    let timeRemaining: Int?
    let voltage_mV: Int?
    let condition: String?
    let overThreshold: Bool
}

struct PowerData {
    let model: String
    let wattage: String
    let connected: Bool
    let load: Double?
    let chip: String
    let memory: String
    let cores: Int
    let memoryPressure: Int
}

struct PowerVisual: View {
    let data: PowerData
    @State private var animateIn = false
    @State private var currentTime = Date()
    
    private let timer = Timer.publish(every: 1, on: .main, in: .common).autoconnect()
    
    private var loadPercentage: Double {
        guard let load = data.load, data.cores > 0 else { return 0 }
        // Load average divided by cores gives us a rough CPU usage percentage
        return min((load / Double(data.cores)) * 100, 100)
    }
    
    private var loadDescription: String {
        guard let load = data.load else { return "Unknown" }
        if loadPercentage < 25 {
            return "System Idle"
        } else if loadPercentage < 50 {
            return "Light Load"
        } else if loadPercentage < 75 {
            return "Moderate Load"
        } else {
            return "Heavy Load"
        }
    }
    
    var body: some View {
        VStack(spacing: 16) {
            // Power Icon with pulse animation for high load
            ZStack {
                Circle()
                    .fill(
                        LinearGradient(
                            colors: [.green.opacity(0.3), .green.opacity(0.1)],
                            startPoint: .topLeading,
                            endPoint: .bottomTrailing
                        )
                    )
                    .frame(width: 90, height: 90)
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
                    .scaleEffect(loadPercentage > 75 ? (animateIn ? 1.05 : 1.0) : 1.0)
                    .animation(.easeInOut(duration: 1).repeatForever(autoreverses: true), value: animateIn)
                
                Image(systemName: "bolt.fill")
                    .font(.system(size: 40, weight: .semibold))
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
            
            // System Info
            VStack(spacing: 6) {
                Text(data.model)
                    .font(.system(size: 18, weight: .semibold, design: .rounded))
                
                // Live clock - the cool part you like!
                Text(currentTime, style: .time)
                    .font(.system(size: 16, weight: .medium, design: .monospaced))
                    .foregroundStyle(.primary)
                    .onReceive(timer) { _ in
                        currentTime = Date()
                    }
            }
            
            // CPU Load Visualization
            if let load = data.load {
                VStack(spacing: 8) {
                    // Load bar with percentage
                    VStack(alignment: .leading, spacing: 4) {
                        HStack {
                            Text("CPU Load")
                                .font(.caption)
                                .foregroundStyle(.secondary)
                            Spacer()
                            Text("\(Int(loadPercentage))%")
                                .font(.caption)
                                .monospacedDigit()
                                .foregroundStyle(loadColor(for: load))
                        }
                        
                        GeometryReader { geometry in
                            ZStack(alignment: .leading) {
                                RoundedRectangle(cornerRadius: 4)
                                    .fill(Color.gray.opacity(0.2))
                                    .frame(height: 8)
                                
                                RoundedRectangle(cornerRadius: 4)
                                    .fill(
                                        LinearGradient(
                                            colors: [loadColor(for: load), loadColor(for: load).opacity(0.7)],
                                            startPoint: .leading,
                                            endPoint: .trailing
                                        )
                                    )
                                    .frame(width: geometry.size.width * (loadPercentage / 100), height: 8)
                                    .animation(.easeInOut(duration: 0.5), value: loadPercentage)
                            }
                        }
                        .frame(height: 8)
                    }
                    
                    Text(loadDescription)
                        .font(.caption2)
                        .foregroundStyle(.secondary)
                        
                    Text("Load Average: \(String(format: "%.2f", load)) / \(data.cores) cores")
                        .font(.caption2)
                        .foregroundStyle(.tertiary)
                }
                .padding(.horizontal)
            }
            
            // System Stats - More meaningful info
            HStack(spacing: 20) {
                // Processor Info
                VStack(spacing: 4) {
                    Image(systemName: "cpu")
                        .font(.system(size: 20))
                        .foregroundStyle(.blue)
                    Text(data.chip)
                        .font(.caption2)
                        .foregroundStyle(.primary)
                        .lineLimit(1)
                    Text("\(data.cores) cores")
                        .font(.caption2)
                        .foregroundStyle(.secondary)
                }
                .frame(maxWidth: .infinity)
                
                // Memory Info with pressure indicator
                VStack(spacing: 4) {
                    ZStack {
                        Image(systemName: "memorychip")
                            .font(.system(size: 20))
                            .foregroundStyle(memoryColor)
                        
                        if data.memoryPressure > 70 {
                            Image(systemName: "exclamationmark.triangle.fill")
                                .font(.system(size: 10))
                                .foregroundStyle(.orange)
                                .offset(x: 10, y: -10)
                        }
                    }
                    Text(data.memory)
                        .font(.caption2)
                        .foregroundStyle(.primary)
                    Text("\(100 - data.memoryPressure)% free")
                        .font(.caption2)
                        .foregroundStyle(memoryColor)
                }
                .frame(maxWidth: .infinity)
                
                // Power Status
                VStack(spacing: 4) {
                    Image(systemName: data.connected ? "bolt.circle.fill" : "bolt.slash.circle")
                        .font(.system(size: 20))
                        .foregroundStyle(data.connected ? .green : .gray)
                    Text("AC Power")
                        .font(.caption2)
                        .foregroundStyle(.primary)
                    Text(data.connected ? "Connected" : "No Power")
                        .font(.caption2)
                        .foregroundStyle(.secondary)
                }
                .frame(maxWidth: .infinity)
            }
        }
        .padding()
        .onAppear {
            withAnimation {
                animateIn = true
            }
        }
    }
    
    private func loadColor(for load: Double) -> Color {
        let percentage = loadPercentage
        switch percentage {
        case 0..<25: return .green
        case 25..<50: return .yellow
        case 50..<75: return .orange
        default: return .red
        }
    }
    
    private var memoryColor: Color {
        let free = 100 - data.memoryPressure
        switch free {
        case 50...100: return .green
        case 30..<50: return .yellow
        case 15..<30: return .orange
        default: return .red
        }
    }
}
