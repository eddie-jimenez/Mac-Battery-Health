//
//  BatteryHealth.swift
//  Mac Battery Analyzer
//
//  Native SwiftUI UI with Export functionality
//

import SwiftUI
import Combine
import Charts
import UniformTypeIdentifiers
#if os(macOS)
import AppKit
#endif

struct BatteryHealthView: View {
    @EnvironmentObject var authManager: AuthenticationManager   // must expose `accessToken: String?`

    // Data state
    @State private var items: [BatteryRow] = []
    @State private var isLoading = false
    @State private var errorMessage: String?
    @State private var showingExportMenu = false

    // Controls
    @State private var caInput: String = "6f80311e-8fca-41c1-8dea-cc92272a8406" // GUID or display name
    @State private var searchText: String = ""
    @State private var minHealth: Double = 0
    @State private var onlyOverThreshold: Bool = false
    @State private var onlyCharging: Bool = false

    // Sorting
    enum SortKey: String, CaseIterable {
        case device, user, health, cycles, full, design, current, charging, extPower, minLeft, mV, condition, updated
    }
    @State private var sortKey: SortKey = .device
    @State private var sortAsc: Bool = true

    // UI constants
    private let headerRowHeight: CGFloat = 30
    private let HPAD: CGFloat = 14
    private static var separator: Color {
        #if os(macOS)
        Color(NSColor.separatorColor)
        #else
        Color.secondary.opacity(0.35)
        #endif
    }
    
    // Fixed column widths - single source of truth
    private let columnWidths: [CGFloat] = [
        200,  // Device
        240,  // User
        70,   // Health%
        70,   // Cycles
        90,   // Full mAh
        90,   // Design mAh
        100,  // Current mAh
        80,   // Charging
        70,   // Ext Pwr
        70,   // Min Left
        80,   // mV
        90,   // Condition
        150   // Updated
    ]

    var body: some View {
        VStack(spacing: 0) {
            header
            Divider()
            content
        }
        .task {
            if items.isEmpty, let token = authManager.accessToken, !token.isEmpty {
                await refresh()
            }
        }
    }

    // MARK: Header

    private var header: some View {
        HStack(spacing: 12) {
            Image(systemName: "bolt.batteryblock")
                .font(.system(size: 22))
                .symbolRenderingMode(.hierarchical)
                .foregroundStyle(.green)
            Text("macOS Battery Health")
                .font(.title3).bold()

            Spacer(minLength: 8)

            TextField("CA GUID or Name", text: $caInput)
                .textFieldStyle(.roundedBorder)
                .frame(minWidth: 320)

            Button {
                Task { await refresh() }
            } label: {
                Label("Refresh", systemImage: "arrow.clockwise.circle.fill")
            }
            .keyboardShortcut("r")
            .buttonStyle(.borderedProminent)
            .disabled(isLoading)
            
            // Export button
            Menu {
                Button {
                    exportToCSV()
                } label: {
                    Label("Export as CSV", systemImage: "doc.text")
                }
                
                Button {
                    exportToHTML()
                } label: {
                    Label("Export as HTML", systemImage: "globe")
                }
            } label: {
                Label("Export", systemImage: "square.and.arrow.up")
            }
            .menuStyle(.borderlessButton)
            .disabled(sortedFiltered.isEmpty)
        }
        .padding(.horizontal, HPAD)
        .padding(.vertical, 8)
    }

    // MARK: Content

    private var content: some View {
        ZStack {
            VStack(spacing: 12) {
                if let errorMessage {
                    Text(errorMessage).foregroundStyle(.red).padding(.top, 8)
                }

                toolbarRow

                ViewThatFits(in: .horizontal) {
                    HStack(spacing: 12) {
                        VStack(spacing: 12) {
                            kpiRow
                            histogram
                                .frame(height: 240)
                                .background(.quaternary.opacity(0.08), in: RoundedRectangle(cornerRadius: 14, style: .continuous))
                        }
                        LocalBatteryHealthView()
                            .frame(width: 320, height: 294)
                    }
                    VStack(spacing: 12) {
                        kpiRow
                        histogram
                            .frame(height: 240)
                            .background(.quaternary.opacity(0.08), in: RoundedRectangle(cornerRadius: 14, style: .continuous))
                        LocalBatteryHealthView()
                            .frame(width: 320, height: 294)
                    }
                }
                .padding(.horizontal, HPAD)

                dataTable
                    .padding(.horizontal, HPAD)
                    .padding(.bottom, 12)
            }

            if isLoading {
                ProgressView("Loading Battery Health…")
                    .padding(40)
                    .background(.ultraThinMaterial, in: RoundedRectangle(cornerRadius: 12, style: .continuous))
            }
        }
    }

    // MARK: Toolbar / Filters

    private var toolbarRow: some View {
        ViewThatFits(in: .horizontal) {
            HStack(spacing: 12) { toolbarCore; Spacer(); rowCount }
            VStack(alignment: .leading, spacing: 8) { toolbarCore; rowCount }
        }
        .padding(.horizontal, HPAD)
        .padding(.top, 10)
    }

    @ViewBuilder
    private var toolbarCore: some View {
        HStack(spacing: 8) {
            Image(systemName: "magnifyingglass").foregroundStyle(.secondary)
            TextField("Filter by device or UPN…", text: $searchText)
                .textFieldStyle(.plain)
                .frame(minWidth: 240)
        }
        .padding(.horizontal, 12)
        .padding(.vertical, 8)
        .background(.quaternary.opacity(0.08), in: RoundedRectangle(cornerRadius: 12, style: .continuous))

        HStack(spacing: 10) {
            Text("Min Health").font(.callout).foregroundStyle(.secondary)
            Slider(value: $minHealth, in: 0...100, step: 1)
                .frame(width: 160)
            Text("\(Int(minHealth))%").monospacedDigit().foregroundStyle(.secondary)
        }
        .padding(.horizontal, 12)
        .padding(.vertical, 8)
        .background(.quaternary.opacity(0.08), in: RoundedRectangle(cornerRadius: 999, style: .continuous))

        Toggle("Over Threshold", isOn: $onlyOverThreshold)
            .toggleStyle(.button).buttonStyle(.bordered)
        Toggle("Charging", isOn: $onlyCharging)
            .toggleStyle(.button).buttonStyle(.bordered)
    }

    private var rowCount: some View {
        Text("\(sortedFiltered.count) / \(items.count) rows")
            .font(.callout).foregroundStyle(.secondary)
    }

    // MARK: KPIs

    private var kpiRow: some View {
        HStack(spacing: 12) {
            kpiCard(title: "Devices", value: "\(sortedFiltered.count)")
            kpiCard(title: "Avg Health", value: avgHealthString)
            kpiCard(title: "≥1000 Cycles", value: "\(sortedFiltered.filter { ($0.cycleCount ?? 0) >= 1000 }.count)")
            kpiCard(title: "On External Power", value: "\(sortedFiltered.filter { $0.externalPower == true }.count)")
        }
    }

    @ViewBuilder
    private func kpiCard(title: String, value: String) -> some View {
        VStack(alignment: .leading, spacing: 6) {
            Text(title).font(.caption).foregroundStyle(.secondary)
            Text(value).font(.title).bold().contentTransition(.numericText())
        }
        .frame(maxWidth: .infinity, minHeight: 76, alignment: .leading)
        .padding(14)
        .background(
            LinearGradient(colors: [.green.opacity(0.15), .clear], startPoint: .topLeading, endPoint: .bottomTrailing)
                .background(.quaternary.opacity(0.08))
        )
        .background(.quaternary.opacity(0.08), in: RoundedRectangle(cornerRadius: 14, style: .continuous))
        .overlay(RoundedRectangle(cornerRadius: 14).stroke(.quaternary.opacity(0.25)))
    }

    // MARK: Histogram

    private var histogram: some View {
        let bins: [Int] = [0,60,70,80,90,100,999]
        let counts = histogramCounts(in: bins)

        return Chart {
            ForEach(0..<counts.count, id: \.self) { i in
                BarMark(
                    x: .value("Health Band", bandLabel(bins: bins, index: i)),
                    y: .value("Devices", counts[i])
                )
                .annotation(position: .top, alignment: .center) {
                    Text("\(counts[i])").font(.caption).foregroundStyle(.secondary)
                }
            }
        }
        .chartYAxisLabel("Battery Health (%)")
        .padding(12)
    }

    private func histogramCounts(in bins: [Int]) -> [Int] {
        var arr = Array(repeating: 0, count: bins.count - 1)
        for r in sortedFiltered {
            guard let h = r.healthPercent else { continue }
            for i in 0..<(bins.count - 1) where h >= bins[i] && h < bins[i+1] {
                arr[i] += 1; break
            }
        }
        return arr
    }

    private func bandLabel(bins: [Int], index: Int) -> String {
        if index == 0 { return "<\(bins[1])" }
        if index == bins.count - 2 { return "\(bins[index])+" }
        return "\(bins[index])–\(bins[index+1]-1)"
    }

    // MARK: - Table with Fixed Column Widths

    private var dataTable: some View {
        VStack(spacing: 0) {
            // Header
            HStack(spacing: 0) {
                Button(action: { tapSort(.device) }) {
                    HStack {
                        Text("Device").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .device {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                        Spacer()
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[0], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.user) }) {
                    HStack {
                        Text("User").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .user {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                        Spacer()
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[1], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.health) }) {
                    HStack {
                        Spacer()
                        Text("Health%").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .health {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[2], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.cycles) }) {
                    HStack {
                        Spacer()
                        Text("Cycles").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .cycles {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[3], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.full) }) {
                    HStack {
                        Spacer()
                        Text("Full mAh").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .full {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[4], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.design) }) {
                    HStack {
                        Spacer()
                        Text("Design mAh").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .design {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[5], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.current) }) {
                    HStack {
                        Spacer()
                        Text("Current mAh").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .current {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[6], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.charging) }) {
                    HStack {
                        Spacer()
                        Text("Charging").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .charging {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                        Spacer()
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[7], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.extPower) }) {
                    HStack {
                        Spacer()
                        Text("Ext Pwr").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .extPower {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                        Spacer()
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[8], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.minLeft) }) {
                    HStack {
                        Spacer()
                        Text("Min Left").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .minLeft {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[9], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.mV) }) {
                    HStack {
                        Spacer()
                        Text("mV").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .mV {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[10], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.condition) }) {
                    HStack {
                        Text("Condition").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .condition {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                        Spacer()
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[11], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
                
                Button(action: { tapSort(.updated) }) {
                    HStack {
                        Text("Updated").font(.caption).foregroundStyle(.secondary)
                        if sortKey == .updated {
                            Image(systemName: sortAsc ? "arrow.up" : "arrow.down")
                                .font(.caption2).foregroundStyle(.secondary)
                        }
                        Spacer()
                    }
                    .padding(.horizontal, 8)
                }
                .buttonStyle(.plain)
                .frame(width: columnWidths[12], height: headerRowHeight)
                .border(Self.separator.opacity(0.28), width: 0.5)
            }
            .padding(.vertical, 6)
            
            Divider()
            
            // Data rows
            ScrollView {
                LazyVStack(spacing: 0) {
                    ForEach(sortedFiltered) { r in
                        HStack(spacing: 0) {
                            // Device
                            HStack {
                                Text(r.deviceName).lineLimit(1).truncationMode(.tail)
                                Spacer()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[0])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // User
                            HStack {
                                Text(r.userPrincipalName ?? "—").lineLimit(1).truncationMode(.tail)
                                Spacer()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[1])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Health%
                            HStack {
                                Spacer()
                                Text(fmtInt(r.healthPercent)).monospacedDigit()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[2])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Cycles
                            HStack {
                                Spacer()
                                Text(fmtInt(r.cycleCount)).monospacedDigit()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[3])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Full mAh
                            HStack {
                                Spacer()
                                Text(fmtNum(r.fullCharge_mAh)).monospacedDigit()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[4])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Design mAh
                            HStack {
                                Spacer()
                                Text(fmtNum(r.design_mAh)).monospacedDigit()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[5])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Current mAh
                            HStack {
                                Spacer()
                                Text(fmtNum(r.current_mAh)).monospacedDigit()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[6])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Charging
                            ZStack { pill(r.isCharging) }
                            .frame(width: columnWidths[7])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Ext Power
                            ZStack { pill(r.externalPower) }
                            .frame(width: columnWidths[8])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Min Left
                            HStack {
                                Spacer()
                                Text(fmtInt(r.timeRemainingMin)).monospacedDigit()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[9])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // mV
                            HStack {
                                Spacer()
                                Text(fmtNum(r.voltage_mV)).monospacedDigit()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[10])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Condition
                            HStack {
                                Text(r.condition ?? "—").lineLimit(1).truncationMode(.tail)
                                Spacer()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[11])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                            
                            // Updated
                            HStack {
                                Text(fmtDate(r.whenISO)).foregroundStyle(.secondary).lineLimit(1).truncationMode(.tail)
                                Spacer()
                            }
                            .padding(.horizontal, 8)
                            .frame(width: columnWidths[12])
                            .border(Self.separator.opacity(0.28), width: 0.5)
                        }
                        .padding(.vertical, 10)
                        .overlay(alignment: .bottom) {
                            Rectangle().fill(Self.separator.opacity(0.15)).frame(height: 0.5)
                        }
                    }
                }
            }
            .background(.quaternary.opacity(0.04), in: RoundedRectangle(cornerRadius: 14, style: .continuous))
            .clipShape(RoundedRectangle(cornerRadius: 14, style: .continuous))
        }
    }

    // MARK: Helpers

    private func tapSort(_ key: SortKey) {
        if sortKey == key { sortAsc.toggle() } else { sortKey = key; sortAsc = true }
    }

    @ViewBuilder
    private func pill(_ state: Bool?) -> some View {
        switch state {
        case true:
            Text("Yes")
                .padding(.horizontal, 8).padding(.vertical, 4)
                .background(Color.green.opacity(0.18), in: Capsule())
                .overlay(Capsule().stroke(Color.green.opacity(0.35)))
        case false:
            Text("No")
                .padding(.horizontal, 8).padding(.vertical, 4)
                .background(Color.red.opacity(0.18), in: Capsule())
                .overlay(Capsule().stroke(Color.red.opacity(0.35)))
        default:
            Text("—").foregroundStyle(.secondary)
        }
    }

    // MARK: Formatting

    private static let numberFormatter: NumberFormatter = {
        let f = NumberFormatter()
        f.numberStyle = .decimal
        f.groupingSeparator = Locale.current.groupingSeparator
        return f
    }()

    private func fmtInt(_ v: Int?) -> String { v.map(String.init) ?? "—" }

    private func fmtNum(_ v: Int?) -> String {
        guard let v else { return "—" }
        return BatteryHealthView.numberFormatter.string(from: NSNumber(value: v)) ?? String(v)
    }

    private func fmtDate(_ iso: String) -> String {
        let f = ISO8601DateFormatter()
        if let d = f.date(from: iso) {
            let out = DateFormatter()
            out.dateStyle = .short
            out.timeStyle = .short
            return out.string(from: d)
        }
        return iso
    }

    // MARK: Derived values

    private var filtered: [BatteryRow] {
        items.filter { r in
            let q = searchText.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
            if !q.isEmpty {
                let nameHit = r.deviceName.lowercased().contains(q)
                let upnHit = (r.userPrincipalName ?? "").lowercased().contains(q)
                if !nameHit && !upnHit { return false }
            }
            if let h = r.healthPercent, h < Int(minHealth) { return false }
            if onlyOverThreshold && r.overThreshold != true { return false }
            if onlyCharging && r.isCharging != true { return false }
            return true
        }
    }

    private var sortedFiltered: [BatteryRow] {
        filtered.sorted { a, b in
            func cmp<T: Comparable>(_ x: T?, _ y: T?) -> Bool {
                switch (x, y) {
                case let (l?, r?): return sortAsc ? (l < r) : (l > r)
                case (nil, _?):     return false
                case (_?, nil):     return true
                default:            return false
                }
            }
            switch sortKey {
            case .device:    return cmp(a.deviceName, b.deviceName)
            case .user:      return cmp(a.userPrincipalName ?? "", b.userPrincipalName ?? "")
            case .health:    return cmp(a.healthPercent, b.healthPercent)
            case .cycles:    return cmp(a.cycleCount, b.cycleCount)
            case .full:      return cmp(a.fullCharge_mAh, b.fullCharge_mAh)
            case .design:    return cmp(a.design_mAh, b.design_mAh)
            case .current:   return cmp(a.current_mAh, b.current_mAh)
            case .charging:  return cmp(a.isCharging == true ? 1 : 0, b.isCharging == true ? 1 : 0)
            case .extPower:  return cmp(a.externalPower == true ? 1 : 0, b.externalPower == true ? 1 : 0)
            case .minLeft:   return cmp(a.timeRemainingMin, b.timeRemainingMin)
            case .mV:        return cmp(a.voltage_mV, b.voltage_mV)
            case .condition: return cmp(a.condition ?? "", b.condition ?? "")
            case .updated:
                let fa = ISO8601DateFormatter().date(from: a.whenISO)
                let fb = ISO8601DateFormatter().date(from: b.whenISO)
                return cmp(fa, fb)
            }
        }
    }

    private var avgHealthString: String {
        let hs = sortedFiltered.compactMap { $0.healthPercent }
        guard !hs.isEmpty else { return "—" }
        let avg = Double(hs.reduce(0,+)) / Double(hs.count)
        return String(format: "%.1f%%", avg)
    }

    // MARK: Export Functions
    
    private func exportToCSV() {
        let csv = generateCSV()
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "yyyy-MM-dd_HHmmss"
        let fileName = "battery_health_\(dateFormatter.string(from: Date())).csv"
        
        #if os(macOS)
        let savePanel = NSSavePanel()
        savePanel.allowedContentTypes = [.commaSeparatedText]
        savePanel.nameFieldStringValue = fileName
        savePanel.begin { response in
            if response == .OK, let url = savePanel.url {
                do {
                    try csv.write(to: url, atomically: true, encoding: .utf8)
                    // Show success alert
                    DispatchQueue.main.async {
                        let alert = NSAlert()
                        alert.messageText = "Export Successful"
                        alert.informativeText = "CSV file has been saved to:\n\(url.lastPathComponent)"
                        alert.alertStyle = .informational
                        alert.addButton(withTitle: "OK")
                        alert.runModal()
                    }
                } catch {
                    // Show error alert
                    DispatchQueue.main.async {
                        let alert = NSAlert()
                        alert.messageText = "Export Failed"
                        alert.informativeText = "Failed to save CSV: \(error.localizedDescription)"
                        alert.alertStyle = .critical
                        alert.addButton(withTitle: "OK")
                        alert.runModal()
                    }
                }
            }
        }
        #endif
    }
    
    private func generateCSV() -> String {
        var csv = "Device,User,Health%,Cycles,Full_mAh,Design_mAh,Current_mAh,Charging,Ext_Power,Min_Left,mV,Condition,Updated\n"
        
        for row in sortedFiltered {
            let device = escapeCSV(row.deviceName)
            let user = escapeCSV(row.userPrincipalName ?? "")
            let health = row.healthPercent.map(String.init) ?? ""
            let cycles = row.cycleCount.map(String.init) ?? ""
            let full = row.fullCharge_mAh.map(String.init) ?? ""
            let design = row.design_mAh.map(String.init) ?? ""
            let current = row.current_mAh.map(String.init) ?? ""
            let charging = row.isCharging == true ? "Yes" : row.isCharging == false ? "No" : ""
            let extPower = row.externalPower == true ? "Yes" : row.externalPower == false ? "No" : ""
            let minLeft = row.timeRemainingMin.map(String.init) ?? ""
            let voltage = row.voltage_mV.map(String.init) ?? ""
            let condition = escapeCSV(row.condition ?? "")
            let updated = escapeCSV(fmtDate(row.whenISO))
            
            csv += "\(device),\(user),\(health),\(cycles),\(full),\(design),\(current),\(charging),\(extPower),\(minLeft),\(voltage),\(condition),\(updated)\n"
        }
        
        return csv
    }
    
    private func escapeCSV(_ field: String) -> String {
        if field.contains(",") || field.contains("\"") || field.contains("\n") {
            return "\"\(field.replacingOccurrences(of: "\"", with: "\"\""))\""
        }
        return field
    }
    
    private func exportToHTML() {
        let html = generateHTML()
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "yyyy-MM-dd_HHmmss"
        let fileName = "battery_health_\(dateFormatter.string(from: Date())).html"
        
        #if os(macOS)
        let savePanel = NSSavePanel()
        savePanel.allowedContentTypes = [.html]
        savePanel.nameFieldStringValue = fileName
        savePanel.begin { response in
            if response == .OK, let url = savePanel.url {
                do {
                    try html.write(to: url, atomically: true, encoding: .utf8)
                    // Show success alert
                    DispatchQueue.main.async {
                        let alert = NSAlert()
                        alert.messageText = "Export Successful"
                        alert.informativeText = "HTML report has been saved to:\n\(url.lastPathComponent)"
                        alert.alertStyle = .informational
                        alert.addButton(withTitle: "OK")
                        alert.runModal()
                    }
                } catch {
                    // Show error alert
                    DispatchQueue.main.async {
                        let alert = NSAlert()
                        alert.messageText = "Export Failed"
                        alert.informativeText = "Failed to save HTML: \(error.localizedDescription)"
                        alert.alertStyle = .critical
                        alert.addButton(withTitle: "OK")
                        alert.runModal()
                    }
                }
            }
        }
        #endif
    }
    
    private func generateHTML() -> String {
        let dateFormatter = DateFormatter()
        dateFormatter.dateStyle = .long
        dateFormatter.timeStyle = .medium
        let exportDate = dateFormatter.string(from: Date())
        
        // Calculate histogram data
        let bins: [Int] = [0,60,70,80,90,100,999]
        let counts = histogramCounts(in: bins)
        let maxCount = counts.max() ?? 1
        
        var html = """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>macOS Battery Health Report</title>
            <style>
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
                    background: linear-gradient(135deg, #1e1e2e 0%, #2d2d44 100%);
                    color: #e0e0e0;
                    padding: 20px;
                    min-height: 100vh;
                }
                .container {
                    max-width: 1400px;
                    margin: 0 auto;
                    background: rgba(30, 30, 46, 0.95);
                    border-radius: 20px;
                    padding: 30px;
                    box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5);
                }
                .header {
                    display: flex;
                    align-items: center;
                    gap: 15px;
                    margin-bottom: 30px;
                    padding-bottom: 20px;
                    border-bottom: 1px solid rgba(255, 255, 255, 0.1);
                }
                .header-icon {
                    width: 40px;
                    height: 40px;
                    background: linear-gradient(135deg, #4ade80, #22c55e);
                    border-radius: 10px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 24px;
                }
                h1 {
                    font-size: 28px;
                    font-weight: 600;
                    background: linear-gradient(135deg, #4ade80, #22c55e);
                    -webkit-background-clip: text;
                    -webkit-text-fill-color: transparent;
                }
                .export-info {
                    margin-left: auto;
                    font-size: 14px;
                    color: #9ca3af;
                }
                .filters {
                    background: rgba(255, 255, 255, 0.05);
                    border-radius: 12px;
                    padding: 15px;
                    margin-bottom: 25px;
                    display: flex;
                    gap: 20px;
                    flex-wrap: wrap;
                    align-items: center;
                }
                .filter-item {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    font-size: 14px;
                }
                .filter-label { color: #9ca3af; }
                .filter-value {
                    background: rgba(74, 222, 128, 0.2);
                    padding: 4px 12px;
                    border-radius: 20px;
                    color: #4ade80;
                }
                .kpi-row {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                    gap: 20px;
                    margin-bottom: 30px;
                }
                .kpi-card {
                    background: linear-gradient(135deg, rgba(74, 222, 128, 0.15), transparent);
                    border-radius: 14px;
                    padding: 20px;
                    border: 1px solid rgba(74, 222, 128, 0.2);
                }
                .kpi-title {
                    font-size: 12px;
                    color: #9ca3af;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                    margin-bottom: 8px;
                }
                .kpi-value {
                    font-size: 32px;
                    font-weight: 700;
                    color: #4ade80;
                }
                .histogram {
                    background: rgba(255, 255, 255, 0.05);
                    border-radius: 14px;
                    padding: 20px;
                    margin-bottom: 30px;
                }
                .histogram-title {
                    font-size: 14px;
                    color: #9ca3af;
                    margin-bottom: 15px;
                }
                .histogram-bars {
                    display: flex;
                    align-items: flex-end;
                    justify-content: space-between;
                    height: 150px;
                    gap: 10px;
                    margin-bottom: 10px;
                }
                .histogram-bar {
                    flex: 1;
                    background: linear-gradient(to top, #4ade80, #22c55e);
                    border-radius: 4px 4px 0 0;
                    position: relative;
                    min-height: 2px;
                    display: flex;
                    align-items: flex-start;
                    justify-content: center;
                    padding-top: 5px;
                }
                .histogram-bar-value {
                    position: absolute;
                    top: -20px;
                    font-size: 12px;
                    color: #9ca3af;
                }
                .histogram-labels {
                    display: flex;
                    justify-content: space-between;
                    gap: 10px;
                }
                .histogram-label {
                    flex: 1;
                    text-align: center;
                    font-size: 11px;
                    color: #6b7280;
                }
                .data-table {
                    background: rgba(255, 255, 255, 0.03);
                    border-radius: 14px;
                    overflow: hidden;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                th {
                    background: rgba(255, 255, 255, 0.05);
                    padding: 12px;
                    text-align: left;
                    font-size: 12px;
                    font-weight: 600;
                    color: #9ca3af;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                    border-bottom: 1px solid rgba(255, 255, 255, 0.1);
                }
                td {
                    padding: 12px;
                    font-size: 14px;
                    border-bottom: 1px solid rgba(255, 255, 255, 0.05);
                }
                tr:hover {
                    background: rgba(255, 255, 255, 0.03);
                }
                .pill {
                    display: inline-block;
                    padding: 3px 10px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: 500;
                }
                .pill-yes {
                    background: rgba(74, 222, 128, 0.2);
                    color: #4ade80;
                    border: 1px solid rgba(74, 222, 128, 0.3);
                }
                .pill-no {
                    background: rgba(239, 68, 68, 0.2);
                    color: #ef4444;
                    border: 1px solid rgba(239, 68, 68, 0.3);
                }
                .health-badge {
                    display: inline-block;
                    padding: 2px 8px;
                    border-radius: 4px;
                    font-weight: 600;
                }
                .health-good { background: rgba(74, 222, 128, 0.2); color: #4ade80; }
                .health-warning { background: rgba(251, 191, 36, 0.2); color: #fbbf24; }
                .health-danger { background: rgba(239, 68, 68, 0.2); color: #ef4444; }
                .text-right { text-align: right; }
                .text-center { text-align: center; }
                .summary {
                    margin-top: 30px;
                    padding: 20px;
                    background: rgba(255, 255, 255, 0.05);
                    border-radius: 12px;
                    font-size: 14px;
                    color: #9ca3af;
                    text-align: center;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <div class="header-icon">⚡</div>
                    <h1>macOS Battery Health Report</h1>
                    <div class="export-info">
                        <div>Exported: \(exportDate)</div>
                        <div style="font-size: 12px; margin-top: 4px;">Total Devices: \(sortedFiltered.count)</div>
                    </div>
                </div>
        """
        
        // Add filters info if any are active
        if minHealth > 0 || onlyOverThreshold || onlyCharging || !searchText.isEmpty {
            html += """
                <div class="filters">
                    <div style="font-size: 12px; color: #6b7280;">Active Filters:</div>
            """
            if minHealth > 0 {
                html += """
                    <div class="filter-item">
                        <span class="filter-label">Min Health:</span>
                        <span class="filter-value">≥\(Int(minHealth))%</span>
                    </div>
                """
            }
            if onlyOverThreshold {
                html += """
                    <div class="filter-item">
                        <span class="filter-label">Cycles:</span>
                        <span class="filter-value">≥1000</span>
                    </div>
                """
            }
            if onlyCharging {
                html += """
                    <div class="filter-item">
                        <span class="filter-label">Status:</span>
                        <span class="filter-value">Charging</span>
                    </div>
                """
            }
            if !searchText.isEmpty {
                html += """
                    <div class="filter-item">
                        <span class="filter-label">Search:</span>
                        <span class="filter-value">\(escapeHTML(searchText))</span>
                    </div>
                """
            }
            html += "</div>"
        }
        
        // KPI Cards
        html += """
            <div class="kpi-row">
                <div class="kpi-card">
                    <div class="kpi-title">Devices</div>
                    <div class="kpi-value">\(sortedFiltered.count)</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-title">Average Health</div>
                    <div class="kpi-value">\(avgHealthString)</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-title">≥1000 Cycles</div>
                    <div class="kpi-value">\(sortedFiltered.filter { ($0.cycleCount ?? 0) >= 1000 }.count)</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-title">On External Power</div>
                    <div class="kpi-value">\(sortedFiltered.filter { $0.externalPower == true }.count)</div>
                </div>
            </div>
        """
        
        // Histogram
        html += """
            <div class="histogram">
                <div class="histogram-title">BATTERY HEALTH DISTRIBUTION</div>
                <div class="histogram-bars">
        """
        
        for i in 0..<counts.count {
            let height = maxCount > 0 ? (Double(counts[i]) / Double(maxCount)) * 100 : 0
            html += """
                    <div class="histogram-bar" style="height: \(height)%;">
                        <div class="histogram-bar-value">\(counts[i])</div>
                    </div>
            """
        }
        
        html += """
                </div>
                <div class="histogram-labels">
        """
        
        for i in 0..<counts.count {
            let label = bandLabel(bins: bins, index: i)
            html += """
                    <div class="histogram-label">\(label)</div>
            """
        }
        
        html += """
                </div>
            </div>
        """
        
        // Data Table
        html += """
            <div class="data-table">
                <table>
                    <thead>
                        <tr>
                            <th>Device</th>
                            <th>User</th>
                            <th class="text-right">Health%</th>
                            <th class="text-right">Cycles</th>
                            <th class="text-right">Full mAh</th>
                            <th class="text-right">Design mAh</th>
                            <th class="text-right">Current mAh</th>
                            <th class="text-center">Charging</th>
                            <th class="text-center">Ext Pwr</th>
                            <th class="text-right">Min Left</th>
                            <th class="text-right">mV</th>
                            <th>Condition</th>
                            <th>Updated</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        for row in sortedFiltered {
            let healthClass = if let h = row.healthPercent {
                h >= 90 ? "health-good" : h >= 70 ? "health-warning" : "health-danger"
            } else { "" }
            
            html += """
                        <tr>
                            <td>\(escapeHTML(row.deviceName))</td>
                            <td>\(escapeHTML(row.userPrincipalName ?? "—"))</td>
                            <td class="text-right">
            """
            
            if let health = row.healthPercent {
                html += """
                                <span class="health-badge \(healthClass)">\(health)%</span>
                """
            } else {
                html += "—"
            }
            
            html += """
                            </td>
                            <td class="text-right">\(row.cycleCount.map(String.init) ?? "—")</td>
                            <td class="text-right">\(fmtNum(row.fullCharge_mAh))</td>
                            <td class="text-right">\(fmtNum(row.design_mAh))</td>
                            <td class="text-right">\(fmtNum(row.current_mAh))</td>
                            <td class="text-center">
            """
            
            if let charging = row.isCharging {
                html += """
                                <span class="pill \(charging ? "pill-yes" : "pill-no")">\(charging ? "Yes" : "No")</span>
                """
            } else {
                html += "—"
            }
            
            html += """
                            </td>
                            <td class="text-center">
            """
            
            if let extPower = row.externalPower {
                html += """
                                <span class="pill \(extPower ? "pill-yes" : "pill-no")">\(extPower ? "Yes" : "No")</span>
                """
            } else {
                html += "—"
            }
            
            html += """
                            </td>
                            <td class="text-right">\(row.timeRemainingMin.map(String.init) ?? "—")</td>
                            <td class="text-right">\(fmtNum(row.voltage_mV))</td>
                            <td>\(escapeHTML(row.condition ?? "—"))</td>
                            <td>\(fmtDate(row.whenISO))</td>
                        </tr>
            """
        }
        
        html += """
                    </tbody>
                </table>
            </div>
            
            <div class="summary">
                Generated by Mac Battery Analyzer • \(sortedFiltered.count) devices • Report created on \(exportDate)
            </div>
        </div>
        </body>
        </html>
        """
        
        return html
    }
    
    private func escapeHTML(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&#39;")
    }

    // MARK: Networking (unchanged from original)

    private var betaBase: String { "https://graph.microsoft.com/beta" }

    private func looksLikeGUID(_ s: String) -> Bool {
        let regex = try! NSRegularExpression(pattern: "^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$")
        return regex.firstMatch(in: s, range: NSRange(location: 0, length: s.utf16.count)) != nil
    }

    private func makeRequest(url: URL, token: String) -> URLRequest {
        var r = URLRequest(url: url)
        r.setValue("Bearer \(token)", forHTTPHeaderField: "Authorization")
        r.setValue("application/json", forHTTPHeaderField: "Accept")
        return r
    }

    private func resolveCAId(_ input: String, token: String) async throws -> String {
        if looksLikeGUID(input) { return input }
        let listURL = URL(string: "\(betaBase)/deviceManagement/deviceCustomAttributeShellScripts?$select=id,displayName,createdDateTime&$top=999")!
        let (data, resp) = try await URLSession.shared.data(for: makeRequest(url: listURL, token: token))
        guard let http = resp as? HTTPURLResponse, (200..<300).contains(http.statusCode) else {
            throw NSError(domain: "BatteryHealth", code: httpStatus(resp), userInfo: [NSLocalizedDescriptionKey: "CA list fetch failed"])
        }
        struct CAList: Decodable { let value: [Item]; struct Item: Decodable { let id: String; let displayName: String; let createdDateTime: String } }
        let list = try JSONDecoder().decode(CAList.self, from: data)
        if let exact = list.value.first(where: { $0.displayName == input }) { return exact.id }
        if let fuzzy = list.value
            .filter({ $0.displayName.lowercased().contains(input.lowercased()) })
            .sorted(by: { $0.createdDateTime > $1.createdDateTime })
            .first { return fuzzy.id }
        throw NSError(domain: "BatteryHealth", code: 404, userInfo: [NSLocalizedDescriptionKey: "Custom Attribute not found: \(input)"])
    }

    private func fetchAllRunStates(caId: String, token: String) async throws -> [CAState] {
        var all: [CAState] = []
        var next: URL? = URL(string:
            "\(betaBase)/deviceManagement/deviceCustomAttributeShellScripts/\(caId)/deviceRunStates" +
            "?$select=id,lastStateUpdateDateTime,resultMessage,runState,errorCode,errorDescription" +
            "&$expand=managedDevice($select=id,deviceName,userPrincipalName,osVersion,userId)" +
            "&$top=200"
        )
        let dec = JSONDecoder()
        while let url = next {
            let (data, resp) = try await URLSession.shared.data(for: makeRequest(url: url, token: token))
            guard let http = resp as? HTTPURLResponse, (200..<300).contains(http.statusCode) else {
                throw NSError(domain: "BatteryHealth", code: httpStatus(resp), userInfo: [NSLocalizedDescriptionKey: "Run states fetch failed (\(httpStatus(resp)))"])
            }
            let page = try dec.decode(CAStatePage.self, from: data)
            all.append(contentsOf: page.value)
            if let n = page.nextLink, let u = URL(string: n) { next = u } else { next = nil }
        }
        return all
    }

    private func parseCSV(_ csv: String) -> (Int?, Int?, Int?, Int?, Int?, Bool?, Bool?, Int?, Int?, String?, Bool?) {
        let parts = csv.split(separator: ",", maxSplits: 10, omittingEmptySubsequences: false)
            .map { String($0).trimmingCharacters(in: .whitespacesAndNewlines) }
        func toInt(_ s: String?) -> Int? { guard let s, !s.isEmpty, s.lowercased() != "none" else { return nil }; return Int(s) }
        func toBool(_ s: String?) -> Bool? {
            guard let s else { return nil }
            switch s.lowercased() { case "true": return true; case "false": return false; default: return nil }
        }
        func toStr(_ s: String?) -> String? { guard let s, !s.isEmpty, s.lowercased() != "none" else { return nil }; return s }
        let h    = parts.count > 0  ? toInt(parts[0]) : nil
        let cyc  = parts.count > 1  ? toInt(parts[1]) : nil
        let fcc  = parts.count > 2  ? toInt(parts[2]) : nil
        let des  = parts.count > 3  ? toInt(parts[3]) : nil
        let cur  = parts.count > 4  ? toInt(parts[4]) : nil
        let ich  = parts.count > 5  ? toBool(parts[5]) : nil
        let ext  = parts.count > 6  ? toBool(parts[6]) : nil
        let trm  = parts.count > 7  ? toInt(parts[7]) : nil
        let volt = parts.count > 8  ? toInt(parts[8]) : nil
        let cond = parts.count > 9  ? toStr(parts[9]) : nil
        let ovr  = parts.count > 10 ? toBool(parts[10]) : nil
        return (h,cyc,fcc,des,cur,ich,ext,trm,volt,cond,ovr)
    }

    private func buildRows(states: [CAState]) -> [BatteryRow] {
        states.compactMap { s in
            guard let id = s.id,
                  let md = s.managedDevice,
                  let devId = md.id,
                  let devName = md.deviceName,
                  let when = s.lastStateUpdateDateTime,
                  let run = s.runState else { return nil }
            let raw = (s.resultMessage ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
            let (h,cyc,fcc,des,cur,ich,ext,trm,volt,cond,ovr) = parseCSV(raw)
            return BatteryRow(
                id: id,
                deviceId: devId,
                deviceName: devName,
                userPrincipalName: md.userPrincipalName,
                osVersion: md.osVersion,
                userId: md.userId,
                whenISO: when,
                runState: run,
                errorCode: s.errorCode,
                errorDescription: s.errorDescription,
                healthPercent: h,
                cycleCount: cyc,
                fullCharge_mAh: fcc,
                design_mAh: des,
                current_mAh: cur,
                isCharging: ich,
                externalPower: ext,
                timeRemainingMin: trm,
                voltage_mV: volt,
                condition: cond,
                overThreshold: ovr,
                rawCsv: raw
            )
        }
    }

    private func fetchAndRender(caInput: String, accessToken: String) async throws -> [BatteryRow] {
        let caId = try await resolveCAId(caInput, token: accessToken)
        let states = try await fetchAllRunStates(caId: caId, token: accessToken)
        return buildRows(states: states)
    }

    private func refresh() async {
        guard let token = authManager.accessToken, !token.isEmpty else {
            errorMessage = "Sign in first."
            return
        }
        errorMessage = nil
        isLoading = true
        defer { isLoading = false }
        do {
            let rows = try await fetchAndRender(caInput: caInput, accessToken: token)
            withAnimation { self.items = rows }
        } catch {
            self.errorMessage = error.localizedDescription
        }
    }

    private func httpStatus(_ resp: URLResponse?) -> Int {
        (resp as? HTTPURLResponse)?.statusCode ?? -1
    }
}

// MARK: - Data Model

struct BatteryRow: Identifiable, Codable {
    let id: String
    let deviceId: String
    let deviceName: String
    let userPrincipalName: String?
    let osVersion: String?
    let userId: String?

    let whenISO: String
    let runState: String
    let errorCode: Int?
    let errorDescription: String?

    // Parsed CSV metrics (11 fields)
    let healthPercent: Int?
    let cycleCount: Int?
    let fullCharge_mAh: Int?
    let design_mAh: Int?
    let current_mAh: Int?
    let isCharging: Bool?
    let externalPower: Bool?
    let timeRemainingMin: Int?
    let voltage_mV: Int?
    let condition: String?
    let overThreshold: Bool?

    // Raw CSV for reference
    let rawCsv: String
}

// MARK: - Graph DTOs

private struct CAStatePage: Decodable {
    let value: [CAState]
    let nextLink: String?
    enum CodingKeys: String, CodingKey { case value; case nextLink = "@odata.nextLink" }
}
private struct CAState: Decodable {
    let id: String?
    let lastStateUpdateDateTime: String?
    let resultMessage: String?
    let runState: String?
    let errorCode: Int?
    let errorDescription: String?
    let managedDevice: ManagedDevice?
}
private struct ManagedDevice: Decodable {
    let id: String?
    let deviceName: String?
    let userPrincipalName: String?
    let osVersion: String?
    let userId: String?
}
