import AppKit
import Foundation
import SwiftUI
import UniformTypeIdentifiers

private let openFilesNotification = Notification.Name("DocumentConverterOpenFiles")

final class AppDelegate: NSObject, NSApplicationDelegate {
    private var pendingURLs: [URL] = []

    func applicationDidFinishLaunching(_ notification: Notification) {
        NSApplication.shared.setActivationPolicy(.regular)
        NSApplication.shared.activate(ignoringOtherApps: true)
        if !pendingURLs.isEmpty {
            NotificationCenter.default.post(name: openFilesNotification, object: pendingURLs)
            pendingURLs.removeAll()
        }
    }

    func application(_ application: NSApplication, open urls: [URL]) {
        if application.isRunning {
            NotificationCenter.default.post(name: openFilesNotification, object: urls)
        } else {
            pendingURLs.append(contentsOf: urls)
        }
    }

    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool {
        true
    }
}

enum FileStatus: String {
    case queued, running, ok, warning, error

    var label: String {
        switch self {
        case .queued: return "待转换"
        case .running: return "转换中"
        case .ok: return "已转换"
        case .warning: return "已转换"
        case .error: return "转换失败"
        }
    }

    var symbol: String {
        switch self {
        case .queued: return "clock"
        case .running: return "arrow.triangle.2.circlepath"
        case .ok: return "checkmark.circle.fill"
        case .warning: return "exclamationmark.circle.fill"
        case .error: return "xmark.circle.fill"
        }
    }

    var color: Color {
        switch self {
        case .queued: return .secondary
        case .running: return .accentColor
        case .ok: return .green
        case .warning: return .orange
        case .error: return .red
        }
    }
}

struct ConversionFile: Identifiable {
    let id = UUID()
    let input: URL
    var status: FileStatus = .queued
    var output: URL?
    var message = ""
}

struct BundledRuntime: Sendable {
    let python: URL
    let script: URL
    let engine: URL
    let pythonHome: URL

    static func locate() throws -> BundledRuntime {
        guard let resources = Bundle.main.resourceURL else {
            throw ConverterError.runtime("找不到应用资源，请重新获取完整的应用。")
        }
        let pythonHome = resources.appendingPathComponent("runtime/python", isDirectory: true)
        let python = pythonHome.appendingPathComponent("bin/python3")
        let script = resources.appendingPathComponent("converter.py")
        let engine = resources.appendingPathComponent("libreoffice/LibreOfficeDev.app/Contents/MacOS/soffice")
        guard FileManager.default.isExecutableFile(atPath: python.path) else {
            throw ConverterError.runtime("应用内的运行环境缺失，请重新获取完整的应用。")
        }
        guard FileManager.default.fileExists(atPath: script.path) else {
            throw ConverterError.runtime("应用内的转换程序缺失，请重新获取完整的应用。")
        }
        guard FileManager.default.isExecutableFile(atPath: engine.path) else {
            throw ConverterError.runtime("应用内的转换引擎缺失，请重新获取完整的应用。")
        }
        return BundledRuntime(python: python, script: script, engine: engine, pythonHome: pythonHome)
    }
}

enum ConverterError: LocalizedError {
    case runtime(String)
    var errorDescription: String? {
        switch self { case .runtime(let text): return text }
    }
}

enum ConversionDestination: Sendable {
    case besideInputs
    case folder(URL)
}

/// Reads pipe data away from the UI thread. Each conversion's output is one JSON
/// line; the engine also sends a start event so the UI does not infer progress.
private func runEngine(
    runtime: BundledRuntime,
    inputs: [URL],
    destination: ConversionDestination,
    onLine: @escaping @Sendable (Data) async -> Void
) async throws -> (Int32, String) {
    let manager = FileManager.default
    let temporaryDirectory = manager.temporaryDirectory.appendingPathComponent("MDw2d-Batch-" + UUID().uuidString, isDirectory: true)
    try manager.createDirectory(at: temporaryDirectory, withIntermediateDirectories: false, attributes: [.posixPermissions: 0o700])
    defer { try? manager.removeItem(at: temporaryDirectory) }
    let manifest = temporaryDirectory.appendingPathComponent("inputs.json")
    let manifestData = try JSONSerialization.data(withJSONObject: inputs.map(\.path))
    guard manager.createFile(atPath: manifest.path, contents: manifestData, attributes: [.posixPermissions: 0o600]) else {
        throw ConverterError.runtime("无法准备转换清单，请检查临时文件夹的访问权限。")
    }
    let process = Process()
    process.executableURL = runtime.python
    process.arguments = [
        "-I", "-B", "-u", runtime.script.path,
        "--engine", runtime.engine.path,
        "--input-list-file", manifest.path
    ]
    switch destination {
    case .besideInputs: process.arguments! += ["--beside-input"]
    case .folder(let directory): process.arguments! += ["--output-dir", directory.path]
    }
    var environment = ProcessInfo.processInfo.environment
    environment["PYTHONHOME"] = runtime.pythonHome.path
    environment["PYTHONNOUSERSITE"] = "1"
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    environment["PYTHONUNBUFFERED"] = "1"
    process.environment = environment
    let stdout = Pipe()
    let stderr = Pipe()
    process.standardOutput = stdout
    process.standardError = stderr

    // stdout is streamed below; drain stderr concurrently to avoid a full-pipe
    // deadlock if an engine writes diagnostics during a long document import.
    let stderrTask = Task.detached(priority: .utility) {
        stderr.fileHandleForReading.readDataToEndOfFile()
    }
    do {
        try process.run()
    } catch {
        try? stdout.fileHandleForWriting.close()
        try? stderr.fileHandleForWriting.close()
        _ = await stderrTask.value
        throw error
    }
    try? stdout.fileHandleForWriting.close()
    try? stderr.fileHandleForWriting.close()

    var pending = Data()
    while true {
        let data = stdout.fileHandleForReading.availableData
        if data.isEmpty { break }
        pending.append(data)
        while let newline = pending.firstIndex(of: 0x0A) {
            let line = pending.prefix(upTo: newline)
            if !line.isEmpty { await onLine(Data(line)) }
            pending.removeSubrange(...newline)
        }
    }
    if !pending.isEmpty { await onLine(pending) }
    process.waitUntilExit()
    let diagnostic = String(data: await stderrTask.value, encoding: .utf8) ?? ""
    return (process.terminationStatus, diagnostic)
}

@MainActor
final class ConverterModel: ObservableObject {
    @Published var files: [ConversionFile] = []
    @Published var outputDirectory: URL?
    @Published var isWorking = false
    @Published var isScanning = false
    @Published var includeSubfolders = true
    @Published private(set) var saveBesideOriginals = false
    @Published private(set) var completedCount = 0
    @Published var dropTargeted = false
    @Published var alertMessage: String?
    @Published var batchMessage = ""

    private var scanTask: Task<FolderScanResult, Error>?
    private var scanToken = UUID()
    private var sourceFolder: URL?
    private var scanNotice = ""
    private var inputIndices: [String: Int] = [:]

    var isBusy: Bool { isWorking || isScanning }
    var canStart: Bool { !isBusy && !files.isEmpty && (saveBesideOriginals || outputDirectory != nil) }
    var outputLabel: String {
        saveBesideOriginals ? "各原文件所在的文件夹" : outputDirectory?.path ?? "请选择转换后 DOCX 文件的保存文件夹"
    }

    func addFiles(_ urls: [URL]) {
        guard !isBusy else { return }
        let directories = urls.filter { (try? $0.resourceValues(forKeys: [.isDirectoryKey]).isDirectory) == true }
        if !directories.isEmpty {
            if urls.count == 1, let folder = directories.first { convertFolder(folder) }
            else { alertMessage = "文件夹请单独添加或拖入；文件和文件夹不能混合加入自动转换。" }
            return
        }
        if files.isEmpty {
            saveBesideOriginals = false
            scanNotice = ""
        }
        var skipped = 0
        var paths = Set(files.map { $0.input.path })
        var additions: [ConversionFile] = []
        for url in urls where url.isFileURL {
            let normalized = url.standardizedFileURL.resolvingSymlinksInPath()
            var directory: ObjCBool = false
            guard FileManager.default.fileExists(atPath: normalized.path, isDirectory: &directory), !directory.boolValue else {
                skipped += 1
                continue
            }
            guard paths.insert(normalized.path).inserted else { continue }
            additions.append(ConversionFile(input: normalized))
        }
        files.append(contentsOf: additions)
        if skipped > 0 { alertMessage = "请选择文件。文件夹或不存在的文件未加入转换列表。" }
        batchMessage = ""
    }

    func chooseFiles() {
        guard !isBusy else { return }
        let panel = NSOpenPanel()
        panel.title = "选择要转换的文档"
        panel.prompt = "添加文档"
        panel.canChooseFiles = true
        panel.canChooseDirectories = false
        panel.allowsMultipleSelection = true
        // Misnamed files are permitted: the engine identifies the actual format.
        if panel.runModal() == .OK { addFiles(panel.urls) }
    }

    func chooseSourceFolder() {
        guard !isBusy else { return }
        let panel = NSOpenPanel()
        panel.title = "选择要自动转换的文件夹"
        panel.prompt = "选择并转换"
        panel.message = "查找 WPS、DOC、RTF、ODT 等文字文档，跳过 DOCX。结果保存在各原文件旁，保留原件。"
        panel.canChooseFiles = false
        panel.canChooseDirectories = true
        panel.canCreateDirectories = false
        panel.allowsMultipleSelection = false
        panel.directoryURL = sourceFolder
        if panel.runModal() == .OK, let folder = panel.url { convertFolder(folder) }
    }

    func convertFolder(_ folder: URL) {
        guard !isBusy else { return }
        let recursive = includeSubfolders
        let token = UUID()
        scanToken = token
        isScanning = true
        batchMessage = recursive ? "正在查找文件夹及子文件夹中的文档…" : "正在查找文件夹中的文档…"
        let worker = Task.detached(priority: .userInitiated) {
            try FolderScanner.scan(folder: folder, includeSubfolders: recursive)
        }
        scanTask = worker
        Task { [weak self] in
            do {
                let result = try await worker.value
                guard let self, self.scanToken == token, self.isScanning else { return }
                self.scanTask = nil
                self.isScanning = false
                self.sourceFolder = folder
                self.saveBesideOriginals = true
                self.files = result.files.map { ConversionFile(input: $0) }
                self.completedCount = 0
                self.inputIndices = [:]
                var notes: [String] = []
                if result.skippedDOCX > 0 { notes.append("已跳过 \(result.skippedDOCX) 个 DOCX") }
                if !result.unreadablePaths.isEmpty { notes.append("另有 \(result.unreadablePaths.count) 个位置无法读取") }
                self.scanNotice = notes.joined(separator: "，")
                if result.files.isEmpty {
                    self.batchMessage = "未找到需要转换的文字文档。" + self.scanNotice
                } else {
                    self.start()
                }
            } catch {
                guard let self, self.scanToken == token, self.isScanning else { return }
                self.scanTask = nil
                self.isScanning = false
                self.batchMessage = error is CancellationError ? "已取消查找，未开始转换。" : "文件夹查找失败，未开始转换。"
                if !(error is CancellationError) { self.alertMessage = error.localizedDescription }
            }
        }
    }

    func cancelScan() {
        guard isScanning else { return }
        scanToken = UUID()
        scanTask?.cancel()
        scanTask = nil
        isScanning = false
        batchMessage = "已取消查找，未开始转换。"
    }

    func chooseOutputDirectory() {
        guard !isBusy else { return }
        let panel = NSOpenPanel()
        panel.title = "选择 DOCX 文件的保存位置"
        panel.prompt = "选择保存位置"
        panel.canChooseFiles = false
        panel.canChooseDirectories = true
        panel.canCreateDirectories = true
        panel.allowsMultipleSelection = false
        panel.directoryURL = outputDirectory
        if panel.runModal() == .OK {
            outputDirectory = panel.url
            saveBesideOriginals = false
        }
    }

    func remove(_ id: UUID) {
        guard !isBusy else { return }
        if let file = files.first(where: { $0.id == id }), [.ok, .warning, .error].contains(file.status) { completedCount -= 1 }
        files.removeAll { $0.id == id }
        batchMessage = ""
    }

    func clear() {
        guard !isBusy else { return }
        files.removeAll()
        saveBesideOriginals = false
        completedCount = 0
        inputIndices = [:]
        scanNotice = ""
        batchMessage = ""
    }

    func acceptDrop(_ providers: [NSItemProvider]) -> Bool {
        guard !isBusy else { return false }
        let accepted = providers.filter { $0.hasItemConformingToTypeIdentifier(UTType.fileURL.identifier) }
        Task { [weak self] in
            var urls: [URL] = []
            for provider in accepted {
                let url: URL? = await withCheckedContinuation { continuation in
                    provider.loadItem(forTypeIdentifier: UTType.fileURL.identifier, options: nil) { item, _ in
                        if let value = item as? URL { continuation.resume(returning: value) }
                        else if let value = item as? Data { continuation.resume(returning: URL(dataRepresentation: value, relativeTo: nil)) }
                        else if let value = item as? String { continuation.resume(returning: URL(string: value)) }
                        else { continuation.resume(returning: nil) }
                    }
                }
                if let url { urls.append(url) }
            }
            self?.addFiles(urls)
        }
        return !accepted.isEmpty
    }

    func start() {
        guard canStart else { return }
        let destination: ConversionDestination
        if saveBesideOriginals { destination = .besideInputs }
        else if let outputDirectory { destination = .folder(outputDirectory) }
        else { return }
        let runtime: BundledRuntime
        do { runtime = try BundledRuntime.locate() }
        catch { alertMessage = error.localizedDescription; return }
        let inputs = files.map(\.input)
        inputIndices = Dictionary(uniqueKeysWithValues: files.enumerated().map { ($0.element.input.path, $0.offset) })
        completedCount = 0
        files = files.map { file in
            var queued = file
            queued.status = .queued
            queued.output = nil
            queued.message = ""
            return queued
        }
        isWorking = true
        batchMessage = "正在转换，原文件保持不变。" + scanNotice
        Task { [self] in
            let outcome = await Task.detached(priority: .userInitiated) { [self] () -> Result<(Int32, String), Error> in
                do {
                    return .success(try await runEngine(runtime: runtime, inputs: inputs, destination: destination) { [self] data in
                        await self.receive(data)
                    })
                } catch { return .failure(error) }
            }.value
            finish(outcome)
        }
    }

    private func receive(_ data: Data) {
        guard let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
              let input = json["input"] as? String,
              let index = inputIndices[URL(fileURLWithPath: input).standardizedFileURL.path] else { return }
        if json["event"] as? String == "start" {
            files[index].status = .running
            return
        }
        guard let rawStatus = json["status"] as? String,
              let status = FileStatus(rawValue: rawStatus), [.ok, .warning, .error].contains(status) else { return }
        if ![.ok, .warning, .error].contains(files[index].status) { completedCount += 1 }
        files[index].status = status
        files[index].message = json["message"] as? String ?? ""
        if status != .error,
           let verification = json["verification"] as? [String: Any],
           let fonts = verification["fonts"] as? [String] {
            let missing = fonts.filter { names in
                !names.split(separator: ";").contains { candidate in
                    NSFont(name: String(candidate).trimmingCharacters(in: .whitespacesAndNewlines), size: 12) != nil
                }
            }
            if !missing.isEmpty {
                files[index].status = .warning
                let shown = missing.prefix(6).joined(separator: "、") + (missing.count > 6 ? "等 \(missing.count) 种" : "")
                files[index].message += " 本机缺少原字体：" + shown + "；显示时可能使用替代字体。"
            }
        }
        if status != .error, let path = json["output"] as? String,
           FileManager.default.fileExists(atPath: path) {
            files[index].output = URL(fileURLWithPath: path)
        }
    }

    private func finish(_ result: Result<(Int32, String), Error>) {
        isWorking = false
        let fallback: String
        switch result {
        case .success(let (exitStatus, diagnostic)):
            fallback = exitStatus == 0 ? "转换程序未返回这个文件的结果。" : "转换程序意外停止。请重新尝试这个文件。"
            if exitStatus != 0 && completedCount == 0 {
                alertMessage = diagnostic.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty ? fallback : String(diagnostic.suffix(1800))
            }
        case .failure(let error):
            fallback = "无法启动转换程序：\(error.localizedDescription)"
            alertMessage = fallback
        }
        for index in files.indices where [.queued, .running].contains(files[index].status) {
            files[index].status = .error
            files[index].message = fallback
        }
        completedCount = files.count
        let successes = files.filter { $0.status == .ok }.count
        let warnings = files.filter { $0.status == .warning }.count
        let failures = files.filter { $0.status == .error }.count
        batchMessage = "已结束：\(successes + warnings) 个已转换，\(failures) 个失败。" + (warnings > 0 ? "转换结果请在 Word 中核对排版。" : "") + scanNotice
    }

    func showInFinder(_ url: URL) { NSWorkspace.shared.activateFileViewerSelecting([url]) }

    func openInWord(_ url: URL) {
        guard let app = NSWorkspace.shared.urlForApplication(withBundleIdentifier: "com.microsoft.Word") else {
            alertMessage = "未找到 Microsoft Word。可以在 Finder 中找到转换结果后，选择其他文字处理工具打开。"
            return
        }
        NSWorkspace.shared.open([url], withApplicationAt: app, configuration: NSWorkspace.OpenConfiguration()) { [weak self] _, error in
            if let error { Task { @MainActor in self?.alertMessage = "Word 无法打开这个文件：\(error.localizedDescription)" } }
        }
    }
}

struct FileRow: View {
    let file: ConversionFile
    @ObservedObject var model: ConverterModel

    var body: some View {
        HStack(alignment: .top, spacing: 14) {
            Image(systemName: "doc.text")
                .font(.system(size: 26))
                .foregroundStyle(.secondary)
                .frame(width: 32)
                .padding(.top, 3)
            VStack(alignment: .leading, spacing: 5) {
                Text(file.input.lastPathComponent).font(.body.weight(.medium)).textSelection(.enabled)
                Text(file.input.deletingLastPathComponent().path)
                    .font(.caption).foregroundStyle(.secondary).lineLimit(1).truncationMode(.middle)
                if !file.message.isEmpty {
                    Text(file.message).font(.callout).foregroundStyle(.secondary).textSelection(.enabled)
                }
                if let output = file.output {
                    HStack(spacing: 14) {
                        Button("在 Word 中打开") { model.openInWord(output) }
                        Button("在 Finder 中显示") { model.showInFinder(output) }
                    }
                    .buttonStyle(.link).font(.callout).padding(.top, 3).disabled(model.isBusy)
                }
            }
            Spacer(minLength: 12)
            HStack(spacing: 6) {
                if file.status == .running { ProgressView().controlSize(.small) }
                else { Image(systemName: file.status.symbol).foregroundStyle(file.status.color) }
                Text(file.status.label).foregroundStyle(file.status.color)
            }
            .font(.callout).frame(width: 88, alignment: .leading).padding(.top, 3)
            Button { model.remove(file.id) } label: { Image(systemName: "xmark").foregroundStyle(.secondary) }
                .buttonStyle(.borderless).help("移除这个文件").disabled(model.isBusy).padding(.top, 3)
        }
        .padding(16)
    }
}

struct ConverterView: View {
    @StateObject private var model = ConverterModel()

    var body: some View {
        VStack(alignment: .leading, spacing: 18) {
            HStack(alignment: .top) {
                VStack(alignment: .leading, spacing: 7) {
                    HStack(alignment: .firstTextBaseline, spacing: 10) {
                        Text("MDw2d").font(.largeTitle.weight(.semibold))
                        Text(Bundle.main.object(forInfoDictionaryKey: "CFBundleShortVersionString") as? String ?? "")
                            .font(.callout).foregroundStyle(.secondary)
                    }
                    Text("把旧 Word、WPS 等文档转成 DOCX")
                        .font(.title3).foregroundStyle(.secondary)
                }
                Spacer()
                if let iconURL = Bundle.main.url(forResource: "AppIcon", withExtension: "icns"),
                   let icon = NSImage(contentsOf: iconURL) {
                    Image(nsImage: icon)
                        .resizable().scaledToFit().frame(width: 64, height: 64)
                        .accessibilityHidden(true)
                } else {
                    Image(systemName: "doc.on.doc")
                        .font(.system(size: 40, weight: .light)).foregroundStyle(Color.accentColor)
                        .padding(.top, 5).accessibilityHidden(true)
                }
            }
            HStack {
                Text("文档").font(.headline)
                if !model.files.isEmpty { Text("\(model.files.count) 个").font(.callout).foregroundStyle(.secondary) }
                Spacer()
                Toggle("包含子文件夹", isOn: $model.includeSubfolders)
                    .toggleStyle(.checkbox).font(.callout).disabled(model.isBusy)
                    .accessibilityIdentifier("includeSubfolders")
                Button("转换文件夹…", action: model.chooseSourceFolder).disabled(model.isBusy)
                    .help("选择后立即转换该文件夹的文档，结果保存在各原文件旁。")
                    .accessibilityIdentifier("convertFolder")
                Button("添加文档…", action: model.chooseFiles).disabled(model.isBusy)
                if model.isScanning {
                    Button("取消查找", action: model.cancelScan)
                } else {
                    Button("清空", action: model.clear).disabled(model.isWorking || model.files.isEmpty)
                }
            }
            Group {
                if model.isScanning {
                    VStack(spacing: 14) {
                        ProgressView()
                        Text("正在查找需要转换的文档…").foregroundStyle(.secondary)
                        Button("取消查找", action: model.cancelScan)
                    }
                    .frame(maxWidth: .infinity, maxHeight: .infinity)
                } else if model.files.isEmpty {
                    VStack(spacing: 12) {
                        Image(systemName: "tray.and.arrow.down").font(.system(size: 38)).foregroundStyle(.secondary)
                        Text("拖入文档或一个文件夹").font(.title3.weight(.medium))
                        Text("文件夹会自动转换 WPS、DOC、RTF、ODT 等文档，跳过 DOCX。\n结果保存在各原文件旁，始终保留原件。")
                            .font(.callout).foregroundStyle(.secondary).multilineTextAlignment(.center)
                        Button("选择文档…", action: model.chooseFiles).padding(.top, 2)
                    }
                    .frame(maxWidth: .infinity, maxHeight: .infinity)
                } else {
                    ScrollView {
                        LazyVStack(spacing: 0) {
                            ForEach(model.files) { file in
                                FileRow(file: file, model: model)
                                if file.id != model.files.last?.id { Divider().padding(.horizontal, 16) }
                            }
                        }
                    }
                }
            }
            .frame(minHeight: 230, maxHeight: .infinity)
            .background(RoundedRectangle(cornerRadius: 12).fill(Color(nsColor: .controlBackgroundColor)))
            .overlay(RoundedRectangle(cornerRadius: 12).strokeBorder(model.dropTargeted ? Color.accentColor : Color(nsColor: .separatorColor), style: StrokeStyle(lineWidth: model.dropTargeted ? 2 : 1, dash: model.files.isEmpty ? [6, 4] : [])))
            .onDrop(of: [UTType.fileURL.identifier], isTargeted: $model.dropTargeted, perform: model.acceptDrop)
            HStack(spacing: 12) {
                Image(systemName: "folder").foregroundStyle(.secondary)
                VStack(alignment: .leading, spacing: 3) {
                    Text("保存位置").font(.callout.weight(.medium))
                    Text(model.outputLabel)
                        .font(.callout).foregroundStyle(.secondary).lineLimit(1).truncationMode(.middle)
                }
                Spacer(minLength: 8)
                Button("选择保存位置…", action: model.chooseOutputDirectory).disabled(model.isBusy)
            }
            .padding(14)
            .background(RoundedRectangle(cornerRadius: 10).fill(Color(nsColor: .controlBackgroundColor)))
            HStack(alignment: .center) {
                VStack(alignment: .leading, spacing: 4) {
                    Text(model.batchMessage.isEmpty ? "全程在本机转换，始终保留原文件。" : model.batchMessage)
                        .font(.callout).foregroundStyle(.secondary)
                    Text("转换后可在 Word 中检查排版。复杂对象和字体可能需要调整。")
                        .font(.caption).foregroundStyle(.secondary)
                }
                Spacer(minLength: 20)
                if model.isWorking { Text("\(model.completedCount)/\(model.files.count)").font(.callout.monospacedDigit()).foregroundStyle(.secondary) }
                Button(model.isWorking ? "正在转换…" : "开始转换", action: model.start)
                    .buttonStyle(.borderedProminent).controlSize(.large).disabled(!model.canStart)
                    .keyboardShortcut(.return, modifiers: [])
            }
        }
        .padding(26)
        .frame(minWidth: 720, idealWidth: 840, minHeight: 560, idealHeight: 680)
        .onReceive(NotificationCenter.default.publisher(for: openFilesNotification)) { notification in
            if let urls = notification.object as? [URL] { model.addFiles(urls) }
        }
        .alert("MDw2d", isPresented: Binding(get: { model.alertMessage != nil }, set: { if !$0 { model.alertMessage = nil } })) {
            Button("好", role: .cancel) { model.alertMessage = nil }
        } message: { Text(model.alertMessage ?? "") }
    }
}

@main
struct DocumentConverterApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) private var appDelegate
    var body: some Scene {
        Window("MDw2d", id: "main") { ConverterView() }
            .windowStyle(.titleBar)
            .defaultSize(width: 840, height: 680)
            .commands {
                CommandGroup(replacing: .newItem) {}
            }
    }
}
