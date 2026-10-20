import Foundation

private enum TestFailure: Error, CustomStringConvertible {
    case failed(String)

    var description: String {
        switch self { case .failed(let message): return message }
    }
}

private func expect(_ condition: @autoclosure () -> Bool, _ message: String) throws {
    if !condition() { throw TestFailure.failed(message) }
}

@main
struct FolderScannerTests {
    static func main() async throws {
        let manager = FileManager.default
        let fixture = manager.temporaryDirectory.resolvingSymlinksInPath().appendingPathComponent("MDw2d folder scan \(UUID().uuidString)", isDirectory: true)
        try manager.createDirectory(at: fixture, withIntermediateDirectories: true)
        defer { try? manager.removeItem(at: fixture) }

        func directory(_ path: String) throws -> URL {
            let url = fixture.appendingPathComponent(path, isDirectory: true)
            try manager.createDirectory(at: url, withIntermediateDirectories: true)
            return url
        }
        @discardableResult
        func file(_ path: String) throws -> URL {
            let url = fixture.appendingPathComponent(path)
            try manager.createDirectory(at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try Data("source remains untouched: \(path)".utf8).write(to: url)
            return url
        }
        func names(_ result: FolderScanResult, relativeTo root: URL) -> [String] {
            let prefixLength = root.path.count + 1
            return result.files.map { String($0.path.dropFirst(prefixLength)) }
        }
        func expectRejected(_ url: URL, _ message: String) throws {
            do {
                _ = try FolderScanner.scan(folder: url)
            } catch is FolderScanError {
                return
            }
            throw TestFailure.failed(message)
        }

        let root = try directory("资料目录 中文")
        let original = try file("资料目录 中文/报告 10.DOC")
        let originalData = try Data(contentsOf: original)
        try file("资料目录 中文/报告 2.wps")
        try file("资料目录 中文/报告 1.docx")
        try file("资料目录 中文/已经转换.DOCX")
        try file("资料目录 中文/notes.txt")
        try file("资料目录 中文/表格.xls")
        try file("资料目录 中文/扫描.pdf")
        try file("资料目录 中文/图片.png")
        try file("资料目录 中文/~$未完成.doc")
        try file("资料目录 中文/.隐藏.doc")
        try file("资料目录 中文/.秘密/不该扫描.doc")
        try file("资料目录 中文/工具.APP/内部.doc")
        try file("资料目录 中文/工具.bundle/内部.doc")
        try file("资料目录 中文/系统.framework/内部.doc")
        try file("资料目录 中文/子目录/报告 2.wps")
        try file("资料目录 中文/子目录/模板.DOT")
        try file("资料目录 中文/子目录/模板.dotx")
        try file("资料目录 中文/子目录/模板.dotm")
        try file("资料目录 中文/子目录/模板.wpt")
        try file("资料目录 中文/子目录/模板.ott")
        try file("资料目录 中文/子目录/宏.docm")
        try file("资料目录 中文/子目录/富文本.RTF")
        try file("资料目录 中文/子目录/开放文档.ODT")
        try file("资料目录 中文/子目录/已转换.docx")
        try file("目录外/不要转换.doc")
        try manager.createSymbolicLink(at: root.appendingPathComponent("连接.doc"), withDestinationURL: fixture.appendingPathComponent("目录外/不要转换.doc"))
        try manager.createSymbolicLink(at: root.appendingPathComponent("外部目录"), withDestinationURL: fixture.appendingPathComponent("目录外"))
        try manager.createSymbolicLink(at: root.appendingPathComponent("子目录/循环"), withDestinationURL: root)
        try manager.createSymbolicLink(at: root.appendingPathComponent("坏链接.doc"), withDestinationURL: fixture.appendingPathComponent("不存在.doc"))

        let recursive = try FolderScanner.scan(folder: root)
        try expect(recursive.files.count == 11, "Recursive scan should include all 11 candidate documents, including same names in distinct folders.")
        try expect(recursive.skippedDOCX == 3, "DOCX must be skipped case-insensitively at every included level.")
        try expect(recursive.unreadablePaths.isEmpty, "Broken links and intentionally excluded entries are not scan errors.")
        let relativeNames = names(recursive, relativeTo: root)
        try expect(relativeNames.contains("子目录/报告 2.wps"), "A same-name file in a subfolder must survive deduplication.")
        try expect(relativeNames.contains("报告 2.wps"), "The root file must survive deduplication.")
        try expect(relativeNames.firstIndex(of: "报告 2.wps")! < relativeNames.firstIndex(of: "报告 10.DOC")!, "Document names must use natural numeric order.")
        let repeated = try FolderScanner.scan(folder: root)
        try expect(recursive.files == repeated.files, "Repeated scans must return a stable order.")
        let ancestorLink = fixture.appendingPathComponent("入口连接", isDirectory: true)
        try manager.createSymbolicLink(at: ancestorLink, withDestinationURL: fixture)
        let ancestorLinkedRoot = ancestorLink.appendingPathComponent("资料目录 中文", isDirectory: true)
        let ancestorLinked = try FolderScanner.scan(folder: ancestorLinkedRoot)
        try expect(ancestorLinked.files == recursive.files, "A symbolic-link ancestor must return the same canonical input URLs as the conversion engine.")
        try expect(ancestorLinked.files.allSatisfy { $0 == $0.resolvingSymlinksInPath().standardizedFileURL }, "Every result URL must be canonical.")
        try expect(Set(recursive.files.map(\.path)).count == recursive.files.count, "Each path must appear once.")
        let scannedOriginalData = try Data(contentsOf: original)
        try expect(scannedOriginalData == originalData, "Scanning must leave source bytes unchanged.")

        let flat = try FolderScanner.scan(folder: root, includeSubfolders: false)
        try expect(names(flat, relativeTo: root) == ["报告 2.wps", "报告 10.DOC"], "Nonrecursive mode must only include files directly inside the selected folder.")
        try expect(flat.skippedDOCX == 2, "Nonrecursive mode must only count DOCX in the selected folder.")

        let empty = try FolderScanner.scan(folder: directory("空目录"))
        try expect(empty.files.isEmpty && empty.skippedDOCX == 0 && empty.unreadablePaths.isEmpty, "An empty folder is a valid, empty scan.")
        try expectRejected(original, "A file cannot be scanned as a directory.")
        try expectRejected(URL(string: "https://example.com/folder")!, "A remote URL cannot be scanned as a local directory.")
        try expectRejected(fixture.appendingPathComponent("不存在"), "A missing root must fail visibly.")
        try expectRejected(root.appendingPathComponent("工具.APP"), "A selected application package cannot be scanned internally.")
        try expectRejected(root.appendingPathComponent("外部目录"), "A selected symbolic link cannot escape the folder boundary.")

        let unreadableDirectory = try directory("受限资料/封闭目录")
        let unreadableFile = try file("受限资料/不能读.doc")
        try file("受限资料/封闭目录/内部.doc")
        try manager.setAttributes([.posixPermissions: 0], ofItemAtPath: unreadableDirectory.path)
        try manager.setAttributes([.posixPermissions: 0], ofItemAtPath: unreadableFile.path)
        defer {
            try? manager.setAttributes([.posixPermissions: 0o700], ofItemAtPath: unreadableDirectory.path)
            try? manager.setAttributes([.posixPermissions: 0o600], ofItemAtPath: unreadableFile.path)
        }
        if !manager.isReadableFile(atPath: unreadableFile.path) {
            let restricted = try FolderScanner.scan(folder: fixture.appendingPathComponent("受限资料"))
            try expect(Set(restricted.unreadablePaths) == Set([unreadableDirectory.path, unreadableFile.path]), "Unreadable files and subfolders must both be reported.")
            try expect(restricted.files.isEmpty, "Unreadable candidates must not be queued for conversion.")
            try expectRejected(unreadableDirectory, "An unreadable root must fail instead of claiming an empty success.")
        } else {
            print("Permission case skipped because this process can read mode-000 files.")
        }

        let cancelledScan = Task.detached {
            withUnsafeCurrentTask { $0?.cancel() }
            return try FolderScanner.scan(folder: root)
        }
        do {
            _ = try await cancelledScan.value
            throw TestFailure.failed("A canceled scan must stop before enumeration.")
        } catch is CancellationError {
            // Expected: task cancellation is separate from a filesystem error.
        }
        print("FolderScanner tests passed (recursive, flat, packages, links, permissions, cancellation, source preservation).")
    }
}
