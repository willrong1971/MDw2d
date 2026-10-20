import Foundation

struct FolderScanResult: Sendable {
    let files: [URL]
    let skippedDOCX: Int
    let unreadablePaths: [String]
}

enum FolderScanError: LocalizedError {
    case invalidFolder(String)
    case symbolicLink(String)
    case package(String)
    case unreadableFolder(String, String)

    var errorDescription: String? {
        switch self {
        case .invalidFolder(let path):
            return "所选位置不是可扫描的文件夹：\(path)"
        case .symbolicLink(let path):
            return "请选择实际文件夹，不能扫描符号链接：\(path)"
        case .package(let path):
            return "不能扫描应用或其他文件包的内部内容：\(path)"
        case .unreadableFolder(let path, let reason):
            return "无法读取所选文件夹：\(path)\n\(reason)"
        }
    }
}

/// Finds candidate word-processing documents without reading or changing them.
/// The conversion engine still identifies the actual format and validates it.
enum FolderScanner {
    static let candidateExtensions: Set<String> = [
        "wps", "doc", "rtf", "odt", "dot", "wpt", "docm", "dotx", "dotm", "ott"
    ]

    // Foundation's package flag can be absent for unregistered applications or
    // bundles, so these recognizable bundles are also treated as single items.
    private static let packageExtensions: Set<String> = [
        "app", "appex", "bundle", "framework", "plugin", "kext", "rtfd",
        "xcodeproj", "xcworkspace", "playground", "wdgt", "pages", "numbers", "key"
    ]
    private static let resourceKeys: Set<URLResourceKey> = [
        .isDirectoryKey, .isRegularFileKey, .isSymbolicLinkKey, .isPackageKey, .isHiddenKey
    ]

    static func scan(folder: URL, includeSubfolders: Bool = true) throws -> FolderScanResult {
        try Task.checkCancellation()
        guard folder.isFileURL else { throw FolderScanError.invalidFolder(folder.absoluteString) }

        let selectedRoot = folder.standardizedFileURL
        let manager = FileManager.default
        let rootValues: URLResourceValues
        do {
            rootValues = try selectedRoot.resourceValues(forKeys: resourceKeys)
        } catch {
            throw FolderScanError.unreadableFolder(selectedRoot.path, error.localizedDescription)
        }
        guard rootValues.isSymbolicLink != true else { throw FolderScanError.symbolicLink(selectedRoot.path) }
        guard rootValues.isDirectory == true else { throw FolderScanError.invalidFolder(selectedRoot.path) }
        guard !isPackage(selectedRoot, values: rootValues) else { throw FolderScanError.package(selectedRoot.path) }

        let root = selectedRoot.resolvingSymlinksInPath().standardizedFileURL
        let canonicalRoot = root.path
        var pending = [root]
        var candidates: [URL] = []
        var seenPaths = Set<String>()
        var unreadablePaths = Set<String>()
        var skippedDOCX = 0

        while let directory = pending.popLast() {
            try Task.checkCancellation()
            // Recheck before descent in case a directory changed since listing.
            // Use a fresh URL so prefetched resource values cannot hide a change.
            let descentValues: URLResourceValues
            do {
                descentValues = try URL(fileURLWithPath: directory.path, isDirectory: true)
                    .resourceValues(forKeys: resourceKeys)
            } catch {
                if directory == root {
                    throw FolderScanError.unreadableFolder(root.path, error.localizedDescription)
                }
                unreadablePaths.insert(directory.path)
                continue
            }
            if descentValues.isSymbolicLink == true {
                if directory == root { throw FolderScanError.symbolicLink(root.path) }
                continue
            }
            guard descentValues.isDirectory == true else {
                if directory == root { throw FolderScanError.invalidFolder(root.path) }
                unreadablePaths.insert(directory.path)
                continue
            }
            if isPackage(directory, values: descentValues) {
                if directory == root { throw FolderScanError.package(root.path) }
                continue
            }
            guard isInside(directory, canonicalRoot: canonicalRoot) else { continue }
            let children: [URL]
            do {
                children = try manager.contentsOfDirectory(
                    at: directory,
                    includingPropertiesForKeys: Array(resourceKeys),
                    options: []
                )
            } catch {
                if directory == root {
                    throw FolderScanError.unreadableFolder(root.path, error.localizedDescription)
                }
                unreadablePaths.insert(directory.path)
                continue
            }

            for item in children {
                try Task.checkCancellation()
                let url = item.standardizedFileURL
                let name = url.lastPathComponent
                guard !name.hasPrefix("."), !name.hasPrefix("~$") else { continue }
                let values: URLResourceValues
                do {
                    values = try url.resourceValues(forKeys: resourceKeys)
                } catch {
                    unreadablePaths.insert(url.path)
                    continue
                }
                guard values.isHidden != true, values.isSymbolicLink != true else { continue }
                guard isInside(url, canonicalRoot: canonicalRoot) else { continue }
                if values.isDirectory == true {
                    if includeSubfolders && !isPackage(url, values: values) {
                        pending.append(url)
                    }
                    continue
                }
                guard values.isRegularFile == true else { continue }
                let extensionName = url.pathExtension.lowercased()
                if extensionName == "docx" {
                    skippedDOCX += 1
                    continue
                }
                guard candidateExtensions.contains(extensionName), seenPaths.insert(url.path).inserted else {
                    continue
                }
                guard manager.isReadableFile(atPath: url.path) else {
                    unreadablePaths.insert(url.path)
                    continue
                }
                candidates.append(url)
            }
        }

        try Task.checkCancellation()
        candidates.sort { naturallyPrecedes(relativePath($0, root: root), relativePath($1, root: root)) }
        return FolderScanResult(
            files: candidates,
            skippedDOCX: skippedDOCX,
            unreadablePaths: unreadablePaths.sorted(by: naturallyPrecedes)
        )
    }

    private static func isPackage(_ url: URL, values: URLResourceValues) -> Bool {
        values.isPackage == true || packageExtensions.contains(url.pathExtension.lowercased())
    }

    private static func isInside(_ url: URL, canonicalRoot: String) -> Bool {
        let path = url.resolvingSymlinksInPath().standardizedFileURL.path
        return path == canonicalRoot || path.hasPrefix(canonicalRoot == "/" ? "/" : canonicalRoot + "/")
    }

    private static func relativePath(_ url: URL, root: URL) -> String {
        let prefix = root.path == "/" ? "/" : root.path + "/"
        return String(url.path.dropFirst(prefix.count))
    }

    private static func naturallyPrecedes(_ left: String, _ right: String) -> Bool {
        let comparison = left.compare(right, options: [.numeric, .caseInsensitive], locale: Locale(identifier: "en_US_POSIX"))
        return comparison == .orderedSame ? left < right : comparison == .orderedAscending
    }
}
