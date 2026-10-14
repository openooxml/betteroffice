import AppKit
import UniformTypeIdentifiers
import WebKit

private struct OpenDocument {
  var url: URL?
  var name: String
}

final class DesktopWebView: WKWebView {
  var openFiles: (([URL]) -> Void)?
  var documentExtension = "docx"

  private func documents(_ sender: NSDraggingInfo) -> [URL] {
    (sender.draggingPasteboard.readObjects(
      forClasses: [NSURL.self], options: [.urlReadingFileURLsOnly: true]) as? [URL] ?? [])
      .filter { $0.pathExtension.lowercased() == documentExtension }
  }

  override func draggingEntered(_ sender: NSDraggingInfo) -> NSDragOperation {
    documents(sender).isEmpty ? super.draggingEntered(sender) : .copy
  }

  override func performDragOperation(_ sender: NSDraggingInfo) -> Bool {
    let files = documents(sender)
    if files.isEmpty { return super.performDragOperation(sender) }
    openFiles?(files)
    return true
  }
}

final class DesktopApp: NSObject, NSApplicationDelegate, NSWindowDelegate, WKNavigationDelegate,
  WKUIDelegate, WKScriptMessageHandlerWithReply, WKDownloadDelegate
{
  private var window: NSWindow!
  private var webView: DesktopWebView!
  private var server: DesktopServer!
  private var documents: [String: OpenDocument] = [:]
  private var active: String?
  private var ready = false
  private var pending: [URL] = []
  private var closing = false
  private var printViews: [WKWebView] = []
  #if DEBUG
    private var ranUITest = false
  #endif
  private var downloads: [ObjectIdentifier: (temporary: URL, destination: URL)] = [:]
  private let format: String
  private let appName: String
  private let resources: URL

  init(format: String, resources: URL) {
    self.format = format
    self.resources = resources
    self.appName =
      "BetterOffice " + (["docx": "Docs", "xlsx": "Sheets", "pptx": "Slides"][format] ?? "Docs")
    super.init()
  }

  func applicationDidFinishLaunching(_ notification: Notification) {
    makeMenus()
    let configuration = WKWebViewConfiguration()
    configuration.userContentController.addScriptMessageHandler(
      self, contentWorld: .page, name: "desktop")
    configuration.preferences.javaScriptCanOpenWindowsAutomatically = true
    configuration.userContentController.addUserScript(
      WKUserScript(
        source:
          "window.print = () => window.webkit.messageHandlers.desktop.postMessage({type: 'print'});",
        injectionTime: .atDocumentStart, forMainFrameOnly: true))
    webView = DesktopWebView(frame: .zero, configuration: configuration)
    webView.navigationDelegate = self
    webView.uiDelegate = self
    webView.documentExtension = format
    webView.openFiles = { [weak self] urls in self?.openFiles(urls) }
    webView.registerForDraggedTypes([.fileURL])
    window = NSWindow(
      contentRect: NSRect(x: 0, y: 0, width: 1180, height: 860),
      styleMask: [.titled, .closable, .miniaturizable, .resizable], backing: .buffered, defer: false
    )
    window.title = appName
    window.minSize = NSSize(width: 620, height: 560)
    window.contentView = webView
    window.delegate = self
    window.center()
    window.makeKeyAndOrderFront(nil)
    NSApp.activate(ignoringOtherApps: true)
    do {
      server = try DesktopServer(resources: resources)
      server.start { [weak self] result in
        switch result {
        case .success(let url): self?.webView.load(URLRequest(url: url))
        case .failure(let error): self?.showError(error)
        }
      }
    } catch { showError(error) }
  }

  func application(_ application: NSApplication, open urls: [URL]) { openFiles(urls) }
  func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool { true }
  func applicationShouldHandleReopen(_ sender: NSApplication, hasVisibleWindows: Bool) -> Bool {
    window?.makeKeyAndOrderFront(nil)
    return true
  }

  func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply {
    if closing || !ready { return .terminateNow }
    requestClose { allowed in sender.reply(toApplicationShouldTerminate: allowed) }
    return .terminateLater
  }

  func windowShouldClose(_ sender: NSWindow) -> Bool {
    if closing || !ready { return true }
    requestClose { [weak self] allowed in
      if allowed {
        self?.closing = true
        sender.close()
      }
    }
    return false
  }

  private func requestClose(_ completion: @escaping (Bool) -> Void) {
    webView.callAsyncJavaScript(
      "return await window.desktop.close()", arguments: [:], in: nil, in: .page
    ) { [weak self] result in
      switch result {
      case .success(let allowed): completion(allowed as? Bool ?? false)
      case .failure(let error):
        self?.showError(error)
        completion(false)
      }
    }
  }

  private func openFiles(_ urls: [URL]) {
    guard let url = urls.first else { return }
    guard ready else {
      pending = [url]
      return
    }
    do {
      let file = try loadFile(url)
      webView.callAsyncJavaScript(
        "await window.desktop.open(file)", arguments: ["file": file], in: nil, in: .page
      ) { [weak self] result in
        if case .failure(let error) = result { self?.showError(error) }
      }
    } catch { showError(error) }
  }

  private func loadFile(_ url: URL) throws -> [String: Any] {
    guard url.isFileURL, url.pathExtension.lowercased() == format else {
      throw message("Choose a .\(format) file for \(appName).")
    }
    let properties = try url.resourceValues(forKeys: [.isRegularFileKey])
    guard properties.isRegularFile == true else {
      throw message("Choose a regular .\(format) file.")
    }
    return register(bytes: try Data(contentsOf: url), name: url.lastPathComponent, url: url)
  }

  private func register(bytes: Data, name: String, url: URL?) -> [String: Any] {
    let id = UUID().uuidString
    documents = documents.filter { $0.key == active }
    documents[id] = OpenDocument(url: url, name: name)
    server.retainDocuments(Set(documents.keys))
    return ["id": id, "name": name, "url": server.put(bytes, id: id).absoluteString]
  }

  private var recentURLs: [URL] {
    (UserDefaults.standard.stringArray(forKey: "recent-\(format)") ?? []).compactMap(
      URL.init(string:)
    ).filter { $0.isFileURL && $0.pathExtension.lowercased() == format }
  }

  private func remember(_ url: URL) {
    var recent = recentURLs.filter { $0 != url }
    recent.insert(url, at: 0)
    UserDefaults.standard.set(
      Array(recent.prefix(5)).map(\.absoluteString), forKey: "recent-\(format)")
    NSDocumentController.shared.noteNewRecentDocumentURL(url)
  }

  private func configuration() -> [String: Any] {
    [
      "format": format, "name": appName,
      "recent": recentURLs.map {
        [
          "id": $0.absoluteString, "name": $0.lastPathComponent,
          "folder": ($0.deletingLastPathComponent().path as NSString).abbreviatingWithTildeInPath,
        ]
      },
    ]
  }

  func userContentController(
    _ userContentController: WKUserContentController, didReceive message: WKScriptMessage,
    replyHandler: @escaping (Any?, String?) -> Void
  ) {
    guard message.frameInfo.isMainFrame,
      message.frameInfo.securityOrigin.protocol == "http",
      message.frameInfo.securityOrigin.host == "127.0.0.1",
      message.frameInfo.securityOrigin.port == server.origin?.port,
      let body = message.body as? [String: Any], let type = body["type"] as? String
    else {
      replyHandler(nil, "Untrusted desktop request")
      return
    }
    do {
      switch type {
      case "configuration": replyHandler(configuration(), nil)
      case "ready":
        ready = true
        replyHandler(nil, nil)
        if !pending.isEmpty {
          let files = pending
          pending = []
          openFiles(files)
        }
        #if DEBUG
          if !ranUITest, let index = CommandLine.arguments.firstIndex(of: "--ui-test"),
            index + 2 < CommandLine.arguments.count
          {
            ranUITest = true
            runUITest(
              script: CommandLine.arguments[index + 1], output: CommandLine.arguments[index + 2])
          }
        #endif
      case "open":
        let panel = NSOpenPanel()
        panel.allowedContentTypes = [UTType(filenameExtension: format)!]
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        panel.beginSheetModal(for: window) { [weak self] response in
          guard let self, response == .OK, let url = panel.url else {
            replyHandler(nil, nil)
            return
          }
          do { replyHandler(try self.loadFile(url), nil) } catch {
            replyHandler(nil, error.localizedDescription)
          }
        }
      case "openRecent":
        guard let id = body["id"] as? String,
          let url = recentURLs.first(where: { $0.absoluteString == id })
        else { throw self.message("Recent file is unavailable.") }
        replyHandler(try loadFile(url), nil)
      case "import":
        guard let name = body["name"] as? String,
          (name as NSString).pathExtension.lowercased() == format,
          let encoded = body["bytes"] as? String,
          let bytes = Data(base64Encoded: encoded)
        else { throw self.message("Invalid file.") }
        replyHandler(
          register(bytes: bytes, name: URL(fileURLWithPath: name).lastPathComponent, url: nil), nil)
      case "activate":
        guard let id = body["id"] as? String, let document = documents[id] else {
          throw self.message("File is no longer available.")
        }
        active = id
        window.title = document.name + " — " + appName
        window.representedURL = document.url
        window.isDocumentEdited = false
        if let url = document.url { remember(url) }
        documents = [id: document]
        server.retainDocuments([id])
        replyHandler(nil, nil)
      case "deactivate":
        active = nil
        documents = [:]
        server.retainDocuments([])
        window.title = appName
        window.representedURL = nil
        window.isDocumentEdited = false
        replyHandler(nil, nil)
      case "state":
        if body["id"] as? String == active {
          window.isDocumentEdited = body["dirty"] as? Bool ?? true
        }
        replyHandler(nil, nil)
      case "confirmClose":
        let alert = NSAlert()
        alert.messageText = "Save changes to \(body["name"] as? String ?? "this file")?"
        alert.informativeText = "Your changes will be lost if you don’t save them."
        alert.addButton(withTitle: "Save")
        alert.addButton(withTitle: "Cancel")
        alert.addButton(withTitle: "Don’t Save")
        alert.beginSheetModal(for: window) { response in
          replyHandler(
            response == .alertFirstButtonReturn
              ? "save" : response == .alertThirdButtonReturn ? "discard" : "cancel", nil)
        }
      case "print":
        if let view = message.webView {
          view.printOperation(with: NSPrintInfo.shared).runModal(
            for: window, delegate: nil, didRun: nil, contextInfo: nil)
        }
        replyHandler(nil, nil)
      case "save": try save(body, reply: replyHandler)
      default: throw self.message("Unknown desktop action.")
      }
    } catch { replyHandler(nil, error.localizedDescription) }
  }

  private func save(_ body: [String: Any], reply: @escaping (Any?, String?) -> Void) throws {
    guard let id = body["id"] as? String, id == active, let document = documents[id],
      let encoded = body["bytes"] as? String,
      let bytes = Data(base64Encoded: encoded)
    else { throw message("The file is no longer open.") }
    let write: (URL) -> Void = { [weak self] url in
      guard let self else {
        reply(nil, "The app closed.")
        return
      }
      do {
        try bytes.write(to: url, options: .atomic)
        self.documents[id] = OpenDocument(url: url, name: url.lastPathComponent)
        self.window.title = url.lastPathComponent + " — " + self.appName
        self.window.representedURL = url
        self.remember(url)
        reply(["name": url.lastPathComponent], nil)
      } catch { reply(nil, error.localizedDescription) }
    }
    if body["saveAs"] as? Bool != true, let url = document.url {
      write(url)
      return
    }
    let panel = NSSavePanel()
    panel.allowedContentTypes = [UTType(filenameExtension: format)!]
    panel.nameFieldStringValue = document.name
    panel.canCreateDirectories = true
    panel.beginSheetModal(for: window) { response in
      guard response == .OK, let url = panel.url else {
        reply(nil, nil)
        return
      }
      write(url)
    }
  }

  private func makeMenus() {
    let main = NSMenu()
    NSApp.mainMenu = main
    let appMenu = NSMenu(title: appName)
    appMenu.addItem(
      withTitle: "About \(appName)",
      action: #selector(NSApplication.orderFrontStandardAboutPanel(_:)), keyEquivalent: "")
    appMenu.addItem(.separator())
    appMenu.addItem(
      withTitle: "Hide \(appName)", action: #selector(NSApplication.hide(_:)), keyEquivalent: "h")
    appMenu.addItem(.separator())
    appMenu.addItem(
      withTitle: "Quit \(appName)", action: #selector(NSApplication.terminate(_:)),
      keyEquivalent: "q")
    addMenu(appMenu, to: main)
    let file = NSMenu(title: "File")
    addCommand("Open…", key: "o", command: "open", to: file)
    addCommand("Close File", key: "w", command: "close", to: file)
    file.addItem(.separator())
    addCommand("Save", key: "s", command: "save", to: file)
    addCommand("Save As…", key: "S", command: "saveAs", to: file)
    file.addItem(.separator())
    addCommand("Print…", key: "p", command: "print", to: file)
    addMenu(file, to: main)
    let edit = NSMenu(title: "Edit")
    addCommand("Undo", key: "z", command: "undo", to: edit)
    addCommand("Redo", key: "Z", command: "redo", to: edit)
    edit.addItem(.separator())
    for (title, key, selector) in [
      ("Cut", "x", "cut:"), ("Copy", "c", "copy:"), ("Paste", "v", "paste:"),
      ("Select All", "a", "selectAll:"),
    ] {
      edit.addItem(withTitle: title, action: Selector(selector), keyEquivalent: key)
    }
    edit.addItem(.separator())
    addCommand("Find…", key: "f", command: "find", to: edit)
    addMenu(edit, to: main)
    let windowMenu = NSMenu(title: "Window")
    windowMenu.addItem(
      withTitle: "Minimize", action: #selector(NSWindow.performMiniaturize(_:)), keyEquivalent: "m")
    windowMenu.addItem(
      withTitle: "Zoom", action: #selector(NSWindow.performZoom(_:)), keyEquivalent: "")
    addMenu(windowMenu, to: main)
    NSApp.windowsMenu = windowMenu
  }

  private func addMenu(_ menu: NSMenu, to main: NSMenu) {
    let item = NSMenuItem(title: menu.title, action: nil, keyEquivalent: "")
    item.submenu = menu
    main.addItem(item)
  }

  private func addCommand(
    _ title: String, key: String, command commandName: String, to menu: NSMenu
  ) {
    let item = NSMenuItem(
      title: title, action: #selector(command(_:)), keyEquivalent: key.lowercased())
    item.keyEquivalentModifierMask = key == key.lowercased() ? [.command] : [.command, .shift]
    item.representedObject = commandName
    item.target = self
    menu.addItem(item)
  }

  @objc private func command(_ sender: NSMenuItem) {
    guard ready, let command = sender.representedObject as? String else { return }
    webView.callAsyncJavaScript(
      "window.desktop.command(command)", arguments: ["command": command], in: nil, in: .page
    ) { [weak self] result in
      if case .failure(let error) = result { self?.showError(error) }
    }
  }

  func webView(
    _ webView: WKWebView, decidePolicyFor navigationAction: WKNavigationAction,
    decisionHandler: @escaping (WKNavigationActionPolicy) -> Void
  ) {
    guard let url = navigationAction.request.url else {
      decisionHandler(.cancel)
      return
    }
    if navigationAction.shouldPerformDownload && url.scheme == "blob" {
      decisionHandler(.download)
      return
    }
    if url.scheme == "about"
      || (url.scheme == "http" && url.host == "127.0.0.1" && url.port == server.origin?.port
        && url.path.hasPrefix("/\(server.token)/"))
    {
      decisionHandler(.allow)
    } else {
      decisionHandler(.cancel)
      if ["https", "http", "mailto"].contains(url.scheme ?? "")
        && navigationAction.navigationType == .linkActivated
      {
        NSWorkspace.shared.open(url)
      }
    }
  }

  func webView(
    _ webView: WKWebView, navigationAction: WKNavigationAction, didBecome download: WKDownload
  ) { download.delegate = self }
  func webView(
    _ webView: WKWebView, navigationResponse: WKNavigationResponse, didBecome download: WKDownload
  ) { download.delegate = self }

  func download(
    _ download: WKDownload, decideDestinationUsing response: URLResponse, suggestedFilename: String,
    completionHandler: @escaping (URL?) -> Void
  ) {
    let panel = NSSavePanel()
    panel.nameFieldStringValue = URL(fileURLWithPath: suggestedFilename).lastPathComponent
    panel.beginSheetModal(for: window) { [weak self] response in
      guard response == .OK, let url = panel.url else {
        completionHandler(nil)
        return
      }
      let temporary = FileManager.default.temporaryDirectory.appendingPathComponent(
        UUID().uuidString)
      self?.downloads[ObjectIdentifier(download)] = (temporary, url)
      completionHandler(temporary)
    }
  }

  func downloadDidFinish(_ download: WKDownload) {
    guard let paths = downloads.removeValue(forKey: ObjectIdentifier(download)) else { return }
    defer { try? FileManager.default.removeItem(at: paths.temporary) }
    do {
      try Data(contentsOf: paths.temporary).write(to: paths.destination, options: .atomic)
    } catch { showError(error) }
  }

  func download(_ download: WKDownload, didFailWithError error: Error, resumeData: Data?) {
    if let paths = downloads.removeValue(forKey: ObjectIdentifier(download)) {
      try? FileManager.default.removeItem(at: paths.temporary)
    }
    showError(error)
  }

  func webView(
    _ webView: WKWebView, createWebViewWith configuration: WKWebViewConfiguration,
    for navigationAction: WKNavigationAction, windowFeatures: WKWindowFeatures
  ) -> WKWebView? {
    guard navigationAction.request.url?.scheme == "about" else { return nil }
    let view = WKWebView(
      frame: NSRect(x: 0, y: 0, width: 800, height: 1000), configuration: configuration)
    view.uiDelegate = self
    view.navigationDelegate = self
    printViews.append(view)
    return view
  }

  func webViewDidClose(_ webView: WKWebView) { printViews.removeAll { $0 === webView } }

  func webView(
    _ webView: WKWebView, runJavaScriptTextInputPanelWithPrompt prompt: String,
    defaultText: String?, initiatedByFrame frame: WKFrameInfo,
    completionHandler: @escaping (String?) -> Void
  ) {
    let alert = NSAlert()
    alert.messageText = prompt
    let field = NSTextField(string: defaultText ?? "")
    field.frame = NSRect(x: 0, y: 0, width: 320, height: 24)
    alert.accessoryView = field
    alert.addButton(withTitle: "OK")
    alert.addButton(withTitle: "Cancel")
    alert.beginSheetModal(for: window) {
      completionHandler($0 == .alertFirstButtonReturn ? field.stringValue : nil)
    }
  }

  func webView(
    _ webView: WKWebView, runOpenPanelWith parameters: WKOpenPanelParameters,
    initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping ([URL]?) -> Void
  ) {
    let panel = NSOpenPanel()
    panel.allowsMultipleSelection = parameters.allowsMultipleSelection
    panel.canChooseDirectories = parameters.allowsDirectories
    panel.beginSheetModal(for: window) { response in
      completionHandler(response == .OK ? panel.urls : nil)
    }
  }

  func webView(
    _ webView: WKWebView, runJavaScriptAlertPanelWithMessage text: String,
    initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping () -> Void
  ) {
    let alert = NSAlert()
    alert.messageText = text
    alert.beginSheetModal(for: window) { _ in completionHandler() }
  }

  func webView(
    _ webView: WKWebView, runJavaScriptConfirmPanelWithMessage text: String,
    initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping (Bool) -> Void
  ) {
    let alert = NSAlert()
    alert.messageText = text
    alert.addButton(withTitle: "OK")
    alert.addButton(withTitle: "Cancel")
    alert.beginSheetModal(for: window) { completionHandler($0 == .alertFirstButtonReturn) }
  }

  #if DEBUG
    private func runUITest(script: String, output: String) {
      do {
        let source = try String(contentsOfFile: script, encoding: .utf8)
        webView.callAsyncJavaScript(source, arguments: [:], in: nil, in: .page) {
          [weak self] result in
          guard let self else { return }
          let resultURL = URL(fileURLWithPath: output + ".json")
          switch result {
          case .failure(let error):
            try? JSONSerialization.data(withJSONObject: ["error": error.localizedDescription])
              .write(to: resultURL)
            self.closing = true
            NSApp.terminate(nil)
          case .success(let value):
            try? JSONSerialization.data(withJSONObject: ["result": value]).write(to: resultURL)
            self.webView.takeSnapshot(with: nil) { image, error in
              if let data = image?.tiffRepresentation, let bitmap = NSBitmapImageRep(data: data),
                let png = bitmap.representation(using: .png, properties: [:])
              {
                try? png.write(to: URL(fileURLWithPath: output + ".png"))
              }
              self.closing = true
              NSApp.terminate(nil)
            }
          }
        }
      } catch { showError(error) }
    }
  #endif

  private func message(_ text: String) -> Error {
    NSError(domain: "BetterOffice", code: 1, userInfo: [NSLocalizedDescriptionKey: text])
  }
  private func showError(_ error: Error) {
    let alert = NSAlert(error: error)
    alert.beginSheetModal(for: window)
  }
}

@main
struct DesktopMain {
  static func main() {
    let arguments = CommandLine.arguments
    let resourcePath = arguments.firstIndex(of: "--resources").flatMap {
      $0 + 1 < arguments.count ? arguments[$0 + 1] : nil
    }
    let format =
      arguments.firstIndex(of: "--format").flatMap {
        $0 + 1 < arguments.count ? arguments[$0 + 1] : nil
      }
      ?? Bundle.main.object(forInfoDictionaryKey: "BetterOfficeFormat") as? String ?? "docx"
    let resources =
      resourcePath.map { URL(fileURLWithPath: $0) }
      ?? Bundle.main.resourceURL!.appendingPathComponent("Editor")
    let app = NSApplication.shared
    let delegate = DesktopApp(format: format, resources: resources)
    app.delegate = delegate
    if let index = arguments.firstIndex(of: "--document"), index + 1 < arguments.count {
      delegate.application(app, open: [URL(fileURLWithPath: arguments[index + 1])])
    }
    app.setActivationPolicy(.regular)
    withExtendedLifetime(delegate) { app.run() }
  }
}
