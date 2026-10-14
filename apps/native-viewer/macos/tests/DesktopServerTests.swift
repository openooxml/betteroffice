import Foundation

@main
struct DesktopServerTests {
  static func main() throws {
    let root = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
    try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
    defer { try? FileManager.default.removeItem(at: root) }
    try Data("<h1>Editor</h1>".utf8).write(to: root.appendingPathComponent("index.html"))
    try Data([0, 97, 115, 109]).write(to: root.appendingPathComponent("engine.wasm"))
    try FileManager.default.createSymbolicLink(
      at: root.appendingPathComponent("outside"),
      withDestinationURL: root.deletingLastPathComponent())
    let server = try DesktopServer(resources: root)
    var address: URL?
    var startupError: Error?
    server.start { result in
      switch result {
      case .success(let url): address = url
      case .failure(let error): startupError = error
      }
    }
    let deadline = Date().addingTimeInterval(10)
    while address == nil && startupError == nil && Date() < deadline {
      RunLoop.current.run(until: Date().addingTimeInterval(0.01))
    }
    if let error = startupError { throw error }
    guard let address else { fatalError("HTTP server did not start") }
    let host = "127.0.0.1:\(address.port!)"
    let base = "/\(server.token)/"
    precondition(
      server.response(method: "GET", target: base + "index.html", host: host).status == 200)
    precondition(
      server.response(method: "GET", target: base + "engine.wasm", host: host).contentType
        == "application/wasm")
    precondition(server.response(method: "GET", target: "/index.html", host: host).status == 404)
    precondition(server.response(method: "GET", target: "?", host: host).status == 404)
    precondition(
      server.response(method: "GET", target: base + "index.html", host: "attacker.example").status
        == 403)
    precondition(
      server.response(method: "POST", target: base + "index.html", host: host).status == 405)
    for path in ["../secret", "%2e%2e/secret", "outside/secret", "missing"] {
      precondition(server.response(method: "GET", target: base + path, host: host).status != 200)
    }
    let bytes = Data([80, 75, 3, 4, 0, 255])
    let document = server.put(bytes, id: "chosen")
    precondition(server.response(method: "GET", target: document.path, host: host).body == bytes)
    server.retainDocuments([])
    precondition(server.response(method: "GET", target: document.path, host: host).status == 404)
    let semaphore = DispatchSemaphore(value: 0)
    var httpSucceeded = false
    URLSession.shared.dataTask(with: address) { data, response, error in
      httpSucceeded =
        error == nil && (response as? HTTPURLResponse)?.statusCode == 200
        && data == Data("<h1>Editor</h1>".utf8)
      semaphore.signal()
    }.resume()
    precondition(semaphore.wait(timeout: .now() + 10) == .success && httpSucceeded)
    print("Desktop server: all checks passed")
  }
}
