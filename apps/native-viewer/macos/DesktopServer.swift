import Foundation
import Network

struct DesktopResponse {
  let status: Int
  let contentType: String
  let body: Data
}

final class DesktopServer {
  let token = UUID().uuidString
  private let resources: URL
  private let queue = DispatchQueue(label: "dev.betteroffice.desktop.http")
  private let listener: NWListener
  private let lock = NSLock()
  private var documents: [String: Data] = [:]
  private(set) var origin: URL?

  init(resources: URL) throws {
    self.resources = resources.standardizedFileURL.resolvingSymlinksInPath()
    let parameters = NWParameters.tcp
    parameters.requiredLocalEndpoint = .hostPort(host: "127.0.0.1", port: .any)
    listener = try NWListener(using: parameters)
  }

  func start(completion: @escaping (Result<URL, Error>) -> Void) {
    listener.stateUpdateHandler = { [weak self] state in
      guard let self else { return }
      switch state {
      case .ready:
        guard let port = self.listener.port,
          let origin = URL(string: "http://127.0.0.1:\(port.rawValue)")
        else { return }
        self.origin = origin
        DispatchQueue.main.async {
          completion(.success(origin.appendingPathComponent(self.token + "/index.html")))
        }
      case .failed(let error): DispatchQueue.main.async { completion(.failure(error)) }
      default: break
      }
    }
    listener.newConnectionHandler = { [weak self] connection in
      guard let self else {
        connection.cancel()
        return
      }
      connection.start(queue: self.queue)
      self.queue.asyncAfter(deadline: .now() + 30) { connection.cancel() }
      self.receive(connection, accumulated: Data())
    }
    listener.start(queue: queue)
  }

  func put(_ bytes: Data, id: String) -> URL {
    lock.lock()
    documents[id] = bytes
    lock.unlock()
    return origin!.appendingPathComponent("\(token)/documents/\(id)")
  }

  func retainDocuments(_ ids: Set<String>) {
    lock.lock()
    documents = documents.filter { ids.contains($0.key) }
    lock.unlock()
  }

  func response(method: String, target: String, host: String) -> DesktopResponse {
    func failure(_ status: Int) -> DesktopResponse {
      DesktopResponse(status: status, contentType: "text/plain", body: Data())
    }
    guard method == "GET" || method == "HEAD" else { return failure(405) }
    guard let origin, host == "127.0.0.1:\(origin.port!)" else { return failure(403) }
    let encodedPath = String(target.split(separator: "?", maxSplits: 1, omittingEmptySubsequences: false)[0])
    guard let path = encodedPath.removingPercentEncoding,
      path.hasPrefix("/\(token)/"), !path.contains("\0")
    else { return failure(404) }
    let relative = String(path.dropFirst(token.count + 2))
    guard !relative.split(separator: "/").contains(".."), !relative.contains("\\") else {
      return failure(403)
    }
    if relative.hasPrefix("documents/") {
      lock.lock()
      let bytes = documents[String(relative.dropFirst(10))]
      lock.unlock()
      guard let bytes else { return failure(404) }
      return DesktopResponse(status: 200, contentType: "application/octet-stream", body: bytes)
    }
    let file = resources.appendingPathComponent(relative).standardizedFileURL
      .resolvingSymlinksInPath()
    guard file.path.hasPrefix(resources.path + "/"),
      let bytes = try? Data(contentsOf: file)
    else { return failure(404) }
    let types = [
      "html": "text/html; charset=utf-8", "js": "text/javascript", "mjs": "text/javascript",
      "css": "text/css", "wasm": "application/wasm", "json": "application/json",
      "svg": "image/svg+xml", "png": "image/png", "ttf": "font/ttf", "woff2": "font/woff2",
      "otf": "font/otf",
    ]
    return DesktopResponse(
      status: 200, contentType: types[file.pathExtension] ?? "application/octet-stream", body: bytes
    )
  }

  private func receive(_ connection: NWConnection, accumulated: Data) {
    connection.receive(minimumIncompleteLength: 1, maximumLength: 16_384) {
      [weak self] data, _, complete, error in
      guard let self, error == nil else {
        connection.cancel()
        return
      }
      var request = accumulated
      if let data { request.append(data) }
      guard request.count <= 32_768 else {
        connection.cancel()
        return
      }
      guard let boundary = request.range(of: Data("\r\n\r\n".utf8)) else {
        if complete { connection.cancel() } else { self.receive(connection, accumulated: request) }
        return
      }
      let text = String(decoding: request[..<boundary.lowerBound], as: UTF8.self)
      let lines = text.components(separatedBy: "\r\n")
      let start = lines[0].split(separator: " ")
      guard start.count == 3 else {
        connection.cancel()
        return
      }
      let host =
        lines.dropFirst().first { $0.lowercased().hasPrefix("host:") }?.dropFirst(5)
        .trimmingCharacters(in: .whitespaces) ?? ""
      let response = self.response(method: String(start[0]), target: String(start[1]), host: host)
      let headers =
        "HTTP/1.1 \(response.status) \(response.status == 200 ? "OK" : "Error")\r\nContent-Type: \(response.contentType)\r\nContent-Length: \(response.body.count)\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nCross-Origin-Resource-Policy: same-origin\r\nCross-Origin-Opener-Policy: same-origin\r\nCross-Origin-Embedder-Policy: require-corp\r\nConnection: close\r\n\r\n"
      var output = Data(headers.utf8)
      if start[0] != "HEAD" { output.append(response.body) }
      connection.send(content: output, completion: .contentProcessed { _ in connection.cancel() })
    }
  }
}
