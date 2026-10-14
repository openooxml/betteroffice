import Foundation

guard CommandLine.arguments.count == 3 else {
    FileHandle.standardError.write(Data("usage: package-icns.swift ICONSET OUTPUT.icns\n".utf8))
    exit(2)
}

let iconset = URL(fileURLWithPath: CommandLine.arguments[1])
let output = URL(fileURLWithPath: CommandLine.arguments[2])
let entries = [
    ("icp4", "icon_16x16.png"),
    ("icp5", "icon_32x32.png"),
    ("ic07", "icon_128x128.png"),
    ("ic08", "icon_256x256.png"),
    ("ic09", "icon_512x512.png"),
    ("ic11", "icon_16x16@2x.png"),
    ("ic12", "icon_32x32@2x.png"),
    ("ic13", "icon_128x128@2x.png"),
    ("ic14", "icon_256x256@2x.png"),
    ("ic10", "icon_512x512@2x.png"),
]

func appendSize(_ count: Int, to data: inout Data) {
    var size = UInt32(count).bigEndian
    Swift.withUnsafeBytes(of: &size) { data.append(contentsOf: $0) }
}

var chunks = Data()
for (type, name) in entries {
    let image = try Data(contentsOf: iconset.appendingPathComponent(name))
    chunks.append(contentsOf: type.utf8)
    appendSize(image.count + 8, to: &chunks)
    chunks.append(image)
}

var icns = Data("icns".utf8)
appendSize(chunks.count + 8, to: &icns)
icns.append(chunks)
try icns.write(to: output, options: .atomic)
