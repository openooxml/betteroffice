import AppKit
import Foundation

let variants = ["doc", "sheet", "deck"]

guard CommandLine.arguments.count == 3, variants.contains(CommandLine.arguments[1]) else {
    FileHandle.standardError.write(Data("usage: generate-icon.swift \(variants.joined(separator: "|")) OUT.png\n".utf8))
    exit(2)
}

let variant = CommandLine.arguments[1]
let side = 1024
let unit = CGFloat(side) / 32

struct Block {
    let x: CGFloat
    let y: CGFloat
    let w: CGFloat
    let h: CGFloat
    let alpha: CGFloat
    let white: Bool

    init(_ x: CGFloat, _ y: CGFloat, _ w: CGFloat, _ h: CGFloat, _ alpha: CGFloat = 1, white: Bool = true) {
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.alpha = alpha
        self.white = white
    }
}

func frame(_ x: CGFloat, _ y: CGFloat, _ w: CGFloat, _ h: CGFloat, _ t: CGFloat, _ alpha: CGFloat = 1) -> [Block] {
    [
        Block(x, y, w, t, alpha),
        Block(x, y + h - t, w, t, alpha),
        Block(x, y + t, t, h - 2 * t, alpha),
        Block(x + w - t, y + t, t, h - 2 * t, alpha),
    ]
}

let blocks: [Block]
switch variant {
case "doc":
    blocks = [
        Block(6, 4, 20, 3),
        Block(6, 11, 20, 3),
        Block(6, 18, 20, 3),
        Block(10, 25, 12, 3, 0.55),
    ]
case "sheet":
    blocks = [
        Block(5, 11, 6, 14),
        Block(13, 5, 6, 20),
        Block(21, 15, 6, 10, 0.55),
        Block(4, 25, 24, 2),
    ]
case "deck":
    blocks = [Block(11, 6, 18, 12, 1, white: false)] + frame(11, 6, 18, 12, 2, 0.55)
        + [Block(7, 10, 18, 12, 1, white: false)] + frame(7, 10, 18, 12, 2)
        + [Block(3, 14, 18, 12, 1, white: false)] + frame(3, 14, 18, 12, 2)
default:
    preconditionFailure("unsupported icon variant")
}

guard
    let context = CGContext(
        data: nil,
        width: side,
        height: side,
        bitsPerComponent: 8,
        bytesPerRow: 0,
        space: CGColorSpaceCreateDeviceRGB(),
        bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
    )
else {
    exit(1)
}

context.setShouldAntialias(false)
context.setFillColor(CGColor(red: 0, green: 0, blue: 0, alpha: 1))
let ground = CGPath(
    roundedRect: CGRect(x: 0, y: 0, width: CGFloat(side), height: CGFloat(side)),
    cornerWidth: unit * 7,
    cornerHeight: unit * 7,
    transform: nil
)
context.addPath(ground)
context.fillPath()

for block in blocks {
    let level: CGFloat = block.white ? 1 : 0
    context.setFillColor(CGColor(red: level, green: level, blue: level, alpha: block.alpha))
    context.fill(
        CGRect(
            x: block.x * unit,
            y: (32 - block.y - block.h) * unit,
            width: block.w * unit,
            height: block.h * unit
        )
    )
}

guard
    let image = context.makeImage(),
    let png = NSBitmapImageRep(cgImage: image).representation(using: .png, properties: [:])
else {
    exit(1)
}

try png.write(to: URL(fileURLWithPath: CommandLine.arguments[2]), options: .atomic)
