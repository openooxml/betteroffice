import { describe, expect, test } from "bun:test";

import { collabUrl, createWopiTransport, type WebSocketLike } from "./transport";

class FakeSocket implements WebSocketLike {
  binaryType = "";
  readyState = 0;
  bufferedAmount = 0;
  onopen: (() => void) | null = null;
  onmessage: ((event: { data: unknown }) => void) | null = null;
  onerror: ((event: unknown) => void) | null = null;
  onclose: ((event: { reason?: string }) => void) | null = null;
  sent: ArrayBuffer[] = [];
  closedWith: Array<[number | undefined, string | undefined]> = [];

  send(data: ArrayBuffer): void {
    this.sent.push(data);
  }

  close(code?: number, reason?: string): void {
    this.closedWith.push([code, reason]);
    this.readyState = 3;
  }
}

describe("collabUrl", () => {
  test("maps http and https origins to the /collab endpoint", () => {
    expect(collabUrl("http://127.0.0.1:3000")).toBe("ws://127.0.0.1:3000/collab");
    expect(collabUrl("https://office.example.com")).toBe("wss://office.example.com/collab");
  });
});

describe("createWopiTransport", () => {
  test("connects to /collab with the handle as the second subprotocol", () => {
    const sockets: FakeSocket[] = [];
    const seen: Array<{ url: string; protocols: string[] }> = [];
    const transport = createWopiTransport({
      handle: "handle-abc",
      origin: "https://office.example.com",
      createSocket: (url, protocols) => {
        seen.push({ url, protocols });
        const socket = new FakeSocket();
        sockets.push(socket);
        return socket;
      },
    });
    transport.connect();
    expect(seen).toEqual([{ url: "wss://office.example.com/collab", protocols: ["bo-wopi.v1", "handle-abc"] }]);
    expect(sockets[0]!.binaryType).toBe("arraybuffer");
  });

  test("forwards binary messages and sends only while open", () => {
    const socket = new FakeSocket();
    const transport = createWopiTransport({
      handle: "h",
      origin: "http://127.0.0.1:3000",
      createSocket: () => socket,
    });
    const events: string[] = [];
    transport.onEvent((event) => events.push(event.type));
    expect(transport.send(Uint8Array.of(1))).toBe(false);
    transport.connect();
    socket.readyState = 1;
    socket.onopen?.();
    expect(transport.send(Uint8Array.of(1, 2))).toBe(true);
    expect(new Uint8Array(socket.sent[0]!)).toEqual(Uint8Array.of(1, 2));
    socket.onmessage?.({ data: new Uint8Array([3]).buffer });
    expect(events).toEqual(["open", "message"]);
    transport.disconnect();
    expect(socket.closedWith).toEqual([[1000, "Disconnected"]]);
  });
});
