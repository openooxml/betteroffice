---
"@betteroffice/xlsx": patch
---

The default collaboration frame limit drops from 64 MiB to 16 MiB to match what the relay accepts and retains, and an oversized frame now raises a protocol error before it is sent instead of closing the socket. Hosts running their own relay can restore the previous ceiling with the `maxFrameBytes` option.
