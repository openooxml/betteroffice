---
"@betteroffice/xlsx": patch
"@betteroffice/agents": patch
---

Conditional formats keep italic text: opening, saving and recovery preserve `<i/>` in differential formats, `DifferentialFormatInput` and `ConditionalFormatInfo.format` gain `italic`, and the agent conditional format tools accept and list it. Recovery data from earlier releases still restores.
