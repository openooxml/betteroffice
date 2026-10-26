---
"@betteroffice/docx": patch
"@betteroffice/python-docx": patch
"@betteroffice/rust-crates": patch
---

Chart data labels in Word documents use the number format the chart's values were saved with when the labels name none, so a percentage chart's labels read 20.8% instead of 0.2, as in Word.
