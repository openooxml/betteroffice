---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
---

The PowerPoint engine is about 3 KB smaller compressed: the slide layout cache compares slide floats with a direct walk instead of a generic serializer. Layout and cache behaviour are unchanged.
