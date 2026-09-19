---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Draw eleven more PowerPoint preset shapes. `donut`, `noSmoking`, `corner`, `foldedCorner`, `mathMultiply`, `bentArrow`, `ribbon`, `ellipseRibbon`, `cloudCallout`, `wedgeEllipseCallout` and `wedgeRoundRectCallout` fell back to a plain rectangle; they now follow their ECMA-376 definitions, including the elliptical arcs `arcTo` measures by polar angle rather than by ellipse parameter.
