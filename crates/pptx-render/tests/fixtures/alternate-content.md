`alternate-content.pptx` is a synthetic, one-slide repro for a shape wrapped in
`mc:AlternateContent` at `p:spTree` level, with no private presentation content.

The shape tree holds four top-level children in this order: a title `p:sp`, a `control` `p:sp`, an
`mc:AlternateContent`, and a caption `p:sp`. The `mc:Choice` requires
`urn:betteroffice:fixture-extension`, declared on the `mc:Choice` itself, and holds one element from
that namespace; the `mc:Fallback` holds the `fallback` `p:sp` — a `#315EFB` rectangle at (690, 150)
px carrying the text `Fallback shape`. No consumer implements that namespace, so the `mc:Fallback`
is the branch every one of them reads, and `control` beside it is an unchanged reference point.

The real-world spelling of this markup is PowerPoint ink: `<mc:Choice Requires="p14">` around a
`p:contentPart`, with a `p:pic` in the `mc:Fallback`. It is not what the fixture ships, because
LibreOffice implements `p14` and takes that branch, which would leave the reference render with
nothing where the fallback shape belongs. The parse tests cover the `p14` spelling directly.

The parse tests cover branch selection and the shape budget; the render test checks the fallback
shape's position, fill and text; the `pptx-edit` test fills the fallback shape through the public
edit API and checks that the fill lands inside the `mc:Fallback` with the `mc:Choice` intact, which
is what keeps parsed shape ordinals and the writer's source ordinals aligned.

`alternate-content.before.png` and `alternate-content.after.png` are the slide rendered at 1280x720
on `d2aaf9cb` and on this branch, with the harness's Liberation fonts. Before, the fallback shape is
absent; after, it matches the LibreOffice reference.
