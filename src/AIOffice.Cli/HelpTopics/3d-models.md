# 3D models (pptx) — v1.3

Embed a glTF binary (`.glb`) or text (`.gltf`) 3D model on a slide as a real
`3DModel` media part. PowerPoint 2019+ renders the model; older viewers fall
back to a poster picture, so AIOffice always writes the embed *behind* a poster
picture and flags the honest `model3d_as_media` warning.

```jsonc
{op:"add", path:"/slide[1]", type:"model3d", props:{
  src:"chair.glb",        // required; sandbox-resolved inside the workspace
  poster:"chair.png",     // optional fallback image (PNG/JPEG)
  x:"3cm", y:"3cm", w:"8cm", h:"6cm",
  name:"Chair"}}
```

- `src` is **required** and **sandbox-resolved** — a path that escapes the
  workspace is `sandbox_denied` and the bytes are never read.
- `poster` (optional) is likewise sandbox-resolved; omit it and AIOffice draws a
  labelled grey placeholder.
- The model lands at `/slide[i]/model3d[@id=N]`; the shape id is stable across
  edits. `get` reports it; `remove` deletes the model + media part.

Result note: every `add model3d` carries the `model3d_as_media` warning, stating
the model was embedded as a 3D part behind a poster fallback. The file validates
OpenXmlValidator-clean.
