#---
name: rebuild-editable-mechanism-ppt
description: Reconstruct flat mechanism-diagram images, screenshots, and exported figures as visually faithful, editable PowerPoint slides. Use for 图片转可编辑PPT、机制图复刻、流程图或论文机制图重绘、照图生成PPT、像素级还原、可修改PPT、editable scientific diagram, or when Codex must match a supplied PNG/JPG/PDF figure while preserving text, shapes, connectors, formulas, tables, and object-level editability. Search the local mechanism-figure corpus for an editable source first; otherwise rebuild with @oai/artifact-tool and verify both fidelity and editability.
---

# Rebuild Editable Mechanism PPT

## Core contract

Reconstruct the supplied figure at the same aspect ratio and make every recoverable element independently editable. Treat visual fidelity and editability as separate acceptance axes.

Never claim that an arbitrary raster image can be both pixel-identical and fully editable. A flat image has already lost fonts, layers, object boundaries, connector semantics, masks, and hidden geometry. Reach both goals only when a matching editable source is found. Otherwise report the remaining uncertainty and any raster-only regions.

Do not place the full reference image above editable objects and call the result editable. Allow a full-slide image only when the user explicitly chooses a visual-only reference mode; label its editability as near zero.

Do not use image generation to redraw source-specific illustrations when exact matching matters. Reuse the supplied source crop or an available source vector as an independent object, or rebuild it natively when practical.

Author or edit the final PPTX only with JavaScript ES modules and `@oai/artifact-tool`. Do not use `python-pptx` or the legacy Python artifact API. Use Python only for analysis and QA.

## Read before authoring

1. Read the installed `Presentations` skill, including its content rules, workspace rules, artifact-tool requirements, template-following rules, and full-slide QA requirements.
2. Before writing any deck code, read the current artifact-tool quick start and API docs named by that skill.
3. Read [reconstruction-contract.md](references/reconstruction-contract.md), [style-profile.md](references/style-profile.md), [scene-schema.md](references/scene-schema.md), and [qa-rubric.md](references/qa-rubric.md).
4. Treat the supplied image as the visual source of truth. Use the bundled style corpus only to resolve ambiguity or learn editable construction patterns; never blend in unrelated layouts.

Set `MECH_SKILL_DIR` to the absolute directory containing this `SKILL.md` before running bundled scripts.

## Select one reconstruction mode

Choose the first applicable mode:

1. **Editable-source match**: Search the user-provided files, `D:\Desktop\机制图`, and `assets/style-corpus`. If a matching PPTX or AI/PDF-derived editable source exists, preserve and reuse its native objects. Validate against the supplied image.
2. **Native-first rebuild**: Use native text, shapes, tables, charts, custom paths, and attached connectors. Make this the default when no source match exists.
3. **High-fidelity hybrid**: Keep all simple content native and use the smallest possible independent raster patches only for complex illustrations, textures, unsupported equations, or effects. Record every patch and its area.
4. **Visual-only reference**: Use a full-slide image only after the user explicitly prioritizes appearance over editability. Do not present it as a successful editable reconstruction.

Do not silently change modes during iteration.

## Workflow

### 1. Establish paths and inspect the source

Follow the `Presentations` skill workspace policy. Preserve the input file. Put analysis, generated `.mjs`, scene JSON, renders, diffs, and QA reports under the task scratch directory; put only final deliverables at the requested destination.

On Windows, normalize the runtime environment before invoking the bundled artifact-tool or Python helpers, especially on user profiles with non-ASCII characters:

```powershell
if (-not $env:HOME) { $env:HOME = $env:USERPROFILE }
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
```

Run:

```powershell
python "$MECH_SKILL_DIR\scripts\analyze_reference.py" <reference-image> --out <tmp>/reference-analysis.json
python "$MECH_SKILL_DIR\scripts\build_reference_index.py" "D:\Desktop\机制图" --out <tmp>/reference-index.json
```

Lock the slide aspect ratio to the source image ratio. Do not force 16:9: the learned corpus spans multiple ratios. Use the source pixel dimensions as the reconstruction coordinate system and map them consistently to slide units.

Check required fonts before layout tuning:

```powershell
powershell -ExecutionPolicy Bypass -File "$MECH_SKILL_DIR\scripts\check_fonts.ps1" -Fonts "Times New Roman,Cambria Math,Arial,Microsoft YaHei,微软雅黑"
```

Report missing fonts before substituting them. Prefer the exact font when installed; otherwise choose the closest measured fallback and log it.

### 2. Search for an editable source

Compare exact SHA-256 first, then perceptual hashes and rendered-slide similarity. Inspect likely PPTX pages rather than trusting filenames. Prefer, in order:

1. the exact originating PPTX page;
2. an AI/PDF source whose objects can be preserved through an allowed editable import path;
3. a closely related editable slide whose geometry and text correspond to the reference;
4. a native-first rebuild.

When a source PPTX supplies the actual page, use the `Presentations` template-following workflow. Do not import an unrelated corpus slide merely because its colors look similar.

### 3. Build a scene model

Create `scene.json` before deck code. Follow [scene-schema.md](references/scene-schema.md). Give every object:

- a stable ID and semantic role;
- source-pixel bounding box and z-order;
- editable class: `native`, `custom-path`, `vector-asset`, or `raster-patch`;
- fill, stroke, typography, and confidence;
- parent container and connector endpoints where applicable.

Run:

```powershell
python "$MECH_SKILL_DIR\scripts\validate_scene.py" <tmp>/scene.json
```

Transcribe all visible text exactly. Preserve case, punctuation, Greek letters, subscripts, superscripts, and mathematical italics. Use rich-text runs and Cambria Math/Unicode for recoverable formulas. If a formula cannot be represented reliably, crop only that formula as a raster patch and disclose it.

### 4. Rebuild with native objects

Create the final slide with plain `.mjs` and `@oai/artifact-tool`.

Apply these ordering rules:

1. slide background;
2. connectors and feedback loops;
3. large containers and regions;
4. nodes, icons, charts, and tables;
5. labels, formulas, annotations, and legends.

Create connectors before nodes so lines remain behind shapes and labels. Attach connectors to semantic endpoints when the runtime supports it; otherwise use explicit stable coordinates and names. Never let a connector cross text unintentionally.

Use native objects for text, rectangles, rounded rectangles, ellipses, lines, arrows, tables, simple charts, and simple icons. Use custom paths for bounded irregular geometry. Use independent image objects only for genuinely complex visuals. Name important objects predictably, such as `region-belief-state`, `node-feedback-delay`, and `edge-controller-to-action`.

Prefer coherent semantic objects over Illustrator-style fragmentation. Do not turn one icon or arrow into dozens of tiny freeform shards merely to increase the object count. Group related objects when the runtime supports reliable native grouping; otherwise use stable prefixes and document the intended group in the scene model.

Match the reference image, not generic presentation defaults. Preserve thin academic strokes, compact labels, and source font sizes even when they are smaller than normal presentation guidance, because this is a figure-reconstruction task with explicit visual guidance.

### 5. Render, compare, and iterate

Render the slide at the source image dimensions. Compare it with the reference:

```powershell
python "$MECH_SKILL_DIR\scripts\compare_render.py" <reference-image> <rendered-slide.png> --out <tmp>/fidelity.json --diff-image <tmp>/difference.png
python "$MECH_SKILL_DIR\scripts\audit_editability.py" <final.pptx> --out <tmp>/editability.json
```

When desktop PowerPoint is available, also render the final deck with Office at the source dimensions and compare that render separately:

```powershell
powershell -ExecutionPolicy Bypass -File "$MECH_SKILL_DIR\scripts\render_with_powerpoint.ps1" -Pptx <final.pptx> -OutputDir <tmp>/office-render -Width <source-width> -Height <source-height>
```

Inspect the full-size render, not only a montage. Fix geometry before color, color before typography, typography before micro-alignment. Re-render after each meaningful change.

For native-first or hybrid work, complete at least two visual correction passes after the first render unless the first render already meets its gate. Stop only when the gate is met or two consecutive passes improve the visual proxy by less than 0.005; in the latter case, report the failed gate and the remaining high-impact differences.

Use the two independent targets in [qa-rubric.md](references/qa-rubric.md). Do not hide low editability behind a high visual score, and do not accept a structurally editable slide that visibly differs from the reference.

### 6. Final QA and delivery

Run all checks required by the `Presentations` skill, including overflow and overlap checks. Inspect every slide at full size in both the artifact-tool render and, when PowerPoint is available, an Office render. Resolve font substitution, clipping, broken transparency, black-background artifacts, detached arrows, and unexpected wrapping.

Deliver:

- the editable `.pptx`;
- a concise fidelity/editability summary;
- a list of raster patches, missing fonts, and known deviations, if any.

Do not deliver the scene JSON, builder code, diff image, or scratch artifacts unless the user asks.

## Learned style corpus

Use [style-profile.md](references/style-profile.md) for the distilled style. Use `assets/style-corpus` only when a concrete example is needed.

Treat these as gold construction references:

- `机制图2.pptx`: native text and shapes with no slide pictures;
- `the sun also risef(2)(2)_formula_text.pptx`: editable formula-text variant;
- `机制图3.pptx`: dense 16:9 mixed scientific layout;
- `the sunalso rise(1).pptx` and `新机制图1.pptx`: related revision pair.

Treat full-slide screenshots and red replacement callouts as QA or instruction overlays, not reusable final style.

## Resource map

- `scripts/analyze_reference.py`: inspect canvas, transparency, palette, and edge density.
- `scripts/build_reference_index.py`: inventory images and PPTX files with hashes and structure.
- `scripts/validate_scene.py`: validate the object scene before authoring.
- `scripts/compare_render.py`: calculate render similarity and an optional QA diff.
- `scripts/audit_editability.py`: inspect PPTX native objects and raster coverage.
- `scripts/check_fonts.ps1`: report installed and missing Windows fonts.
- `scripts/render_with_powerpoint.ps1`: export a read-only Office render for final QA.
- `assets/style-corpus/manifest.json`: explain the bundled exemplars and provenance.
