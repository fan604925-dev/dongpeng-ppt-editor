---
name: dongpeng-ppt-editor
description: Create, edit, inspect, render, and validate editable Dongpeng corporate PowerPoint presentations with strict fidelity to the official red-and-white template. Use for Dongpeng PPT or .pptx work including creating decks from the standard template; adding, duplicating, deleting, or reordering slides; editing text, images, shapes, tables, charts, and icons; preserving brand layouts; saving PPTX files; exporting PDF or slide images; and checking fonts, colors, overflow, cropping, and brand elements.
---

# Dongpeng PPT Editor

Use the official template as the source of truth. Prefer duplicating registered template slides over drawing pages from scratch.

## Required environment

- Run on Windows with desktop Microsoft PowerPoint installed.
- Invoke PowerPoint automation through the bundled PowerShell scripts.
- Request GUI/PowerPoint execution approval when the environment requires it.
- Keep `assets/东鹏集团PPT标准模板.pptx` read-only in practice. Never overwrite it.

## Core workflow

1. Read [references/design-rules.md](references/design-rules.md) before creating or restyling slides.
2. Read [references/layout-catalog.md](references/layout-catalog.md) and choose a registered layout for every new slide.
3. Inspect an existing deck before editing:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\inspect-ppt.ps1" -InputPath "deck.pptx" -OutputPath "inspection.json"
   ```

4. Create an operations JSON file outside the Skill directory. Follow [references/operations-schema.md](references/operations-schema.md).
5. Apply operations to a new output path:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\edit-ppt.ps1" -SpecPath "operations.json"
   ```

6. Render PNG previews and export PDF with either the edit specification or explicit export operations.
7. Validate the resulting deck:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\validate-ppt.ps1" -InputPath "output.pptx" -OutputPath "validation.json"
   ```

8. Visually inspect every rendered slide. Fix overflow, bad cropping, inconsistent alignment, missing brand chrome, or unreadable contrast before delivery.

## Creation rules

- Start a new deck with `new_deck: true`, then add pages through `insert_template_slide`.
- Use `layout_id` values from `assets/layout-map.json`.
- Use semantic `role` selectors when available. Inspect the slide and use an explicit selector when a role is not registered.
- Build narrative rhythm with red ceremony slides and white information slides. For decks of at least five pages, include at least two red and two white pages.
- Keep the official cover, section, footer, logo, and tagline assets unchanged.

## Editing rules

- Write to a new output file unless the user explicitly authorizes overwriting.
- Preserve the page size at 960 × 540 pt and the 16:9 ratio.
- Use Microsoft YaHei for Chinese and Arial or Microsoft YaHei for Latin text and numbers.
- Use `#E21413` as the only primary accent. Use only registered red tints and neutral grays.
- Keep body text at 12 pt or larger. Use 9 pt only for the small upper-left section label and 8.4 pt only for copyright.
- Prefer `cover` for room/scene photographs and `contain` for product cutouts, texture boards, screenshots, and diagrams.
- Never stretch images unless the user explicitly requests it.
- Preserve straight corners and flat composition. Do not introduce generic SaaS cards, neon gradients, or unrelated icon styles.
- Update existing template charts and tables when possible instead of rebuilding their styling.

## Script selection

- Use `inspect-ppt.ps1` to discover slide IDs, tags, text, paths, types, and geometry.
- Use `edit-ppt.ps1` for slide, text, image, shape, table, save, PDF, and PNG operations.
- Use `validate-ppt.ps1` for structural and visual-risk checks.
- Read `assets/layout-map.json` when resolving a semantic role or adding a template layout.

## Delivery

Deliver the editable `.pptx`, the exported `.pdf` when requested, and preview images when visual verification matters. Report any validation warnings that could not be resolved.

