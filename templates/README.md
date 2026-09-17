# Templates

`everglades.pptx` in this folder is Communications’ official **2023 Everglades Foundation Presentation Template**. The rebrand engine inspects it as the server template.

The deck is 16 sample-slide layouts (one dummy `DEFAULT` master layout). **Green and blue are brand-palette colors on those slides** — Sawgrass Lime `C1D451` and Water Teal `00ACBF`, documented on the staff notes slide — not Design-tab variants, extra masters, or a converter colorway overlay. Convert clones the prototypes as drawn. The staff UI shows a design picker only if a later template actually encodes two selectable options.

If you are given a `.potx` instead, drop it here as `everglades.potx` — `rebrand/inspect_template.py` rewrites it to `.pptx`. `npm run template` builds a python-pptx starter for engine development only; do not use it in place of this official file.
