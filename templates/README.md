# Templates

`everglades.pptx` in this folder is Communications’ official **2023 Everglades Foundation Presentation Template**. The rebrand engine inspects it as the server template.

The deck is 16 sample-slide layouts (one dummy `DEFAULT` master layout). It contains **both** brand-palette colors — **Green** is Sawgrass Lime `C1D451`, **Blue** is Water Teal `00ACBF` — documented on the staff notes slide. Staff pick one in the converter; convert applies that colorway to the whole rebuilt deck. Neither option is dropped from the template.

If you are given a `.potx` instead, drop it here as `everglades.potx` — `rebrand/inspect_template.py` rewrites it to `.pptx`. `npm run template` builds a python-pptx starter for engine development only; do not use it in place of this official file.
