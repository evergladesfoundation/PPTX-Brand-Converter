# Templates

`everglades.pptx` in this folder is Communications’ official **2023 Everglades Foundation Presentation Template**. The rebrand engine inspects it as the server template.

The deck is a set of sample-slide layouts with **green** (sawgrass lime `C1D451`) and **blue** (water teal `00ACBF`) colorways. Staff pick a colorway in the converter UI before convert; inspect, plan, rebuild, and QA all use that palette’s layout map and theme tokens.

If you are given a `.potx` instead, drop it here as `everglades.potx` — `rebrand/inspect_template.py` rewrites it to `.pptx`. `npm run template` builds a python-pptx starter for engine development only; do not use it in place of this official file.
