# Modern PPT Transitions (PowerPoint 2013+ "Exciting" Gallery)

Three files work together:

- **transitions-modern.sh** — Build script.
- **transitions-modern.pptx** — 19-slide deck.
- **transitions-modern.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-modern.sh
# → transitions-modern.pptx
```

## What this trio shows

The 12 "Exciting" / "Dynamic Content" presets PowerPoint 2013 added
under one shared OOXML element: `<p15:prstTrans prst="..."/>`. All
twelve live in the `p15` namespace (Office 2012/main); none has its own
`p:`-namespace element. officecli writes each one inside an
`mc:AlternateContent` wrapper with an inline `<p:fade/>` fallback so
pre-2013 PowerPoint plays a graceful fade instead of nothing.

## The 12 presets

| CLI token | UI name | Notes |
|---|---|---|
| `fallOver` | Fall Over | direction-sensitive |
| `drape` | Drape | direction-sensitive |
| `curtains` | Curtains | symmetric |
| `wind` | Wind | direction-sensitive |
| `prestige` | Prestige | symmetric |
| `fracture` | Fracture | symmetric |
| `crush` | Crush | symmetric |
| `peelOff` | Peel Off | direction-sensitive |
| `pageCurlDouble` | Page Curl (double) | direction-sensitive |
| `pageCurlSingle` | Page Curl (single) | direction-sensitive |
| `airplane` | Airplane | direction-sensitive |
| `origami` | Origami | direction-sensitive |

Token spelling matches the OOXML `prst` attribute (lowerCamelCase).
Input is case-insensitive (`transition=PageCurlDouble` and
`pagecurldouble` both work), but `Get` returns the canonical
lowerCamelCase form.

## Direction (-in / -out)

```bash
officecli set deck.pptx /slide[N] --prop transition=pageCurlDouble
officecli set deck.pptx /slide[N] --prop transition=pageCurlDouble-out
officecli set deck.pptx /slide[N] --prop transition=wind-out
```

- Default (no suffix) = `-in` (no invX/invY attributes written).
- `-out` sets `invX="1" invY="1"` — flips the transition's directional
  axis. Visually affects the direction-sensitive presets in the table
  above; symmetric presets (curtains, fracture, crush, prestige) parse
  the suffix but render unchanged.
- Any other direction is rejected:
  ```
  Error: Transition 'fallOver' only accepts -in or -out (got '-up').
  ```

## OOXML representation

```xml
<mc:AlternateContent>
  <mc:Choice Requires="p15">
    <p:transition xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">
      <p15:prstTrans prst="pageCurlDouble" invX="1" invY="1"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

## See also

- [transitions-shapes.md](transitions-shapes.md) — Box (also a `p15:prstTrans` element) lives there alongside circle/diamond/zoom.
- [transitions-dynamic.md](transitions-dynamic.md) — the older 2010 "Exciting" gallery (`p14:` namespace transitions: vortex / switch / flip / ferris / ...).
- [transitions-morph.md](transitions-morph.md) — Morph (2016+), a separate `p159:` namespace element.
