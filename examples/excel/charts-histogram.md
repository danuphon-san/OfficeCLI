# Histogram Charts Showcase

This demo consists of three files that work together:

- **charts-histogram.py** — Python script that calls `officecli` to generate the workbook. Each chart command is shown as a copyable shell command in the comments.
- **charts-histogram.xlsx** — The generated workbook with 3 sheets and 12 histograms total.
- **charts-histogram.md** — This file. Maps each sheet to the features it demonstrates.

## Regenerate

```bash
cd examples/excel
python3 charts-histogram.py
# → charts-histogram.xlsx
```

## Why a dedicated histogram showcase?

Histograms are Excel's cx-namespace "extended" chart type. The binning layer
(`layoutPr/binning`) is where all the interesting knobs live — auto vs
explicit count, bin width, interval-closed side, outlier cut-offs — and
getting them right takes some care because Excel rejects the file entirely
if the XML uses the wrong form of `cx:binCount` / `cx:binSize`. This file
exercises every binning knob officecli exposes plus the common styling
props, so you can copy-paste from whichever row most matches the shape you
want.

## Chart Sheets

### Sheet: 1-Binning Modes

Four histograms demonstrating how to control binning.

```bash
# 1. Auto-binning (Excel picks the bin count)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop title="Test Scores (auto bins)" \
  --prop series1="Scores:45,52,58,61,63,...,97,99"

# 2. Explicit bin count
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Samples:<values>" \
  --prop binCount=10

# 3. Explicit bin width
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Sales:120,135,148,...,700" \
  --prop binSize=50

# 4. Underflow / overflow cut-offs (outlier fencing)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Latency:12,28,40,...,1200" \
  --prop underflowBin=50 \
  --prop overflowBin=500
```

**Features:** `chartType=histogram`, auto-binning (default), `binCount=N`, `binSize=W`, `underflowBin=N`, `overflowBin=M`

Notes:
- If both `binCount` and `binSize` are given, `binCount` wins.
- `underflowBin` / `overflowBin` group everything below/above the cut-off into a single `<N` or `>M` bar.
- Histograms default `gapWidth=0` (bars touch) to match Excel's native output.

### Sheet: 2-Styled Histograms

Four histograms demonstrating styling and presentation tweaks.

```bash
# 1. Single fill color + axis titles + value labels above bars
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Age:18,19,20,...,62" \
  --prop binCount=8 \
  --prop fill=4472C4 \
  --prop xAxisTitle="Age (years)" \
  --prop yAxisTitle="Count" \
  --prop dataLabels=true

# 2. Left-closed intervals [a,b) instead of default (a,b]
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="RT:10,12,15,...,100" \
  --prop binSize=10 \
  --prop intervalClosed=l \
  --prop fill=70AD47

# 3. Dense bell curve with category-axis gridlines on
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Samples:<120 gaussian values>" \
  --prop binCount=20 \
  --prop fill=ED7D31 \
  --prop xGridlines=true \
  --prop xAxisTitle="Value" \
  --prop yAxisTitle="Frequency"

# 4. Minimal (tick labels and gridlines off)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Orders:5,8,12,...,180" \
  --prop binCount=6 \
  --prop fill=7030A0 \
  --prop tickLabels=false \
  --prop gridlines=false
```

**Features:** `fill=HEX` (single-color histograms), `xAxisTitle` / `yAxisTitle`, `dataLabels`, `intervalClosed=l` (left-closed), `xGridlines`, `gridlines=false`, `tickLabels=false`

Notes:
- `fill=HEX` is the one-series shortcut; for multi-series extended charts use `colors=A,B,C`.
- `intervalClosed=l` affects how boundary values are assigned — relevant when values fall exactly on a bin edge.
- `gridlines` controls the value-axis gridlines (on by default); `xGridlines` controls the category-axis gridlines (off by default).
- `tickLabels=false` hides the bin range labels like `[100, 200]` on the x-axis.

### Sheet: 3-Typography & Colors

Four histograms showing the per-run text styling and axis/gridline color
knobs that turn a raw Excel default into a presentation-ready chart.
All four use the same vocabulary as regular cChart (`ChartHelper.Builder`),
so any knob you learn here transfers to `column`/`bar`/`line`/etc.

```bash
# 1. Title styling only (color / size / bold / font)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop title="Revenue Distribution" \
  --prop title.color=C00000 \
  --prop title.size=18 \
  --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop series1="Revenue:<values>" \
  --prop binCount=10 --prop fill=C00000

# 2. Axis title styling (applies to both X and Y axis titles)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop title="Session Duration" \
  --prop xAxisTitle="Minutes" --prop yAxisTitle="Users" \
  --prop axisTitle.color=4472C4 \
  --prop axisTitle.size=12 \
  --prop axisTitle.font="Helvetica Neue" \
  --prop series1="Minutes:<values>" --prop binCount=8 --prop fill=4472C4

# 3. Axis tick labels (compound "size:color:fontname") + gridline color
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop title="Latency Histogram" \
  --prop xAxisTitle="ms" --prop yAxisTitle="Requests" \
  --prop axisfont="10:555555:Helvetica Neue" \
  --prop gridlineColor=E0E0E0 \
  --prop series1="Latency:<values>" --prop binCount=12 --prop fill=70AD47

# 4. All four knobs combined — magazine-look chart
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop title="Bell Curve (magazine look)" \
  --prop title.color=2F5597 --prop title.size=16 --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=666666 --prop axisTitle.size=11 \
  --prop axisTitle.font="Helvetica Neue" \
  --prop axisfont="9:888888:Helvetica Neue" \
  --prop gridlineColor=EEEEEE \
  --prop series1="Samples:<values>" --prop binCount=15 --prop fill=4472C4
```

**Features:** `title.color` / `title.size` / `title.bold` / `title.font`,
`axisTitle.color` / `axisTitle.size` / `axisTitle.font`,
`axisfont="size:color:fontname"` (compound form for tick labels on both axes),
`gridlineColor` / `xGridlineColor`

Notes:
- `title.*` styles the chart title run (`<a:r><a:rPr>`). Accepts hex colors, `pt` sizes, truthy bold values, and any installed font family.
- `axisTitle.*` styles BOTH axis title runs at once — cx doesn't split X/Y title styling.
- `axisfont` is the compound form shared with regular cChart (see `ChartHelper.Builder.ApplyAxisTextProperties`). Format is `"size:color:fontname"`, all fields optional.
- `gridlineColor` controls the value-axis major gridlines; `xGridlineColor` controls the category-axis ones (only applies when `xGridlines=true`).
- Under the hood these keys are parsed by the shared `BuildDefaultRunPropertiesFromCompoundSpec` and `ApplyRunStyleProperties` helpers, so the same vocabulary works across regular cChart and cx extended charts.

### Sheet: 5-Parity Knobs

Six histograms demoing the "parity" knobs — vocabulary that used to exist
only in regular cChart and is now available on cx histograms. Each chart
isolates ONE knob combination so you can see exactly what each feature
emits in the OOXML.

```bash
# 1. Locked value-axis range (axismin / axismax / majorunit)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Scores:<values>" \
  --prop binCount=15 --prop fill=4472C4 \
  --prop axismin=0 --prop axismax=25 --prop majorunit=5

# 2. Drop shadows (series.shadow / title.shadow)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop title="Drop Shadows" --prop title.bold=true \
  --prop title.shadow="808080-4-45-3-50" \
  --prop series1="Sales:<values>" --prop binCount=15 --prop fill=ED7D31 \
  --prop series.shadow="000000-6-45-3-35"

# 3. Card look (plotareafill / chartareafill / plotarea.border / chartarea.border)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop title="Card Look" \
  --prop series1="Latency:<values>" --prop binCount=18 --prop fill=4472C4 \
  --prop plotareafill=FAFBFC --prop plotarea.border=D0D7DE:1 \
  --prop chartareafill=FFFFFF --prop chartarea.border=E5E5E5:0.75

# 4. Minimal axes (axisline / valaxis.visible=false)
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Events:<values>" --prop binSize=10 --prop fill=70AD47 \
  --prop axisline=888888:1.25 \
  --prop valaxis.visible=false \
  --prop gridlines=false

# 5. Data label format + overlaid legend + legend font
officecli add data.xlsx /Sheet --type chart \
  --prop chartType=histogram \
  --prop series1="Revenue:<values>" --prop binCount=10 --prop fill=C00000 \
  --prop dataLabels=true --prop datalabels.numfmt=0 \
  --prop legend=top --prop legend.overlay=true \
  --prop legendfont="10:555555:Helvetica Neue"
```

**Features:** `axismin` / `axismax` / `majorunit` / `minorunit` (value-axis
scaling), `axis.visible` / `cataxis.visible` / `valaxis.visible` (axis @hidden
flag), `axisline` / `cataxis.line` / `valaxis.line` (axis spine outline),
`plotareafill` / `chartareafill` (background solid fills), `plotarea.border`
/ `chartarea.border` (background outlines), `series.shadow` / `title.shadow`
(a:outerShdw effects, format `"COLOR-BLUR-ANGLE-DIST-OPACITY"`),
`legend.overlay` (bool), `legendfont` (compound `"size:color:fontname"`),
`datalabels.numfmt` (Excel format code like `0.0%`, `#,##0`).

Notes:
- All parity knobs share their vocabulary verbatim with regular cChart
  (`column` / `bar` / `line`), so knowledge transfers both ways. Under the
  hood both code paths call the same `ChartHelper.BuildOutlineElement` and
  `DrawingEffectsHelper.BuildOuterShadow` helpers.
- `axismin`/`axismax`/`majorunit`/`minorunit` land as string-typed
  attributes on `cx:valScaling`; the CLI still parses them as invariant
  doubles so NaN/Infinity are rejected.
- `cx:plotArea/cx:spPr` must appear AFTER all axes per `CT_PlotArea`
  schema — the builder handles the placement for you.

## Histogram Property Reference

| Property | Default | Notes |
|---|---|---|
| `chartType` | — | Must be `histogram` |
| `title` | — | Chart title |
| `series1` | — | `"name:v1,v2,v3,..."` — raw values, not pre-binned |
| `binCount` | auto | Integer: force exactly N bins |
| `binSize` | auto | Number: force fixed bin width |
| `intervalClosed` | `r` | `r` = (a,b], `l` = [a,b) |
| `underflowBin` | — | Group values < N into a single `<N` bar |
| `overflowBin` | — | Group values > M into a single `>M` bar |
| `gapWidth` | `0` | Space between bars (0 = touching) |
| `fill` | — | Single-color shortcut (HEX) |
| `colors` | — | Comma list of HEX (multi-series) |
| `dataLabels` | `false` | `true` puts value count above each bar |
| `datalabels.numfmt` | — | Excel format code for data labels (e.g. `0.0`, `0.00%`, `#,##0`) |
| `xAxisTitle` | — | Category-axis title |
| `yAxisTitle` | — | Value-axis title |
| `gridlines` | `true` | Value-axis major gridlines |
| `xGridlines` | `false` | Category-axis major gridlines |
| `tickLabels` | `true` | Show bin range labels on x-axis |
| `axismin` / `min` | — | Value-axis lower bound (numeric) |
| `axismax` / `max` | — | Value-axis upper bound (numeric) |
| `majorunit` | — | Value-axis major gridline interval |
| `minorunit` | — | Value-axis minor gridline interval |
| `axis.visible` / `axis.delete` | — | `false` hides both axes (`axis.delete` is the inverse alias) |
| `cataxis.visible` | — | Category-axis @hidden flag (false = hidden) |
| `valaxis.visible` | — | Value-axis @hidden flag (false = hidden) |
| `axisline` | — | Axis spine: `"color"` / `"color:width"` / `"color:width:dash"` / `"none"` (both axes) |
| `cataxis.line` | — | Category-axis spine only |
| `valaxis.line` | — | Value-axis spine only |
| `plotareafill` / `plotfill` | — | Plot-area solid background color (hex) |
| `plotarea.border` / `plotborder` | — | Plot-area outline (`color` / `color:width` / `color:width:dash` / `none`) |
| `chartareafill` / `chartfill` | — | Chart-area solid background color (hex) |
| `chartarea.border` / `chartborder` | — | Chart-area outline |
| `series.shadow` | — | Outer shadow on bars: `"COLOR-BLUR-ANGLE-DIST-OPACITY"` or `none` |
| `legend` | — | `top` / `bottom` / `left` / `right` / `none` |
| `legend.overlay` | `false` | Legend floats on top of plot area when `true` |
| `legendfont` | — | Compound `"size:color:fontname"` — legend text styling |
| `title.color` | — | Chart title color (hex) |
| `title.size` | — | Chart title font size (pt) |
| `title.bold` | — | Chart title bold (`true`/`false`) |
| `title.font` | — | Chart title font family |
| `title.shadow` | — | Chart title drop shadow: `"COLOR-BLUR-ANGLE-DIST-OPACITY"` or `none` |
| `axisTitle.color` | — | Axis title color (both X and Y) |
| `axisTitle.size` | — | Axis title font size (pt) |
| `axisTitle.font` | — | Axis title font family |
| `axisTitle.bold` | — | Axis title bold |
| `axisfont` | — | Compound tick-label styling: `"size:color:fontname"` |
| `gridlineColor` | — | Value-axis major gridline color (hex) |
| `xGridlineColor` | — | Category-axis major gridline color (requires `xGridlines=true`) |

## Inspect the Generated File

```bash
officecli query charts-histogram.xlsx chart
officecli get charts-histogram.xlsx "/1-Binning Modes/chart[1]"
officecli view charts-histogram.xlsx "/2-Styled Histograms"
```
