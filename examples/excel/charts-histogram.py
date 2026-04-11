#!/usr/bin/env python3
"""
Histogram Charts Showcase — every binning mode, every styling knob,
and a visual gallery of the iconic distribution shapes.

Generates: charts-histogram.xlsx (4 sheets, 16 histograms)

Usage:
  python3 charts-histogram.py
"""

import subprocess, os, atexit, random, math

FILE = "charts-histogram.xlsx"

def cli(cmd):
    """Run: officecli <cmd>"""
    r = subprocess.run(f"officecli {cmd}", shell=True, capture_output=True, text=True)
    out = (r.stdout or "").strip()
    if out:
        for line in out.split("\n"):
            if line.strip():
                print(f"  {line.strip()}")
    if r.returncode != 0:
        err = (r.stderr or "").strip()
        if err and "UNSUPPORTED" not in err and "process cannot access" not in err:
            print(f"  ERROR: {err}")

if os.path.exists(FILE):
    os.remove(FILE)

cli(f'create "{FILE}"')
cli(f'open "{FILE}"')
atexit.register(lambda: cli(f'close "{FILE}"'))

# --------------------------------------------------------------------------
# Deterministic sample generators — same seed, same file every regeneration.
# --------------------------------------------------------------------------
random.seed(42)
BELL = sorted(round(random.gauss(75, 12), 1) for _ in range(120))
BELL_CSV = ",".join(str(v) for v in BELL)

# Bimodal: two cohorts (e.g. beginners ~55, experts ~88) glued together.
random.seed(7)
BIMODAL = sorted(
    [round(random.gauss(55, 6), 1) for _ in range(60)] +
    [round(random.gauss(88, 5), 1) for _ in range(60)]
)
BIMODAL_CSV = ",".join(str(v) for v in BIMODAL)

# Right-skewed / log-normal: classic income shape, long tail to the right.
random.seed(11)
SKEWED_R = sorted(round(math.exp(random.gauss(3.2, 0.55)), 1) for _ in range(150))
SKEWED_R_CSV = ",".join(str(v) for v in SKEWED_R)

# Left-skewed: retirement ages — most cluster high, few retire early.
random.seed(23)
SKEWED_L = sorted(round(75 - math.exp(random.gauss(1.6, 0.6)), 1) for _ in range(120))
SKEWED_L_CSV = ",".join(str(v) for v in SKEWED_L)

# Uniform: sample IDs drawn evenly across a range.
random.seed(31)
UNIFORM = sorted(round(random.uniform(0, 100), 1) for _ in range(140))
UNIFORM_CSV = ",".join(str(v) for v in UNIFORM)

# Heavy-tailed (Pareto / power-law): API latencies — most are fast, tiny
# fraction are catastrophic. Perfect target for underflow/overflow fencing.
random.seed(47)
HEAVY = sorted(round(random.paretovariate(1.6) * 20, 1) for _ in range(160))
HEAVY_CSV = ",".join(str(v) for v in HEAVY)


# ==========================================================================
# Sheet: 1-Binning Modes
#
# Four histograms showing the raw binning vocabulary with no styling:
# auto / binCount / binSize / underflow+overflow. This sheet is the
# "Rosetta stone" — once these four work for you, all of Excel's
# histogram binning is accessible.
# ==========================================================================
print("\n--- 1-Binning Modes ---")
cli(f'set "{FILE}" /Sheet1 --prop name="1-Binning Modes"')

# --------------------------------------------------------------------------
# Chart 1: Auto-binning (Excel's default — no binCount, no binSize)
#
# officecli add charts-histogram.xlsx "/1-Binning Modes" --type chart \
#   --prop chartType=histogram \
#   --prop title="Test Scores (auto bins)" \
#   --prop series1="Scores:45,52,58,..." \
#   --prop x=0 --prop y=0 --prop width=13 --prop height=18
#
# Features: chartType=histogram, auto-binning (Excel picks bin count)
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Binning Modes" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Test Scores (auto bins)"'
    f' --prop series1=Scores:45,52,58,61,63,65,67,68,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,97,99'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Explicit bin count (binCount=10)
#
# Features: binCount=N forces exactly N bins regardless of range.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Binning Modes" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Bell Curve (binCount=10)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=10'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: Explicit bin width (binSize=50)
#
# Features: binSize=W forces bins of fixed width W.
#   binCount wins if both are given.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Binning Modes" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Sales ($, binSize=50)"'
    f' --prop series1=Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700'
    f' --prop binSize=50'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Underflow / overflow bins (outlier cut-offs)
#
# Features: underflowBin=N / overflowBin=M group values <N / >M into a
#   single "<N" or ">M" bar — classic outlier fencing.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/1-Binning Modes" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Latency (ms) with <50 and >500 cut-offs"'
    f' --prop series1=Latency:12,28,40,60,75,82,95,110,125,140,155,170,185,200,220,240,260,285,310,340,380,420,460,510,620,750,900,1200'
    f' --prop underflowBin=50'
    f' --prop overflowBin=500'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet: 2-Styled Histograms
#
# Four histograms covering every styling knob: fill, axis titles, data
# labels, interval-closed side, gridlines (both axes + colors), tick-label
# hiding, gap width, legend, and bold axis titles.
# ==========================================================================
print("\n--- 2-Styled Histograms ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Styled Histograms"')

# --------------------------------------------------------------------------
# Chart 1: Custom fill + axis titles + value labels + legend=top
#
# Features: fill=HEX, xAxisTitle, yAxisTitle, dataLabels, legend=top.
#   The series name ("Age") is picked up by the top legend.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Styled Histograms" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Age Distribution"'
    f' --prop series1=Age:18,19,20,21,21,22,22,23,23,24,24,25,25,26,27,28,29,30,31,33,35,38,42,45,50,55,62'
    f' --prop binCount=8'
    f' --prop fill=4472C4'
    f' --prop xAxisTitle="Age (years)"'
    f' --prop yAxisTitle="Count"'
    f' --prop dataLabels=true'
    f' --prop legend=top'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Left-closed intervals + xGridlines + xGridlineColor
#
# Features: intervalClosed=l ([a,b) half-open) + xGridlines=true +
#   xGridlineColor=hex to color the category-axis gridlines.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Styled Histograms" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Response Times [a,b) - left-closed"'
    f' --prop series1=RT:10,12,15,18,20,22,25,28,30,33,35,38,40,42,45,48,50,55,60,65,70,75,80,90,100'
    f' --prop binSize=10'
    f' --prop intervalClosed=l'
    f' --prop fill=70AD47'
    f' --prop xGridlines=true'
    f' --prop xGridlineColor=B4E0A0'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: Dense bell curve with gapWidth + bold axis titles
#
# Features: gapWidth=20 (bars no longer touch — useful when you want
#   histogram bars to look column-chart-ish), axisTitle.bold=true.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Styled Histograms" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Bell Curve (120 samples, gapWidth=20)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=20'
    f' --prop fill=ED7D31'
    f' --prop gapWidth=20'
    f' --prop xAxisTitle="Value"'
    f' --prop yAxisTitle="Frequency"'
    f' --prop axisTitle.bold=true'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Minimal — tickLabels + gridlines off (sparkline feel)
#
# Features: tickLabels=false hides bin range labels on X,
#   gridlines=false hides the value-axis gridlines.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/2-Styled Histograms" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Order Value (minimal)"'
    f' --prop series1=Orders:5,8,12,15,18,22,25,28,32,35,40,45,50,58,65,75,88,105,130,180'
    f' --prop binCount=6'
    f' --prop fill=7030A0'
    f' --prop tickLabels=false'
    f' --prop gridlines=false'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet: 3-Typography & Colors
#
# Per-run text styling and axis/gridline colors. Same vocabulary as regular
# cChart (ChartHelper.Builder.ApplyRunStyleProperties), so everything you
# learn here transfers to column/bar/line/etc.
# ==========================================================================
print("\n--- 3-Typography & Colors ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Typography & Colors"')

# --------------------------------------------------------------------------
# Chart 1: Title styling alone — color/size/bold/font
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Typography & Colors" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Revenue Distribution"'
    f' --prop title.color=C00000'
    f' --prop title.size=18'
    f' --prop title.bold=true'
    f' --prop title.font="Helvetica Neue"'
    f' --prop series1=Revenue:{BELL_CSV}'
    f' --prop binCount=10'
    f' --prop fill=C00000'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Axis title styling — color/size/font
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Typography & Colors" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Session Duration"'
    f' --prop series1=Minutes:2,3,4,5,5,6,6,7,7,8,9,10,11,12,14,16,18,20,22,25,28,32,36,42,50,58,65,72,85,100'
    f' --prop binCount=8'
    f' --prop fill=4472C4'
    f' --prop xAxisTitle="Minutes"'
    f' --prop yAxisTitle="Users"'
    f' --prop axisTitle.color=4472C4'
    f' --prop axisTitle.size=12'
    f' --prop axisTitle.font="Helvetica Neue"'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: axisfont compound form + value-axis gridline color
#
# `axisfont="size:color:fontname"` is the same compound form regular cChart
# uses (ChartHelper.Builder.ApplyAxisTextProperties). Styles BOTH axes' tick
# labels in one go.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Typography & Colors" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Latency Histogram"'
    f' --prop series1=Latency:12,28,40,60,75,82,95,110,125,140,155,170,185,200,220,240,260,285,310,340,380,420,460,510,620,750,900,1200'
    f' --prop binCount=12'
    f' --prop fill=70AD47'
    f' --prop xAxisTitle="ms"'
    f' --prop yAxisTitle="Requests"'
    f' --prop axisfont="10:555555:Helvetica Neue"'
    f' --prop gridlineColor=E0E0E0'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: All typography knobs together — "magazine look"
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/3-Typography & Colors" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Bell Curve (magazine look)"'
    f' --prop title.color=2F5597'
    f' --prop title.size=16'
    f' --prop title.bold=true'
    f' --prop title.font="Helvetica Neue"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=15'
    f' --prop fill=4472C4'
    f' --prop xAxisTitle="Score"'
    f' --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=666666'
    f' --prop axisTitle.size=11'
    f' --prop axisTitle.font="Helvetica Neue"'
    f' --prop axisfont="9:888888:Helvetica Neue"'
    f' --prop gridlineColor=EEEEEE'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet: 4-Distribution Gallery
#
# A 2x3 visual gallery of the *canonical* distribution shapes you'll see
# in real-world data. Same styling vocabulary as Sheet 3, different data.
# The goal: pattern-recognition. If you ever see one of these shapes in
# production telemetry, you know what you're looking at.
#
#   ┌────────────┬────────────┐
#   │  Bimodal   │ Right-skew │
#   ├────────────┼────────────┤
#   │   Uniform  │ Left-skew  │
#   ├────────────┼────────────┤
#   │ Heavy-tail │   Normal   │
#   └────────────┴────────────┘
#
# Each chart uses a distinct color from the default Excel palette to keep
# the gallery scan-friendly. All four chart areas use the same magazine-
# grade axisfont + gridlineColor so the gallery reads as a cohesive set.
# ==========================================================================
print("\n--- 4-Distribution Gallery ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Distribution Gallery"')

# Shared typography "house style" — applied to every chart in the gallery
# so the page reads as a consistent dashboard rather than a grab-bag.
STYLE = (
    ' --prop title.color=1F2937'
    ' --prop title.size=14'
    ' --prop title.bold=true'
    ' --prop title.font="Helvetica Neue"'
    ' --prop axisTitle.color=6B7280'
    ' --prop axisTitle.size=10'
    ' --prop axisTitle.font="Helvetica Neue"'
    ' --prop axisfont="9:6B7280:Helvetica Neue"'
    ' --prop gridlineColor=EEEEEE'
)

# --------------------------------------------------------------------------
# Chart 1: Bimodal — two cohorts (beginners + experts) fused together.
#   Classic sign that your data has two populations you should split.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Distribution Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Bimodal — two hidden populations"'
    f' --prop series1=Score:{BIMODAL_CSV}'
    f' --prop binCount=20'
    f' --prop fill=4472C4'
    f' --prop xAxisTitle="Test score"'
    f' --prop yAxisTitle="Count"'
    f'{STYLE}'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Right-skewed (log-normal) — classic income/file-size/wait-time
#   shape. Mean >> median; long tail pulls right.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Distribution Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Right-skewed — log-normal (income)"'
    f' --prop series1=Income:{SKEWED_R_CSV}'
    f' --prop binCount=18'
    f' --prop fill=ED7D31'
    f' --prop xAxisTitle="Monthly income ($k)"'
    f' --prop yAxisTitle="People"'
    f'{STYLE}'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: Uniform — every value equally likely. Flat top is the tell.
#   binSize rather than binCount to emphasize the "flat floor".
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Distribution Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Uniform — flat floor"'
    f' --prop series1=Draws:{UNIFORM_CSV}'
    f' --prop binSize=10'
    f' --prop fill=70AD47'
    f' --prop xAxisTitle="Random draw (0-100)"'
    f' --prop yAxisTitle="Count"'
    f'{STYLE}'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Left-skewed — retirement ages. Most cluster near the cap,
#   tail stretches toward earlier ages.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Distribution Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Left-skewed — retirement ages"'
    f' --prop series1=Age:{SKEWED_L_CSV}'
    f' --prop binCount=16'
    f' --prop fill=7030A0'
    f' --prop xAxisTitle="Age at retirement"'
    f' --prop yAxisTitle="Count"'
    f'{STYLE}'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 5: Heavy-tailed (Pareto / power-law) — API latencies. Most fast,
#   tiny fraction catastrophic. Uses overflowBin to fence the long tail
#   so the interesting bulk stays readable.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Distribution Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Heavy-tailed — API latency (overflow=500)"'
    f' --prop series1=Latency:{HEAVY_CSV}'
    f' --prop binSize=25'
    f' --prop overflowBin=500'
    f' --prop fill=C00000'
    f' --prop xAxisTitle="Latency (ms)"'
    f' --prop yAxisTitle="Requests"'
    f'{STYLE}'
    f' --prop x=0 --prop y=38 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 6: Normal (bell curve) — the reference shape. Same BELL series
#   used throughout this file so you can compare the same data with
#   different bin strategies across sheets.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/4-Distribution Gallery" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Normal — bell curve (reference)"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=18'
    f' --prop fill=2F5597'
    f' --prop xAxisTitle="Value"'
    f' --prop yAxisTitle="Count"'
    f'{STYLE}'
    f' --prop x=14 --prop y=38 --prop width=13 --prop height=18')

# ==========================================================================
# Sheet: 5-Parity Knobs
#
# These charts demo the cx-extended-chart knobs that used to exist only in
# regular cChart and are now available on cx histograms (axis scaling,
# axis visibility/line, plot & chart area fill/border, series & title
# shadows, legend overlay & font, data-label number format). Each chart
# highlights ONE new knob combo so you can diff against a plain histogram
# to see exactly what each knob emits.
# ==========================================================================
print("\n--- 5-Parity Knobs ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-Parity Knobs"')

# --------------------------------------------------------------------------
# Chart 1: axismin / axismax / majorunit — lock the value-axis range.
#
# Useful when you want two histograms in a dashboard to share a Y scale,
# or when the auto-scaled axis makes outliers unreadable. Mirrors regular
# cChart's `axismin` / `axismax` / `majorunit` vocabulary exactly.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Parity Knobs" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Locked Y Scale (axismin/max/majorunit)"'
    f' --prop series1=Scores:{BELL_CSV}'
    f' --prop binCount=15'
    f' --prop fill=4472C4'
    f' --prop axismin=0'
    f' --prop axismax=25'
    f' --prop majorunit=5'
    f' --prop x=0 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 2: Series & title drop shadows (series.shadow / title.shadow).
#
# Format: "COLOR-BLUR-ANGLE-DIST-OPACITY". Same vocabulary as regular
# cChart — under the hood both paths call DrawingEffectsHelper.BuildOuterShadow.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Parity Knobs" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Drop Shadows (series.shadow + title.shadow)"'
    f' --prop title.color=2F5597'
    f' --prop title.size=14'
    f' --prop title.bold=true'
    f' --prop "title.shadow=808080-4-45-3-50"'
    f' --prop series1=Sales:{BELL_CSV}'
    f' --prop binCount=15'
    f' --prop fill=ED7D31'
    f' --prop "series.shadow=000000-6-45-3-35"'
    f' --prop x=14 --prop y=0 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 3: plot area + chart area fill & border (plotareafill / chartareafill
# / plotarea.border / chartarea.border). Polished "card" look useful for
# embedding charts into reports.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Parity Knobs" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Card Look (plotarea + chartarea fill/border)"'
    f' --prop title.color=333333'
    f' --prop title.size=13'
    f' --prop series1=Latency:{HEAVY_CSV}'
    f' --prop binCount=18'
    f' --prop fill=4472C4'
    f' --prop plotareafill=FAFBFC'
    f' --prop "plotarea.border=D0D7DE:1"'
    f' --prop chartareafill=FFFFFF'
    f' --prop "chartarea.border=E5E5E5:0.75"'
    f' --prop overflowBin=500'
    f' --prop x=0 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 4: Axis line styling + visibility (axisline / valaxis.visible).
#
# axisline accepts "color" / "color:width" / "color:width:dash" / "none".
# valaxis.visible=false hides ONLY the value axis (cataxis.visible hides
# only the category axis; axis.visible hides both).
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Parity Knobs" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Minimal Axes (axisline + valaxis.visible=false)"'
    f' --prop series1=Events:{UNIFORM_CSV}'
    f' --prop binSize=10'
    f' --prop fill=70AD47'
    f' --prop "axisline=888888:1.25"'
    f' --prop valaxis.visible=false'
    f' --prop gridlines=false'
    f' --prop x=14 --prop y=19 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 5: Data-label number format + overlaid legend with styled font
# (datalabels.numfmt / legend.overlay / legendfont).
#
# datalabels.numfmt uses Excel's standard format codes ("0", "0.0",
# "0.00%", "#,##0"). legendfont uses the compound "size:color:fontname"
# form shared with axisfont. legend.overlay=true lets the legend float
# on top of the plot area (useful when chart space is tight).
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Parity Knobs" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Label Format + Overlaid Legend"'
    f' --prop title.color=C00000'
    f' --prop title.size=13'
    f' --prop title.bold=true'
    f' --prop series1=Revenue:{BELL_CSV}'
    f' --prop binCount=10'
    f' --prop fill=C00000'
    f' --prop dataLabels=true'
    f' --prop "datalabels.numfmt=0"'
    f' --prop legend=top'
    f' --prop legend.overlay=true'
    f' --prop "legendfont=10:555555:Helvetica Neue"'
    f' --prop x=0 --prop y=38 --prop width=13 --prop height=18')

# --------------------------------------------------------------------------
# Chart 6: Everything — all new knobs stacked onto one presentation-grade
# "magazine hero" chart. This is what the cx-histogram pipeline can do
# end-to-end after the parity port.
# --------------------------------------------------------------------------
cli(f'add "{FILE}" "/5-Parity Knobs" --type chart'
    f' --prop chartType=histogram'
    f' --prop title="Everything Combined"'
    f' --prop title.color=1F2937'
    f' --prop title.size=16'
    f' --prop title.bold=true'
    f' --prop "title.font=Helvetica Neue"'
    f' --prop "title.shadow=808080-4-45-3-40"'
    f' --prop series1=Samples:{BELL_CSV}'
    f' --prop binCount=20'
    f' --prop fill=2F5597'
    f' --prop "series.shadow=000000-5-45-3-30"'
    f' --prop axismin=0'
    f' --prop axismax=25'
    f' --prop majorunit=5'
    f' --prop xAxisTitle="Score" --prop yAxisTitle="Count"'
    f' --prop axisTitle.color=4B5563'
    f' --prop axisTitle.size=11'
    f' --prop axisTitle.bold=true'
    f' --prop "axisfont=9:6B7280:Helvetica Neue"'
    f' --prop "axisline=6B7280:1"'
    f' --prop gridlineColor=EEEEEE'
    f' --prop plotareafill=FAFBFC'
    f' --prop "plotarea.border=D0D7DE:1"'
    f' --prop chartareafill=FFFFFF'
    f' --prop "chartarea.border=E5E5E5:0.75"'
    f' --prop dataLabels=true'
    f' --prop "datalabels.numfmt=0"'
    f' --prop legend=top'
    f' --prop "legendfont=9:6B7280:Helvetica Neue"'
    f' --prop x=14 --prop y=38 --prop width=13 --prop height=18')

print(f"\nDone! Generated: {FILE}")
print("  5 sheets, 22 histograms total")
print("  Sheet 1 (1-Binning Modes):        auto / binCount / binSize / underflow+overflow")
print("  Sheet 2 (2-Styled Histograms):    fill+legend / xGridlineColor / gapWidth+bold / minimal")
print("  Sheet 3 (3-Typography & Colors):  title.* / axisTitle.* / axisfont / all-combined")
print("  Sheet 4 (4-Distribution Gallery): bimodal / right-skew / uniform / left-skew / heavy-tail / normal")
print("  Sheet 5 (5-Parity Knobs):         axis scaling / shadows / area fill+border / axisline / datalabels.numfmt / everything")
