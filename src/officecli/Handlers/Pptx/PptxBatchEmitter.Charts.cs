// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(chart-data-string): mirrors WordBatchEmitter.Charts.cs —
    // emit a semantic `data="Name1:v1,v2;Name2:v3,v4"` string reconstructed
    // from series children that AddChart re-builds at replay. The embedded
    // xlsx (ppt/embeddings/Microsoft_Excel_Worksheet.xlsx) is lossy on
    // round-trip: formulas, conditional formatting, defined names from the
    // source workbook are dropped. Same trade-off as docx — chart visual
    // round-trips, chart workbook does not.
    private static void EmitChart(PowerPointHandler ppt, DocumentNode chartNode,
                                  string parentSlidePath, List<BatchItem> items,
                                  SlideEmitContext ctx)
    {
        // depth=1 so series children materialize with their name/values.
        var fullChart = ppt.Get(chartNode.Path, depth: 1);
        var props = FilterEmittableProps(fullChart.Format);
        // Strip Get-only keys AddChart neither expects nor accepts.
        props.Remove("id");
        props.Remove("seriesCount");

        // Scatter/bubble charts intrinsically carry TWO c:valAx (X and Y are
        // both value-axes — no category axis), which the Reader's
        // "multi-valAx ⇒ secondary axis" heuristic mistakes for a combo
        // primary+secondary pair. Re-emitting `secondaryAxis=1,2` on a scatter
        // forces ApplySecondaryAxis at replay, which retags one series's axIds
        // and (because plotArea now contains two ScatterCharts with disjoint
        // series binds) gets detected as `chartType=combo` on the next Get.
        // Drop the spurious key for these chart types — primary/secondary is
        // not a meaningful concept on a scatter or bubble plot.
        if (props.TryGetValue("chartType", out var chartTypeStr)
            && (chartTypeStr.Equals("scatter", StringComparison.OrdinalIgnoreCase)
                || chartTypeStr.Equals("bubble", StringComparison.OrdinalIgnoreCase)))
        {
            props.Remove("secondaryAxis");
        }

        // Reconstruct AddChart's data="Name1:v1,v2;..." input from the
        // series children (each carries `name` + `values` Format keys).
        var seriesParts = new List<string>();
        if (fullChart.Children != null)
        {
            foreach (var s in fullChart.Children)
            {
                if (s.Type != "series") continue;
                if (!s.Format.TryGetValue("name", out var nObj) || nObj == null) continue;
                if (!s.Format.TryGetValue("values", out var vObj) || vObj == null) continue;
                var name = nObj.ToString() ?? "";
                var vals = vObj.ToString() ?? "";
                if (name.Length == 0 || vals.Length == 0) continue;
                seriesParts.Add($"{name}:{vals}");
            }
        }

        // Waterfall round-trip: BuildWaterfallChart encodes the user's delta
        // input into 3 stacked-bar series (Base/Increase/Decrease) with
        // cumulative values. Re-feeding those 3 series on replay doubles the
        // running total — Builder would re-encode the already-encoded data.
        // Reverse the encoding here: each category's delta is `inc[i]` if
        // inc[i] != 0 else `-dec[i]`; emit a single series under the chart's
        // own name (or "Series 1") so AddChart's waterfall path takes over.
        // Per-category names are recovered from `categories=`.
        if (props.TryGetValue("chartType", out var ctype)
            && ctype.Equals("waterfall", StringComparison.OrdinalIgnoreCase)
            && fullChart.Children != null)
        {
            var byName = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in fullChart.Children)
            {
                if (s.Type != "series") continue;
                if (!s.Format.TryGetValue("name", out var nObj)) continue;
                if (!s.Format.TryGetValue("values", out var vObj)) continue;
                var nm = nObj?.ToString() ?? "";
                var vs = (vObj?.ToString() ?? "")
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => double.TryParse(t.Trim(),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var d) ? d : 0.0)
                    .ToArray();
                byName[nm] = vs;
            }
            if (byName.TryGetValue("Increase", out var inc)
                && byName.TryGetValue("Decrease", out var dec)
                && inc.Length == dec.Length
                && inc.Length > 0)
            {
                var deltas = new double[inc.Length];
                for (int i = 0; i < inc.Length; i++)
                    deltas[i] = inc[i] != 0 ? inc[i] : -dec[i];
                var deltaStr = string.Join(",",
                    deltas.Select(d => d.ToString("G",
                        System.Globalization.CultureInfo.InvariantCulture)));
                seriesParts = new List<string> { $"Waterfall:{deltaStr}" };
                // Strip per-series color/style props that referred to the
                // encoded triplet — Builder re-applies increase/decrease/
                // total colors from explicit chart-level keys.
            }
        }

        if (seriesParts.Count > 0)
            props["data"] = string.Join(";", seriesParts);

        // Per-series style round-trip: NodeBuilder emits color/lineWidth/
        // lineDash/marker/smooth on each series child Format, but the chart-
        // level `add` step has no series sub-nodes to attach those to. The
        // chart Setter accepts dotted per-series keys (series{N}.color,
        // series{N}.lineWidth, ...) — re-flatten them here so a chart with
        // an explicit series color (#C00000 darkred etc.) round-trips
        // instead of falling back to the DefaultSeriesColors palette.
        // Skipped for waterfall (handled via increase/decrease/totalColor
        // chart-level props; emitting series1.color=transparent for "Base"
        // would fight Builder's NoFill encoding).
        var isWaterfall = props.TryGetValue("chartType", out var ctForSeries)
            && ctForSeries.Equals("waterfall", StringComparison.OrdinalIgnoreCase);
        if (!isWaterfall && fullChart.Children != null)
        {
            int seriesIdx = 0;
            foreach (var s in fullChart.Children)
            {
                if (s.Type != "series") continue;
                seriesIdx++;
                foreach (var key in new[] { "color", "lineWidth", "lineDash",
                    "marker", "markerSize", "smooth", "outlineColor",
                    "transparency" })
                {
                    if (s.Format.TryGetValue(key, out var val) && val != null)
                    {
                        var sval = val.ToString();
                        if (!string.IsNullOrEmpty(sval))
                            props[$"series{seriesIdx}.{key}"] = sval;
                    }
                }
            }
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "chart",
            Props = props.Count > 0 ? props : null,
        });
    }
}
