// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeCli.Core;

/// <summary>
/// Builder for cx:chart (Office 2016 extended chart types):
/// funnel, treemap, sunburst, boxWhisker, histogram, waterfall (native).
/// </summary>
internal static class ChartExBuilder
{
    internal static readonly HashSet<string> ExtendedChartTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "funnel", "treemap", "sunburst", "boxwhisker", "histogram"
    };

    internal static bool IsExtendedChartType(string chartType)
    {
        var normalized = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
        return ExtendedChartTypes.Contains(normalized);
    }

    /// <summary>
    /// Build a cx:chartSpace for an extended chart type.
    /// </summary>
    internal static CX.ChartSpace BuildExtendedChartSpace(
        string chartType,
        string? title,
        string[]? categories,
        List<(string name, double[] values)> seriesData,
        Dictionary<string, string> properties)
    {
        var normalized = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");

        var chartSpace = new CX.ChartSpace();

        // 1. Build ChartData
        var chartData = new CX.ChartData();
        for (int si = 0; si < seriesData.Count; si++)
        {
            var data = BuildDataBlock((uint)si, normalized, categories, seriesData[si].values);
            chartData.AppendChild(data);
        }
        chartSpace.AppendChild(chartData);

        // 2. Build Chart
        var chart = new CX.Chart();

        if (!string.IsNullOrEmpty(title))
        {
            var chartTitle = new CX.ChartTitle();
            chartTitle.AppendChild(new CX.Text(
                new CX.RichTextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US" },
                            new Drawing.Text(title))))));
            chart.AppendChild(chartTitle);
        }

        var plotArea = new CX.PlotArea();
        var plotAreaRegion = new CX.PlotAreaRegion();

        var layoutId = normalized switch
        {
            "funnel" => "funnel",
            "treemap" => "treemap",
            "sunburst" => "sunburst",
            "boxwhisker" => "boxWhisker",
            "histogram" => "clusteredColumn",
            _ => "funnel"
        };

        for (int si = 0; si < seriesData.Count; si++)
        {
            var series = new CX.Series { LayoutId = new EnumValue<CX.SeriesLayout>(
                ParseSeriesLayout(layoutId)) };
            series.AppendChild(new CX.Text(
                new CX.TextData(
                    new CX.Formula(""),
                    new CX.VXsdstring(seriesData[si].name))));
            series.AppendChild(new CX.DataId { Val = (uint)si });

            // Chart-type specific layoutPr
            var layoutPr = BuildLayoutProperties(normalized, properties, seriesData[si].values.Length);
            if (layoutPr != null)
                series.AppendChild(layoutPr);

            plotAreaRegion.AppendChild(series);
        }

        plotArea.AppendChild(plotAreaRegion);

        // Add axes for chart types that need them
        if (normalized is "boxwhisker" or "histogram")
        {
            plotArea.AppendChild(new CX.Axis(new CX.CategoryAxisScaling()) { Id = 0 });
            plotArea.AppendChild(new CX.Axis(new CX.ValueAxisScaling()) { Id = 1 });
        }

        chart.AppendChild(plotArea);
        chartSpace.AppendChild(chart);

        return chartSpace;
    }

    private static CX.Data BuildDataBlock(uint id, string chartType, string[]? categories, double[] values)
    {
        var data = new CX.Data { Id = id };

        // String dimension for categories (if provided)
        if (categories != null && chartType is "funnel" or "treemap" or "sunburst" or "boxwhisker")
        {
            var strDim = new CX.StringDimension { Type = CX.StringDimensionType.Cat };
            var strLvl = new CX.StringLevel { PtCount = (uint)categories.Length };
            for (int i = 0; i < categories.Length; i++)
                strLvl.AppendChild(new CX.ChartStringValue(categories[i]) { Index = (uint)i });
            strDim.AppendChild(strLvl);
            data.AppendChild(strDim);
        }

        // Numeric dimension
        var numType = chartType is "treemap" or "sunburst"
            ? CX.NumericDimensionType.Size
            : CX.NumericDimensionType.Val;
        var numDim = new CX.NumericDimension { Type = numType };
        var numLvl = new CX.NumericLevel { PtCount = (uint)values.Length, FormatCode = "General" };
        for (int i = 0; i < values.Length; i++)
            numLvl.AppendChild(new CX.NumericValue(values[i].ToString("G")) { Idx = (uint)i });
        numDim.AppendChild(numLvl);
        data.AppendChild(numDim);

        return data;
    }

    private static CX.SeriesLayoutProperties? BuildLayoutProperties(
        string chartType, Dictionary<string, string> properties, int valueCount)
    {
        switch (chartType)
        {
            case "treemap":
            {
                var lp = new CX.SeriesLayoutProperties();
                var parentLayout = properties.GetValueOrDefault("parentLabelLayout") ?? "overlapping";
                lp.AppendChild(new CX.ParentLabelLayout
                {
                    ParentLabelLayoutVal = parentLayout.ToLowerInvariant() switch
                    {
                        "none" => CX.ParentLabelLayoutVal.None,
                        "banner" => CX.ParentLabelLayoutVal.Banner,
                        _ => CX.ParentLabelLayoutVal.Overlapping
                    }
                });
                return lp;
            }
            case "boxwhisker":
            {
                var lp = new CX.SeriesLayoutProperties();
                lp.AppendChild(new CX.SeriesElementVisibilities
                {
                    MeanLine = false, MeanMarker = true,
                    Nonoutliers = false, Outliers = true
                });
                var method = properties.GetValueOrDefault("quartileMethod") ?? "exclusive";
                lp.AppendChild(new CX.Statistics
                {
                    QuartileMethod = method.ToLowerInvariant() switch
                    {
                        "inclusive" => CX.QuartileMethod.Inclusive,
                        _ => CX.QuartileMethod.Exclusive
                    }
                });
                return lp;
            }
            case "histogram":
            {
                // cx:layoutPr > cx:binning (empty for auto-bin; child cx:binCount
                // OR cx:binSize for explicit bin count/width). `cx:aggregation`
                // is for Pareto charts and causes Excel to render the whole
                // dataset as a single bar.
                //
                // NOTE: the Open XML SDK models cx:binCount as a leaf text
                // element (BinCountXsdunsignedInt → `<cx:binCount>5</cx:binCount>`),
                // but real Excel writes it as an empty element with a `val`
                // attribute (`<cx:binCount val="5"/>`). SDK's form is schema-
                // valid per the generated type metadata but Excel rejects the
                // whole file with "We found a problem with some content"
                // and deletes the drawing. Same applies to cx:binSize. Work
                // around by appending a raw OpenXmlUnknownElement carrying
                // the correct form.
                const string cxNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
                var lp = new CX.SeriesLayoutProperties();
                var binning = new CX.Binning();

                // intervalClosed: "r" (default, bins are (a,b]) or "l" (bins are [a,b))
                var intervalClosed = properties.GetValueOrDefault("intervalClosed") ?? "r";
                binning.IntervalClosed = intervalClosed.ToLowerInvariant() switch
                {
                    "l" => CX.IntervalClosedSide.L,
                    _   => CX.IntervalClosedSide.R,
                };

                // underflow / overflow: cut-off values for outlier bins
                if (properties.TryGetValue("underflowBin", out var underflow))
                    binning.Underflow = underflow;
                if (properties.TryGetValue("overflowBin", out var overflow))
                    binning.Overflow = overflow;

                // binCount (explicit count) XOR binSize (explicit width). If
                // both are given, binCount wins (it's the more common knob).
                if (properties.TryGetValue("binCount", out var binCountStr) &&
                    uint.TryParse(binCountStr, out var binCount))
                {
                    var binCountEl = new OpenXmlUnknownElement("cx", "binCount", cxNs);
                    binCountEl.SetAttribute(new OpenXmlAttribute("val", "", binCount.ToString()));
                    binning.AppendChild(binCountEl);
                }
                else if (properties.TryGetValue("binSize", out var binSizeStr) &&
                         double.TryParse(binSizeStr, System.Globalization.NumberStyles.Float,
                             System.Globalization.CultureInfo.InvariantCulture, out var binSize))
                {
                    var binSizeEl = new OpenXmlUnknownElement("cx", "binSize", cxNs);
                    binSizeEl.SetAttribute(new OpenXmlAttribute("val", "",
                        binSize.ToString("G", System.Globalization.CultureInfo.InvariantCulture)));
                    binning.AppendChild(binSizeEl);
                }

                lp.AppendChild(binning);
                return lp;
            }
            default:
                return null;
        }
    }

    private static CX.SeriesLayout ParseSeriesLayout(string layoutId)
    {
        return layoutId switch
        {
            "funnel" => CX.SeriesLayout.Funnel,
            "treemap" => CX.SeriesLayout.Treemap,
            "sunburst" => CX.SeriesLayout.Sunburst,
            "boxWhisker" => CX.SeriesLayout.BoxWhisker,
            "clusteredColumn" => CX.SeriesLayout.ClusteredColumn,
            "paretoLine" => CX.SeriesLayout.ParetoLine,
            "regionMap" => CX.SeriesLayout.RegionMap,
            _ => CX.SeriesLayout.Funnel
        };
    }

    /// <summary>
    /// Detect if a cx:chartSpace contains an extended chart type and return the type name.
    /// </summary>
    internal static string? DetectExtendedChartType(CX.ChartSpace chartSpace)
    {
        var series = chartSpace.Descendants<CX.Series>().FirstOrDefault();
        var layoutId = series?.LayoutId?.InnerText;
        if (layoutId == null) return null;
        return layoutId switch
        {
            "funnel" => "funnel",
            "treemap" => "treemap",
            "sunburst" => "sunburst",
            "boxWhisker" => "boxWhisker",
            "clusteredColumn" => "histogram",
            "paretoLine" => "pareto",
            "regionMap" => "regionMap",
            _ => layoutId
        };
    }
}
