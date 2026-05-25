// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

/// <summary>
/// Phase 2 of the trackChange redesign — automatic OLD-value capture on
/// <c>set + trackChange.*</c>. This file is the single home for
/// rPrChange / pPrChange capture on the Set path; do NOT inline equivalent
/// logic into the per-element Set branches.
///
/// Wiring contract (see <see cref="WordHandler.Set(string, Dictionary{string, string})"/>):
///   1. resolve the target element
///   2. call <see cref="BeginTrackChangeIfRequested"/>, replace the user's
///      properties dict with the returned (stripped) dict before dispatching
///      to SetElement
///   3. after SetElement returns successfully, invoke the returned
///      <c>wrapAction</c> (no-op when trackChange.* was not present)
///
/// Scope:
///   - Run (rPrChange) and Paragraph (pPrChange) only — tcPr/trPr/tblPr/sectPr
///     PrChange are Phase 5 work and are explicitly rejected here.
///   - RTL cascade keys (font.cs / bold.cs / italic.cs / size.cs) are
///     rejected when combined with trackChange.* — the cascade writes
///     across runs and would smear the rPrChange snapshot.
///   - An element that already carries a pending rPrChange/pPrChange is
///     rejected (caller must accept/reject the existing one first).
/// </summary>
public partial class WordHandler
{
    /// <summary>
    /// Sentinel for the no-op case (no trackChange.* keys in the input).
    /// Returning the same Dictionary instance lets the caller short-circuit
    /// without allocating a copy.
    /// </summary>
    private static readonly Action _trackChangeNoop = static () => { };

    /// <summary>
    /// Inspect <paramref name="properties"/> for trackChange.* keys. When
    /// present, snapshot the element's current rPr/pPr clone and return:
    ///   - <c>stripped</c>: copy of <paramref name="properties"/> with all
    ///     trackChange.* keys removed (so downstream Set helpers don't see
    ///     them and don't surface them as unsupported);
    ///   - <c>wrapAction</c>: builds and appends the rPrChange/pPrChange
    ///     (containing the snapshot) to the now-mutated rPr/pPr. The caller
    ///     must invoke this AFTER the Set succeeds.
    ///
    /// When no trackChange.* keys are present, returns the input dict
    /// unchanged and a no-op action (cheap fast path).
    /// </summary>
    private (Dictionary<string, string> stripped, Action wrapAction)
        BeginTrackChangeIfRequested(OpenXmlElement element, Dictionary<string, string> properties)
    {
        if (!HasTrackChangeKey(properties))
            return (properties, _trackChangeNoop);

        // ---- guard 1: only Run and Paragraph supported ----
        if (element is not Run && element is not Paragraph)
            throw new InvalidOperationException(
                "trackChange capture on set is only supported for run and paragraph elements "
                + "(rPrChange / pPrChange); tcPr/trPr/tblPr/sectPr PrChange are not yet implemented.");

        // ---- guard 2: RTL cascade props would smear the snapshot ----
        foreach (var k in properties.Keys)
        {
            var lk = k.ToLowerInvariant();
            if (lk is "font.cs" or "font.complexscript" or "font.complex"
                  or "bold.cs" or "italic.cs" or "size.cs"
                  or "font.bold.cs" or "font.italic.cs" or "font.size.cs"
                  or "boldcs" or "italiccs" or "sizecs")
                throw new InvalidOperationException(
                    "RTL cascade properties are not supported with trackChange yet");
        }

        // ---- extract trackChange.* sub-keys (case-insensitive) ----
        string? tcAuthor = null, tcDate = null, tcId = null;
        foreach (var (k, v) in properties)
        {
            var lk = k.ToLowerInvariant();
            if (lk == "trackchange.author") tcAuthor = v;
            else if (lk == "trackchange.date") tcDate = v;
            else if (lk == "trackchange.id") tcId = v;
        }
        // Defaults per Phase-2 spec.
        var author = string.IsNullOrEmpty(tcAuthor) ? "OfficeCLI" : tcAuthor!;
        DateTime date = DateTime.UtcNow;
        if (!string.IsNullOrEmpty(tcDate) && DateTime.TryParse(tcDate, out var parsed))
            date = parsed;
        var idStr = string.IsNullOrEmpty(tcId) ? GenerateRevisionId() : tcId!;

        // ---- strip trackChange.* from the dict passed to SetElement ----
        var stripped = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in properties)
        {
            var lk = k.ToLowerInvariant();
            if (lk == "trackchange" || lk.StartsWith("trackchange.", StringComparison.Ordinal))
                continue;
            stripped[k] = v;
        }

        // ---- snapshot + plan the wrap based on element kind ----
        if (element is Run run)
        {
            // Reject when an unresolved rPrChange is already pending: writing
            // a second one would either overwrite the first (silent loss of
            // history) or stack two snapshots that point at unrelated state.
            var existingRPr = run.GetFirstChild<RunProperties>();
            if (existingRPr?.GetFirstChild<RunPropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending rPrChange; accept/reject existing first");

            // Snapshot: deep-clone the current rPr's CHILDREN into a
            // <w:rPr> body, which we will host inside a
            // <w:rPrChange><w:rPr>...</w:rPr></w:rPrChange>. Schema-wise the
            // inner element is the bare <w:rPr> (ECMA-376 §17.13.5.31) — the
            // SDK exposes it via the PreviousRunProperties strongly-typed
            // subclass; using a plain RunProperties round-trips as an
            // unknown sibling element. Use PreviousRunProperties.
            var snapshotInner = new PreviousRunProperties();
            if (existingRPr != null)
            {
                foreach (var child in existingRPr.ChildElements)
                {
                    if (child is RunPropertiesChange) continue;
                    snapshotInner.AppendChild(child.CloneNode(true));
                }
            }

            Action wrap = () =>
            {
                // After SetElement: rPr now holds the new values. Append
                // <w:rPrChange w:id=... w:author=... w:date=...>
                //   <w:rPr>{snapshot}</w:rPr>
                // </w:rPrChange>
                var rPr = run.GetFirstChild<RunProperties>()
                          ?? run.PrependChild(new RunProperties());
                var rprChange = new RunPropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                rprChange.AppendChild(snapshotInner);
                // Schema CT_RPr places rPrChange last; AppendChild is correct.
                rPr.AppendChild(rprChange);
            };
            return (stripped, wrap);
        }
        else
        {
            var para = (Paragraph)element;
            var existingPPr = para.ParagraphProperties;
            if (existingPPr?.GetFirstChild<ParagraphPropertiesChange>() != null)
                throw new InvalidOperationException(
                    "element already has a pending pPrChange; accept/reject existing first");

            // Snapshot the pPr. The strongly-typed child class for the <w:pPr>
            // inside <w:pPrChange> is ParagraphPropertiesExtended in
            // DocumentFormat.OpenXml 3.x — NOT PreviousParagraphProperties
            // (despite the parallel naming with PreviousRunProperties used by
            // rPrChange). Confirmed empirically: writing PreviousParagraphProperties
            // round-trips to ParagraphPropertiesExtended after save+reload,
            // breaking strongly-typed reads. Use ParagraphPropertiesExtended on
            // write so write and read see the same SDK type.
            var previous = new ParagraphPropertiesExtended();
            if (existingPPr != null)
            {
                foreach (var child in existingPPr.ChildElements)
                {
                    if (child is ParagraphPropertiesChange) continue;
                    previous.AppendChild(child.CloneNode(true));
                }
            }

            Action wrap = () =>
            {
                var pPr = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                var pprChange = new ParagraphPropertiesChange
                {
                    Author = author,
                    Date = date,
                    Id = idStr,
                };
                pprChange.AppendChild(previous);
                // Schema CT_PPr places pPrChange last; AppendChild is correct.
                pPr.AppendChild(pprChange);
            };
            return (stripped, wrap);
        }
    }

    private static bool HasTrackChangeKey(Dictionary<string, string> properties)
    {
        foreach (var k in properties.Keys)
        {
            var lk = k.ToLowerInvariant();
            if (lk == "trackchange" || lk.StartsWith("trackchange.", StringComparison.Ordinal))
                return true;
        }
        return false;
    }
}
