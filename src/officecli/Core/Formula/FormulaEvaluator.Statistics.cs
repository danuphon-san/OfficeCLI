// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Statistical distribution wrappers ====================
    //
    // Thin wrappers over the special-function core (FormulaEvaluator.Special-
    // Functions.cs). Each takes the same arguments Excel does; the `cumulative`
    // flag selects CDF vs PDF. These are the normal- / gamma- / chi-squared- /
    // Poisson- / exponential-family functions whose engines are erf and the
    // regularized incomplete gamma integral.

    private static double N(List<object> a, int i, double def = 0) =>
        i < a.Count && a[i] is FormulaResult r ? r.AsNumber() : def;

    private static bool Cumulative(List<object> a, int i) =>
        i < a.Count && a[i] is FormulaResult r && r.AsNumber() != 0;

    // NORM.DIST(x, mean, sd, cumulative) / NORMDIST.
    private static FormulaResult? EvalNormDist(List<object> args)
    {
        if (args.Count < 4) return null;
        double x = N(args, 0), mean = N(args, 1), sd = N(args, 2);
        if (sd <= 0) return FormulaResult.Error("#NUM!");
        double z = (x - mean) / sd;
        return FR(Cumulative(args, 3) ? NormCdf(z) : NormPdf(z) / sd);
    }

    // NORM.S.DIST(z, cumulative).
    private static FormulaResult? EvalNormSDist(List<object> args)
    {
        if (args.Count < 1) return null;
        double z = N(args, 0);
        return FR(Cumulative(args, 1) ? NormCdf(z) : NormPdf(z));
    }

    // CONFIDENCE.NORM(alpha, sd, size) — half-width of the normal CI.
    private static FormulaResult? EvalConfidenceNorm(List<object> args)
    {
        if (args.Count < 3) return null;
        double alpha = N(args, 0), sd = N(args, 1), size = N(args, 2);
        if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) return FormulaResult.Error("#NUM!");
        return FR(InvNormCdf(1 - alpha / 2) * sd / Math.Sqrt(size));
    }

    // ERF(lower, [upper]) — error function, optionally over an interval.
    private static FormulaResult? EvalErf(List<object> args)
    {
        if (args.Count < 1) return null;
        double lower = N(args, 0);
        return args.Count >= 2 ? FR(Erf(N(args, 1)) - Erf(lower)) : FR(Erf(lower));
    }

    // GAMMA(x) — Γ(x); poles at 0 and the negative integers.
    private static FormulaResult? EvalGamma(List<object> args)
    {
        if (args.Count < 1) return null;
        double x = N(args, 0);
        if (x <= 0 && x == Math.Floor(x)) return FormulaResult.Error("#NUM!");
        var g = Gamma(x);
        return double.IsNaN(g) ? FormulaResult.Error("#NUM!") : FR(g);
    }

    // GAMMA.DIST(x, alpha, beta, cumulative) / GAMMADIST.
    private static FormulaResult? EvalGammaDist(List<object> args)
    {
        if (args.Count < 4) return null;
        double x = N(args, 0), alpha = N(args, 1), beta = N(args, 2);
        if (x < 0 || alpha <= 0 || beta <= 0) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 3)) return FR(RegGammaP(alpha, x / beta));
        double pdf = Math.Exp((alpha - 1) * Math.Log(x) - x / beta - alpha * Math.Log(beta) - GammaLn(alpha));
        return FR(pdf);
    }

    // CHISQ.DIST(x, df, cumulative) — chi-squared = gamma(df/2, 2).
    private static FormulaResult? EvalChisqDist(List<object> args)
    {
        if (args.Count < 3) return null;
        double x = N(args, 0), df = N(args, 1);
        if (x < 0 || df < 1) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 2)) return FR(RegGammaP(df / 2, x / 2));
        double pdf = Math.Exp((df / 2 - 1) * Math.Log(x) - x / 2 - (df / 2) * Math.Log(2) - GammaLn(df / 2));
        return FR(pdf);
    }

    // POISSON.DIST(x, mean, cumulative) / POISSON.
    private static FormulaResult? EvalPoisson(List<object> args)
    {
        if (args.Count < 3) return null;
        double k = Math.Floor(N(args, 0)), mean = N(args, 1);
        if (k < 0 || mean < 0) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 2)) return FR(RegGammaQ(k + 1, mean));   // CDF P(X≤k)
        return FR(Math.Exp(-mean + k * Math.Log(mean) - GammaLn(k + 1)));
    }

    // EXPON.DIST(x, lambda, cumulative) / EXPONDIST.
    private static FormulaResult? EvalExpon(List<object> args)
    {
        if (args.Count < 3) return null;
        double x = N(args, 0), lambda = N(args, 1);
        if (x < 0 || lambda <= 0) return FormulaResult.Error("#NUM!");
        return FR(Cumulative(args, 2) ? 1 - Math.Exp(-lambda * x) : lambda * Math.Exp(-lambda * x));
    }
}
