using System;

namespace OfficeTranslate.Core
{
    // Normalizes Excel COM range arrays for bulk reads. Lives in Core (rather
    // than the Excel add-in) so it can be unit tested without referencing the
    // Office interop assemblies.
    //
    // Background: Range.Value2 / Formula / HasFormula return a 1-based
    // object[,] for multi-cell ranges but a scalar for a single cell, and real
    // Excel COM has been observed returning a scalar DBNull for HasFormula on
    // a multi-cell range. Every method here is total: out-of-range and
    // unexpected shapes degrade to safe defaults instead of throwing.
    internal static class RangeGridHelper
    {
        // Converts a COM value (1-based object[,] or scalar) into a 0-based
        // grid. Callers must use GetLength() with 0-based indexing only.
        internal static object[,] Normalize(object raw)
        {
            if (raw is object[,] grid)
            {
                var r0 = grid.GetLowerBound(0);
                var c0 = grid.GetLowerBound(1);
                var rows = grid.GetUpperBound(0) - r0 + 1;
                var cols = grid.GetUpperBound(1) - c0 + 1;
                if (r0 == 0 && c0 == 0) return grid;
                var normalized = new object[rows, cols];
                for (var r = 0; r < rows; r++)
                    for (var c = 0; c < cols; c++)
                        normalized[r, c] = grid[r0 + r, c0 + c];
                return normalized;
            }
            return new object[,] { { raw } };
        }

        // True when the cell holds a formula. Prefers the per-cell HasFormula
        // flags when COM returned a well-formed grid; otherwise falls back to
        // detecting a leading "=" in the Formula text. Out-of-range indexes
        // return false.
        //
        // Known limitation of the fallback: a text cell literally starting
        // with "=" (typed with a leading apostrophe) is indistinguishable from
        // a formula here and will be skipped. Such cells are rare, and skipping
        // (fail-safe) beats mistranslating a formula's text.
        internal static bool IsFormulaCell(object[,] formulaFlags, object[,] formulas, int r, int c)
        {
            if (r >= 0 && c >= 0 &&
                r < formulaFlags.GetLength(0) && c < formulaFlags.GetLength(1) &&
                formulaFlags[r, c] is bool isFormula)
                return isFormula;
            if (r >= 0 && c >= 0 &&
                r < formulas.GetLength(0) && c < formulas.GetLength(1))
                return formulas[r, c] is string text &&
                    text.StartsWith("=", StringComparison.Ordinal);
            return false;
        }
    }
}
