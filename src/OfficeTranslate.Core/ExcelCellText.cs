using System;

namespace OfficeTranslate.Core
{
    public static class ExcelCellText
    {
        public static string Literal(string text)
        {
            if (text.Length > 32767) throw new InvalidOperationException();
            // Excel consumes this prefix on Value2 assignment and keeps the entire value as text.
            return "'" + text;
        }
    }
}
