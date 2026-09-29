using System.Collections.Generic;
using System.Text;

namespace OfficeTranslate.Core
{
// R2: one image gets ONE combined side-note. Previously every degraded
// region created its own note at the same position and each creation
// deleted the previous one, so only the last region's translation
// survived. Hosts collect one entry per degraded region
// (reason + "\r" + translation) and call Combine once.
public static class ImageOverlayNotes
{
public const string Header = "";

public static string Combine(IReadOnlyList<string> entries)
{
var builder = new StringBuilder(Header);
if (entries!= null)
{
foreach (var entry in entries)
{
builder.Append("\r\r");
builder.Append(entry);
}
}
return builder.ToString();
}
}
}
