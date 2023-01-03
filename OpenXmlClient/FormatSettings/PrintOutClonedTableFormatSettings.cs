using System.ComponentModel;

namespace OpenXmlClient.FormatSettings;

[Description("Set spaces between tables, repeat heading, page breaking")]
public class PrintOutClonedTableFormatSettings : PrintOutTableFormatSettings
{
    [Description("Page break settings between cloned tables")]
    public bool IsSetPageBreak { get; set; }

    public int NumberOfLinesBetweenPages { get; set; }
}