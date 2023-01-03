using System.ComponentModel;

namespace OpenXmlClient.FormatSettings;

public class RepeatHeadingRowSettings
{
    [Description("Start row position number for repeat heading on each page")]
    public int StartRepeatHeadingRowNumber { get; set; }
    [Description("End row position number for repeat heading on each page")]
    public int EndRepeatHeadingRowNumber { get; set; }
}