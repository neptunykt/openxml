namespace OpenXmlClient.FormatSettings;

public class TableValueOutputFormat
{
    public string DateTimeFormat { get; set; }
    public string DecimalFormat { get; set; }

    public static void SetOutputFormat(TableValueOutputFormat parentOutputFormat,
        TableValueOutputFormat childOutputFormat)
    {
        if (parentOutputFormat == null)
        {
            return;
        }
            
        childOutputFormat.DecimalFormat = childOutputFormat.DecimalFormat ?? parentOutputFormat.DecimalFormat;
        childOutputFormat.DateTimeFormat = childOutputFormat.DateTimeFormat ?? parentOutputFormat.DateTimeFormat;
    }
}