namespace OpenXmlClient.FormatSettings;

 public class OutputFormat : TableValueOutputFormat
  {
    public string FontName { get; set; }

    public int FontSize { get; set; }

    public bool? IsBold { get; set; }

    public bool? IsItalic { get; set; }

    public bool? IsSymbol { get; set; }

    public bool? IsUnderLine { get; set; }

    public string FontColor { get; set; }

    public string BackGroundColor { get; set; }

    public static void SetOutputFormat(
      OutputFormat parentOutputFormat,
      OutputFormat childOutputFormat)
    {
      if (parentOutputFormat == null)
        return;
      childOutputFormat.DecimalFormat = childOutputFormat.DecimalFormat ?? parentOutputFormat.DecimalFormat;
      childOutputFormat.FontColor = childOutputFormat.FontColor ?? parentOutputFormat.FontColor;
      childOutputFormat.FontName = childOutputFormat.FontName ?? parentOutputFormat.FontName;
      childOutputFormat.FontSize = childOutputFormat.FontSize != 0 ? childOutputFormat.FontSize : parentOutputFormat.FontSize;
      OutputFormat outputFormat1 = childOutputFormat;
      bool? nullable1 = childOutputFormat.IsBold;
      bool? nullable2 = nullable1 ?? parentOutputFormat.IsBold;
      outputFormat1.IsBold = nullable2;
      OutputFormat outputFormat2 = childOutputFormat;
      nullable1 = childOutputFormat.IsItalic;
      bool? nullable3 = nullable1 ?? parentOutputFormat.IsItalic;
      outputFormat2.IsItalic = nullable3;
      OutputFormat outputFormat3 = childOutputFormat;
      nullable1 = childOutputFormat.IsSymbol;
      bool? nullable4 = nullable1 ?? childOutputFormat.IsSymbol;
      outputFormat3.IsSymbol = nullable4;
      childOutputFormat.BackGroundColor = childOutputFormat.BackGroundColor ?? parentOutputFormat.BackGroundColor;
      childOutputFormat.DateTimeFormat = childOutputFormat.DateTimeFormat ?? parentOutputFormat.DateTimeFormat;
      OutputFormat outputFormat4 = childOutputFormat;
      nullable1 = childOutputFormat.IsUnderLine;
      bool? nullable5 = nullable1 ?? parentOutputFormat.IsUnderLine;
      outputFormat4.IsUnderLine = nullable5;
    }
  }