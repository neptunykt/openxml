using System.Collections.ObjectModel;
using System.ComponentModel;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.Table;
using OpenXmlClient.Models.Text;

namespace OpenXmlClient.Classes.Models;

public class TableRenderModel
{
    [Description("Table rows render data, key is table rows tag name")]
    public IDictionary<string, RowsRenderPayload> TableRowsData { get; set; }
    [Description("Table text replace data by tag name, key is tag name for text replacement")]
    public IDictionary<string, TableTextModel> TextReplaceData { get; set; }
    [Description("Table text generator replace data by tag name, key is tag name for text array replacement")]
    public IDictionary<string, Collection<RunModel>> TextGeneratorReplaceData { get; set; } 
        
    public PrintOutTableFormatSettings PrintOutTableFormatSettings { get; set; }

    public IDictionary<string, IDictionary<string, InnerRowsRenderPayload>> InnerTablesStorage { get; set; }

}