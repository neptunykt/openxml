using System.ComponentModel;
using OpenXmlClient.FormatSettings;

namespace OpenXmlClient.Classes.Models;

public class ClonedTableRenderModel
{
    [Description("Tag name place to insert cloned tables")]
    public string InsertTablePlaceTagName { get; set; }

    [Description("Table render data")]
    public TableRenderModel CopyTableRenderStorage { get; set; }    
    public PrintOutClonedTableFormatSettings PrintOutClonedTableFormatSettings { get; set; }
        
    public string ClonedTableTagName { get; set; }

    public OutputFormat OutputFormat { get; set; }
        
    [Description("First key is inner table tag")]
        
    public IDictionary<string, TableRenderModel> InnerTableFillStorage { get; set; }

}