using System.ComponentModel;
using OpenXmlClient.FormatSettings;

namespace OpenXmlClient.Classes.Models;

public class RowsRenderPayload
{
    [Description("XML Serialized DataTable")]
    public string Payload { get; set; }
        
    [Description("Column names ordered for output")]
    public ICollection<string> ColumnNames { get; set; }
        
    [Description("Value output Format data")]
    public TableValueOutputFormat TableValueOutputFormat { get; set; } 
        
    [Description("Remove inner borders in column absolute positions")]
    public ICollection<int> RemoveInnerBordersInColumnAbsolutePositions { get; set; }
        
    [Description("Remove entire row if no records")]
    public bool IsRemoveEntireRowIfNoRecords { get; set; }

}