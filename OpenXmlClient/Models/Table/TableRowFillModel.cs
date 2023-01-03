using System.ComponentModel;
using OpenXmlClient.FormatSettings;

namespace OpenXmlClient.Models.Table;

public class TableRowFillModel
{
    public TableValueOutputFormat TableValueOutputFormat
    {
        get;
        set;
    }

    [Description("Ordered column names")]
    public HashSet<string> ColumnNames { get; set; }
      
    [Description("Data for filling rows")]
    public IEnumerable<dynamic> TableData
    {
        get;
        set;
    }

    [Description("Remove inner borders in column absolute position")]
    public ICollection<int> RemoveInnerBordersInColumnAbsolutePositions { get; set; } 
    [Description("Remove entire row if now records")]
    public bool RemoveEntireRowIfNoRecords { get; set; }
}