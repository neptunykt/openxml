using System.Collections.ObjectModel;
using System.ComponentModel;
using OpenXmlClient.Models.Text;

namespace OpenXmlClient.Models.Table;

public class TableFillStorageModel
{
    [Description("table fill storage data, key is place rows tag")]
    public IDictionary<string, TableRowFillModel> TableRowsFillStorage { get; set; }
     
    [Description("table fill storage data, key is replace text tag")]
    public IDictionary<string, TableTextModel> TableFillTextReplaceStorage { get; set; }
   
    [Description("table fill storage data, key is replace text generator tag")]
    public IDictionary<string, Collection<RunModel>> TableFillTextGeneratorReplaceStorage { get; set; }

    public TableFillStorageModel()
    {
        TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
        TableFillTextReplaceStorage = new Dictionary<string, TableTextModel>();
        InnerTableFillStorage = new Dictionary<string, IDictionary<string, InnerTableRowFillModel>>();
        TableFillTextGeneratorReplaceStorage = new Dictionary<string, Collection<RunModel>>();
    }
      
    [Description("table fill storage data, key is inner table tag")]
    public IDictionary<string, IDictionary<string, InnerTableRowFillModel>> InnerTableFillStorage { get; set; }
}