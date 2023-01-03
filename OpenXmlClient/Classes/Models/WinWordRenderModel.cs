using System.Collections.ObjectModel;
using System.ComponentModel;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.HeaderFooter;
using OpenXmlClient.Models.Text;

#pragma warning disable CS8618

namespace OpenXmlClient.Classes.Models;

public class WinWordRenderModel
{
    // Tables storage for render
        
    public IDictionary<string, TableRenderModel> TablesStorage { get; set; }
    public IDictionary<string, ClonedTableRenderModel> ClonedTablesStorage { get; set; }
    
    public IDictionary<string, ICollection<Collection<RunModel>>> NumberingTextStorage { get; set; }
    public IDictionary<string, TextModel> GlobalTextReplacementStorage { get; set; }

    [Description("Text generator storage")]
    public IDictionary<string, IEnumerable<RunModel>> GlobalTextGeneratorStorage { get; set; }
        
    public HeaderFooterStorage HeaderFooterStorage { get; set; }
        
        
    public static OutputFormat OutputFormat { get; set; }
        
    public static TableValueOutputFormat TableValueOutputFormat { get; set; }
}