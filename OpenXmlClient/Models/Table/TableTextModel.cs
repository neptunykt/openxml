using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.Text;

namespace OpenXmlClient.Models.Table;

public class TableTextModel : BaseTextModel
{
    public TableTextModel()
    {
    }

    public TableTextModel(TableValueOutputFormat tableValueOutputFormat)
    {
        TableValueOutputFormat = tableValueOutputFormat;
    }


    public TableValueOutputFormat TableValueOutputFormat { get; set; }

    public bool RemoveEntireRowIfNoValue { get; set; }
}
