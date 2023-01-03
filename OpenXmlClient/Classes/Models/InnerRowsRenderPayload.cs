using OpenXmlClient.Models.Table;

namespace OpenXmlClient.Classes.Models;

public class InnerRowsRenderPayload : RowsRenderPayload
{
    public bool IsRemoveInnerTableIfNoRecords { get; set; }
    public bool IsRemoveEntireOuterRowIfNoRecords { get; set; }
}