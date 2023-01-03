namespace OpenXmlClient.Models.Table;

public class InnerTableRowFillModel : TableRowFillModel
{
    public bool IsRemoveInnerTableIfNoRecords { get; set; }
    public bool IsRemoveEntireOuterRowIfNoRecords { get; set; }
}