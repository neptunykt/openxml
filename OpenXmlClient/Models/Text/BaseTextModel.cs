using System.ComponentModel;

namespace OpenXmlClient.Models.Text;

public class BaseTextModel
{
    [Description("Replace text")]
    public dynamic Value { get; set; }
}