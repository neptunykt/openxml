using OpenXmlClient.Models.Text;

namespace OpenXmlClient.Models.HeaderFooter;

public class HeaderFooterStorage
{
    public IDictionary<string,TextModel> HeaderReplaceDictionary { get; set; } 
    public IDictionary<string,TextModel> FooterReplaceDictionary { get; set; } 
}