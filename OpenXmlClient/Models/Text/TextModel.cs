using OpenXmlClient.FormatSettings;

namespace OpenXmlClient.Models.Text;

public class TextModel : BaseTextModel
{
    private OutputFormat _outputFormat { get; set; }

    public TextModel()
    {
            
    }
    public TextModel(OutputFormat outputFormat)
    {
        _outputFormat = outputFormat;
    }
        

    public OutputFormat OutputFormat { get; set; }
}