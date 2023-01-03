using OpenXmlClient.FormatSettings;

namespace OpenXmlClient.Models.Text;

public class RunModel : BaseTextModel
{
    private OutputFormat _outputFormat { get; set; }

    public bool IsCarriageReturn { get; set; }
        
    public bool IsTabulation { get; set; }
    
    public bool IsSymbol { get; set; }
        

    public RunModel()
    {
    }

    public RunModel(OutputFormat outputFormat)
    {
        _outputFormat = outputFormat;
    }
    

    public OutputFormat OutputFormat { get; set; }
       

    public void SetOutputFormat(OutputFormat outputFormat) => _outputFormat = outputFormat;
}