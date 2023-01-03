using OpenXmlClient.Classes.Models;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models;

namespace OpenXmlServer.WinWordRenderer.TextRenderer;

public class WinWordGlobalTextRenderer
{
    private readonly WinWordTextService _winWordTextService;
    public WinWordGlobalTextRenderer(WinWordTextService winWordTextService)
    {
        _winWordTextService = winWordTextService;
    }
    public void Render(WinWordRenderModel winWordRenderModel)
    {
        if (winWordRenderModel.GlobalTextReplacementStorage == null ||
            !winWordRenderModel.GlobalTextReplacementStorage.Any())
        {
            return;
        }

        foreach (var (textTagKey, textModel) in winWordRenderModel.GlobalTextReplacementStorage)
        {
            if (textModel.OutputFormat == null)
            {
                textModel.OutputFormat = new OutputFormat();
            }

            if (textModel.Value == null)
            {
                // удаляем вместе с параграфом
                var paragraph = _winWordTextService.GetParagraphsFromBodyByTagName(textTagKey).FirstOrDefault();
                if (paragraph == null)
                {
                    throw new Exception(
                        $"LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_TEXT_TAG: {textTagKey}");
                }
                _winWordTextService.RemoveOpenXmlElement(paragraph);
            }
            OutputFormat.SetOutputFormat(WinWordRenderModel.OutputFormat, textModel.OutputFormat);
            var textValue = WinWordTextService.FillValue(textModel.Value,textModel.OutputFormat);
            _winWordTextService.GlobalReplaceText(textTagKey, textValue);
        }
    }
        
           
}