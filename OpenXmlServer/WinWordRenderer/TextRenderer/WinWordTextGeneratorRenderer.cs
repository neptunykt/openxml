using OpenXmlClient.Classes.Models;
using OpenXmlClient.Models;

namespace OpenXmlServer.WinWordRenderer.TextRenderer;

public class WinWordTextGeneratorRenderer
{
    private readonly WinWordTextService _winWordTextService;

    public WinWordTextGeneratorRenderer(WinWordTextService winWordTextService)
    {
        _winWordTextService = winWordTextService;
    }

    public void Render(WinWordRenderModel winWordRenderModel)
    {
        if (winWordRenderModel.GlobalTextGeneratorStorage == null ||
            !winWordRenderModel.GlobalTextGeneratorStorage.Any())
        {
               
            return;
        }

        foreach (var (textTagKey, runModels) in winWordRenderModel.GlobalTextGeneratorStorage)
        {
            var paragraph = _winWordTextService.GetParagraphsFromBodyByTagName(textTagKey).FirstOrDefault();
            if (paragraph == null)
            {
                throw new Exception($"LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_TEXT_TAG: {textTagKey}");
            }

            var runList = runModels.ToList();
            if (runList.Count == 0)
            {
                // тут удаляем вместе с параграфом если ничего нет
                _winWordTextService.RemoveOpenXmlElement(paragraph);
                continue;
            }
            foreach (var runModel in runList)
            {
                _winWordTextService.CreateRunInParagraph(paragraph, runModel, runModel.OutputFormat);
            }
            // удаляем тег если есть
            _winWordTextService.GlobalReplaceText(textTagKey,"");
        }
           
    }
}