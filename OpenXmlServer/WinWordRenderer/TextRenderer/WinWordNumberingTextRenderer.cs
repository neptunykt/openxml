using OpenXmlClient.Classes.Models;

namespace OpenXmlServer.WinWordRenderer.TextRenderer;

public class WinWordNumberingTextRenderer
{
    private readonly WinWordTextService _winWordTextService;

    public WinWordNumberingTextRenderer(WinWordTextService winWordTextService)
    {
        _winWordTextService = winWordTextService;
    }

    
    public void Render(WinWordRenderModel winWordRenderModel)
    {
        if (winWordRenderModel.NumberingTextStorage == null ||
            !winWordRenderModel.NumberingTextStorage.Any())
        {
               
            return;
        }

        foreach (var (textTagKey, runArray) in winWordRenderModel.NumberingTextStorage)
        {
            var paragraph = _winWordTextService.GetParagraphsFromBodyByTagName(textTagKey).FirstOrDefault();
            if (paragraph == null)
            {
                throw new Exception($"LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_TEXT_TAG: {textTagKey}");
            }

            var runList = runArray.Reverse().ToList();
            if (runList.Count == 0)
            {
                // тут удаляем вместе с параграфом если ничего нет
                _winWordTextService.RemoveOpenXmlElement(paragraph);
                continue;
            }
            foreach (var runModels in runList)
            {
                _winWordTextService.CreateNumberingWithRuns(paragraph, runModels, textTagKey);
            }
            // удаляем эталонный список
            _winWordTextService.RemoveOpenXmlElement(paragraph);
        }
           
    }
}