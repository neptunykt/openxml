using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlClient.Models.HeaderFooter;
using OpenXmlClient.Models.Text;
using OpenXmlServer.WinWordRenderer.TableRenderer;
using OpenXmlServer.WinWordRenderer.TextRenderer;

namespace OpenXmlServer.WinWordRenderer.HeaderFooterRenderer;

    public class WinWordHeaderFooterRenderer
    {
        private readonly WordprocessingDocument _wordprocessingDocument;

        public WinWordHeaderFooterRenderer(WordprocessingDocument wordprocessingDocument)
        {
            _wordprocessingDocument = wordprocessingDocument;
        }


        /// <summary>
        /// Method for replacing by tag
        /// </summary>
        /// <param name="headerFooterStorage"></param>
        public void Render(HeaderFooterStorage headerFooterStorage)
        {
            IEnumerable<OpenXmlPart> headerParts = null;
            IEnumerable<OpenXmlPart> footerParts = null;
            if (headerFooterStorage == null)
            {
                return;
            }

            if (headerFooterStorage.HeaderReplaceDictionary?.Count > 0)
            {
                headerParts = _wordprocessingDocument?.MainDocumentPart?.HeaderParts;
            }
            else if (headerFooterStorage.FooterReplaceDictionary?.Count > 0)
            {
                footerParts = _wordprocessingDocument?.MainDocumentPart?.FooterParts;
            }

            ReplaceParts(headerParts, headerFooterStorage.HeaderReplaceDictionary);
            ReplaceParts(footerParts, headerFooterStorage.FooterReplaceDictionary);
        }

        /// <summary>
        /// Replace in Header jr footer
        /// </summary>
        /// <param name="parts"></param>
        /// <param name="replaceStorage"></param>
        private static void ReplaceParts(IEnumerable<OpenXmlPart> parts, IDictionary<string, TextModel> replaceStorage)
        {
            var partsList = parts?.ToList();
            if (partsList == null || partsList.Count == 0)
            {
                return;
            }

            foreach (var rootElement in partsList.Select(p=>p.RootElement))
            {
                if (rootElement == null)
                {
                    continue;
                }

                var textList = rootElement
                    .Descendants<Text>().ToList();
                foreach (var currentText in textList.Where(currentText => replaceStorage.Keys.Contains(currentText.Text)))
                {
                    var textModel = replaceStorage[currentText.Text];
                    currentText.Text = WinWordTextService.FillValue(textModel.Value, textModel.OutputFormat);
                }
                ReplaceInTables(rootElement, replaceStorage);
            }
        }
        /// <summary>
        /// Replace in tables
        /// </summary>
        /// <param name="rootElement"></param>
        /// <param name="replaceStorage"></param>
        private static void ReplaceInTables(OpenXmlPartRootElement rootElement, IDictionary<string, TextModel> replaceStorage)
        {
            var tableList = rootElement.Descendants<Table>().ToList();
            if (tableList.Count <= 0)
            {
                return;
            }

            foreach (var tableRow in tableList.Select(table => table.Descendants<TableRow>().ToList())
                         .SelectMany(tableRows => tableRows))
            {
                if (replaceStorage.Count == 0)
                {
                    return;
                }
                foreach (var (keyTag, textModel) in replaceStorage)
                {
                    WinWordTableService.ReplaceTextInRow(tableRow, keyTag,
                        WinWordTextService.FillValue(textModel.Value, textModel.OutputFormat));
                }
            }
        }
        
    }