using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.Text;

namespace OpenXmlServer.WinWordRenderer.TextRenderer;

    public class WinWordTextService : OpenXmlElementService
    {
        public WinWordTextService(Body body) : base(body)
        {
        }

        /// <summary>
        /// Get paragraph index
        /// </summary>
        /// <param name="key"></param>
        /// <param name="paragraphs"></param>
        /// <returns></returns>
        public static int GetParagraphPosition(IEnumerable<Paragraph> paragraphs, string key = null)
        {
            var paragraphList = paragraphs?.ToList();
            var count = paragraphList?.Count;
            if (count == null || count == 0)
            {
                return 0;
            }

            if (key == null)
            {
                return count.Value;
            }

            for (var i = 0; i < count; i++)
            {
                if (IsParagraphContainsKey(paragraphList[i], key))
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// Имеется ли ключ в Параграфах
        /// </summary>
        /// <param name="para"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static bool IsParagraphContainsKey(Paragraph para, string key)
        {
            var runList = para.Descendants<Run>().ToList();
            if (runList.Count == 0)
            {
                return false;
            }

            return runList.Select(run => run.Descendants<Text>()).Any(textList => IsTextContainsKey(textList, key));
        }


        /// <summary>
        /// Is paragraph contains key
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static bool IsParagraphsContainsKey(IEnumerable<Paragraph> paragraphs, string key)
        {
            var paragraphList = paragraphs.ToList();
            return paragraphList.Any() && paragraphList.Any(paragraph => IsParagraphContainsKey(paragraph, key));
        }

        /// <summary>
        /// Get Text by tagName
        /// </summary>
        /// <param name="textItems"></param>
        /// <param name="tagName"></param>
        /// <returns></returns>
        public static Text GetTextContainsKey(IEnumerable<Text> textItems, string tagName)
        {
            var textList = textItems.ToList();
            return textList.FirstOrDefault(p => p.Text == tagName);
        }


        /// <summary>
        /// Get Run list with contains key
        /// </summary>
        /// <param name="para"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static IEnumerable<Run> GetRunContainsKey(Paragraph para, string key)
        {
            var runList = para.Descendants<Run>().ToList();
            var runs = new List<Run>();
            if (runList.Count == 0)
            {
                return runs;
            }

            foreach (var run in runList)
            {
                var textItems = run.Descendants<Text>();
                var text = GetTextContainsKey(textItems, key);
                if (text == null)
                {
                    continue;
                }

                runs.Add(run);
                return runs;
            }

            return runs;
        }

        /// <summary>
        /// Check if Text contains key
        /// </summary>
        /// <param name="textItems"></param>
        /// <param name="tagName"></param>
        /// <returns></returns>
        private static bool IsTextContainsKey(IEnumerable<Text> textItems, string tagName)
        {
            var textList = textItems.ToList();
            if (textList.Count == 0)
            {
                return false;
            }

            return textList.Any(text => text.Text == tagName);
        }

        /// <summary>
        /// Global text replacement
        /// </summary>
        /// <param name="tagName"></param>
        /// <param name="replaceText"></param>
        public void GlobalReplaceText(string tagName, string replaceText)
        {
            var paragraphs = _body.Descendants<Paragraph>();
            ReplaceText(paragraphs, tagName, replaceText);
        }

        /// <summary>
        /// Global text replacement
        /// </summary>
        /// <param name="tagName"></param>
        public void GlobalRemoveTextWithParagraph(string tagName)
        {
            var paragraphs = _body.Descendants<Paragraph>();

            foreach (var paragraph in paragraphs.Where(x =>
                         IsParagraphContainsKey(x, tagName)))
            {
                paragraph.Remove();
            }
        }


        /// <summary>
        /// Replace text in paragraphs
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="tagName"></param>
        /// <param name="replaceText"></param>
        public static void ReplaceText(IEnumerable<Paragraph> paragraphs, string tagName, string replaceText)
        {
            foreach (var paragraph in paragraphs)
            {
                ReplaceText(paragraph, tagName, replaceText);
            }
        }

        /// <summary>
        /// Replace text in paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="tagName"></param>
        /// <param name="replaceText"></param>
        private static void ReplaceText(Paragraph paragraph, string tagName, string replaceText)
        {
            foreach (var run in paragraph.Descendants<Run>())
            {
                var textList = run.Descendants<Text>().ToList();
                foreach (var text in textList.Where(text => text.Text.Contains(tagName)))
                {
                    text.Text = text.Text.Replace(tagName, replaceText);
                }
            }
        }


        /// <summary>
        ///  Вставка текст с возвратом параграфа
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="insertText"></param>
        /// <returns></returns>
        public static Paragraph InsertTextAfterParagraph(Paragraph paragraph, string insertText)
        {
            var newParagraph = new Paragraph();
            var text = new Text();
            text.Text = insertText;
            var run = new Run();
            run.AddChild(text);
            newParagraph.AddChild(run);
            return paragraph.InsertAfterSelf(newParagraph);
        }

        /// <summary>
        /// Находим параграфы по айдишке
        /// </summary>
        /// <param name="tagName"></param>
        /// <returns></returns>
        public IEnumerable<Paragraph> GetParagraphsFromBodyByTagName(string tagName)
        {
            var result = new List<Paragraph>();
            var paragraphs = _body?.Descendants<Paragraph>().ToList();
            var paragraphList = paragraphs?.ToList();
            var count = paragraphList?.Count;
            if (count == null || count == 0)
            {
                return result;
            }

            for (var i = 0; i < count; i++)
            {
                if (IsParagraphContainsKey(paragraphList[i], tagName))
                {
                    result.Add(paragraphList[i]);
                    return result;
                }
            }

            return result;
        }

        /// <summary>
        /// Находим параграфы по айдишке
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="tagName"></param>
        /// <returns></returns>
        public static IEnumerable<Paragraph> GetParagraphsByTagName(IEnumerable<Paragraph> paragraphs, string tagName)
        {
            var result = new List<Paragraph>();
            var paragraphList = paragraphs.ToList();
            var count = paragraphList.Count;
            if (count == 0)
            {
                return result;
            }

            for (var i = 0; i < count; i++)
            {
                if (!IsParagraphContainsKey(paragraphList[i], tagName))
                {
                    continue;
                }

                result.Add(paragraphList[i]);
                return result;
            }

            return result;
        }
        
        /// <summary>
        /// Insert element before paragraph
        /// </summary>
        /// <param name="newChild"></param>
        /// <param name="paragraph"></param>
        /// <typeparam name="T"></typeparam>
        public void InsertBeforeParagraph<T>(T newChild, Paragraph paragraph) where T : OpenXmlElement =>
            _body.InsertBefore(newChild, paragraph);


        /// <summary>
        /// Format for string Output
        /// </summary>
        /// <param name="value"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static string FillValue(object value, TableValueOutputFormat format)
        {
            if (format == null || format.DateTimeFormat == null || format.DecimalFormat == null)
            {
                throw new Exception("LOAN_CORPORATE_PRINT_FORM_SERVICE/FORMAT_SET");
            }

            switch (value)
            {
                case DateTime dateTimeValue:
                    return dateTimeValue.ToString(format.DateTimeFormat);
                case decimal decimalValue:
                    return string.Format(format.DecimalFormat, decimalValue);
                default:
                    return value?.ToString();
            }
        }


        public void CreateNumberingWithRuns(Paragraph paragraph, ICollection<RunModel> runs, string textTagKey)
        {
            // копируем numbering properties
            var paragraphNumbering = paragraph.ParagraphProperties?.NumberingProperties;
            var paragraphIndentation = paragraph.ParagraphProperties?.Indentation;
            var paragraphTabs = paragraph.ParagraphProperties?.Tabs;
            var paragraphJustification = paragraph.ParagraphProperties?.Justification;
            if (paragraphNumbering == null)
            {
                throw new Exception($"LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_NUMBERING: {textTagKey}");
            }
            var newParagraph = new Paragraph();
            newParagraph.ParagraphProperties = new ParagraphProperties();
            newParagraph.ParagraphProperties.NumberingProperties = paragraphNumbering.Clone() as NumberingProperties;
            newParagraph.ParagraphProperties.Indentation = paragraphIndentation?.Clone() as Indentation;
            newParagraph.ParagraphProperties.Tabs = paragraphTabs?.Clone() as Tabs;
            newParagraph.ParagraphProperties.Justification = paragraphJustification?.Clone() as Justification;
            foreach (var run in runs)
            {
                CreateRunInParagraph(newParagraph, run, run.OutputFormat);
            }

            paragraph.InsertAfterSelf(newParagraph);

        }
        public void CreateRunInParagraph(Paragraph paragraph, RunModel runModel,
            OutputFormat outputFormat = null)
        {
            RemoveOpenXmlElement(paragraph.ParagraphProperties?.ParagraphStyleId);
            if (runModel.IsSymbol)
            {
                FillSymbol(paragraph, runModel);
                return;
            }
            var run = paragraph.AppendChild(new Run());
            if (runModel.IsCarriageReturn)
            {
                run.AppendChild(new CarriageReturn());
                return;
            }

            if (runModel.IsTabulation)
            {
                run.AppendChild(new TabChar());
                return;
            }

            if (outputFormat != null)
            {
                runModel.SetOutputFormat(outputFormat);
            }

            // Fill run text
            run.AppendChild(
                new Text(FillValue(runModel.Value, runModel.OutputFormat))
                {
                    // preserve spacing between runs
                    Space = SpaceProcessingModeValues.Preserve
                });
            var runProp = new RunProperties();
            var runFont = new RunFonts { Ascii = outputFormat?.FontName };
            var fontSize = outputFormat?.FontSize * 2;

            if (outputFormat?.IsItalic == true)
            {
                runProp.AppendChild(new Italic());
            }

            if (outputFormat?.IsUnderLine == true)
            {
                var underLine = new Underline { Val = UnderlineValues.Single };
                runProp.AppendChild(underLine);
            }

            if (outputFormat?.IsBold == true)
            {
                runProp.AppendChild(new Bold());
            }

            if (!string.IsNullOrEmpty(outputFormat?.FontColor))
            {
                runProp.AppendChild(new Color { Val = outputFormat.FontColor });
            }
            
            if (!string.IsNullOrEmpty(outputFormat?.BackGroundColor))
            {
                runProp.AppendChild(new Shading
                {
                    Color = "auto",
                    Fill = outputFormat?.BackGroundColor,
                    Val = ShadingPatternValues.Clear
                });
            }

            var size = new FontSize { Val = new StringValue(fontSize.ToString()) };
            runProp.AppendChild(runFont);
            runProp.AppendChild(size);
            run.PrependChild(runProp);
        }
        private static void FillSymbol(Paragraph paragraph, RunModel runModel)
        {
            var newRun = new Run();
            paragraph.Append(newRun.ChildElements.Append(new SymbolChar { Font = "Wingdings", Char = runModel.Value }));
        }
    }