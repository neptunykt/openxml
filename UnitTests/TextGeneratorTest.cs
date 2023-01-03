using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlClient.Classes;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.HeaderFooter;
using OpenXmlClient.Models.Text;
using OpenXmlServer.WinWordRenderer;
using Xunit;

namespace UnitTests
{
        public class TextGeneratorTest
    {
        [Fact]
        public void TextRender()
        {
            // Arrange
            // Text global tags
            const string plannedSignedOnTag = "PLANNEDSIGNINON";
            const string companiesContractTag = "COMPANIESCONTRACT";
            const string removeNumberingTag = "REMOVENUMBERING";
            const string fillNumberingTag = "FILLNUMBERING";
            const string symbolOutputTag = "SYMBOLOUTPUT";
            const string numberingTextTag = "NUMBERINGTEST";


            const string fileName = "TextGeneratorTest.docx";
            var saveFileName = "TextGeneratorTestResult.docx";
            // Задаем формат вывода
            var globalOutputFormat = new OutputFormat();
            globalOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalOutputFormat.DecimalFormat = "{0:00.00}";
            globalOutputFormat.FontName = "Times New Roman";
            globalOutputFormat.FontSize = 11;
            globalOutputFormat.IsBold = false;
            globalOutputFormat.IsItalic = false;
            globalOutputFormat.FontColor = OpenXmlFontColor.Black;

            var globalTableValueOutputFormat = new TableValueOutputFormat();
            globalTableValueOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalTableValueOutputFormat.DecimalFormat = "{0:0.00}";
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            var headerFooterStorage = new HeaderFooterStorage();
            headerFooterStorage.HeaderReplaceDictionary = new Dictionary<string, TextModel>();
            headerFooterStorage.HeaderReplaceDictionary.Add(plannedSignedOnTag, new TextModel { Value = DateTime.Now });
            // удаляем (пункт 1.3 будет удален вместе с его нумерацией)
            dataTransformer.FillGlobalTextReplacement(new Dictionary<string, TextModel>
            {
                {
                    // для удаления параграфа отправляем пустой TextModel
                    removeNumberingTag, new TextModel()
                },
                // Пункт 1.3 будет теперь этот текст
                {
                    fillNumberingTag, new TextModel
                    {
                        Value = @"Глобальная замена с использованием стиля метки, 
первоначально параграф имел значение нумерации 1.4., но сейчас имеет нумерацию 1.3."
                    }
                },
                {
                    plannedSignedOnTag, new TextModel
                    {
                        Value = DateTime.Now,
                        OutputFormat = new OutputFormat { DateTimeFormat = "dd MMMM yyyy" }
                    }
                }
            });


            dataTransformer.FillHeaderFooterReplacement(headerFooterStorage);

            // Создаем массив текста
            var runList = new List<RunModel>
            {
                new RunModel
                {
                    OutputFormat = new OutputFormat
                    {
                        IsBold = true
                    },
                    Value = "Акционерный Коммерческий Банк «Алмазэргиэнбанк» Акционерное общество "
                },
                new RunModel
                {
                    OutputFormat = new OutputFormat
                    {
                        IsBold = false
                    },
                    Value = "именуемый в дальнейшем \"БАНК\" и"
                },
                // перевод строки
                new RunModel
                {
                    OutputFormat = new OutputFormat(),
                    IsCarriageReturn = true
                },
                // Табуляция
                new RunModel
                {
                    OutputFormat = new OutputFormat(),
                    IsTabulation = true
                },
                // Задаем отдельный формат
                // Задание формата текста - наклонный, 12 пикселей, красного цвета
                new RunModel
                {
                    OutputFormat = new OutputFormat
                    {
                        IsItalic = true,
                        FontColor = OpenXmlFontColor.Red,
                        BackGroundColor = OpenXmlFontColor.Yellow
                    },
                    Value = "Компания \"ООО Рога и копыта\" шрифтом текста 12 пикселей, Times New Roman, текст после перевода строки наклонный и красного цвета"
                },
                new RunModel
                {
                    Value = ", заключили договор на получение кредита."
                }
            };


            var symbolList = new List<RunModel>();
            var textSymbolRun = new RunModel { Value = "Вывод специальных символов: " };
            symbolList.Add(textSymbolRun);
            var symbolRun1 = new RunModel { IsSymbol = true, Value = "F09F" };
            var symbolRun2 = new RunModel { IsSymbol = true, Value = "F031" };
            var symbolRun3 = new RunModel { IsSymbol = true, Value = "F032" };
            var symbolRun4 = new RunModel { IsSymbol = true, Value = "F033" };
            var symbolRun5 = new RunModel { IsSymbol = true, Value = "F034" };
            var symbolRun6 = new RunModel { IsSymbol = true, Value = "F035" };
            var symbolRun7 = new RunModel { IsSymbol = true, Value = "F036" };
            var symbolRun8 = new RunModel { IsSymbol = true, Value = "F037" };
            var symbolRun9 = new RunModel { IsSymbol = true, Value = "F038" };
            var symbolRun10 = new RunModel { IsSymbol = true, Value = "F039" };
            symbolList.Add(symbolRun1);
            symbolList.Add(symbolRun2);
            symbolList.Add(symbolRun3);
            symbolList.Add(symbolRun4);
            symbolList.Add(symbolRun5);
            symbolList.Add(symbolRun6);
            symbolList.Add(symbolRun7);
            symbolList.Add(symbolRun8);
            symbolList.Add(symbolRun9);
            symbolList.Add(symbolRun10);

            var numberingTextArray = new Collection<Collection<RunModel>>();
            var runCollection1 = new Collection<RunModel>
            {
                new RunModel
                {
                    Value = "Это пункт 1.7. с учетом удаленного пункта и этот текст выведен жирным",
                    OutputFormat = new OutputFormat { IsBold = true }
                },
                new RunModel { Value = " этот текст будет обычным" }
            };
            numberingTextArray.Add(runCollection1);

            var runCollection2 = new Collection<RunModel>
            {
                new RunModel
                {
                    Value = "Это пункт 1.8. с учетом удаленного пункта и этот текст выведен жирным",
                    OutputFormat = new OutputFormat { IsBold = true }
                },
                new RunModel { Value = " этот текст будет обычным" }
            };
            numberingTextArray.Add(runCollection2);
            dataTransformer.FillNumberingTextStorage(new Dictionary<string, ICollection<Collection<RunModel>>>
            {
                {
                    numberingTextTag, numberingTextArray
                }
            });

            dataTransformer.FillGlobalTextGeneratorStorage(new Dictionary<string, IEnumerable<RunModel>>
            {
                { companiesContractTag, runList },
                { symbolOutputTag, symbolList }
            });
            var winWordModel = dataTransformer.GetWinWordModel();
            // Рендерим
            var folderName = "TestWinWord";

            WinWordTestRenderer.Render(fileName, winWordModel, folderName, saveFileName);
            var directoryPath = WinWordTestRenderer.GetDirectoryPath(folderName);
            File.ReadAllBytes(Path.Combine(directoryPath, saveFileName));
            var byteArray = File.ReadAllBytes(Path.Combine(directoryPath, fileName));
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    // Assert   
                    Assert.NotNull(wordDocument);
                }
            }
        }
    }
}