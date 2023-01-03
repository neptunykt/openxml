using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using OpenXmlClient.Classes;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.Table;
using OpenXmlClient.Models.Text;
using OpenXmlServer.WinWordRenderer;
using UnitTests.Models;
using Xunit;

namespace UnitTests
{
    public class TableTest
    {
           // Default table test
        [Fact]
        public void DefaultTest()
        {
            // Arrange
            const string tableTag1Name = "TABLETAGA";
            const string insertTableRowsTag1Name = "INSERTROWS1";
            // Insert TextTag
            const string insertTextTag1Name = "INSERTTEXT1";
            const string insertTextTag2Name = "INSERTTEXT2";
            const string fileName = "TableTest.docx";
            var saveFileName = "DefaultTableResult.docx";
            var folderName = "TestWinWord";
            var salaryListFile = "SalaryList.json";
            // Настройка вывода для вставляемых строк

            var globalOutputFormat = new OutputFormat();
            globalOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalOutputFormat.DecimalFormat = "{0:00.00}";
            globalOutputFormat.FontName = "Times New Roman";
            globalOutputFormat.FontSize = 12;
            globalOutputFormat.IsBold = false;
            globalOutputFormat.IsItalic = false;
            globalOutputFormat.FontColor = OpenXmlFontColor.Black;

            var globalTableValueOutputFormat = new TableValueOutputFormat();
            globalTableValueOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalTableValueOutputFormat.DecimalFormat = "{0:00.00}";
            // ReSharper disable once EntityNameCapturedOnly.Local
            SalarySchedulerFakeClass salarySchedulerFakeClass;
            IEnumerable<IEnumerable<SalarySchedulerFakeClass>> salaryScheduler;
            using (var reader = new StreamReader(Path.Combine(folderName, salaryListFile)))
            {
                var json = reader.ReadToEnd();
                salaryScheduler =
                    JsonConvert.DeserializeObject<IEnumerable<IEnumerable<SalarySchedulerFakeClass>>>(json);
            }

            // ReSharper disable once GenericEnumeratorNotDisposed
            var enumerable = salaryScheduler as IEnumerable<SalarySchedulerFakeClass>[] ?? salaryScheduler?.ToArray();
            var tableRowFillModel1 = new TableRowFillModel();
            tableRowFillModel1.ColumnNames = new HashSet<string>();
            tableRowFillModel1.TableData = enumerable?[0] ?? Array.Empty<SalarySchedulerFakeClass>();
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var tableStorage = new TableFillStorageModel();
            tableStorage.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>
            {
                {insertTableRowsTag1Name,tableRowFillModel1 }
            };
            
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            // Заполняем замены текста в таблице
            tableStorage.TableFillTextReplaceStorage = new Dictionary<string, TableTextModel>();
            tableStorage.TableFillTextReplaceStorage.Add(insertTextTag1Name,
                new TableTextModel { Value = "Заглушка для ФИО" });
            // Удаляем текстовую метку если нет данных
            tableStorage.TableFillTextReplaceStorage.Add(insertTextTag2Name,
                new TableTextModel());
            dataTransformer.FillTable(tableStorage, tableTag1Name);
            var winWordModel = dataTransformer.GetWinWordModel();
            // Рендерим это делает сервер
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


        [Fact]
        public void RemoveTableIfNoDataTest()
        {
            // Arrange
            const string tableTag1Name = "TABLETAGA";
            const string insertTableRowsTag1Name = "INSERTROWS1";
            const string fileName = "TableTest.docx";
            var saveFileName = "RemoveTableResult.docx";
            var folderName = "TestWinWord";
            var salaryListFile = "SalaryList.json";
            // Настройка вывода для вставляемых строк

            var globalOutputFormat = new OutputFormat();
            globalOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalOutputFormat.DecimalFormat = "{0:00.00}";
            globalOutputFormat.FontName = "Times New Roman";
            globalOutputFormat.FontSize = 12;
            globalOutputFormat.IsBold = false;
            globalOutputFormat.IsItalic = false;
            globalOutputFormat.FontColor = OpenXmlFontColor.Black;

            var globalTableValueOutputFormat = new TableValueOutputFormat();
            globalTableValueOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalTableValueOutputFormat.DecimalFormat = "{0:00.00}";
            // ReSharper disable once EntityNameCapturedOnly.Local
            SalarySchedulerFakeClass salarySchedulerFakeClass;
            IEnumerable<IEnumerable<SalarySchedulerFakeClass>> salaryScheduler;
            using (var reader = new StreamReader(Path.Combine(folderName, salaryListFile)))
            {
                var json = reader.ReadToEnd();
                salaryScheduler =
                    JsonConvert.DeserializeObject<IEnumerable<IEnumerable<SalarySchedulerFakeClass>>>(json);
            }

            var tableRowFillModel1 = new TableRowFillModel();
            var tableStorage1 = new TableFillStorageModel();
            // условие, что нет данных
           
                // ReSharper disable once GenericEnumeratorNotDisposed
                var enumerable = salaryScheduler as IEnumerable<SalarySchedulerFakeClass>[] ??
                                 salaryScheduler?.ToArray();
                var enumerator = enumerable?.GetEnumerator();

                tableRowFillModel1.ColumnNames = new HashSet<string>();

                // Задаем порядок вывода столбцов по наименованию
                tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
                tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
                tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
                tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
                // Здесь должен быть цикл по значениям
           
                tableStorage1.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
                tableStorage1.TableRowsFillStorage.Add(insertTableRowsTag1Name, tableRowFillModel1);
                enumerator?.MoveNext();
                // Заполняем замены текста в таблице
                tableStorage1.TableFillTextReplaceStorage = new Dictionary<string, TableTextModel>();

                var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            // удаляем таблицу если нет данных
            var tableFormatSettings = new TableFormatSettings();
            tableFormatSettings.PrintOutTableFormatSettings = new PrintOutTableFormatSettings();
            tableFormatSettings.PrintOutTableFormatSettings.RemoveTable = true;
            dataTransformer.FillTable(tableStorage1, tableTag1Name, tableFormatSettings);
            var winWordModel = dataTransformer.GetWinWordModel();
            // Рендерим это делает сервер
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
        
        
          [Fact]
        public void TextGeneratorTableTest()
        {
            // Arrange
            const string tableTag1Name = "TABLETAGA";
            const string tableTextGeneratorTag1 = "TEXTGENERATORTAG1";
            const string tableTextGeneratorTag2 = "TEXTGENERATORTAG2";
            const string fileName = "TableTest.docx";
            var saveFileName = "TextGeneratorTableResult.docx";
            var folderName = "TestWinWord";
            // Настройка вывода для вставляемых строк

            var globalOutputFormat = new OutputFormat();
            globalOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalOutputFormat.DecimalFormat = "{0:00.00}";
            globalOutputFormat.FontName = "Times New Roman";
            globalOutputFormat.FontSize = 12;
            globalOutputFormat.IsBold = false;
            globalOutputFormat.IsItalic = false;
            globalOutputFormat.FontColor = OpenXmlFontColor.Black;

            var globalTableValueOutputFormat = new TableValueOutputFormat();
            globalTableValueOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalTableValueOutputFormat.DecimalFormat = "{0:00.00}";
            // ReSharper disable once EntityNameCapturedOnly.Local
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            // удаляем таблицу если нет данных
            var tableFormatSettings = new TableFormatSettings();
            tableFormatSettings.PrintOutTableFormatSettings = new PrintOutTableFormatSettings();
            tableFormatSettings.PrintOutTableFormatSettings.RemoveTable = true;
            var textGeneratorDictionary = new Dictionary<string, Collection<RunModel>>();
            var runCollection = new Collection<RunModel>
            {
                new RunModel
                {
                    Value = "Этот текст выведен жирным",
                    OutputFormat = new OutputFormat { IsBold = true }
                },
                new RunModel { Value = " этот текст будет обычным" },
                new RunModel{Value = " этот текст будет красным на желтом фоне", OutputFormat = new OutputFormat
                {
                    FontColor = OpenXmlFontColor.Red,
                    BackGroundColor = OpenXmlFontColor.Yellow
                }}
            };
            textGeneratorDictionary.Add(tableTextGeneratorTag1, runCollection);
            textGeneratorDictionary.Add(tableTextGeneratorTag2, new Collection<RunModel>());
            
            dataTransformer.FillTable(new TableFillStorageModel
            {
             TableFillTextGeneratorReplaceStorage  = textGeneratorDictionary
            },
                tableTag1Name
                );
            var winWordModel = dataTransformer.GetWinWordModel();
            // Рендерим это делает сервер
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