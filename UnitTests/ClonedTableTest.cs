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
    public class ClonedTableTest
    {
        // Default table clone test
        [Fact]
        public void DefaultTest()
        {
            // Arrange
            const string cloneTableTagName = "CLONETABLETAGA";
            const string cloneTablePlaceTagName = "CLONETABLEPLACETAGA";
            const string insertCloneTableRowsTagName = "INSERTROWS";
            const string insertTextGeneratorTagName = "INSERTTEXTGENERATOR";
            const string fileName = "ClonedTableTest.docx";
            var saveFileName = "DefaultClonedTableResult.docx";
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
            var enumerator = enumerable?.GetEnumerator();
            var tableRowFillModel1 = new TableRowFillModel();
            tableRowFillModel1.ColumnNames = new HashSet<string>();

            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var list = new List<TableFillStorageModel>();
            // Здесь должен быть цикл по значениям
            var clonedTableStorage1 = new TableFillStorageModel();
            clonedTableStorage1.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage1.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel1);
            enumerator?.MoveNext();
            tableRowFillModel1.TableData = enumerable?[0];
            var tableRowFillModel2 = new TableRowFillModel();
            tableRowFillModel2.ColumnNames = new HashSet<string>();
            tableRowFillModel2.TableData = enumerable?[1];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage2 = new TableFillStorageModel();
            clonedTableStorage2.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage2.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel2);
            clonedTableStorage2.TableFillTextGeneratorReplaceStorage = new Dictionary<string, Collection<RunModel>>();
            // удаляем запись
            clonedTableStorage2.TableFillTextGeneratorReplaceStorage.Add(insertTextGeneratorTagName,new Collection<RunModel>());            

            var tableRowFillModel3 = new TableRowFillModel();
            tableRowFillModel3.ColumnNames = new HashSet<string>();
            tableRowFillModel3.TableData = enumerable?[2];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage3 = new TableFillStorageModel();
            clonedTableStorage3.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage3.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel3);
            clonedTableStorage3.TableFillTextGeneratorReplaceStorage = new Dictionary<string, Collection<RunModel>>
            {
                {insertTextGeneratorTagName,new Collection<RunModel>
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
                }}
            };
           
            // Задаем вывод строк в метку
            list.Add(clonedTableStorage1);
            list.Add(clonedTableStorage2);
            list.Add(clonedTableStorage3);
            // Готовим данные для отправки
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            dataTransformer.FillClonedTable(list, cloneTableTagName, cloneTablePlaceTagName);
            var winWordModel = dataTransformer.GetWinWordModel();
            var tableStorage = new TableFillStorageModel();
            tableStorage.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
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

        // Set custom format output
         [Fact]
        public void SetCustomFormatTest()
        {
            // Arrange
            const string cloneTableTagName = "CLONETABLETAGA";
            const string cloneTablePlaceTagName = "CLONETABLEPLACETAGA";
            const string insertCloneTableRowsTagName = "INSERTROWS";
            const string fileName = "ClonedTableTest.docx";
            var saveFileName = "CustomFormatClonedTableResult.docx";
            var folderName = "TestWinWord";
            var salaryListFile = "SalaryList.json";
            // Настройка вывода для вставляемых строк

            var globalOutputFormat = new OutputFormat();
            globalOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            globalOutputFormat.DecimalFormat = "{0:0.00}";
            globalOutputFormat.FontName = "Times New Roman";
            globalOutputFormat.FontSize = 12;
            globalOutputFormat.IsBold = false;
            globalOutputFormat.IsItalic = false;
            globalOutputFormat.FontColor = OpenXmlFontColor.Black;

            var globalTableValueOutputFormat = new TableValueOutputFormat();
            globalTableValueOutputFormat.DateTimeFormat = "dd.MM.yyyy";
            // Дополнение нулем, если в первом разряде нет
            globalTableValueOutputFormat.DecimalFormat = "{0:0.00}";
            // ReSharper disable once EntityNameCapturedOnly.Local
            SalarySchedulerFakeClass salarySchedulerFakeClass;
            IEnumerable<IEnumerable<SalarySchedulerFakeClass>> salaryScheduler;
            using (var reader = new StreamReader(Path.Combine(folderName, salaryListFile)))
            {
                var json = reader.ReadToEnd();
                salaryScheduler =
                    JsonConvert.DeserializeObject<IEnumerable<IEnumerable<SalarySchedulerFakeClass>>>(json);
            }
            var list = new List<TableFillStorageModel>();
            // ReSharper disable once GenericEnumeratorNotDisposed
            var enumerable = salaryScheduler as IEnumerable<SalarySchedulerFakeClass>[] ?? salaryScheduler?.ToArray();
            var enumerator = enumerable?.GetEnumerator();
            var tableRowFillModel1 = new TableRowFillModel();
            tableRowFillModel1.ColumnNames = new HashSet<string>();

            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            // Custom format set for one table
            tableRowFillModel1.TableValueOutputFormat = new OutputFormat
            {
                DateTimeFormat = "dddd, dd MMMM yyyy",
                // не дополнять нулем и оставить один разряд после запятой
                DecimalFormat = "{0:0.0}"
            };
            
         
            // Здесь должен быть цикл по значениям
            var clonedTableStorage1 = new TableFillStorageModel();
            clonedTableStorage1.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage1.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel1);
            enumerator?.MoveNext();
            tableRowFillModel1.TableData = enumerable?[0];
            var tableRowFillModel2 = new TableRowFillModel();
            tableRowFillModel2.ColumnNames = new HashSet<string>();
            tableRowFillModel2.TableData = enumerable?[1];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage2 = new TableFillStorageModel();
            clonedTableStorage2.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage2.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel2);

            var tableRowFillModel3 = new TableRowFillModel();
            tableRowFillModel3.ColumnNames = new HashSet<string>();
            tableRowFillModel3.TableData = enumerable?[2];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage3 = new TableFillStorageModel();
            clonedTableStorage3.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage3.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel3);
            // Задаем вывод строк в метку
            list.Add(clonedTableStorage1);
            list.Add(clonedTableStorage2);
            list.Add(clonedTableStorage3);
            // Готовим данные для отправки
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            dataTransformer.FillClonedTable(list, cloneTableTagName, cloneTablePlaceTagName);
            var winWordModel = dataTransformer.GetWinWordModel();
            var tableStorage = new TableFillStorageModel();
            tableStorage.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
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

        
        // Set page breaks between cloned tables
        [Fact]
        public void SetPageBreaksTest()
        {
            // Arrange
            const string cloneTableTagName = "CLONETABLETAGA";
            const string cloneTablePlaceTagName = "CLONETABLEPLACETAGA";
            const string insertCloneTableRowsTagName = "INSERTROWS";
            const string fileName = "ClonedTableTest.docx";
            var saveFileName = "SetPageBreaksClonedTableResult.docx";
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
            var enumerator = enumerable?.GetEnumerator();
            var tableRowFillModel1 = new TableRowFillModel();
            tableRowFillModel1.ColumnNames = new HashSet<string>();

            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var list = new List<TableFillStorageModel>();
            // Здесь должен быть цикл по значениям
            var clonedTableStorage1 = new TableFillStorageModel();
            clonedTableStorage1.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage1.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel1);
            enumerator?.MoveNext();
            tableRowFillModel1.TableData = enumerable?[0];
            var tableRowFillModel2 = new TableRowFillModel();
            tableRowFillModel2.ColumnNames = new HashSet<string>();
            tableRowFillModel2.TableData = enumerable?[1];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage2 = new TableFillStorageModel();
            clonedTableStorage2.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage2.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel2);

            var tableRowFillModel3 = new TableRowFillModel();
            tableRowFillModel3.ColumnNames = new HashSet<string>();
            tableRowFillModel3.TableData = enumerable?[2];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage3 = new TableFillStorageModel();
            clonedTableStorage3.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage3.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel3);

            // Задаем вывод строк в метку
            list.Add(clonedTableStorage1);
            list.Add(clonedTableStorage2);
            list.Add(clonedTableStorage3);
            var clonedTableFormatSettings = new ClonedTableFormatSettings();
            clonedTableFormatSettings.PrintOutClonedTableFormatSettings = new PrintOutClonedTableFormatSettings { IsSetPageBreak = true };
            // Готовим данные для отправки
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            dataTransformer.FillClonedTable(list, cloneTableTagName, cloneTablePlaceTagName, clonedTableFormatSettings);
            var winWordModel = dataTransformer.GetWinWordModel();
            var tableStorage = new TableFillStorageModel();
            tableStorage.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
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
        
        // Set table header on each page for cloned tables
          [Fact]
        public void SetHeaderOnEachPageTest()
        {
            // Arrange
            const string cloneTableTagName = "CLONETABLETAGA";
            const string cloneTablePlaceTagName = "CLONETABLEPLACETAGA";
            const string insertCloneTableRowsTagName = "INSERTROWS";
            const string fileName = "ClonedTableTest.docx";
            var saveFileName = "SetHeaderOnEachPageClonedTableResult.docx";
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
            var enumerator = enumerable?.GetEnumerator();
            var tableRowFillModel1 = new TableRowFillModel();
            tableRowFillModel1.ColumnNames = new HashSet<string>();

            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var list = new List<TableFillStorageModel>();
            // Здесь должен быть цикл по значениям
            var clonedTableStorage1 = new TableFillStorageModel();
            clonedTableStorage1.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage1.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel1);
            enumerator?.MoveNext();
            tableRowFillModel1.TableData = enumerable?[0];
            var tableRowFillModel2 = new TableRowFillModel();
            tableRowFillModel2.ColumnNames = new HashSet<string>();
            tableRowFillModel2.TableData = enumerable?[1];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage2 = new TableFillStorageModel();
            clonedTableStorage2.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage2.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel2);

            var tableRowFillModel3 = new TableRowFillModel();
            tableRowFillModel3.ColumnNames = new HashSet<string>();
            tableRowFillModel3.TableData = enumerable?[2];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage3 = new TableFillStorageModel();
            clonedTableStorage3.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage3.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel3);

            // Задаем вывод строк в метку
            list.Add(clonedTableStorage1);
            list.Add(clonedTableStorage2);
            list.Add(clonedTableStorage3);
            var clonedTableFormatSettings = new ClonedTableFormatSettings();
            clonedTableFormatSettings.PrintOutClonedTableFormatSettings = new PrintOutClonedTableFormatSettings { IsSetPageBreak = true };
            clonedTableFormatSettings.PrintOutClonedTableFormatSettings.RepeatHeadingRowSettings =
                new RepeatHeadingRowSettings
                {
                    StartRepeatHeadingRowNumber = 1,
                    EndRepeatHeadingRowNumber = 1
                };
            // Готовим данные для отправки
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            dataTransformer.FillClonedTable(list, cloneTableTagName, cloneTablePlaceTagName, clonedTableFormatSettings);
            var winWordModel = dataTransformer.GetWinWordModel();
            var tableStorage = new TableFillStorageModel();
            tableStorage.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
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
        
        // Fill inner tables
        [Fact]
          public void FillInnerTableTest()
        {
            // Arrange
            const string cloneTableTagName = "CLONETABLETAGA";
            const string cloneTablePlaceTagName = "CLONETABLEPLACETAGA";
            const string insertCloneTableRowsTagName = "INSERTROWS";
            const string cloneInnerTablePlaceTagName = "CLONEINNERTABLETAGA";
            const string insertCloneInnerTableRowsTagName = "INSERTINNERROWS";
            const string fileName = "ClonedInnerTable.docx";
            var saveFileName = "ClonedInnerTableResult.docx";
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
            // ------ первая таблица ------
            // ReSharper disable once GenericEnumeratorNotDisposed
            var enumerable = salaryScheduler as IEnumerable<SalarySchedulerFakeClass>[] ?? salaryScheduler?.ToArray();
            var enumerator = enumerable?.GetEnumerator();
            var tableRowFillModel1 = new TableRowFillModel();
            tableRowFillModel1.ColumnNames = new HashSet<string>();

            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var list = new List<TableFillStorageModel>();
            // Здесь должен быть цикл по значениям
            var clonedTableStorage1 = new TableFillStorageModel();
            clonedTableStorage1.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage1.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel1);
            enumerator?.MoveNext();
            tableRowFillModel1.TableData = enumerable?[0];
            
            // Заполнение внутренней таблицы

            clonedTableStorage1.InnerTableFillStorage =
                new Dictionary<string, IDictionary<string, InnerTableRowFillModel>>();
            var innerTableRowFillModel1 = new InnerTableRowFillModel();
            innerTableRowFillModel1.ColumnNames = new HashSet<string>();
            innerTableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            innerTableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            innerTableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            innerTableRowFillModel1.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            innerTableRowFillModel1.TableData = enumerable?[0];
            var innerTableRowFillStorage1 = new Dictionary<string, InnerTableRowFillModel>();
            innerTableRowFillStorage1.Add(insertCloneInnerTableRowsTagName, innerTableRowFillModel1);
            clonedTableStorage1.InnerTableFillStorage.Add(cloneInnerTablePlaceTagName, innerTableRowFillStorage1);
            // ------ вторая таблица ------
            var tableRowFillModel2 = new TableRowFillModel();
            tableRowFillModel2.ColumnNames = new HashSet<string>();
            tableRowFillModel2.TableData = enumerable?[1];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel2.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            var clonedTableStorage2 = new TableFillStorageModel();
            clonedTableStorage2.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage2.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel2);
            // удаляем только айдишку внутренней таблицы если нет данных (по умолчанию)
            clonedTableStorage2.InnerTableFillStorage =
                new Dictionary<string, IDictionary<string, InnerTableRowFillModel>>();
            var innerTableRowFillStorage2 = new Dictionary<string, InnerTableRowFillModel>();
            // добавляем пустой tableRowFillModel
            innerTableRowFillStorage2.Add(insertCloneInnerTableRowsTagName, new InnerTableRowFillModel());
            clonedTableStorage2.InnerTableFillStorage.Add(cloneInnerTablePlaceTagName, innerTableRowFillStorage2);
            // ------ третья таблица ------
            var tableRowFillModel3 = new TableRowFillModel();
            tableRowFillModel3.ColumnNames = new HashSet<string>();
            tableRowFillModel3.TableData = enumerable?[2];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel3.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            
            var clonedTableStorage3 = new TableFillStorageModel();
            clonedTableStorage3.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage3.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel3);
            // удаляем внутренюю таблицу если нет данных
            clonedTableStorage3.InnerTableFillStorage =
                new Dictionary<string, IDictionary<string, InnerTableRowFillModel>>();
            var innerTableRowFillStorage3 = new Dictionary<string, InnerTableRowFillModel>();
            // добавляем пустой tableRowFillModel (удаляем только пустую строку с тегом вывода строк если нет записей)
            innerTableRowFillStorage3.Add(insertCloneInnerTableRowsTagName, new InnerTableRowFillModel{RemoveEntireRowIfNoRecords = true});
            clonedTableStorage3.InnerTableFillStorage.Add(cloneInnerTablePlaceTagName, innerTableRowFillStorage3);
            // ------ четвертая таблица ------
            var tableRowFillModel4 = new TableRowFillModel();
            tableRowFillModel4.ColumnNames = new HashSet<string>();
            tableRowFillModel4.TableData = enumerable?[3];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel4.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel4.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel4.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel4.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            
            var clonedTableStorage4 = new TableFillStorageModel();
            clonedTableStorage4.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage4.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel3);
            // удаляем внутренюю таблицу если нет данных
            clonedTableStorage4.InnerTableFillStorage =
                new Dictionary<string, IDictionary<string, InnerTableRowFillModel>>();
            var innerTableRowFillStorage4 = new Dictionary<string, InnerTableRowFillModel>();
            // добавляем пустой tableRowFillModel (удаляем всю внешнюю строку если нет записей)
            innerTableRowFillStorage4.Add(insertCloneInnerTableRowsTagName, new InnerTableRowFillModel{IsRemoveEntireOuterRowIfNoRecords = true});
            clonedTableStorage4.InnerTableFillStorage.Add(cloneInnerTablePlaceTagName, innerTableRowFillStorage4);
            
            // ------ пятая таблица ------
            var tableRowFillModel5 = new TableRowFillModel();
            tableRowFillModel5.ColumnNames = new HashSet<string>();
            tableRowFillModel5.TableData = enumerable?[3];
            // задаем порядок вывода столбцов по наименованию
            tableRowFillModel5.ColumnNames.Add(nameof(salarySchedulerFakeClass.Fio));
            tableRowFillModel5.ColumnNames.Add(nameof(salarySchedulerFakeClass.DateSignOn));
            tableRowFillModel5.ColumnNames.Add(nameof(salarySchedulerFakeClass.WorkedHours));
            tableRowFillModel5.ColumnNames.Add(nameof(salarySchedulerFakeClass.Sum));
            
            var clonedTableStorage5 = new TableFillStorageModel();
            clonedTableStorage5.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
            clonedTableStorage5.TableRowsFillStorage.Add(insertCloneTableRowsTagName, tableRowFillModel3);
            // удаляем внутренюю таблицу если нет данных
            clonedTableStorage5.InnerTableFillStorage =
                new Dictionary<string, IDictionary<string, InnerTableRowFillModel>>();
            var innerTableRowFillStorage5 = new Dictionary<string, InnerTableRowFillModel>();
            // добавляем пустой tableRowFillModel (удаляем внутреннюю таблицу если нет данных)
            innerTableRowFillStorage5.Add(insertCloneInnerTableRowsTagName, new InnerTableRowFillModel{IsRemoveInnerTableIfNoRecords = true});
            clonedTableStorage5.InnerTableFillStorage.Add(cloneInnerTablePlaceTagName, innerTableRowFillStorage5);
            
            // Задаем вывод строк в метку
            list.Add(clonedTableStorage1);
            list.Add(clonedTableStorage2);
            list.Add(clonedTableStorage3);
            list.Add(clonedTableStorage4);
            list.Add(clonedTableStorage5);
            // Готовим данные для отправки
            var dataTransformer = new DataTransformer(globalOutputFormat, globalTableValueOutputFormat);
            dataTransformer.FillClonedTable(list, cloneTableTagName, cloneTablePlaceTagName);
            var winWordModel = dataTransformer.GetWinWordModel();
            var tableStorage = new TableFillStorageModel();
            tableStorage.TableRowsFillStorage = new Dictionary<string, TableRowFillModel>();
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