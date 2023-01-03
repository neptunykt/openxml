using System.Collections.ObjectModel;
using System.Data;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlClient.Classes.Models;
using OpenXmlClient.FormatSettings;
using OpenXmlServer.WinWordRenderer.TextRenderer;

namespace OpenXmlServer.WinWordRenderer.TableRenderer;

    public class WinWordTableRenderer
    {
        private readonly WinWordTableService _winWordTableService;
        private readonly WinWordTextService _winWordTextService;

        public WinWordTableRenderer(WinWordTableService winWordTableService, WinWordTextService winWordTextService)
        {
            _winWordTableService = winWordTableService;
            _winWordTextService = winWordTextService;
        }

        /// <summary>
        /// Clone tables
        /// </summary>
        /// <param name="winWordRenderModel"></param>
        /// <exception cref="ServiceException"></exception>
        public void Clone(WinWordRenderModel winWordRenderModel)
        {
            if (winWordRenderModel?.ClonedTablesStorage == null || winWordRenderModel.ClonedTablesStorage.Count == 0)
            {
                return;
            }

            // key is for one table
            var keyValuePair = winWordRenderModel.ClonedTablesStorage.FirstOrDefault();
            var insertTablePlaceTagName = keyValuePair.Value.InsertTablePlaceTagName;
            var clonedTableTagName = keyValuePair.Value.ClonedTableTagName;
            var tableIndex = 0;
            var clonedTableCount = winWordRenderModel.ClonedTablesStorage.Count;
            var isLastTable = false;
            foreach (var (tableTagNameKey, valueTableData) in winWordRenderModel.ClonedTablesStorage)
            {
                tableIndex++;
                if (tableIndex == clonedTableCount)
                {
                    isLastTable = true;
                }
                // находим таблицу
                var tableToClone = _winWordTableService.GetTablesByTableTagName(valueTableData.ClonedTableTagName)
                    .FirstOrDefault();
                if (tableToClone == null)
                {
                    throw new Exception(
                        $"LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_CLONED_TABLE_TAG_NAME: {valueTableData.ClonedTableTagName}");
                }
                var clonedTable = tableToClone.CloneNode(true) as Table;
                var searchedParagraph = _winWordTextService
                    .GetParagraphsFromBodyByTagName(valueTableData.InsertTablePlaceTagName).FirstOrDefault();
                if (searchedParagraph == null)
                {
                    throw new Exception(
                        $"LOANCORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_PLACE_TO_INSERT_CLONED_TABLE: {valueTableData.InsertTablePlaceTagName}");
                }

                SetLinesOrPageBreaksBetweenClonedTables(searchedParagraph, clonedTable, valueTableData, isLastTable);
                // заменяем TableTagName у склонированной таблицы
                _winWordTableService.ReplaceTableTagName(valueTableData.ClonedTableTagName, tableTagNameKey);
                
                FillTableRenderStorage(winWordRenderModel, valueTableData, tableTagNameKey);
            }

            // Удаляем уже ненужную метку для вставки таблицы
            _winWordTextService.GlobalRemoveTextWithParagraph(insertTablePlaceTagName);
            // Удаляем таблицу для клонирования
            _winWordTableService.RemoveTableByTagName(clonedTableTagName);
        }

        /// <summary>
        /// Fill WinWordModel for new tableTagName
        /// </summary>
        /// <param name="winWordRenderModel"></param>
        /// <param name="clonedTableRenderModel"></param>
        /// <param name="tableTagName"></param>
        private static void FillTableRenderStorage(WinWordRenderModel winWordRenderModel,
            ClonedTableRenderModel clonedTableRenderModel, string tableTagName)
        {
            if (winWordRenderModel.TablesStorage == null)
            {
                winWordRenderModel.TablesStorage = new Dictionary<string, TableRenderModel>();
            }

            var copyTableRenderModel = clonedTableRenderModel.CopyTableRenderStorage;
            copyTableRenderModel.PrintOutTableFormatSettings = clonedTableRenderModel.PrintOutClonedTableFormatSettings;
            if (copyTableRenderModel.InnerTablesStorage == null)
            {
                copyTableRenderModel.InnerTablesStorage = new Dictionary<string, IDictionary<string, InnerRowsRenderPayload>>();
            }

            if (!winWordRenderModel.ClonedTablesStorage.ContainsKey(tableTagName))
            {
                throw new Exception($"LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_TABLE_TAG: {tableTagName}");
            }

            if (winWordRenderModel.ClonedTablesStorage[tableTagName]
                    .InnerTableFillStorage != null && winWordRenderModel.ClonedTablesStorage[tableTagName]
                    .InnerTableFillStorage.Count > 0)
            {
                foreach (var (_, innerTableValues) in winWordRenderModel.ClonedTablesStorage[tableTagName]
                             .InnerTableFillStorage)
                {
                    if (innerTableValues?.InnerTablesStorage == null)
                    {
                        continue;
                    }

                    foreach (var (innerTableTag, innerTableValueStorage) in innerTableValues.InnerTablesStorage)
                    {
                        copyTableRenderModel.InnerTablesStorage.Add(innerTableTag, innerTableValueStorage);
                    }

                }
            }
            winWordRenderModel.TablesStorage.Add(tableTagName, copyTableRenderModel);
        }

        /// <summary>
        /// Main method for filling rows into tag name
        /// </summary>
        /// <param name="table"></param>
        /// <param name="tableRenderModel"></param>
        /// <exception cref="Exception"></exception>
        private static void RenderTableRows(Table table, TableRenderModel tableRenderModel)
        {
            SetTableHeaderOnEachPage(table, tableRenderModel.PrintOutTableFormatSettings);
            var payloadStore = tableRenderModel.TableRowsData;

            if (payloadStore == null)
            {
                return;
            }

            foreach (var (keyRowsPlaceTagName, payloadTableValue) in payloadStore)
            {
                // key в Dictionary нужен для того, чтобы найти позицию в таблице куда выводить строки
                if (!payloadStore.ContainsKey(keyRowsPlaceTagName))
                {
                    throw new Exception(
                        $"LOANCORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_TABLE_ROWS_PLACE_TAG_NAME: {keyRowsPlaceTagName}");
                }

                var (columnStart, tableRows, runs) =
                    WinWordTableService.GetTableRowColumnIndexAndTextByTagName(table, keyRowsPlaceTagName);
                if (payloadTableValue?.Payload != null)
                {
                    var serializer = new XmlSerializer(typeof(DataTable));
                    using (var sw = new StringReader(payloadTableValue.Payload))
                    {
                        var dataTable = serializer.Deserialize(sw) as DataTable;
                        var dataTableRows = dataTable?.Rows;
                        var enumerable = tableRows as TableRow[] ?? tableRows.ToArray();
                        if (dataTableRows == null || dataTableRows.Count == 0 || tableRows == null ||
                            enumerable.Length == 0)
                        {
                            throw new Exception("LOANCORPORATE_PRINT_FORM_SERVICE/DATA_ROWS_COUNT_IS_NULL");
                        }


                        RenderRows(columnStart, enumerable, runs, keyRowsPlaceTagName, payloadTableValue,
                            dataTable);
                    }

                    // удаляем ненужную метку для вставки вместе со строкой (это если есть записи)
                    WinWordTableService.RemoveTableRowByTableAndTableRowsPlaceTagName(table, keyRowsPlaceTagName);
                }
                else if (payloadTableValue != null && payloadTableValue.IsRemoveEntireRowIfNoRecords)
                {
                    // удаляем ненужную метку для вставки строк (строку не удаляем если не указано)
                    WinWordTableService.RemoveTableRowByTableAndTableRowsPlaceTagName(table, keyRowsPlaceTagName);
                }
                else
                {
                    // чистим метку для вставки
                    WinWordTableService.ClearTableRowsPlaceTagName(table, keyRowsPlaceTagName);
                }
            }
        }

        /// <summary>
        /// Render rows
        /// </summary>
        /// <param name="columnStart"></param>
        /// <param name="tableRows"></param>
        /// <param name="runs"></param>
        /// <param name="keyRowsPlaceTagName"></param>
        /// <param name="payloadTableValue"></param>
        /// <param name="dataTable"></param>
        private static void RenderRows(int columnStart, IEnumerable<TableRow> tableRows, IEnumerable<Run> runs,
            string keyRowsPlaceTagName, RowsRenderPayload payloadTableValue, DataTable dataTable)
        {
            var tableRowsList = tableRows.ToList();

            var dataTableRows = dataTable?.Rows;
            var tableRow = tableRowsList.FirstOrDefault();
            var runProperties = runs.FirstOrDefault()?.RunProperties;
            var columnNamesCollection = payloadTableValue.ColumnNames;
            var rowsCount = dataTableRows?.Count;
            if (rowsCount == null)
            {
                return;
            }

            for (var i = 0; i < rowsCount; i++)
            {
                // создаем новую строку
                var newRow = tableRowsList.FirstOrDefault()?.CloneNode(true);
                // удаляем тег в строке
                WinWordTableService.RemoveTagInRow(newRow as TableRow, keyRowsPlaceTagName);
                WinWordTableService.RemoveInnerBorderInColumns(newRow as TableRow, i, rowsCount.Value,
                    payloadTableValue.RemoveInnerBordersInColumnAbsolutePositions);
                if (tableRow == null || newRow == null || runProperties == null)
                {
                    continue;
                }

                tableRow.InsertBeforeSelf(newRow);
                var newCells = newRow.Descendants<TableCell>().ToList();
                var columnIndex = columnStart;
                // вытаскиваем стиль в указанном Runs
                foreach (var columnName in columnNamesCollection)
                {
                    var value = WinWordTextService.FillValue(dataTableRows[i][columnName],
                        payloadTableValue.TableValueOutputFormat);
                    var paragraphProperties = newCells[columnIndex].Descendants<Paragraph>().FirstOrDefault()
                        ?.ParagraphProperties;
                    var textValue = value ?? "";
                    WinWordTableService.RemoveParagraphInCells(newCells, columnIndex);
                    var run = new Run();
                    run.AppendChild(new Text(textValue));
                    // Клонируем стили
                    run.PrependChild(runProperties.Clone() as RunProperties);
                    var paragraph = new Paragraph();
                    // Клонируем стили параграфа
                    paragraph.AppendChild(run);
                    FillParagraphProperties(paragraphProperties, paragraph);
                    newCells[columnIndex].AppendChild(paragraph);
                    columnIndex++;
                }
            }
        }


        /// <summary>
        /// Copy paragraph properties
        /// </summary>
        /// <param name="paragraphProperties"></param>
        /// <param name="paragraph"></param>
        private static void FillParagraphProperties(ParagraphProperties paragraphProperties, Paragraph paragraph)
        {
            if (paragraphProperties == null)
            {
                return;
            }

            if (paragraph.ParagraphProperties == null)
            {
                paragraph.ParagraphProperties = new ParagraphProperties();
            }

            paragraph.PrependChild(paragraphProperties.Clone() as ParagraphProperties);
        }


        /// <summary>
        /// Fill Tables method
        /// </summary>
        /// <param name="winWordRenderModel"></param>
        public void Render(WinWordRenderModel winWordRenderModel)
        {
            if (!(winWordRenderModel.TablesStorage?.Count > 0))
            {
                return;
            }

            foreach (var (keyTableTag, tableValue) in winWordRenderModel.TablesStorage)
            {
                var table = _winWordTableService.GetTablesByTableTagName(keyTableTag).FirstOrDefault();
                if (table != null)
                {
                    if (IsRemoveTable(tableValue))
                    {
                        // удаляем таблицу
                        _winWordTableService.RemoveTableByTagName(keyTableTag);
                        continue;
                    }
                    
                    RenderTableRows(table, tableValue);
                    RenderInnerTables(table, tableValue);
                    // Заполняем текстовые метки в таблице
                    var tableRows = table.Descendants<TableRow>();
                    ReplaceTagTextGeneratorInTable(tableRows, tableValue);
                    ReplaceTextInTable(tableRows, tableValue);
                }

                // удаляем тег таблицы
                WinWordTableService.RemoveTableTagName(table);
            }
        }

        /// <summary>
        /// Replace tag in table
        /// </summary>
        /// <param name="tableRows"></param>
        /// <param name="tableRenderModel"></param>
        private static void ReplaceTextInTable(IEnumerable<TableRow> tableRows, TableRenderModel tableRenderModel)
        {
            // находим строку 
            var tableRowsList = tableRows.ToList();
            if (tableRenderModel?.TextReplaceData == null)
            {
                return;
            }

            foreach (var (textReplaceTag, textReplaceData) in tableRenderModel.TextReplaceData)
            {
                var tableRow = WinWordTableService.GetRowsWithTextTag(tableRowsList, textReplaceTag).FirstOrDefault();
                if (tableRow == null)
                {
                    continue;
                }

                if (textReplaceData.Value != null)
                {
                    // Заменяем текст
                    var strValue =
                        WinWordTextService.FillValue(textReplaceData.Value, textReplaceData.TableValueOutputFormat);
                    WinWordTableService.ReplaceTextInRow(tableRow, textReplaceTag, strValue);
                }
                else if (textReplaceData.Value == null && textReplaceData.RemoveEntireRowIfNoValue)
                {
                    // Удаляем полностью строку
                    tableRow.Remove();
                }
                else if (textReplaceData.Value == null && !textReplaceData.RemoveEntireRowIfNoValue)
                {
                    // Просто удаляем тег
                    WinWordTableService.RemoveTagInRow(tableRow, textReplaceTag);
                }
            }
        }
        
        
        
        /// <summary>
        /// Replace tag in table
        /// </summary>
        /// <param name="tableRows"></param>
        /// <param name="tableRenderModel"></param>
        public void ReplaceTagTextGeneratorInTable(IEnumerable<TableRow> tableRows, TableRenderModel tableRenderModel)
        {
            // находим строку 
            var tableRowsList = tableRows.ToList();
            if (tableRenderModel?.TextGeneratorReplaceData == null)
            {
                return;
            }

            foreach (var (textReplaceTag, textReplaceData) in tableRenderModel.TextGeneratorReplaceData)
            {
                var tableRow = WinWordTableService.GetRowsWithTextTag(tableRowsList, textReplaceTag).FirstOrDefault();
                if (tableRow == null)
                {
                    throw new Exception($"NOT_FOUND_TABLE_TEXT_GENERATOR_TAG/{textReplaceTag}");
                }
                _winWordTableService.ReplaceTagTextGeneratorInRow(tableRow, textReplaceTag, textReplaceData);

            }
        }
        
        
        

        /// <summary>
        /// Set lines or page breaks between cloned tables
        /// </summary>
        /// <param name="searchedParagraph"></param>
        /// <param name="clonedTable"></param>
        /// <param name="clonedTableRenderModel"></param>
        /// <param name="isLastTable"></param>
        private void SetLinesOrPageBreaksBetweenClonedTables(Paragraph searchedParagraph, Table clonedTable,
            ClonedTableRenderModel clonedTableRenderModel, bool isLastTable)
        {
            if (clonedTableRenderModel.PrintOutClonedTableFormatSettings == null && !isLastTable)
            {
                var paragraph = new Paragraph();
                searchedParagraph.InsertBeforeSelf(paragraph);
                _winWordTextService.InsertBeforeParagraph(clonedTable, searchedParagraph);
                return;
            }
            if (clonedTableRenderModel.PrintOutClonedTableFormatSettings == null && isLastTable
                || clonedTableRenderModel.PrintOutClonedTableFormatSettings != null && isLastTable)
            {
                _winWordTextService.InsertBeforeParagraph(clonedTable, searchedParagraph);
                return;
            }

            if (clonedTableRenderModel.PrintOutClonedTableFormatSettings != null &&
                clonedTableRenderModel.PrintOutClonedTableFormatSettings.NumberOfLinesBetweenPages > 0)
            {
                var tablesCount = clonedTableRenderModel.PrintOutClonedTableFormatSettings.NumberOfLinesBetweenPages;
                for (var i = 1; i <= tablesCount; i++)
                {
                    var run = new Run();
                    var paragraph = new Paragraph();
                    paragraph.AddChild(run);
                    searchedParagraph.InsertBeforeSelf(paragraph);
                }
                _winWordTextService.InsertBeforeParagraph(clonedTable, searchedParagraph);
                return;
            }


            if (clonedTableRenderModel.PrintOutClonedTableFormatSettings.IsSetPageBreak)
            {
                InsertPageBreakOnClonedTable(searchedParagraph, clonedTable);
            }

            
        }

        /// <summary>
        /// Insert page break on cloned tables
        /// </summary>
        /// <param name="searchedParagraph"></param>
        /// <param name="clonedTable"></param>
        private void InsertPageBreakOnClonedTable(Paragraph searchedParagraph, Table clonedTable)
        {
            var pageBreak = new Break { Type = BreakValues.Page };
            var newRun = new Run();
            newRun.AddChild(pageBreak);
            var newParagraph = new Paragraph();
            newParagraph.AddChild(newRun);
            searchedParagraph.InsertBeforeSelf(newParagraph);
            _winWordTextService.InsertBeforeParagraph(clonedTable, searchedParagraph);
        }

        /// <summary>
        /// Set table header on each page
        /// </summary>
        /// <param name="table"></param>
        /// <param name="printOutTableFormatSettings"></param>
        /// <exception cref="ServiceException"></exception>
        private static void SetTableHeaderOnEachPage(Table table, PrintOutTableFormatSettings printOutTableFormatSettings)
        {
            if (printOutTableFormatSettings?.RepeatHeadingRowSettings == null)
            {
                return;
            }

            var rows = table.Descendants<TableRow>().ToArray();
            var rowsCount = rows.Length;
            if (rowsCount < printOutTableFormatSettings.RepeatHeadingRowSettings.EndRepeatHeadingRowNumber)
            {
                throw new Exception("LOAN_CORPORATE_APPLICATION_SERVICE/WRONG_END_REPEATING_HEADING_ROW_NUMBER");
            }

            for (var i = printOutTableFormatSettings.RepeatHeadingRowSettings.StartRepeatHeadingRowNumber; i <=
                 printOutTableFormatSettings.RepeatHeadingRowSettings.EndRepeatHeadingRowNumber; i++)
            {
                var tableHeaderRowProps = new TableRowProperties(
                    new CantSplit() { Val = OnOffOnlyValues.On },
                    new TableHeader() { Val = OnOffOnlyValues.On }
                );
                rows[i].AppendChild(tableHeaderRowProps);
            }


        }

        
        /// <summary>
        /// Render inner tables
        /// </summary>
        /// <param name="outerTable"></param>
        /// <param name="tableRenderModel"></param>

        private static void RenderInnerTables(Table outerTable,TableRenderModel tableRenderModel)
        {
            var innerTableStorage = tableRenderModel.InnerTablesStorage;
            if (innerTableStorage == null || innerTableStorage.Count == 0)
            {
                return;
            }

            foreach (var (innerTableTag, innerTableRowsValue) in innerTableStorage)
            {
                foreach (var (innerRowPlaceTag, innerRowsRenderPayload) in innerTableRowsValue)
                {
                    // если нет записей
                    if (innerRowsRenderPayload.Payload == null && innerRowsRenderPayload.IsRemoveEntireOuterRowIfNoRecords)
                    {
                        WinWordTableService.RemoveTableRowByTableAndTableRowsPlaceTagName(outerTable, innerRowPlaceTag);
                        continue;
                    }

                    if (innerRowsRenderPayload.Payload == null && innerRowsRenderPayload.IsRemoveInnerTableIfNoRecords)
                    {
                        var innerTableToRemove = WinWordTableService.GetInnerTableInTableByInnerTableTagName(outerTable, innerTableTag);
                        innerTableToRemove.Remove();
                        continue;
                    }
                   
                    var innerTable = WinWordTableService.GetInnerTableInTableByInnerTableTagName(outerTable, innerTableTag);
                    var innerTableRenderModel = new TableRenderModel();
                    innerTableRenderModel.TableRowsData = innerTableRowsValue.ToDictionary(p => p.Key, s =>
                        new RowsRenderPayload
                        {
                            ColumnNames = s.Value.ColumnNames,
                            RemoveInnerBordersInColumnAbsolutePositions =
                                s.Value.RemoveInnerBordersInColumnAbsolutePositions,
                            Payload = s.Value.Payload,
                            IsRemoveEntireRowIfNoRecords = s.Value.IsRemoveEntireRowIfNoRecords,
                            TableValueOutputFormat = s.Value.TableValueOutputFormat
                        });
                    RenderTableRows(innerTable, innerTableRenderModel);
                    WinWordTableService.RemoveTableTagName(innerTable);
                }
            }
        }

        /// <summary>
        /// Should remove table if no records
        /// </summary>
        /// <param name="tableRenderModel"></param>
        /// <returns></returns>
        private static bool IsRemoveTable(TableRenderModel tableRenderModel)
        {
            if (tableRenderModel?.PrintOutTableFormatSettings?.RemoveTable == null 
                || !tableRenderModel.PrintOutTableFormatSettings.RemoveTable)
            {
                return false;
            }
            

            if (tableRenderModel.PrintOutTableFormatSettings.RemoveTable &&
                (tableRenderModel.InnerTablesStorage == null || tableRenderModel.InnerTablesStorage.Count == 0) &&
                (tableRenderModel.TextReplaceData == null || tableRenderModel.TextReplaceData.Count == 0) &&
                (tableRenderModel.TableRowsData == null || tableRenderModel.TableRowsData.Count == 0) &&
                tableRenderModel.TextGeneratorReplaceData == null || tableRenderModel.TextGeneratorReplaceData.Count == 0)
            {
                return true;
            }

            bool isRemoveTableValues;
            bool isRemoveTableTextReplacement;
            
            // проверка всех TableValue
            var isRemoveInnerTableValues = new Collection<bool>();
            if (tableRenderModel.InnerTablesStorage?.Count > 0)
            {
                foreach (var (_, innerTableValue) in tableRenderModel.InnerTablesStorage)
                {
                    isRemoveInnerTableValues.Add(innerTableValue.Select(p => p.Value?.Payload).All(s => s == null));
                }
            }
            else
            {
                isRemoveInnerTableValues.Add(true);
            }

            if (tableRenderModel.TextReplaceData?.Count > 0)
            {
                isRemoveTableTextReplacement =
                    tableRenderModel.TextReplaceData.Select(p => p.Value?.Value).All(s => s == null);
            }
            else
            {
                isRemoveTableTextReplacement = true;
            }

            if (tableRenderModel.TableRowsData.Count > 0)
            {
                isRemoveTableValues = tableRenderModel.TableRowsData.Select(p => p.Value?.Payload).All(s => s == null);
            }
            else
            {
                isRemoveTableValues = true;
            }

            if (isRemoveTableValues && isRemoveTableTextReplacement && isRemoveInnerTableValues.All(p => true) &&
                tableRenderModel.PrintOutTableFormatSettings.RemoveTable)
            {
                return true;
            }

            return false;
        }
        
    }