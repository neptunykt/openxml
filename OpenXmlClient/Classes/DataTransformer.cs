using System.Collections.ObjectModel;
using OpenXmlClient.Classes.Models;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.HeaderFooter;
using OpenXmlClient.Models.Table;
using OpenXmlClient.Models.Text;

#pragma warning disable CS8601
#pragma warning disable CS8625
#pragma warning disable CS8604

namespace OpenXmlClient.Classes;

public class DataTransformer
{
    private readonly WinWordRenderModel _winWordRenderModel;

    /// <summary>
    /// Return transformed model
    /// </summary>
    /// <returns></returns>
    public WinWordRenderModel GetWinWordModel() => _winWordRenderModel;

    public DataTransformer(OutputFormat globalOutputFormat, TableValueOutputFormat tableValueOutputFormat)
    {
        _winWordRenderModel = new WinWordRenderModel();
        if (!CheckFillingGlobalFormat(globalOutputFormat))
        {
            throw new Exception("LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FILLED_GLOBAL_OUTPUT_FORMAT");
        }

        if (!CheckFillingValueGlobalFormat(tableValueOutputFormat))
        {
            throw new Exception("LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FILLED_VALUE_GLOBAL_OUTPUT_FORMAT");
        }

        WinWordRenderModel.OutputFormat = globalOutputFormat;
        WinWordRenderModel.TableValueOutputFormat = tableValueOutputFormat;
    }

    /// <summary>
    /// Check filling global output format
    /// </summary>
    /// <param name="outputFormat"></param>
    /// <returns></returns>
    private static bool CheckFillingGlobalFormat(OutputFormat outputFormat) =>
        !string.IsNullOrEmpty(outputFormat?.DecimalFormat) &&
        !string.IsNullOrEmpty(outputFormat.FontColor) &&
        !string.IsNullOrEmpty(outputFormat.FontName) &&
        !string.IsNullOrEmpty(outputFormat.DateTimeFormat) &&
        outputFormat.IsBold != null &&
        outputFormat.IsItalic != null;


    /// <summary>
    /// Check filling global output format
    /// </summary>
    /// <param name="tableValueOutputFormat"></param>
    /// <returns></returns>
    private static bool CheckFillingValueGlobalFormat(TableValueOutputFormat tableValueOutputFormat) =>
        !string.IsNullOrEmpty(tableValueOutputFormat?.DecimalFormat) &&
        !string.IsNullOrEmpty(tableValueOutputFormat.DateTimeFormat);

    /// <summary>
    /// Fill Cloned table table rows
    /// </summary>
    /// <param name="clonedTableStorages"></param>
    /// <param name="clonedTableTagName"></param>
    /// <param name="insertTablePlaceTagName"></param>
    /// <param name="clonedTableFormatSettings"></param>
    public void FillClonedTable(IEnumerable<TableFillStorageModel> clonedTableStorages,
        string clonedTableTagName, string insertTablePlaceTagName,
        ClonedTableFormatSettings clonedTableFormatSettings = null)
    {
        if (_winWordRenderModel.ClonedTablesStorage == null)
        {
            _winWordRenderModel.ClonedTablesStorage = new Dictionary<string, ClonedTableRenderModel>();
        }

        // тут только одна таблица заполняется
        var index = 1;
        foreach (var clonedTableStorage in clonedTableStorages)
        {
            var copyTableRenderModel = new ClonedTableRenderModel();
            copyTableRenderModel.InsertTablePlaceTagName = insertTablePlaceTagName;
            copyTableRenderModel.ClonedTableTagName = clonedTableTagName;
            copyTableRenderModel.CopyTableRenderStorage = new TableRenderModel();
            copyTableRenderModel.CopyTableRenderStorage.TableRowsData =
                new Dictionary<string, RowsRenderPayload>();
            copyTableRenderModel.CopyTableRenderStorage.TextReplaceData =
                new Dictionary<string, TableTextModel>();
            copyTableRenderModel.CopyTableRenderStorage.TextGeneratorReplaceData =
                new Dictionary<string, Collection<RunModel>>();
            copyTableRenderModel.PrintOutClonedTableFormatSettings =
                clonedTableFormatSettings?.PrintOutClonedTableFormatSettings;
            // Проставляем 
            if (clonedTableStorage.TableRowsFillStorage.Count > 0)
            {
                foreach (var (tableRowsTagNameKey, value) in clonedTableStorage.TableRowsFillStorage)
                {
                    var rowsRenderPayLoad =
                        TableRowDataTableSerializer.FillRows(value,
                            clonedTableFormatSettings?.TableValueOutputFormat);
                    copyTableRenderModel.CopyTableRenderStorage.TableRowsData.Add(tableRowsTagNameKey,
                        rowsRenderPayLoad);
                }
            }


            if (clonedTableStorage.TableFillTextReplaceStorage != null &&
                clonedTableStorage.TableFillTextReplaceStorage.Count > 0)
            {
                // ставим формат 
                FillTableLocalOutputFormat(clonedTableStorage.TableFillTextReplaceStorage,
                    clonedTableFormatSettings?.OutputFormat);
                FillTableLocalOutputFormat(clonedTableStorage.TableFillTextReplaceStorage,
                    WinWordRenderModel.OutputFormat);
                foreach (var (key, value) in clonedTableStorage.TableFillTextReplaceStorage)
                {
                    copyTableRenderModel.CopyTableRenderStorage.TextReplaceData.Add(key, value);
                }
            }

            if (clonedTableStorage.TableFillTextGeneratorReplaceStorage != null &&
                clonedTableStorage.TableFillTextGeneratorReplaceStorage.Count > 0)
            {
                FillTextGeneratorOutputFormat(clonedTableStorage.TableFillTextGeneratorReplaceStorage);
                foreach (var (key, value) in clonedTableStorage.TableFillTextGeneratorReplaceStorage)
                {
                    copyTableRenderModel.CopyTableRenderStorage.TextGeneratorReplaceData.Add(key, value);
                }
            }

            // Проставка уникальной cloned tableTagName
            _winWordRenderModel.ClonedTablesStorage.Add($"{clonedTableTagName}{index}", copyTableRenderModel);
            FillClonedInnerTables(clonedTableStorage.InnerTableFillStorage, $"{clonedTableTagName}{index}",
                clonedTableFormatSettings);
            index++;
        }
    }


    /// <summary>
    /// Fill table table rows
    /// </summary>
    /// <param name="tableStorage"></param>
    /// <param name="tableTagName"></param>
    /// <param name="tableFormatSettings"></param>
    public void FillTable(TableFillStorageModel tableStorage, string tableTagName,
        TableFormatSettings tableFormatSettings = null)
    {
        if (_winWordRenderModel.TablesStorage == null)
        {
            _winWordRenderModel.TablesStorage = new Dictionary<string, TableRenderModel>();
        }

        var tableRenderModel = new TableRenderModel();
        tableRenderModel.PrintOutTableFormatSettings = tableFormatSettings?.PrintOutTableFormatSettings;
        foreach (var (tableRowTagNameKey, tableRowFillValue) in tableStorage.TableRowsFillStorage)
        {
            var payload =
                TableRowDataTableSerializer.FillRows(tableRowFillValue,
                    tableFormatSettings?.TableValueOutputFormat);

            if (tableRenderModel.TableRowsData == null)
            {
                tableRenderModel.TableRowsData = new Dictionary<string, RowsRenderPayload>();
            }


            tableRenderModel.TableRowsData.Add(tableRowTagNameKey, payload);
        }

        if (tableRenderModel.TextReplaceData == null)
        {
            tableRenderModel.TextReplaceData = new Dictionary<string, TableTextModel>();
        }

        foreach (var (replaceTagNameKey, replaceText) in tableStorage.TableFillTextReplaceStorage)
        {
            // замена формата
            if (replaceText.TableValueOutputFormat == null)
            {
                replaceText.TableValueOutputFormat = new TableValueOutputFormat();
            }

            FillTableTextReplacementLocalOutputFormat(replaceText, tableFormatSettings?.TableValueOutputFormat);
            FillTableTextReplacementLocalOutputFormat(replaceText, WinWordRenderModel.TableValueOutputFormat);

            tableRenderModel.TextReplaceData.Add(replaceTagNameKey, replaceText);
        }
        FillTableTextGeneratorStorage(tableStorage.TableFillTextGeneratorReplaceStorage);
        tableRenderModel.TextGeneratorReplaceData = tableStorage.TableFillTextGeneratorReplaceStorage;
        _winWordRenderModel.TablesStorage.Add(tableTagName, tableRenderModel);
    }

    /// <summary>
    /// Fill global text replacement storage
    /// </summary>
    /// <param name="globalTextReplacement"></param>
    public void FillGlobalTextReplacement(IDictionary<string, TextModel> globalTextReplacement)
    {
        if (_winWordRenderModel != null)
        {
            FillGlobalDictionaryOutputFormat(globalTextReplacement);
            _winWordRenderModel.GlobalTextReplacementStorage = globalTextReplacement;
        }
    }

    private static void FillGlobalDictionaryOutputFormat(IDictionary<string, TextModel> globalTextReplacement)
    {
        if (globalTextReplacement == null || globalTextReplacement.Count <= 0)
        {
            return;
        }

        foreach (var (_, value) in globalTextReplacement)
        {
            if (value.OutputFormat == null)
            {
                value.OutputFormat = new OutputFormat();
            }

            OutputFormat.SetOutputFormat(WinWordRenderModel.OutputFormat, value.OutputFormat);
        }
    }

    /// <summary>
    /// Fill header footer storage
    /// </summary>
    /// <param name="headerFooterStorage"></param>
    public void FillHeaderFooterReplacement(HeaderFooterStorage headerFooterStorage)
    {
        if (_winWordRenderModel == null)
        {
            return;
        }

        FillGlobalDictionaryOutputFormat(headerFooterStorage.FooterReplaceDictionary);
        FillGlobalDictionaryOutputFormat(headerFooterStorage.HeaderReplaceDictionary);
        _winWordRenderModel.HeaderFooterStorage = headerFooterStorage;
    }

    /// <summary>
    /// Fill Global text generator storage
    /// </summary>
    /// <param name="textGeneratorStorage"></param>
    public void FillGlobalTextGeneratorStorage(IDictionary<string, IEnumerable<RunModel>> textGeneratorStorage)
    {
        if (_winWordRenderModel == null)
        {
            return;
        }

        foreach (var (_, list) in textGeneratorStorage)
        {
            foreach (var item in list)
            {
                item.OutputFormat ??= new OutputFormat();

                OutputFormat.SetOutputFormat(WinWordRenderModel.OutputFormat, item.OutputFormat);
            }
        }

        _winWordRenderModel.GlobalTextGeneratorStorage = textGeneratorStorage;
    }


    public void FillNumberingTextStorage(IDictionary<string, ICollection<Collection<RunModel>>> numberingTextStorage)
    {
        if (_winWordRenderModel == null)
        {
            return;
        }

        foreach (var (_, list) in numberingTextStorage)
        {
            foreach (var items in list)
            {
                foreach (var item in items)
                {
                    item.OutputFormat ??= new OutputFormat();
                    OutputFormat.SetOutputFormat(WinWordRenderModel.OutputFormat, item.OutputFormat);
                }
            }
        }

        _winWordRenderModel.NumberingTextStorage = numberingTextStorage;
    }

    private void FillTableTextGeneratorStorage(IDictionary<string, Collection<RunModel>> tableTextGeneratorStorage)
    {
        if (tableTextGeneratorStorage == null || tableTextGeneratorStorage.Count == 0)
        {
            return;
        }
        foreach (var (textGeneratorKey, _) in tableTextGeneratorStorage)
        {
            if (!tableTextGeneratorStorage.ContainsKey(textGeneratorKey))
            {
                throw new Exception($"NOT_FOUND_TABLE_TEXT_GENERATOR_TAG/{textGeneratorKey}");
            }
            
        }

        FillTextGeneratorOutputFormat(tableTextGeneratorStorage);

    }

    private void FillTextGeneratorOutputFormat(IDictionary<string, Collection<RunModel>> textGeneratorStorage)
    {
        if (textGeneratorStorage == null || textGeneratorStorage.Count == 0)
        {
            return;
        }

        foreach (var (_, list) in textGeneratorStorage)
        {
            foreach (var item in list)
            {
                item.OutputFormat ??= new OutputFormat();

                OutputFormat.SetOutputFormat(WinWordRenderModel.OutputFormat, item.OutputFormat);
            }
        }
    }

    private static void FillTableLocalOutputFormat(IDictionary<string, TableTextModel> textReplacement,
        OutputFormat parentOutputFormat)
    {
        foreach (var (_, textModel) in textReplacement)
        {
            if (textModel.TableValueOutputFormat == null)
            {
                textModel.TableValueOutputFormat = new TableValueOutputFormat();
            }

            TableValueOutputFormat.SetOutputFormat(parentOutputFormat, textModel.TableValueOutputFormat);
        }
    }

    private static void FillTableTextReplacementLocalOutputFormat(TableTextModel tableTextReplacement,
        TableValueOutputFormat parentOutputFormat) =>
        TableValueOutputFormat.SetOutputFormat(parentOutputFormat, tableTextReplacement.TableValueOutputFormat);


    /// <summary>
    /// Fill Cloned inner tables
    /// </summary>
    /// <param name="innerTableStorage"></param>
    /// <param name="clonedTableTagName"></param>
    /// <param name="clonedTableFormatSettings"></param>
    private void FillClonedInnerTables(
        IDictionary<string, IDictionary<string, InnerTableRowFillModel>> innerTableStorage,
        string clonedTableTagName,
        ClonedTableFormatSettings clonedTableFormatSettings = null)
    {
        if (innerTableStorage == null || innerTableStorage.Count == 0)
        {
            return;
        }

        // если нет таблицы, то создаем
        if (!_winWordRenderModel.ClonedTablesStorage.ContainsKey(clonedTableTagName))
        {
            _winWordRenderModel.ClonedTablesStorage.Add(clonedTableTagName, new ClonedTableRenderModel());
        }


        var clonedTableRenderModel = _winWordRenderModel.ClonedTablesStorage[clonedTableTagName];
        if (clonedTableRenderModel.InnerTableFillStorage == null)
        {
            clonedTableRenderModel.InnerTableFillStorage = new Dictionary<string, TableRenderModel>();
        }

        var innerTableRenderModel = new TableRenderModel();
        innerTableRenderModel.InnerTablesStorage =
            new Dictionary<string, IDictionary<string, InnerRowsRenderPayload>>();
        foreach (var (innerTableTag, tableRowStorage) in innerTableStorage)
        {
            if (tableRowStorage == null)
            {
                continue;
            }

            clonedTableRenderModel.OutputFormat = clonedTableFormatSettings?.OutputFormat;
            var innerRowsDictionary = new Dictionary<string, InnerRowsRenderPayload>();

            foreach (var (rowKey, rowFillValue) in tableRowStorage)
            {
                var innerRowsRenderPayLoad = new InnerRowsRenderPayload();
                var rowsRenderPayLoad = TableRowDataTableSerializer.FillRows(rowFillValue,
                    clonedTableFormatSettings?.TableValueOutputFormat);
                innerRowsRenderPayLoad.Payload = rowsRenderPayLoad.Payload;
                innerRowsRenderPayLoad.ColumnNames = rowsRenderPayLoad.ColumnNames;
                innerRowsRenderPayLoad.TableValueOutputFormat = rowsRenderPayLoad.TableValueOutputFormat;
                innerRowsRenderPayLoad.IsRemoveEntireRowIfNoRecords = rowsRenderPayLoad.IsRemoveEntireRowIfNoRecords;
                innerRowsRenderPayLoad.RemoveInnerBordersInColumnAbsolutePositions =
                    rowsRenderPayLoad.RemoveInnerBordersInColumnAbsolutePositions;
                innerRowsRenderPayLoad.IsRemoveInnerTableIfNoRecords = rowFillValue.IsRemoveInnerTableIfNoRecords;
                innerRowsRenderPayLoad.IsRemoveEntireOuterRowIfNoRecords =
                    rowFillValue.IsRemoveEntireOuterRowIfNoRecords;
                innerRowsDictionary.Add(rowKey, innerRowsRenderPayLoad);
            }

            innerTableRenderModel.InnerTablesStorage.Add(innerTableTag, innerRowsDictionary);
        }

        clonedTableRenderModel.InnerTableFillStorage.Add(clonedTableTagName, innerTableRenderModel);
    }
}