using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlClient.Models.Text;
using OpenXmlServer.WinWordRenderer.TextRenderer;

namespace OpenXmlServer.WinWordRenderer.TableRenderer;

public class WinWordTableService : OpenXmlElementService
{
    private readonly WinWordTextService _winWordTextService;

    public WinWordTableService(WinWordTextService winWordTextService, Body body) : base(body)
    {
        _winWordTextService = winWordTextService;
    }

    /// <summary>
    /// Get column number and Runs list
    /// </summary>
    /// <param name="tableRow"></param>
    /// <param name="tagName"></param>
    /// <returns></returns>
    public static (int, IEnumerable<Run>) GetColumnNumberAndRunContainsTag(TableRow tableRow, string tagName)
    {
        var cells = tableRow.Descendants<TableCell>().ToList();
        for (var i = 0; i < cells.Count; i++)
        {
            var paragraphs = cells[i].Descendants<Paragraph>().ToList();
            if (paragraphs.Count <= 0)
            {
                continue;
            }

            foreach (var paragraph in paragraphs)
            {
                var result = WinWordTextService.GetRunContainsKey(paragraph, tagName).ToList();
                if (result.Count > 0)
                {
                    return (i, result);
                }
            }
        }

        return (-1, new List<Run>());
    }

    /// <summary>
    /// Get tableRow, Column index and Text by tagName
    /// </summary>
    /// <param name="table"></param>
    /// <param name="tableRowsPlaceTagName"></param>
    /// <returns></returns>
    public static (int, IEnumerable<TableRow>, IEnumerable<Run>) GetTableRowColumnIndexAndTextByTagName(Table table,
        string tableRowsPlaceTagName)
    {
        var tableRows = new List<TableRow>();
        foreach (var row in table.Descendants<TableRow>())
        {
            // находим tagName в строке
            var (columnNumber, runs) = GetColumnNumberAndRunContainsTag(row, tableRowsPlaceTagName);
            if (columnNumber != -1)
            {
                tableRows.Add(row);
                return (columnNumber, tableRows, runs);
            }
        }

        // не нашли
        throw new Exception($"LOAN_CORPORATE_PRINT_FORM_SERVICE/NOT_FOUND_TABLE_ROW_TAG: {tableRowsPlaceTagName}");
    }


    /// <summary>
    /// Get table by table tag name
    /// </summary>
    /// <param name="tableTagName"></param>
    /// <returns></returns>
    public IEnumerable<Table> GetTablesByTableTagName(string tableTagName)
    {
        var tables = new List<Table>();
        var tableList = _winWordTextService.GetOpenXmlElement<Table>().ToList();
        if (tableList.Count == 0)
        {
            return tables;
        }

        foreach (var table in tableList)
        {
            // ищем только в первой строке
            var row = table.GetFirstChild<TableRow>();
            if (row?.InnerText == tableTagName)
            {
                tables.Add(table);
                return tables;
            }
        }

        return tables;
    }


    /// <summary>
    /// Insert Tag Table Name at table first row
    /// </summary>
    /// <param name="tableTagName"></param>
    /// <param name="table"></param>
    public static void InsertTableTagName(string tableTagName, Table table)
    {
        var text = new Text();
        text.Text = tableTagName;
        var run = new Run();
        run.AppendChild(text);
        var paragraph = new Paragraph();
        paragraph.AppendChild(run);
        var tableCell = new TableCell();
        tableCell.AppendChild(paragraph);
        var tableRow = new TableRow();
        tableRow.AppendChild(tableCell);
        var row = table.GetFirstChild<TableRow>();
        row?.InsertBeforeSelf(tableRow);
    }

    /// <summary>
    /// Remove Table tag name with first row
    /// </summary>
    /// <param name="table"></param>
    public static void RemoveTableTagName(Table table)
    {
        var row = table.GetFirstChild<TableRow>();
        table.RemoveChild(row);
    }

    /// <summary>
    /// Remove table row by table and tableRowsPlaceTagName
    /// </summary>
    /// <param name="table"></param>
    /// <param name="tableRowsPlaceTagName"></param>
    public static void RemoveTableRowByTableAndTableRowsPlaceTagName(Table table,
        string tableRowsPlaceTagName)
    {
        if (table == null)
        {
            return;
        }

        var (_, tableRows, _) = GetTableRowColumnIndexAndTextByTagName(table, tableRowsPlaceTagName);
        var tableRow = tableRows.FirstOrDefault();
        if (tableRow != null)
        {
            table.RemoveChild(tableRow);
        }
    }

    /// <summary>
    /// Clear TableRowsPlaceTagName from table without remove row
    /// </summary>
    /// <param name="table"></param>
    /// <param name="tableRowsPlaceTagName"></param>
    public static void ClearTableRowsPlaceTagName(Table table,
        string tableRowsPlaceTagName)
    {
        if (table == null)
        {
            return;
        }

        var rows = table.Descendants<TableRow>().ToList();
        if (rows.Count > 0)
        {
            foreach (var row in rows)
            {
                ReplaceTextInRow(row, tableRowsPlaceTagName, "");
            }
        }
    }


    /// <summary>
    /// Remove table by tag name
    /// </summary>
    /// <param name="tableTagName"></param>
    public void RemoveTableByTagName(string tableTagName)
    {
        var table = GetTablesByTableTagName(tableTagName).FirstOrDefault();
        if (table != null)
        {
            RemoveOpenXmlElement(table);
        }
    }

    /// <summary>
    /// Get row with text tag
    /// </summary>
    /// <param name="tableRows"></param>
    /// <param name="tag"></param>
    /// <returns></returns>
    public static IEnumerable<TableRow> GetRowsWithTextTag(IEnumerable<TableRow> tableRows, string tag) =>
        (from tableRow in tableRows
            let paragraphs = tableRow.Descendants<Paragraph>()
            where WinWordTextService.IsParagraphsContainsKey(paragraphs, tag)
            select tableRow).ToList();


    /// <summary>
    /// Replace Text in Row
    /// </summary>
    /// <param name="row"></param>
    /// <param name="textTagName"></param>
    /// <param name="replaceText"></param>
    public static void ReplaceTextInRow(TableRow row, string textTagName, string replaceText)
    {
        if (row == null)
        {
            return;
        }

        var cells = row.Descendants<TableCell>().ToList();
        if (cells.Count == 0)
        {
            return;
        }

        foreach (var cell in cells)
        {
            var paragraphs = cell.Descendants<Paragraph>().ToList();
            if (paragraphs.Count == 0)
            {
                continue;
            }

            WinWordTextService.ReplaceText(paragraphs, textTagName, replaceText);
        }
    }


    public void ReplaceTagTextGeneratorInRow(TableRow row, string textTagName, ICollection<RunModel> runList)
    {
        if (row == null)
        {
            return;
        }

        if (runList.Count == 0)
        {
            // удаляем строку если нету
            row.Remove();
            return;
        }

        var cells = row.Descendants<TableCell>().ToList();
        if (cells.Count == 0)
        {
            return;
        }

        foreach (var cell in cells)
        {
            var paragraphs = cell.Descendants<Paragraph>().ToList();
            if (paragraphs.Count == 0)
            {
                continue;
            }

            var paragraph = WinWordTextService.GetParagraphsByTagName(paragraphs, textTagName)?.FirstOrDefault();
            if (paragraph == null)
            {
                continue;
            }
            // Вставляем
            foreach (var runModel in runList)
            {
                _winWordTextService.CreateRunInParagraph(paragraph, runModel, runModel.OutputFormat);
            }
            // удаляем тег если есть
            _winWordTextService.GlobalReplaceText(textTagName,"");
            return;
        }

        throw new Exception($"NOT_FOUND_TEXT_GENRATOR_TAG/{textTagName}");
    }


/// <summary>
/// Replace table tag name
/// </summary>
/// <param name="tableTagName"></param>
/// <param name="newTableTagName"></param>
public void ReplaceTableTagName(string tableTagName, string newTableTagName)
{
    var table = GetTablesByTableTagName(tableTagName).FirstOrDefault();
    if (table != null)
    {
        var firstRow = table.Descendants<TableRow>().FirstOrDefault();
        ReplaceTextInRow(firstRow, tableTagName, newTableTagName);
    }
}


public static void RemoveParagraphInCells(IEnumerable<TableCell> cells, int index)
{
    var cellsList = cells.ToList();
    if (!cellsList[index].Any())
    {
        return;
    }

    var paragraphs = cellsList[index].Descendants<Paragraph>().ToList();
    foreach (var paragraph in paragraphs)
    {
        paragraph.Remove();
    }
}

/// <summary>
/// Remove tag in row
/// </summary>
/// <param name="row"></param>
/// <param name="tagName"></param>
public static void RemoveTagInRow(TableRow row, string tagName)
{
    var cells = row.Descendants<TableCell>().ToList();
    foreach (var paragraphs in cells.Select(cell => cell.Descendants<Paragraph>()))
    {
        WinWordTextService.ReplaceText(paragraphs, tagName, "");
    }
}

/// <summary>
/// Remove Inner border in columns
/// </summary>
/// <param name="row"></param>
/// <param name="rowIndex"></param>
///   /// <param name="rowsCount"></param>
/// <param name="removeInnerBordersInColumnAbsolutePositions"></param>
public static void RemoveInnerBorderInColumns(TableRow row, int rowIndex, int rowsCount,
    ICollection<int> removeInnerBordersInColumnAbsolutePositions)
{
    if (removeInnerBordersInColumnAbsolutePositions == null)
    {
        return;
    }

    var cells = row.Descendants<TableCell>().ToList();
    foreach (var cellIndex in removeInnerBordersInColumnAbsolutePositions)
    {
        var cellItem = cells[cellIndex - 1];
        if (cellItem.TableCellProperties == null)
        {
            cellItem.TableCellProperties = new TableCellProperties();
        }

        if (cellItem.TableCellProperties.TableCellBorders == null)
        {
            cellItem.TableCellProperties.TableCellBorders = new TableCellBorders();
        }

        // убираем нижний бордер
        if (rowIndex != rowsCount - 1)
        {
            cellItem.TableCellProperties.TableCellBorders.BottomBorder = new BottomBorder();
            cellItem.TableCellProperties.TableCellBorders.BottomBorder.Val =
                new EnumValue<BorderValues>(BorderValues.Nil);
            cellItem.TableCellProperties.TableCellBorders.TopBorder = new TopBorder();
            cellItem.TableCellProperties.TableCellBorders.TopBorder.Val =
                new EnumValue<BorderValues>(BorderValues.Nil);
        }
        else
        {
            cellItem.TableCellProperties.TableCellBorders.TopBorder = new TopBorder();
            cellItem.TableCellProperties.TableCellBorders.TopBorder.Val =
                new EnumValue<BorderValues>(BorderValues.Nil);
        }


        if (rowIndex == 0)
        {
            continue;
        }

        // чистим Runs в  параграфах
        var paragraphs = cells[cellIndex - 1].Descendants<Paragraph>().ToList();
        foreach (var paragraph in paragraphs)
        {
            paragraph.RemoveAllChildren();
        }
    }
}

/// <summary>
/// Get inner table from outerTable by inner table tag
/// </summary>
/// <param name="outerTable"></param>
/// <param name="innerTableTag"></param>
/// <returns></returns>
/// <exception cref="ServiceException"></exception>
public static Table GetInnerTableInTableByInnerTableTagName(Table outerTable, string innerTableTag)
{
    var tableRows = outerTable.Descendants<TableRow>();
    foreach (var tableRow in tableRows)
    {
        var tables = GetTableFromCells(tableRow.Descendants<TableCell>(), innerTableTag).ToList();
        if (tables.Count > 0)
        {
            return tables.FirstOrDefault();
        }
    }

    throw new Exception($"LOAN_CORPORATE_PRINT_FORM_SERIVCE/NOT_FOUND_INNER_TABLE_TAG: {innerTableTag}");
}


/// <summary>
/// Get table from cells if exist
/// </summary>
/// <param name="cells"></param>
/// <param name="innerTableTag"></param>
/// <returns></returns>
private static IEnumerable<Table> GetTableFromCells(IEnumerable<TableCell> cells, string innerTableTag)
{
    var cellList = cells.ToList();
    return (from cell in cellList
        select cell.Descendants<Table>().FirstOrDefault()
        into table
        where table != null
        where GetRowsWithTextTag(table.Descendants<TableRow>(), innerTableTag).FirstOrDefault() != null
        select table).ToList();
}

}