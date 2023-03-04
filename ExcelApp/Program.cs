// See https://aka.ms/new-console-template for more information

using System.Data;
using ExcelApp;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

var type = 16;

var workBook = NPOIHelper.GetWorkbook(@$"d:\Downloads\{type}.xlsx");
var dictColumns = new Dictionary<string, string>();
var dictCustomColumns = new SortedDictionary<string, string>();

var dtList = new List<DataTable>();
for (int i = 0; i < workBook.NumberOfSheets; i++)
{
    var sheet = workBook.GetSheetAt(i);
    var row = sheet.GetRow(0);
    var lastCellNum = row.LastCellNum;
    // var columnCount = row.GetCell()
    
    var dt = GetDataTable(workBook, i, ref dictColumns, ref dictCustomColumns);
    dtList.Add(dt);
}

foreach (var item in dictCustomColumns)
{
    dictColumns.Add(item.Key, item.Value);
}

Console.WriteLine(dictColumns.Keys);
Console.WriteLine(dictColumns.Values);

var newWorkBook = ToWorkbook(dtList);
using (FileStream fs = new FileStream(@$"d:\Downloads\{type}_new.xlsx", FileMode.Create, FileAccess.Write))
{
    newWorkBook.Write(fs);
}

Console.WriteLine("Hello, World!");

DataTable GetDataTable(IWorkbook book, int index, ref Dictionary<string, string> dictCommonColumn, ref SortedDictionary<string, string> dictCustomColumn)
{
    DataTable table = new DataTable();
    ISheet sheet = book.GetSheetAt(index);
    IRow headerRow = sheet.GetRow(0);//读取第一行（头）
    var secondRow = sheet.GetRow(1);
    if (headerRow == null)
        return null;

    var cellIndexList = new List<int>();

    //列头
    int cellCount = headerRow.LastCellNum;
    for (int i = 1; i < cellCount; i++)
    {
        var cellValue = headerRow.GetCell(i).ToString()?.Trim() ?? "";
        var columnDesc = secondRow.GetCell(i).ToString()?.Trim() ?? "";

        if (string.IsNullOrEmpty(cellValue))
        {
            continue;
        }
        if (cellValue == "SHTXT40" && columnDesc.StartsWith("英文描述"))
        {
            cellValue = "SHTXT40_EN";
        }

        if (cellValue == "KLASSE18" && columnDesc.StartsWith("设备自定义分类"))
        {
            cellValue = "KLASSE18_CUSTOM";
        }

        if (cellValue.StartsWith("PMC") || cellValue.StartsWith("PMN"))
        {
            dictCustomColumn.TryAdd(cellValue, columnDesc);
        }
        else
        {
            dictCommonColumn.TryAdd(cellValue, columnDesc);
        }

        DataColumn column = new DataColumn(cellValue);
        table.Columns.Add(column);

        cellIndexList.Add(i);
    }

    //行内容
    for (int i = 5; i <= sheet.LastRowNum; i++)
    {
        IRow row = sheet.GetRow(i);
        if (row == null)
            continue;
        DataRow dataRow = table.NewRow();

        var columnIndex = 0;
        foreach (var cellIndex in cellIndexList)
        {
            dataRow[columnIndex++] = row.GetValue(cellIndex);
        }

        table.Rows.Add(dataRow);
    }
    return table;
}

IWorkbook ToWorkbook(List<DataTable> dataList)
{
    var book = new XSSFWorkbook();

    ISheet sheet = book.CreateSheet();
    IRow headerRow = sheet.CreateRow(0);
    IRow secondRow = sheet.CreateRow(1);

    var dictColumnIndex = new Dictionary<string, int>();

    int index = 0;
    foreach (var item in dictColumns)
    {
        headerRow.CreateCell(index).SetCellValue(item.Key);
        secondRow.CreateCell(index).SetCellValue(item.Value);

        dictColumnIndex.Add(item.Key, index);

        index++;
    }

    var rowIndex = 2;
    foreach (var dataTable in dataList)
    {
        for (int i = 0; i < dataTable.Rows.Count; i++)
        {
            DataRow row = dataTable.Rows[i];
            IRow dataRow = sheet.CreateRow(rowIndex);
            for (int j = 0; j < dataTable.Columns.Count; j++)
            {
                var columnName = dataTable.Columns[j].ColumnName;
                if (dictColumnIndex.ContainsKey(columnName))
                {
                    dataRow.CreateCell(dictColumnIndex[columnName]).SetCellValue(row[j].ToString()?.Trim());
                }
            }

            rowIndex++;
        }
    }

    return book;
}