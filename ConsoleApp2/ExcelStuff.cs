using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ConsoleApp2;

public static class ExcelStuff
{
    public static DataTable ReadExcelToDataTable(string fileName, string? sheetName = null)
    {
        DataTable dataTable = new DataTable();
        Console.WriteLine(DateTime.Now);
        using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
        {
            XSSFWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet sheet = string.IsNullOrWhiteSpace(sheetName) ? workbook.GetSheetAt(0) : workbook.GetSheet(sheetName);

            // Get header row
            IRow headerRow = sheet.GetRow(0);
            foreach (ICell headerCell in headerRow.Cells)
            {
                dataTable.Columns.Add(headerCell.ToString());
            }

            // Get data rows
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = dataTable.NewRow();

                for (int j = row.FirstCellNum; j < headerRow.Cells.Count; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        dataRow[j] = row.GetCell(j).ToString();
                    }
                }

                dataTable.Rows.Add(dataRow);
            }
        }

        Console.WriteLine(DateTime.Now);

        var firstrow = dataTable.Rows[dataTable.Rows.Count - 1].ItemArray;

        foreach (var o in firstrow)
        {
            Console.WriteLine(o);
        }

        Console.WriteLine($"DataTable has {dataTable.Rows.Count} records");
        Console.WriteLine(DateTime.Now);
        return dataTable;
    }

    public static DataTable GenerateDataTable(int records = 10)
    {
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("Column1");
        dataTable.Columns.Add("Column2");
        dataTable.Columns.Add("Column3");
        dataTable.Columns.Add("cheese");
        dataTable.Columns.Add("crackers");
        for (int i = 0; i < records; i++)
        {
            DataRow dataRow = dataTable.NewRow();
            dataRow[0] = $"A{i}";
            dataRow[1] = $"B{i}";
            dataRow[2] = $"C{i}";
            dataRow[3] = $"D{i}";
            dataRow[4] = $"E{i}";
            dataTable.Rows.Add(dataRow);
        }

        return dataTable;
    }
}