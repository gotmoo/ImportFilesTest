using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Text;
using Microsoft.Data.SqlClient;
using NPOI.HSSF.Record;
using NPOI.OpenXmlFormats.Shared;
using static ConsoleApp2.DbFunctions;
using static ConsoleApp2.ExcelStuff;

//...
var inFile = @"c:\PowerShell\DlpExcel.xlsx";
string? sheetName = "";
Console.WriteLine(Path.GetExtension(inFile).ToLowerInvariant());

var tableName = Path.GetFileNameWithoutExtension(inFile);

DataTable? dataTable;
switch (Path.GetExtension(inFile).ToLowerInvariant())
{
    case ".csv":
        dataTable = ReadCsvFile(inFile);
        
        break;
    case ".xlsx":
        dataTable = ReadExcelToDataTable(inFile, sheetName);
        break;
    default: return;
}


Console.WriteLine(inFile);
Console.WriteLine($"Reading Excel into Datatable: {DateTime.Now}");
//var myDt = ReadExcelToDataTable(inFile);


Console.WriteLine($"Checking columns and database: {DateTime.Now}");

var columnMapping = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
foreach (DataColumn myDtColumn in dataTable.Columns)
{
    columnMapping[(myDtColumn.ColumnName)] = "";
}

const string connectionString = "server=(LocalDb)\\MSSQLLocalDB;database=RandomTest;Integrated Security=SSPI;";

columnMapping = MapColumnsToDb(tableName, connectionString, columnMapping);
Console.WriteLine($"Uploading data: {DateTime.Now}");

BulkUploadDatatable(tableName, dataTable, connectionString, columnMapping);
Console.WriteLine($"Done: {DateTime.Now}");

static DataTable ReadCsvFile(string inFile)
{
// Create a new DataTable
    DataTable table = new DataTable();

// Open the CSV file using a StreamReader
    using StreamReader reader = new StreamReader(inFile);
// Read the header row and add columns to the DataTable
    string[] headers = reader.ReadLine().Split(',');
    foreach (string header in headers)
    {
        table.Columns.Add(header);
    }

    // Read the remaining rows and add them to the DataTable
    while (!reader.EndOfStream)
    {
        string[] rows = ReadQuotedCsvLine(reader);
        DataRow row = table.NewRow();
        for (int i = 0; i < headers.Length; i++)
        {
            row[i] = rows[i];
        }

        table.Rows.Add(row);
    }

    return table;
}

static string[] ReadQuotedCsvLine(StreamReader reader)
{
    List<string> items = new List<string>();
    StringBuilder sb = new StringBuilder();
    bool inQuotes = false;
    while (true)
    {
        int nextChar = reader.Read();
        if (nextChar == -1)
        {
            break;
        }

        char c = (char)nextChar;
        if (c == '\"')
        {
            inQuotes = !inQuotes;
            if (!inQuotes && reader.Peek() != ',' && reader.Peek() != '\n')
            {
                sb.Append(c);
            }
        }
        else if (c == ',')
        {
            if (inQuotes)
            {
                sb.Append(c);
            }
            else
            {
                items.Add(sb.ToString().Trim());
                sb.Clear();
            }
        }
        else if (c == '\r' || c == '\n')
        {
            if (inQuotes)
            {
                sb.Append(c);
            }
            else
            {
                items.Add(sb.ToString().Trim());
                sb.Clear();
                if (c == '\r' && reader.Peek() == '\n')
                {
                    reader.Read();
                }

                break;
            }
        }
        else
        {
            sb.Append(c);
        }
    }

    return items.ToArray();
}