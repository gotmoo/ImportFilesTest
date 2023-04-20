using System.Data;
using Microsoft.Data.SqlClient;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ConsoleApp2;

public static class DbFunctions
{
    public static void BulkUploadDatatable(string table, DataTable dataTable, string connString,
        Dictionary<string, string> colMap)
    {
        SqlConnection myConn = new SqlConnection(connString);
        SqlCommand myCommand = new SqlCommand($"Truncate Table dbo.[{table}]", myConn);
        try
        {
            myConn.Open();
            var reader = myCommand.ExecuteNonQuery();
        }
        catch (System.Exception ex)
        {
            Console.WriteLine("Something went wrong");
            Console.WriteLine(ex.ToString());
        }
        finally
        {
            if (myConn.State == ConnectionState.Open)
            {
                myConn.Close();
            }
        }
        using var copy = new SqlBulkCopy(connString);
        copy.DestinationTableName = $"dbo.[{table}]";
        copy.BulkCopyTimeout = 60;
        // Add mappings so that the column order doesn't matter
        foreach (var columnMappingKey in colMap.Keys.Where(columnMappingKey =>
                     !string.IsNullOrWhiteSpace(colMap[columnMappingKey])))
        {
            copy.ColumnMappings.Add(columnMappingKey, colMap[columnMappingKey]);
        }

        copy.WriteToServer(dataTable);
    }

    public static Dictionary<string, string> MapColumnsToDb(string table, string connString,
        Dictionary<string, string> colMap)
    {
        var tableCheckQuery = $$"""
                        SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
                        FROM INFORMATION_SCHEMA.COLUMNS
                        WHERE TABLE_NAME = N'{{table}}'
                        order by ORDINAL_POSITION
                        """;
        SqlConnection myConn = new SqlConnection(connString);
        SqlCommand myCommand = new SqlCommand(tableCheckQuery, myConn);
        try
        {
            myConn.Open();
            var reader = myCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    var colName = reader.GetString(2);
                    if (colMap.ContainsKey(colName))
                    {
                        colMap[colName] = colName;
                    }
                }
            }
            else
            {
                reader.Close();
                var createTableQuery = $$"""
                            CREATE TABLE [dbo].[{{table}}](
                            {{string.Join(", ", colMap.Keys.Select(x => $"[{x}] [varchar](max) NULL"))}}
                            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
                            """;
                myCommand.CommandText = createTableQuery;

                myCommand.ExecuteNonQuery();
                foreach (var columnMappingKey in colMap.Keys)
                {
                    colMap[columnMappingKey] = columnMappingKey;
                }
            }
        }
        catch (System.Exception ex)
        {
            Console.WriteLine("Something went wrong");
            Console.WriteLine(ex.ToString());
        }
        finally
        {
            if (myConn.State == ConnectionState.Open)
            {
                myConn.Close();
            }
        }

        if (colMap.Values.Any(string.IsNullOrWhiteSpace))
        {
            Console.WriteLine("Not all columns in the excel file exist in the database. These columns are ignored:");
            var missingColumns = string.Join("; ", colMap.Keys.Where(x => string.IsNullOrWhiteSpace(colMap[x])));
            Console.WriteLine(missingColumns);
        }

        return colMap;
    }
}