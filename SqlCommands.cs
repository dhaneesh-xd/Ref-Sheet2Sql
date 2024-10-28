using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    internal class SqlCommands
    {
        string db;
        public SqlCommands()
        {
            
        }
        static void CreateTableAndInsertData(string databaseName, DataTable sheetData)
        {
            string connectionString = $"Server=your_server;Database={databaseName};Integrated Security=true;";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = $"CREATE TABLE [{sheetData.TableName}] (";
                foreach (DataColumn column in sheetData.Columns)
                {
                    createTableQuery += $"[{column.ColumnName}] NVARCHAR(MAX),";
                }
                createTableQuery = createTableQuery.TrimEnd(',') + ")";
                OleDbCommand createTableCommand = new OleDbCommand(createTableQuery, connection);
                createTableCommand.ExecuteNonQuery();
                foreach (DataRow row in sheetData.Rows)
                {
                    string insertQuery = $"INSERT INTO [{sheetData.TableName}] VALUES (";
                    foreach (var item in row.ItemArray)
                    {  
                        insertQuery += $"'{item}',";
                    }
                    insertQuery = insertQuery.TrimEnd(',') + ")";
                    OleDbCommand insertCommand = new OleDbCommand(insertQuery, connection);
                    insertCommand.ExecuteNonQuery();
                }
            }
        }
    }
}
