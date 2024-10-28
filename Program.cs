using System.Data;
using System.Data.OleDb;
using TestConsole;


TestTest testTest = new TestTest();
testTest.Test();
public class TestTest
{
    public void Test()
    {
        ClassForTest classForTest = new ClassForTest();
        string connectionString = classForTest.GetConnectionString(@"");

        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            connection.Open();
            DataTable sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            DataSet dataSet = new DataSet();
            foreach (DataRow sheet in sheets.Rows)
            {
                string sheetName = sheet["TABLE_NAME"].ToString();
                DataTable sheetData = new DataTable(sheetName);
                using (OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM [{sheetName}]", connection))
                {
                    adapter.Fill(sheetData);
                }
                dataSet.Tables.Add(sheetData);
            }
            string cleanSheetName = Path.GetFileNameWithoutExtension(connectionString).Replace("$", "").Replace("'", "").Split(';')[0].Split('.')[0];
            Console.WriteLine(cleanSheetName);
            foreach (DataTable table in dataSet.Tables)
            {
                Console.WriteLine($"Data from sheet: {table.TableName.Trim()}");
                foreach (DataColumn col in table.Columns)
                {
                    Console.Write(col.ColumnName + "\t");
                }
                Console.WriteLine();
                foreach (DataRow row in table.Rows)
                {
                    foreach (var item in row.ItemArray)
                    {
                        Console.Write(item + "\t");
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
            }
        }
    }
}
