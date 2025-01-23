using SAPbobsCOM;

namespace DemoDIAPI.Helpers
{
    public static class DatabaseHelper
    {
        public static bool Connect(Company Company)
        {
            Console.WriteLine($"Trying to connect to database: {Company.CompanyDB}");

            var result = Company.Connect();
            if (result != 0)
            {
                Console.WriteLine(Company.GetLastErrorDescription());
                return false;
            }
            else
            {
                Console.WriteLine("Connection correct!");
                return true;
            }

        }
        public static bool IsInDatabase(Company company, string value, string table, string field)
        {
            var query = $"""
                SELECT {field} 
                FROM {table}
                WHERE
                    {field} = '{value}'
                """;

            Console.WriteLine($"Executing query: {query}");
            var recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(query);

            var exists = !recordset.EoF;
            if (exists)
            {//EXISTS ON THE DATABASE
                Console.WriteLine($"Item {value} exists in the database.\n");
            }
            else
            {//DONT EXISTS ON THE DATABASE
                Console.WriteLine($"Item {value} does not exist in the database.\n");
            }

            Utilities.Release(recordset);
            return exists;
        }
    }
}