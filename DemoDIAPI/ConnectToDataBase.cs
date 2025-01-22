namespace DemoDIAPI
{
    public static class ConnectToDataBase
    {
        public static bool Connect(SAPbobsCOM.Company Company)
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
    }
}