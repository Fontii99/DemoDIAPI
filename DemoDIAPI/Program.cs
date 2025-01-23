using DemoDIAPI;
using DemoDIAPI.Classes;
using DemoDIAPI.Helpers;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using SAPbobsCOM;

Console.WriteLine("Example DIAPI");

//Create company object from settings
var Company = new SAPbobsCOM.Company
{
    Server = "ESONEPC0JFN2T",
    UserName = "manager",
    Password = "seidor",
    DbServerType = BoDataServerTypes.dst_MSSQL2019,
    CompanyDB = "SBODemoES2",
    DbUserName = "sa",
    DbPassword = "SAPB1Admin",
};
bool connected = DatabaseHelper.Connect(Company);

var CRUD = 0; //0 Add, 1 Delete, 2 Update

if (connected)
{
    var reader = new ExcelReader();
    var data = reader.ReadExcelFile(@"C:\\Users\\pfont2\\source\\repos\\DemoDIAPI\\DemoDIAPI\\Data\\SAP_Business_One_Data.xlsx");

    List<Order> orders = new List<Order>();
    ExcelData prevLinesOrder = null;
    Order currentOrder = null;
    DemoDIAPI.Classes.Item item;
    foreach (var line in data)
    {
        //Create a new item to SAP
        item = new DemoDIAPI.Classes.Item
        {
            ItemCode = line.ItemCode,
            ItemName = line.ItemName,
            ItemGroup = line.ItemGroup
        };

        try
        {
            Company.StartTransaction();
            var itemHelper = new ItemHelper();
            itemHelper.ProcessItems(Company, item, CRUD);
        }
        catch
        {
            Console.WriteLine("Transaction failed.");
        }
;
        // Check if this is a new order (new DocNum)
        if (prevLinesOrder == null || prevLinesOrder.DocNum != line.DocNum)
        {
            // Check the previous order before creating a new one (only if it's not the first iteration)
            if (prevLinesOrder != null)
            {
                try
                { 
                    Company.StartTransaction();
                    var OrderHelper = new OrderHelper();
                    OrderHelper.ProcessOrder(Company, currentOrder);
                }
                catch
                {
                    Console.WriteLine("Transaction failed.");
                }
;
            }

            // Create a new client to SAP
            var client = new Client
            {
                CardName = line.CardName,
                CardCode = line.CardCode,
                FederalTaxId = line.FederalTaxId
            };

            try
            {
                Company.StartTransaction();
                var clientHelper = new ClientHelper();
                clientHelper.ProcessClient(Company, client, CRUD);
            }
            catch
            {
                Console.WriteLine("Transaction failed.");
            }
;
            // Create a new order
            currentOrder = new Order
            {
                DocNum = line.DocNum,
                DocDate = line.DocDate,
                CardCode = client.CardCode,
                Description = line.Comments,
                orderLine = new List<ExcelData>()
            };

            //Add the first line to the order
            currentOrder.orderLine.Add(line);
            // Add the order to the orders list
            orders.Add(currentOrder);
        }
        else
        {
            // If it's the same order, just add the item to the existing order
            currentOrder.orderLine.Add(line);
        }

        // Update previous line tracking
        prevLinesOrder = line;
    }

    // Don't forget to process the last order after the loop
    if (currentOrder != null)
    {
        try
        {
            Company.StartTransaction();
            var OrderHelper = new OrderHelper();
            OrderHelper.ProcessOrder(Company, currentOrder);
        }
        catch
        {
            Console.WriteLine("Transaction failed.");
        }
    }



    if (Company.Connected)
        Company.Disconnect();

    Utilities.Release(Company);
}
else
{
    Console.WriteLine("I'm not connected");
}