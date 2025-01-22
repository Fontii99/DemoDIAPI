using DemoDIAPI;
using DemoDIAPI.Classes;
using DocumentFormat.OpenXml.ExtendedProperties;
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
bool connected = ConnectToDataBase.Connect(Company);

if (connected){
    var reader = new ExcelReader();
    var data = reader.ReadExcelFile(@"C:\\Users\\pfont2\\source\\repos\\DemoDIAPI\\DemoDIAPI\\Data\\SAP_Business_One_Data.xlsx");

    var clients = new List<Client>();
    var items = new List<Item>();

    foreach (var line in data)
    {
        clients.Add(new Client(line.CardCode, line.CardName, line.FederalTaxId));
        items.Add(new Item(line.ItemCode, line.ItemName, line.ItemGroup));
    }
    int CRUD = 1; //0 Add, 1 Delete
    var itemHelper = new ItemHelper();
    itemHelper.ProcessItems(Company, items, CRUD);



    if (Company.Connected)
        Company.Disconnect();

    Utilities.Release(Company);
}
else
{
    Console.WriteLine("I'm not connected");
}




