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

string testString;
Console.Write("Enter I for Invoices or O for Orders. E for exit:\n");
testString = Console.ReadLine().ToUpper();


switch (testString)
{
    case "O":
        {
            if (connected)
            {
                var reader = new ExcelReader();
                var data = reader.ReadExcelFile(@"C:\\Users\\pfont2\\source\\repos\\DemoDIAPI\\DemoDIAPI\\Data\\SAP_Business_One_Data.xlsx");

                List<Order> orders = new List<Order>();
                ExcelData prevLinesOrder = null;
                Order currentOrder = null;
                foreach (var line in data)
                {
                    try
                    {
                        // Start a single transaction for the entire process
                        if (!Company.InTransaction)
                        {
                            Company.StartTransaction();
                        }

                        // 1. Process Item - don't rollback if it already exists
                        var item = new DemoDIAPI.Classes.Item
                        {
                            ItemCode = line.ItemCode,
                            ItemName = line.ItemName,
                            ItemGroup = line.ItemGroup
                        };

                        var itemHelper = new ItemHelper();
                        bool itemProcessed = itemHelper.ProcessItems(Company, item, CRUD);
                        if (!itemProcessed)
                        {
                            throw new Exception("Item creation failed");
                        }

                        // 2. Process Client (only for new DocNum) - don't rollback if it already exists
                        if (prevLinesOrder == null || prevLinesOrder.DocNum != line.DocNum)
                        {
                            var client = new Client
                            {
                                CardName = line.CardName,
                                CardCode = line.CardCode,
                                FederalTaxId = line.FederalTaxId
                            };

                            var clientHelper = new ClientHelper();
                            bool clientProcessed = clientHelper.ProcessClient(Company, client, CRUD);
                            if (!clientProcessed)
                            {
                                throw new Exception("Client creation failed");
                            }

                            // 3. Create new order
                            currentOrder = new Order
                            {
                                DocNum = line.DocNum,
                                DocDate = line.DocDate,
                                CardCode = client.CardCode,
                                Description = line.Comments,
                                orderLine = new List<ExcelData>()
                            };
                            currentOrder.orderLine.Add(line);
                            orders.Add(currentOrder);

                            // Process previous order if exists
                            if (prevLinesOrder != null)
                            {
                                var OrderHelper = new OrderHelper();
                                if (!OrderHelper.ProcessOrder(Company, currentOrder))
                                {
                                    throw new Exception("Order creation failed");
                                }

                                // If order is created successfully, commit the transaction
                                Company.EndTransaction(BoWfTransOpt.wf_Commit);
                                Console.WriteLine("Order created successfully!");

                                // Start new transaction for next order
                                Company.StartTransaction();
                            }
                        }
                        else
                        {
                            // Add line to existing order
                            currentOrder.orderLine.Add(line);
                        }

                        // Update previous line tracking
                        prevLinesOrder = line;
                    }
                    catch (Exception ex)
                    {
                        // Any failure will trigger rollback
                        if (Company.InTransaction)
                        {
                            Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                        }
                        Console.WriteLine($"Transaction failed: {ex.Message}");
                    }
                }

                // Process the last order if exists
                if (currentOrder != null && currentOrder.orderLine.Any())
                {
                    try
                    {
                        var OrderHelper = new OrderHelper();
                        if (!OrderHelper.ProcessOrder(Company, currentOrder))
                        {
                            throw new Exception("Final order creation failed");
                        }

                        // Commit the final transaction
                        if (Company.InTransaction)
                        {
                            Company.EndTransaction(BoWfTransOpt.wf_Commit);
                            Console.WriteLine("Final order created successfully!");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (Company.InTransaction)
                        {
                            Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                        }
                        Console.WriteLine($"Final transaction failed: {ex.Message}");
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
            break;
        }
    case "I":
        {
            var reader = new ExcelReaderInvoices();
            var invoices = reader.ReadExcelFile(@"C:\\Users\\pfont2\\source\\repos\\DemoDIAPI\\DemoDIAPI\\Data\\Facturas.xlsx");

            foreach (var invoice in invoices)
            {
                try
                {
                    Company.StartTransaction();
                    var invoiceHelper = new InvoiceHelper();
                    invoiceHelper.ProcessInvoice(Company, invoice);
                }
                catch
                {
                    Console.WriteLine("Transaction failed.");
                }
        }
            break;
        }
    case "E":
        {
            Console.WriteLine("Ending process");
            break;
        }
    default:
        {
            Console.WriteLine("Please enter a correct option");
            break;
        }
}
