using DemoDIAPI.Classes;
using DocumentFormat.OpenXml.Spreadsheet;
using SAPbobsCOM;

namespace DemoDIAPI.Helpers
{
    public class ClientHelper
    {
        private string table = "OCRD";
        private string field = "CardCode";
        public bool ProcessClient(Company company, Client client, int CRUD)
        {
            switch (CRUD)
            {
                case 0:
                    {
                        if (DatabaseHelper.IsInDatabase(company, client.CardCode, table, field))
                        {
                            Console.WriteLine($"The client {client.CardCode} already exists");
                            return true;
                        }
                        else
                        {
                            if (AddClientToDatabase(company, client))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                case 1:
                    {
                        if (DatabaseHelper.IsInDatabase(company, client.CardCode, table, field))
                        {
                            DeleteClientToDatabase(company, client);
                            Console.WriteLine($"The client {client.CardCode} deleted successfully!");
                            return true;
                        }
                        else
                        {
                            Console.WriteLine($"The client {client.CardCode} delete failed!");
                            return false;
                        }
                    }
                case 2:
                    {
                        if (DatabaseHelper.IsInDatabase(company, client.CardCode, table, field))
                        {
                            UpdateClientToDatabase(company, client);
                            Console.WriteLine($"The item {client.CardCode} updated successfully!");
                            return true;
                        }
                        else
                        {
                            Console.WriteLine($"The item {client.CardCode} update failed!");
                            return false;
                        }
                    }

            }
            return false;
        }
        private bool AddClientToDatabase(Company company, Client client)
        {
            var newClient = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            newClient.CardCode = client.CardCode;
            newClient.CardName = client.CardName;
            newClient.FederalTaxID = client.FederalTaxId;

            if (newClient.Add() != 0)
            {
                Console.WriteLine($"Error creating {newClient.CardCode}");
                Utilities.Release(newClient);
                return false;
            }
            else
            {
                Utilities.Release(newClient);
                return true;
            }

        }
        private bool DeleteClientToDatabase(Company company, Client client)
        {
            var newClient = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            newClient.GetByKey(client.CardCode); ;
            var result = newClient.Remove();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
                Utilities.Release(newClient);
                return false;
            }
            else
            {
                Console.WriteLine($"Client {newClient.CardCode} deleted!\n");
                Utilities.Release(newClient);
                return true;
            }
        }
        private bool UpdateClientToDatabase(Company company, Client client)
        {
            var newClient = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            newClient.GetByKey(client.CardCode);

            newClient.CardName = "UPDATED NAME";
            var result = newClient.Update();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
                Utilities.Release(newClient);
                return false;
            }
            else
            {
                Console.WriteLine($"Client {newClient.CardCode} updated!\n");
                Utilities.Release(newClient);
                return true;
            }
        }
    }
}
