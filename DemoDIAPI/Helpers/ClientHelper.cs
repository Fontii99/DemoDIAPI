using DemoDIAPI.Classes;
using DocumentFormat.OpenXml.Spreadsheet;
using SAPbobsCOM;

namespace DemoDIAPI.Helpers
{
    public class ClientHelper
    {
        private string table = "OCRD";
        private string field = "CardCode";
        public void ProcessClient(Company company, Client client, int CRUD)
        {
            switch (CRUD)
            {
                case 0:
                    {
                        if (DatabaseHelper.IsInDatabase(company, client.CardCode, table, field))
                        {
                            Console.WriteLine($"The client {client.CardCode} already exists");
                        }
                        else
                        {
                            AddClientToDatabase(company, client);
                        }
                        break;
                    }
                case 1:
                    {
                        if (DatabaseHelper.IsInDatabase(company, client.CardCode, table, field))
                        {
                            DeleteClientToDatabase(company, client);
                            Console.WriteLine($"The client {client.CardCode} deleted successfully!");
                        }
                        else
                        {
                            Console.WriteLine($"The client {client.CardCode} delete failed!");
                        }
                        break;
                    }
                case 2:
                    {
                        if (DatabaseHelper.IsInDatabase(company, client.CardCode, table, field))
                        {
                            UpdateClientToDatabase(company, client);
                            Console.WriteLine($"The item {client.CardCode} updated successfully!");
                        }
                        else
                        {
                            Console.WriteLine($"The item {client.CardCode} update failed!");
                        }
                        break;
                    }

            }
        }
        private void AddClientToDatabase(Company company, Client client)
        {
            var newClient = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            newClient.CardCode = client.CardCode;
            newClient.CardName = client.CardName;
            newClient.FederalTaxID = client.FederalTaxId;

            if (newClient.Add() != 0)
            {
                Console.WriteLine($"Error creating {newClient.CardCode}");
                company.EndTransaction(BoWfTransOpt.wf_RollBack); ;
            }
            Utilities.Release(newClient);
        }
        private void DeleteClientToDatabase(Company company, Client client)
        {
            var newClient = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            newClient.GetByKey(client.CardCode); ;
            var result = newClient.Remove();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
            }
            else
            {
                Console.WriteLine($"Client {newClient.CardCode} deleted!\n");
            }

            Utilities.Release(newClient);
        }
        private void UpdateClientToDatabase(Company company, Client client)
        {
            var newClient = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            newClient.GetByKey(client.CardCode);

            newClient.CardName = "UPDATED NAME";
            var result = newClient.Update();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
            }
            else
            {
                Console.WriteLine($"Client {newClient.CardCode} updated!\n");
            }

            Utilities.Release(newClient);
        }
    }
}
