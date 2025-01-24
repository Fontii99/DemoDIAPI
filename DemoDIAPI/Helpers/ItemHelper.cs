using DemoDIAPI.Classes;
using DocumentFormat.OpenXml.Wordprocessing;
using SAPbobsCOM;

namespace DemoDIAPI.Helpers
{
    public class ItemHelper
    {
        private string table = "OITM";
        private string field = "ItemCode";
        public void ProcessItems(Company company, Item item, int CRUD)
        {
            switch (CRUD)
            {
                case 0:
                    {
                        if (DatabaseHelper.IsInDatabase(company, item.ItemCode, table, field))
                        {
                            Console.WriteLine($"The item {item.ItemCode} already exists");
                        }
                        else
                        {
                            AddItemToDatabase(company, item);
                        }
                        break;
                    }
                case 1:
                    {
                        if (DatabaseHelper.IsInDatabase(company, item.ItemCode, table, field))
                        {
                            DeleteItemToDatabase(company, item);
                            Console.WriteLine($"The item {item.ItemCode} deleted successfully!");
                        }
                        else
                        {
                            Console.WriteLine($"The item {item.ItemCode} delete failed!");
                        }
                        break;
                    }
                case 2:
                    {
                        if (DatabaseHelper.IsInDatabase(company, item.ItemCode, table, field))
                        {
                            UpdateItemToDatabase(company, item);
                            Console.WriteLine($"The item {item.ItemCode} updated successfully!");
                        }
                        else
                        {
                            Console.WriteLine($"The item {item.ItemCode} update failed!");
                        }
                        break;
                    }

            }
        }

        private void AddItemToDatabase(Company company, Item item)
        {
            var newItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);
            newItem.ItemCode = item.ItemCode;
            newItem.ItemName = item.ItemName;
            newItem.ItemsGroupCode = item.ItemGroup;
            //newItem.UoMGroupEntry = -1;
            newItem.DefaultWarehouse = "01";

            if (newItem.Add() != 0)
            {
                {
                    Console.WriteLine($"Error creating {newItem}");
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
            Utilities.Release(newItem);
        }
        private void DeleteItemToDatabase(Company company, Item item)
        {
            var newItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);

            newItem.GetByKey(item.ItemCode);
            var result = newItem.Remove();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
            }
            else
            {
                Console.WriteLine($"Item {item.ItemCode} deleted!\n");
            }

            Utilities.Release(newItem);
        }
        private void UpdateItemToDatabase(Company company, Item item)
        {
            var newItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);

            newItem.GetByKey(item.ItemCode);

            newItem.ItemName = "UPDATED DESCRIPTION";
            var result = newItem.Update();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
            }
            else
            {
                Console.WriteLine($"Item {item.ItemCode} updated!\n");
            }

            Utilities.Release(newItem);
        }
    }
}
