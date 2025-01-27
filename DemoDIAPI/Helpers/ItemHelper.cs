using DemoDIAPI.Classes;
using DocumentFormat.OpenXml.Wordprocessing;
using SAPbobsCOM;

namespace DemoDIAPI.Helpers
{
    public class ItemHelper
    {
        private string table = "OITM";
        private string field = "ItemCode";
        public bool ProcessItems(Company company, Item item, int CRUD)
        {
            switch (CRUD)
            {
                case 0:
                    {
                        if (DatabaseHelper.IsInDatabase(company, item.ItemCode, table, field))
                        {
                            Console.WriteLine($"The item {item.ItemCode} already exists");
                            return true;
                        }
                        else
                        {
                            if (AddItemToDatabase(company, item))
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
                        if (DatabaseHelper.IsInDatabase(company, item.ItemCode, table, field))
                        {
                            DeleteItemToDatabase(company, item);
                            Console.WriteLine($"The item {item.ItemCode} deleted successfully!");
                            return true;
                        }
                        else
                        {
                            Console.WriteLine($"The item {item.ItemCode} delete failed!");
                            return false;
                        }
                    }
                case 2:
                    {
                        if (DatabaseHelper.IsInDatabase(company, item.ItemCode, table, field))
                        {
                            UpdateItemToDatabase(company, item);
                            Console.WriteLine($"The item {item.ItemCode} updated successfully!");
                            return true;
                        }
                        else
                        {
                            Console.WriteLine($"The item {item.ItemCode} update failed!");
                            return false;
                        }
                    }
            }
            return false;
        }

        private bool AddItemToDatabase(Company company, Item item)
        {
            var newItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);
            newItem.ItemCode = item.ItemCode;
            newItem.ItemName = item.ItemName;
            newItem.ItemsGroupCode = item.ItemGroup;
            //newItem.UoMGroupEntry = -1;
            newItem.DefaultWarehouse = "01";

            if (newItem.Add() != 0)
            {
                Console.WriteLine($"Error creating {newItem}");
                Utilities.Release(newItem);
                return false;
            }
            else
            {
                Utilities.Release(newItem);
                return true;
            }

        }
        private bool DeleteItemToDatabase(Company company, Item item)
        {
            var newItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);

            newItem.GetByKey(item.ItemCode);
            var result = newItem.Remove();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
                Utilities.Release(newItem);
                return false;
            }
            else
            {
                Console.WriteLine($"Item {item.ItemCode} deleted!\n");
                Utilities.Release(newItem);
                return true;
            }
        }
        private bool UpdateItemToDatabase(Company company, Item item)
        {
            var newItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);

            newItem.GetByKey(item.ItemCode);

            newItem.UoMGroupEntry = -1;
            var result = newItem.Update();

            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
                Utilities.Release(newItem);
                return false;

            }
            else
            {
                Console.WriteLine($"Item {item.ItemCode} updated!\n");
                Utilities.Release(newItem);
                return true;
            }

        }
    }
}
