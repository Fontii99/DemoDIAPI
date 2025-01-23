using SAPbobsCOM;

namespace DemoDIAPI.Classes
{
    public class ItemHelper
    {
        string table = "OITM";
        string field = "";
        public void ProcessItems(Company company, List<Item> items,int CRUD)
        {
            foreach (var item in items)
            {
                switch (CRUD)
                {
                case 0:
                    {
                        if (IsItemInDatabase(company, item.ItemCode))
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
                        if (IsItemInDatabase(company, item.ItemCode))
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
                        if (IsItemInDatabase(company, item.ItemCode))
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
        }
        
        private bool IsItemInDatabase(Company company, string itemCode)
        {
            var query = $"""
                SELECT T0.ItemCode 
                FROM OITM T0
                WHERE
                    T0.ItemCode = '{itemCode}'
                """;

            Console.WriteLine($"Executing query: {query}");
            var recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(query);

            var exists = !recordset.EoF;
            if (exists)
            {//EXISTS ON THE DATABASE
                Console.WriteLine($"Item {itemCode} exists in the database.\n");
            }else
            {//DONT EXISTS ON THE DATABASE
                Console.WriteLine($"Item {itemCode} does not exist in the database.\n");
            }

            Utilities.Release(recordset);
            return exists;
        }

        private void AddItemToDatabase(Company company, Item item)
        {
            var newItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);
            newItem.ItemCode = item.ItemCode;
            newItem.ItemName = item.ItemName;
            newItem.ItemsGroupCode = item.ItemGroup;
            newItem.ItemType = ItemTypeEnum.itItems;

            var result = newItem.Add();
            if (result != 0)
            {
                Console.WriteLine($"ERROR: {company.GetLastErrorDescription()}\n");
            }
            else
            {
                Console.WriteLine($"Item {item.ItemCode} creation correct!\n");
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
