using SAPbobsCOM;

namespace DemoDIAPI.Classes
{
    public class ItemHelper
    {
        public void ProcessItems(Company company, List<Item> items,int CRUD)
        {
            foreach (var item in items)
            {
                if (CRUD == 0) //ADD ITEM
                {
                    if (IsItemInDatabase(company, item.ItemCode))
                    {
                        Console.WriteLine($"The item {item.ItemCode} already exists");
                    }
                    else
                    {
                        AddItemToDatabase(company, item);
                    }
                }
                if (CRUD == 1) //DELETE ITEM
                {
                    DeleteItemToDatabase(company, item);
                    if (!IsItemInDatabase(company, item.ItemCode))
                    {
                        Console.WriteLine($"The item {item.ItemCode} deleted successfully!");
                    }
                    else
                    {
                        Console.WriteLine($"The item {item.ItemCode} delete failed!");
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
            Console.WriteLine($"____________________\n{newItem.ToString()}\n______________________________");
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
    }
}
