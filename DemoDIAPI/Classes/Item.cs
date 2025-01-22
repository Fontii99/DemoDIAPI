namespace DemoDIAPI.Classes
{
    public class Item
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public int ItemGroup { get; set; }
        public Item(string itemCode, string itemName, int itemGroup)
        {
            ItemCode = itemCode;
            ItemName = itemName;
            ItemGroup = itemGroup;
        }

        public override string ToString()
        {
            return $"ItemCode: {ItemCode}, ItemName: {ItemName}, ItemGroup: {ItemGroup}";
        }
    }
}