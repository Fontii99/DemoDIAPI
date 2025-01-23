namespace DemoDIAPI.Classes
{
    public class Client
    {
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string FederalTaxId { get; set; }

        public Client(string cardCode, string cardName, string federalTaxId)
        {
            CardCode = cardCode;
            CardName = cardName;
            FederalTaxId = federalTaxId;
        }
        public Client() { }

        public override string ToString()
        {
            return $"CardCode: {CardCode}, CardName: {CardName}, FederalTaxId: {FederalTaxId}";
        }
    }
}