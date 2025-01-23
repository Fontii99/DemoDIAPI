using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace DemoDIAPI.Classes
{
    public class Order
    {
        public string DocNum {  get; set; }

        public DateTime DocDate { get; set; }
        public string CardCode { get; set; }
        public string Description { get; set; }
        public List<ExcelData> orderLine { get; set; }

        public Order()
        {
        }

        public Order(string docNum, DateTime docDate, string cardCode, List<ExcelData> orderLine, string description)
        {
            DocNum = string.Empty;
            DocDate = DateTime.Now;
            CardCode = cardCode;
            Description = description;
            this.orderLine = orderLine;
        }
        public override string ToString()
        {
            return $"DocNum: {DocNum}, CardCode: {CardCode}, Description: {Description}, OrderLine: {orderLine.ToString()}";
        }
    }
}
