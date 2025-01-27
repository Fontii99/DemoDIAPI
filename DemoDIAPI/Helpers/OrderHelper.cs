using DemoDIAPI.Classes;
using SAPbobsCOM;

namespace DemoDIAPI.Helpers
{
    public class OrderHelper
    {
        public bool ProcessOrder(Company company, Order order)
        {

            return (AddOrderToDatabase(company, order));

        }
        private bool AddOrderToDatabase(Company company, Order order)
        {
            var newOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);
            newOrder.CardCode = order.CardCode;
            newOrder.DocDate = order.DocDate;
            newOrder.DocDueDate = order.DocDate;       
            newOrder.Comments = order.Description;
            foreach (var line in order.orderLine)
            {
                newOrder.Lines.ItemCode = line.ItemCode;
                newOrder.Lines.Quantity = (double)line.Quantity;
                newOrder.Lines.Price = (double)line.Price;
                newOrder.Lines.DiscountPercent = (double)line.Discount;
                newOrder.Lines.UoMEntry = 1;
                newOrder.Lines.Add();
            }

            if(newOrder.Add()==0)
            {
                string docEntry = company.GetNewObjectKey();
                Utilities.Release(newOrder);
                return true;
            }
            else
            {
                Console.WriteLine($"Order not created properly: {company.GetLastErrorDescription()}");
                Utilities.Release(newOrder);
                return false;
            }


        }
    }
}
