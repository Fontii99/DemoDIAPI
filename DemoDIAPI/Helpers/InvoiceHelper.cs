using DemoDIAPI.Classes;
using DocumentFormat.OpenXml.Wordprocessing;
using SAPbobsCOM;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DemoDIAPI.Helpers
{
    public class InvoiceHelper
    {
        public void ProcessInvoice(Company company, ExcelDataInvoices invoiceData)
        {

            AddInvoiceToDatabase(company, invoiceData);

        }
        private void AddInvoiceToDatabase(Company company, ExcelDataInvoices invoiceData)
        {
            var newInvoice = (Documents)company.GetBusinessObject(BoObjectTypes.oInvoices);
            var order = DatabaseHelper.IsInDatabase(company, invoiceData.CardCode, invoiceData.ItemCode, invoiceData.Quantity);
            if (order[0] != -1)
            {
                newInvoice.DocDate = DateTime.Now;
                newInvoice.DocDueDate = DateTime.Now;
                newInvoice.CardCode = invoiceData.CardCode;
                newInvoice.Lines.ItemCode = invoiceData.ItemCode;
                newInvoice.Lines.Quantity = invoiceData.Quantity;
                newInvoice.Lines.BaseEntry = order[0];
                newInvoice.Lines.BaseType = 17;
                newInvoice.Lines.BaseLine = order[1];
                if (newInvoice.Add() == 0)
                {
                    string docEntry = company.GetNewObjectKey();
                    Console.WriteLine($"DocNum:{docEntry}");
                    company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
                else
                {
                    Console.WriteLine($"Invoice not created properly: {company.GetLastErrorDescription()}");
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
            else
            {
                newInvoice.DocDate = DateTime.Now;
                newInvoice.DocDueDate = DateTime.Now;
                newInvoice.CardCode = invoiceData.CardCode;
                newInvoice.Lines.ItemCode = invoiceData.ItemCode;
                newInvoice.Lines.Quantity = invoiceData.Quantity;
                if (newInvoice.Add() == 0)
                {
                    string docEntry = company.GetNewObjectKey();
                    Console.WriteLine($"DocNum:{docEntry}");
                    company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
                else
                {
                    Console.WriteLine($"Invoice not created properly: {company.GetLastErrorDescription()}");
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
        }
    }
}