namespace DemoDIAPI.Classes
{
    using ClosedXML.Excel;
    public class ExcelDataInvoices
    {
        public string CardCode { get; set; }
        public string ItemCode { get; set; }
        public int Quantity { get; set; }

        public override string ToString()
        {
            return $"Customer: {CardCode}, Item: {ItemCode}, Quantity: {Quantity}";
        }
    }

    public class ExcelReaderInvoices
    {
        public List<ExcelDataInvoices> ReadExcelFile(string filePath)
        {
            var invoices = new List<ExcelDataInvoices>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();

                foreach (var row in rows.Skip(1)) // Skip header row
                {
                    try
                    {
                        var invoice = new ExcelDataInvoices
                        {
                            CardCode = row.Cell(1).GetString(),
                            ItemCode = row.Cell(2).GetString(),
                            Quantity = int.Parse(row.Cell(3).GetString())
                        };
                        invoices.Add(invoice);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error reading row: {row.RowNumber()}. Error: {ex.Message}");
                    }
                }
            }
            return invoices;
        }
    }
}