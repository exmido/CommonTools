using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

//ConvertOpenXml
static partial class ConvertOpenXml
{
    //Header
    public class Header
    {
        public int index = 0;
        public string name = string.Empty;
        public string filter = string.Empty;
        public string format = string.Empty;
    }

    //Table
    public class Table
    {
        public OrderedDictionary<int, Header> headers = new OrderedDictionary<int, Header>();
        public List<Dictionary<string, object>> datas = new List<Dictionary<string, object>>();
    }

    //ToData
    static public Dictionary<string, Table> ToData(string excelFilePath, string filter = "", bool headerOnly = false)
    {
        var tables = new Dictionary<string, Table>();

        // Load the Excel document
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(excelFilePath, false))
        {
            // File path
            Console.WriteLine(excelFilePath);

            // Get the workbook part and the sheets
            WorkbookPart workbookPart = document.WorkbookPart;
            if (workbookPart == null)
                return null;

            Sheets sheets = workbookPart.Workbook.Sheets;
            if (sheets == null)
                return null;

            foreach (Sheet sheet in sheets)
            {
                if (string.IsNullOrEmpty(sheet.Name) || !char.IsLetter(sheet.Name.ToString()[0]))
                    continue;

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                OrderedDictionary<int, Header> headers = GetHeaders(document, sheetData, filter);
                if (headers.Count <= 0)
                    continue;

                // Sheet name
                Console.WriteLine("    " + sheet.Name);

                int indexMax = headers.Last().Key;

                List<Dictionary<string, object>> datas = null;

                // Iterate through rows and collect data
                foreach (Row row in sheetData.Elements<Row>())
                {
                    // Skip first row
                    if (datas == null)
                    {
                        datas = new List<Dictionary<string, object>>();
                        continue;
                    }

                    var rowData = new Dictionary<string, object>();

                    for (int i = 0; i <= indexMax; ++i)
                    {
                        var cell = row.Elements<Cell>().ElementAtOrDefault(i);
                        if (cell == null)
                            continue;

                        Header header;
                        if (headers.TryGetValue(GetCellIndexNumber(cell.CellReference), out header))
                        {
                            if (string.IsNullOrEmpty(header.format))
                            {
                                header.format = (string)GetCellValue(document, cell, "s");
                            }
                            else
                            {
                                object cellValue = GetCellValue(document, cell, header.format);
                                if (cellValue != null)
                                    rowData[header.name] = cellValue;
                            }
                        }
                    }

                    if (headerOnly)
                        break;

                    if (rowData.Count > 0)
                        datas.Add(rowData);
                }

                // Add to table
                Table table = new Table();
                table.headers = headers;

                if (datas != null)
                    table.datas = datas;

                tables.Add(sheet.Name, table);
            }
        }

        return tables;
    }

    //GetCellIndexNumber
    static private int GetCellIndexNumber(string cellReference)
    {
        string cellIndexString = string.Concat(cellReference.TakeWhile(c => c < '0' || c > '9'));

        int index = (cellIndexString[cellIndexString.Length - 1] - 'A');

        for (int i = cellIndexString.Length - 2; i >= 0; --i)
            index += index * ('Z' - 'A') + (cellIndexString[i] - 'A');

        return index;
    }

    //GetHeaders
    static private OrderedDictionary<int, Header> GetHeaders(SpreadsheetDocument document, SheetData sheetData, string filter)
    {
        OrderedDictionary<int, Header> headers = new OrderedDictionary<int, Header>();

        var firstRow = sheetData.Elements<Row>().FirstOrDefault();
        if (firstRow == null)
            return headers;

        foreach (Cell cell in firstRow.Elements<Cell>())
        {
            var s = ((string)GetCellValue(document, cell)).Split(":");

            if (!string.IsNullOrEmpty(filter) && s.Length > 1 && s[1].Intersect(filter).Count() <= 0)
                continue;

            Header header = new Header();
            header.index = GetCellIndexNumber(cell.CellReference);
            header.name = s[0];

            if (s.Length > 1)
                header.filter = s[1];

            try
            {
                headers.Add(header.index, header); // Excel column headers
            }
            catch (Exception ex)
            {
                Console.WriteLine(cell.CellReference.ToString() + " - " + ex.ToString());
            }
        }

        return headers;
    }

    //GetCellValue
    static private object GetCellValue(SpreadsheetDocument document, Cell cell, string fmt = "s")
    {
        if (document == null || cell == null)
            return null;

        string cellValue = string.Empty;

        // Check if cell is a shared string (i.e., it refers to a shared string table)
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            int index = int.Parse(cell.CellValue.Text);
            SharedStringTablePart sharedStringTablePart = document.WorkbookPart.SharedStringTablePart;
            cellValue = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).Text.Text;
        }
        else
        {
            cellValue = cell.CellValue?.Text;
        }

        if (cellValue != string.Empty)
        {
            List<string> values = new List<string>();

            try
            {
                bool toList = fmt.Contains('[');
                if (toList)
                {
                    int startIndex = fmt.IndexOf('[');
                    int endIndex = fmt.LastIndexOf(']');

                    values = cellValue.Split(fmt.Substring(startIndex + 1, endIndex - startIndex - 1)).ToList<string>();
                    fmt = fmt.Substring(0, startIndex).Trim();
                }

                switch (fmt)
                {
                    case "b":
                        return toList ? values.ConvertAll(bool.Parse) : bool.Parse(cellValue);
                    case "i32":
                        return toList ? values.ConvertAll(decimal.Parse).ConvertAll(d => (int)d) : (int)decimal.Parse(cellValue);
                    case "u32":
                        return toList ? values.ConvertAll(decimal.Parse).ConvertAll(d => (uint)d) : (uint)decimal.Parse(cellValue);
                    case "i64":
                        return toList ? values.ConvertAll(decimal.Parse).ConvertAll(d => (long)d) : (long)decimal.Parse(cellValue);
                    case "u64":
                        return toList ? values.ConvertAll(decimal.Parse).ConvertAll(d => (ulong)d) : (ulong)decimal.Parse(cellValue);
                    case "f":
                        return toList ? values.ConvertAll(decimal.Parse).ConvertAll(d => (float)d) : (float)decimal.Parse(cellValue);
                    case "d":
                        return toList ? values.ConvertAll(decimal.Parse).ConvertAll(d => (double)d) : (double)decimal.Parse(cellValue);
                    case "s":
                        return toList ? values : cellValue;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(cell.CellReference.ToString() + " - " + ex.ToString());
            }
        }

        return null;
    }
}
