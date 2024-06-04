# BOOL_v.1.3.2

   private static void PrintSummary(int rowIndex, Excel.Worksheet sheet)
        {
            //Excel.Range rg = sheet.Range[sheet.Cells[rowIndex + 3, 3], sheet.Cells[rowIndex + 23, 5]];
            Excel.Range rg = sheet.Range[sheet.Cells[rowIndex + 3, 3], sheet.Cells[rowIndex + 23, 5]];
            var rg = sheet.Range[sheet.Cells[rowIndex + 3, 3], sheet.Cells[rowIndex + 23, 5]];
            System.Data.DataTable data = new System.Data.DataTable();
            var data = new DataTable();
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Empty", typeof(string));
            data.Columns.Add("Unit", typeof(string));
