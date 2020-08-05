using System;

namespace GPSExif
{
    static class CreateExcelDoc
    {
        private static ExcelApp.Application app = null;
        private static ExcelApp.Workbook workbook = null;
        private static ExcelApp.Worksheet worksheet = null;

        public static ExcelApp.Application CreateDoc()
        {
            try
            {
                app = new ExcelApp.Application();
                app.Visible = true;
                workbook = app.Workbooks.Add(1);
                worksheet = (ExcelApp.Worksheet)workbook.Sheets[1];
            }
            catch (Exception)
            {
                Console.Write("Error");
            }
            return app;
        }

        public static void CreateHeaders(int row, int col, string htext)
        {
            worksheet.Cells[row, col] = htext;
        }
        public static void addData(int row, int col, string data)
        {
            worksheet.Cells[row, col] = data;
        }

    }
}
