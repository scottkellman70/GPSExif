using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GPSExif
{
    class CreateXLS
    {
        private Microsoft.Office.Interop.Excel.Application app = null;
        private Microsoft.Office.Interop.Excel.Workbook workbook = null;
        public Microsoft.Office.Interop.Excel.Worksheet picWorksheet = null;
        public Microsoft.Office.Interop.Excel.Worksheet vidWorksheet = null;
        public Microsoft.Office.Interop.Excel.Worksheet audWorksheet = null;
        public Microsoft.Office.Interop.Excel.Worksheet docWorksheet = null;

        public CreateXLS()
        {
            createDoc();
        }
        public void createDoc()
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Add(1);
                picWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                vidWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[2];
                audWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[3];
                docWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[4];

            }
            catch (Exception)
            {
                Console.Write("Error");
            }
        }

        public void CreatePicHeaders(int row, int col, string htext, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = htext;
        }
        public void AddPicData(int row, int col, string data, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = data;
        }

        /******************************************************************************************************************/

        public void CreateVidHeaders(int row, int col, string htext, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = htext;
        }
        public void AddVidData(int row, int col, string data, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = data;
        }

        /******************************************************************************************************************/

        public void CreateAudHeaders(int row, int col, string htext, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = htext;
        }
        public void AddAudData(int row, int col, string data, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = data;
        }

        /******************************************************************************************************************/
        public void CreateDocHeaders(int row, int col, string htext, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = htext;
        }
        public void AddDocData(int row, int col, string data, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = data;
        }

    }
}
