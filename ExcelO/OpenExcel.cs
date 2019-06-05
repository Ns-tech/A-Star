using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelO
{
    public class OpenExcel
    {
        public DataExcel ReadData()
        {
            DataExcel dt = new DataExcel();
            
            string path = Path.Combine(Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName, "Database.xlsx");

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet1.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<Point> ps = new List<Point>(); 
            if (colCount == 3)
                for (int i = 2; i <= rowCount; i++)
                { 
                    ps.Add(new Point {
                        Id = xlRange.Cells[i, 1].Value2.ToString(),
                        Lat = (double)xlRange.Cells[i, 2].Value2,
                        Lng = (double)xlRange.Cells[i, 3].Value2
                    });
                    Console.Write("\rRead data point excel:{0}", i - 1); 
                }
            Console.Write("\n");
            dt.points = ps;
            xlRange = xlWorksheet2.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count; 
            List<DirectLink> dls = new List<DirectLink>(); 
            if (colCount == 3)
                for (int i = 2; i <= rowCount; i++)
                {
                    dls.Add(new DirectLink
                    {
                        CurrentPoint = xlRange.Cells[i, 1].Value2.ToString(),
                        NextPoint =  xlRange.Cells[i, 2].Value2.ToString(),
                        Distance =  (int)xlRange.Cells[i, 3].Value2
                    });
                    Console.Write("\rRead data link excel:{0}", i - 1);
                     
                }
            dt.directLinks = dls;
            Console.Write("\n");
            xlWorkbook.Close();
            return dt;
        }

    }
}
