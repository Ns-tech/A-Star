using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelO
{
   public class DataExcel
    {
        public List<Point> points { get; set; }
        public List<DirectLink> directLinks { get; set; }
        public DataExcel()
        {
            points = new List<Point>();
            directLinks = new List<DirectLink>();
        }
    }
}
