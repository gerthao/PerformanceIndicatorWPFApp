using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PerformanceIndicatorWPFApp
{
    class Handler
    {
        public Excel.Application App { get; set; }
        public Excel.Workbook Book { get; set; }
        public Excel.Worksheet Sheet { get; set; }
        public Excel.Range Range { get; set; }
        public int LastRow { get; set; }
        public string FileName { get; set; }

        public Handler(string path, int index = 1)
        {
            App = new Excel.Application();
            Book = App.Workbooks.Open(path);
            Sheet = Book.Sheets[index];
            LastRow = Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
        }

        public BindingList<Report> ReportList()
        {
            BindingList<Report> list = new BindingList<Report>();
            for (int index = 2; index <= LastRow; index++)
            {
                System.Array MyValues = (System.Array)Sheet.get_Range("A" +
                   index.ToString(), "D" + index.ToString()).Cells.Value;
                list.Add(new Report
                {
                    Name = MyValues.GetValue(1, 1).ToString()
                });
            }
            return list;
        }

    }
}
