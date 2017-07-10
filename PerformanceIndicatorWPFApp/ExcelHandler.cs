using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;

namespace PerformanceIndicatorWPFApp
{

    class ExcelHandler
    {
        public enum Months { Jan=1, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec }
        public enum Days { Sun=1, Mon, Tue, Wed, Thu, Fri, Sat}

        public Excel.Application App { get; set; }
        public Excel.Workbook Book { get; set; }
        public Excel.Worksheet Sheet { get; set; }
        public Excel.Range Range { get; set; }
        public int LastRow { get; set; }
        public string FileName { get; set; }

        public ExcelHandler(string path, int index = 1)
        {
            if(App != null || Book != null || Sheet != null)
            {
                Close();
            }
            App = new Excel.Application();
            Book = App.Workbooks.Open(path);
            Sheet = Book.Sheets[index];
            LastRow = Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
        }
        ~ExcelHandler(){
            Close();
        }
        public Excel.Range GetCell(int row, int col) => (Excel.Range)Sheet.Cells[row, col];
        public void BuildWorksheet(ref Excel.Worksheet referenceSheet, int count, DateTime startDate, int daysRange)
        {
            if (referenceSheet == null) return;
            var newSheet = new Excel.Worksheet();
            for(int i = 1; i <= 52; i++)
            {
                String[] sheetName = referenceSheet.Name.Split(' ');
                referenceSheet.Copy(referenceSheet, newSheet);
                newSheet.Name = $"{sheetName[0]} {(Int32.Parse(sheetName[1]) + 1) }";
                newSheet = Book.Sheets.Add();
            }
            return;
        }
        public string ToJSONString()
        {
            StringBuilder json = new StringBuilder();
            String tab = "\t";
            bool newTable = true;
            int tablevel = 0;
            try {
                json.AppendLine(@"{");
                tablevel++;

                for (int j = tablevel; j > 0; j--) json.Append(tab);
                json.AppendLine("\"Tables\" : [");
                tablevel++;

                //for (int j = tablevel; j > 0; j--) json.Append(tab);
                //json.AppendLine(@"{");
                //tablevel++;
                for (int i = 4; i < LastRow; i++)
                {
                    Excel.Range cell = GetCell(i, 2);
                    if (cell.Value2 == null){ continue; }
                    if (cell.Font.Bold && newTable)
                    {
                        newTable = false;
                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine(@"{");
                        tablevel++;

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine($"\"Name\" : \"{cell.Value2}\"");

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine("\"Rows\" : [");
                        tablevel++;
                    } else if(cell.Font.Bold && !newTable)
                    {
                       
                        for (int j = tablevel-1; j > 0; j--) json.Append(tab);
                        json.AppendLine("]");
                        tablevel--;

                        newTable = true;
                        for (int j = tablevel-1; j > 0; j--) json.Append(tab);
                        json.AppendLine(@"},");
                        tablevel--;
                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine(@"{");
                        tablevel++;
                        newTable = false;

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine($"\"Name\" : \"{cell.Value2}\"");

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine("\"Rows\" : [");
                        tablevel++;
                    }
                    else
                    {
                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        //json.Append($"\"Row\" : ");
                        json.AppendLine(@"{");
                        tablevel++;

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.Append("\"Name\" : ");
                        json.AppendLine($"\"{(cell as Excel.Range).Value2}\"");

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.Append("\"Value\" : ");
                        json.AppendLine($"\"{ GetCell(i, 5).Value2}\",");

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.Append("\"NumberFormat\" : ");
                        json.AppendLine($"\"{ GetCell(i, 5).NumberFormat}\",");

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.Append("\"Color\" : ");
                        json.AppendLine($"\"{ GetCell(i, 5).DisplayFormat.Interior.Color}\"");

                        tablevel--;
                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine(@"},");

                    }
                }
                if (!newTable)
                {
                    tablevel--;
                    for (int j = tablevel; j > 0; j--) json.Append(tab);
                    json.AppendLine(@"}");
                    tablevel--;
                }
                for (int j = tablevel; j > 0; j--) json.Append(tab);
                json.AppendLine(@"]");
                tablevel--;

                for (int j = tablevel; j > 0; j--) json.Append(tab);
                json.AppendLine(@"}");
            } catch (Exception e)
            {
                Close();
                return $"Exception:  {e.ToString()}";
            }
            return json.ToString();
        }
        public void Close()
        {
            //try
            //{
                Book.Close(true, null, null);
                App.Quit();

                Marshal.ReleaseComObject(Sheet);
                Marshal.ReleaseComObject(Book);
                Marshal.ReleaseComObject(App);

                LastRow = -1;
                Sheet = null;
                Book = null;
                App = null;
            //}
            //catch(Exception e)
            //{
            //    throw e;
            //}
        }
        private class MishandledExcelException : Exception{
            public MishandledExcelException(ExcelHandler _handler)
            {
                _handler.Close();
            }
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

        public void ToJSONFile(string json)
        {
            string fileName = @"C:\Users\gthao\Desktop\jsonfile.json";
            try
            {
                using(StreamWriter writer = File.CreateText(fileName))
                {
                    writer.Write(json);
                    writer.Flush();
                    writer.Close();
                }
            } catch (Exception e)
            {
                Console.Write(e.Message);
            }
        }
    }
}
