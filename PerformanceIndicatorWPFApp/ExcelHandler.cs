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
        public enum Months { Jan=5, Feb, Mar, Qrt1, PrevQrt1, Apr, May, Jun, Qrt2, PrevQrt2, Jul, Aug, Sep, Qrt3, PrevQrt3, Oct, Nov, Dec, Qrt4, PrevQrt4, }
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
                json.Append("\"Document\" : ");
                json.AppendLine($"\"{Book.Name}\",");

                for (int j = tablevel; j > 0; j--) json.Append(tab);
                json.Append("\"Report\" : ");
                json.AppendLine($"\"{GetCell(1, 1).Text}\",");

                for (int j = tablevel; j > 0; j--) json.Append(tab);
                json.Append("\"Plan\" : ");
                json.AppendLine($"\"{GetCell(3, 1).Text}\",");

                for (int j = tablevel; j > 0; j--) json.Append(tab);
                json.Append("\"Year\" : ");
                json.AppendLine($"\"{GetCell(4, 1).Text}\",");


                for (int j = tablevel; j > 0; j--) json.Append(tab);
                json.AppendLine("\"Tables\" : [");
                tablevel++;

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
                        json.AppendLine($"\"TableName\" : \"{cell.Value2}\",");

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
                        json.AppendLine($"\"TableName\" : \"{cell.Value2}\",");

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine("\"Rows\" : [");
                        tablevel++;
                    }
                    else
                    {
                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine(@"{");
                        tablevel++;

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.Append("\"RowName\" : ");
                        json.AppendLine($"\"{(cell as Excel.Range).Value2}\",");

                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine("\"Data\" : [");
                        tablevel++;

                        for (int k = 4; k <= 25; k++)
                        {
                            for (int j = tablevel; j > 0; j--) json.Append(tab);
                            json.Append(@"{  ");

                            //Excel.Range temp = Sheet.Range[("D" + i), ("G" + i)] as Excel.Range;
                            //Excel.Range temp2 = Sheet.Range[("I" + i), ("L" + i)] as Excel.Range;
                            //Excel.Range temp3 = Sheet.Range[("N" + i), ("Q" + i)] as Excel.Range;
                            //Excel.Range temp4 = Sheet.Range[("S" + i), ("V" + i)] as Excel.Range;

                            json.Append("\"CellAddress\" : ");
                            json.Append($"\"{GetCell(i, k).Address}\",  ");
                            json.Append("\"Month\" : ");
                            json.Append($"\"{ GetCell(8, k).Value2}\",  ");
                            json.Append("\"Value\" : ");
                            json.Append($"\"{ GetCell(i, k).Value2}\",  ");
                            json.Append("\"NumberFormat\" : ");
                            json.Append($"\"{ GetCell(i, k).NumberFormat}\",  ");
                            json.Append("\"HasForumla\" : ");
                            json.Append($"{ GetCell(i, k).HasFormula},  ");
                            json.Append("\"Formula\" : ");
                            json.Append($"\"{ GetCell(i, k).Formula}\",  ");
                            json.Append("\"Color\" : ");
                            json.Append($"\"{ GetCell(i, k).Interior.Color}\"  ");

                            json.Append(@"}");

                            if (k == 25) json.AppendLine();
                            else json.AppendLine(",");
                        }
                        tablevel--;
                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        json.AppendLine("]");

                        tablevel--;
                        for (int j = tablevel; j > 0; j--) json.Append(tab);
                        if(GetCell(i+1, 2).Font.Bold || GetCell(i+1, 2).Value2 == null) json.AppendLine(@"}");
                        else json.AppendLine(@"},");

                    }
                }
                if (!newTable)
                {
                    for (int j = tablevel-1; j > 0; j--) json.Append(tab);
                    json.AppendLine(@"]");
                    tablevel--;

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
