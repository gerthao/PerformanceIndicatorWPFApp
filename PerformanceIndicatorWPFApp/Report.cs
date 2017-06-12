using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PerformanceIndicatorWPFApp
{
    class Report
    {
        private static int count;
        public int ID { get; set; }
        public string Name { get; set; }
        public BusinessContact Contact{get; set;}
        public Report() { }
        public Report(string name){
            ID = ++count;
            Name = name;
        }
    }
    public class ReportFactory
    {
        public static Report Create()
        {

        }
    }
}
