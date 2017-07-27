using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PerformanceIndicatorWPFApp
{
    public abstract class Report
    {
        protected DateTime ExcelBaseDate = new DateTime(month: 12, day: 30, year: 1899);
        protected Dictionary<string, string> Data;
        public abstract string ToJson();
    }
}
