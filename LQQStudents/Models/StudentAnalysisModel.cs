using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LQQStudents.Models
{
    public class StudentAnalysisModel
    {
        public string StudentId { get; set; }

        public DateTime AnalysisDate { get; set; }

        public int Year { get { return AnalysisDate.Year; } }

        public int Month { get { return AnalysisDate.Month; } }

        public string YearAndMonth { get { return AnalysisDate.ToString("yyyy-MM"); } }

        public int TotalClassCount { get; set; }

        public int TotalConsumption { get; set; }

        public int LeftClassCount { get { return TotalClassCount - TotalConsumption; } }

        public string OrderId { get; set; }

        public DateTime? BuyTime { get; set; }
    }
}
