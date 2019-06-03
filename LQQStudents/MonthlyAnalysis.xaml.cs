using Microsoft.Win32;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LQQStudents
{
    /// <summary>
    /// MonthlyAnalysis.xaml 的交互逻辑
    /// </summary>
    public partial class MonthlyAnalysis : Window
    {
        public MonthlyAnalysis()
        {
            InitializeComponent();
        }

        private List<ConsumptionInfo> consumInfos = null;
        private List<OrderInfo> orderInfos = null;
        private List<OrderInfo> newOrders = null;

        /// <summary>
        /// load data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == true)
            {
                XSSFWorkbook wk = new XSSFWorkbook(file.OpenFile());
                lbMessage.Content = "begin loadData";
                orderInfos = MainWindow.DealWithSheet1(wk);
                lvOrders.ItemsSource = orderInfos;
                consumInfos = MainWindow.DealWithSheet2(wk);
                lbMessage.Content = "end loadData";
            }
        }

        /// <summary>
        /// analysis
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            const int threshold = 40;
            var total = new List<StudentSumConsumption>();
            for (int i = 1; i < 11; i++)
            {
                int currentConsumption = threshold * i;
                var current = FindMonthlyConsumption(currentConsumption);
                foreach (var item in current)
                {
                    var currentStudentClass = orderInfos.Where(o => o.MemberId == item.MemberId && o.PayTime < item.First40.AddDays(31)).Sum(_ => _.Count);
                    if (currentStudentClass >= (currentConsumption + threshold))
                    {
                        item.LeftClass = currentStudentClass - currentConsumption;
                    }
                }

                total.AddRange(current);
            }

            lvOrders.ItemsSource = total.OrderBy(t=>t.MemberId).ToList();
        }

        private List<StudentSumConsumption> FindMonthlyConsumption(int threshold)
        {
            var thsMonthConsumptions = consumInfos.OrderBy(_ => _.ComsumDate);
            var student_Count = new Dictionary<string, int>();
            var student_Datetime = new List<StudentSumConsumption>();
            //const int threadhold = 40;
            foreach (var item in thsMonthConsumptions)
            {
                if (student_Datetime.Any(_=>_.MemberId == item.MemberId))
                {
                    continue;
                }

                if (!student_Count.ContainsKey(item.MemberId))
                {
                    student_Count.Add(item.MemberId, 0);
                }

                student_Count[item.MemberId] += item.Count;
                if (student_Count[item.MemberId] >= threshold)
                {
                    student_Datetime.Add(new StudentSumConsumption(item.MemberId, item.ComsumDate) { Threshold = threshold });
                }
            }

            return student_Datetime;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var result = lvOrders.ItemsSource as List<StudentSumConsumption>;
            MainWindow.SaveFile(result);
        }
    }

    public class StudentSumConsumption
    {
        public string MemberId { get; set; }

        public DateTime First40 { get; set; }

        public int LeftClass { get; set; }

        public int Threshold { get; set; }

        public StudentSumConsumption(string memberId, DateTime f40)
        {
            this.MemberId = memberId;
            First40 = f40;
        }
    }
}
