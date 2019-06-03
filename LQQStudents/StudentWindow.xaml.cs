using LQQStudents.Models;
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
using System.Windows.Shapes;

namespace LQQStudents
{
    /// <summary>
    /// StudentWindow.xaml 的交互逻辑
    /// </summary>
    public partial class StudentWindow : Window
    {
        private List<ConsumptionInfo> consumInfos = null;
        private List<OrderInfo> orderInfos = null;
       // private List<OrderInfo> newOrders = null;
        private List<StudentAnalysisModel> analysisModels = null;

        public StudentWindow()
        {
            InitializeComponent();
        }

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
                dgSheet1.ItemsSource = orderInfos;
                consumInfos = MainWindow.DealWithSheet2(wk);
                dgSheet2.ItemsSource = consumInfos;
                lbMessage.Content = "end loadData";
            }
        }
        readonly DateTime BeginDate = new DateTime(2018, 1, 1);
        readonly DateTime EndDate = new DateTime(2018, 11, 1);
        /// <summary>
        /// analysis
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            analysisModels = new List<StudentAnalysisModel>();
            DateTime currentDate = BeginDate;
            while (currentDate <= EndDate)
            {
                var theseOrders = orderInfos.Where(o => o.PayTime < currentDate);
                var thisMonthModels = from o in theseOrders
                                      group o by o.MemberId into g
                                      select new
                                      {
                                          StudentId = g.Key,
                                          TotalClassCount = g.Sum(_ => _.Count)
                                      };
                var consums = consumInfos.Where(c => c.ComsumDate < currentDate);
                var thisMonthConsum = from c in consums
                                      group c by c.MemberId into cc
                                      select new
                                      {
                                          StudentId = cc.Key,
                                          ConsumTotal = cc.Sum(_ => _.Count)
                                      };

                var models = (from m in thisMonthModels
                             join c in thisMonthConsum on m.StudentId equals c.StudentId
                             into temp
                             from t in temp.DefaultIfEmpty()
                             select new StudentAnalysisModel()
                             {
                                 StudentId = m.StudentId,
                                 TotalClassCount = m.TotalClassCount,
                                 TotalConsumption = t == null? 0 : t.ConsumTotal,
                                 AnalysisDate = currentDate,
                             }).ToList();

                var nextMonth = currentDate.AddMonths(1);

                var nextMonthOrders = orderInfos.Where(o => o.PayTime > currentDate && o.PayTime < nextMonth);
                models.ForEach(m =>
                {
                    var sequalOrder = nextMonthOrders.FirstOrDefault(_ => _.MemberId == m.StudentId);
                    if(sequalOrder != null)
                    {
                        m.OrderId = sequalOrder.OrderId;
                        m.BuyTime = sequalOrder.PayTime;
                    }
                });

                analysisModels.AddRange(models);

                currentDate = nextMonth;
            }

            analysisModels = analysisModels.OrderBy(_ => _.AnalysisDate).OrderBy(_ => _.StudentId).ToList();
            dgSheet2.ItemsSource = analysisModels;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            MainWindow.SaveFile<StudentAnalysisModel>(analysisModels);
        }
    }
}
