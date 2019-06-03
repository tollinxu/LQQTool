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
    /// WindowLeft40.xaml 的交互逻辑
    /// </summary>
    public partial class WindowLeft40 : Window
    {
        public WindowLeft40()
        {
            InitializeComponent();
        }
        private List<ConsumptionInfo> consumInfos = null;
        private List<OrderInfo> orderInfos = null;
        private List<OrderInfo> newOrders = null;
        /// <summary>
        /// load
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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            int leftCount = 40;
            int.TryParse(tbLeftCount.Text, out leftCount);
            lbMessage.Content = "begin analsys。。。";
            
            newOrders = orderInfos.Where(_ => _.IsNew).ToList();

            var result = consumInfos.OrderBy(_ => _.ComsumDate);

            var studentDict = new Dictionary<string, DateTime>();
            var studentConsumcount = new Dictionary<string, int>();
            foreach (var item in result)
            {
                if (studentDict.ContainsKey(item.MemberId))
                {
                    continue;
                }
                var order = newOrders.FirstOrDefault(o => o.MemberId == item.MemberId);
                if (order == null)
                {
                    continue;
                }

                if (!studentConsumcount.ContainsKey(item.MemberId))
                {
                    studentConsumcount.Add(item.MemberId, 0);
                }
                studentConsumcount[item.MemberId] += item.Count;
                if (order.Count - studentConsumcount[item.MemberId] < leftCount)
                {
                    studentDict.Add(item.MemberId, item.ComsumDate);
                    order.FinishYearMonth = item.ComsumDate.ToString("yyyy-MM-dd");
                }
            }

            foreach (var order in newOrders.Where(_ => !string.IsNullOrWhiteSpace(_.FinishYearMonth)))
            {
                //var yearMonth = DateTime.Parse(order.FinishYearMonth);
                var beginDate = DateTime.Parse(order.FinishYearMonth);// new DateTime(int.Parse(yearMonth[0]), int.Parse(yearMonth[1]), 1);
                var threthholdTime = beginDate.AddMonths(2);
                var extendOrder = orderInfos.FirstOrDefault(o => o.OrderId != order.OrderId && o.MemberId == order.MemberId && o.PayTime < threthholdTime && o.PayTime > order.PayTime);
                if (extendOrder != null)
                {
                    order.IsExtend = true;
                    order.ExtendOrderId = extendOrder.OrderId;
                    order.ExtendTime = extendOrder.PayTime;
                }
                var threthholdTwoTime = beginDate.AddMonths(2);

                var extendTwoOrder = orderInfos.FirstOrDefault(o => o.OrderId != order.OrderId && o.MemberId == order.MemberId && o.PayTime < threthholdTwoTime && o.PayTime > beginDate);
                if (extendTwoOrder != null)
                {
                    order.IsNextTwoExtend = true;
                    order.ExtendTwoOrderId = extendTwoOrder.OrderId;
                    order.ExtendTwoTime = extendTwoOrder.PayTime;
                }
            }
            dgSheet2.ItemsSource = newOrders;
            lbMessage.Content = "analsys finish";
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            MainWindow.SaveFile(newOrders);
        }
    }
}
