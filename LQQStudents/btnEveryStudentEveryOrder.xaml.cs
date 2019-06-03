using Microsoft.Win32;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    /// btnEveryStudentEveryOrder.xaml 的交互逻辑
    /// </summary>
    public partial class btnEveryStudentEveryOrder : Window
    {
        private List<ConsumptionInfo> consumInfos = null;
        private List<OrderInfo> orderInfos = null;
        private List<OrderInfo> newOrders = null;

        public btnEveryStudentEveryOrder()
        {
            InitializeComponent();
        }

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
                //lvCon.ItemsSource = consumInfos;
                lbMessage.Content = "end loadData";
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            lbMessage.Content = "begin analsys";
            newOrders = orderInfos.ToList();

            var result = consumInfos.OrderBy(_ => _.ComsumDate);

            var students = newOrders.Select(o => o.MemberId).Distinct();

            var studentConsumcount = new ConcurrentDictionary<string, int>();

            var totalCount = students.Count();
            var current = 0;
            progress.Maximum = totalCount;
            Task.Factory.StartNew(() =>
            {
                foreach (var studentId in students)
                {
                    current++;
                    var orders = newOrders.Where(o => o.MemberId == studentId).ToArray();
                    if (orders == null || !orders.Any())
                    {
                    }
                    else
                    {
                        try
                        {
                            int index = 0;
                            int length = orders.Length;
                            var items = result.Where(r => r.MemberId == studentId).OrderBy(_ => _.ComsumDate);
                            foreach (var item in items)
                            {
                                if (index >= length)
                                {
                                    break;
                                }
                                var order = orders[index];
                                if ((index + 1) < length)
                                {
                                    var nextOrder = orders[index + 1];
                                    order.IsExtend = true;
                                    order.ExtendOrderId = nextOrder.OrderId;
                                    order.ExtendTime = nextOrder.PayTime;
                                }

                                if (EachOrderEachStudent(studentConsumcount, item, order))
                                {
                                    index++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                    }

                    this.Dispatcher.BeginInvoke(new Action(() => progress.Value = current));
                }

                newOrder.ItemsSource = newOrders;
                lbMessage.Content = "analsys finish";
            });

            //foreach (var order in newOrders.Where(_ => !string.IsNullOrWhiteSpace(_.FinishYearMonth)))
            //{
            //    var yearMonth = order.FinishYearMonth.Split('-');
            //    var beginDate = new DateTime(int.Parse(yearMonth[0]), int.Parse(yearMonth[1]), 1);
            //    var threthholdTime = beginDate.AddMonths(1);
            //    var extendOrder = orderInfos.FirstOrDefault(o => o.OrderId != order.OrderId && o.MemberId == order.MemberId && o.PayTime < threthholdTime);
            //    if (extendOrder != null)
            //    {
            //        order.IsExtend = true;
            //        order.ExtendOrderId = extendOrder.OrderId;
            //        order.ExtendTime = extendOrder.PayTime;
            //    }
            //    var threthholdTwoTime = beginDate.AddMonths(2);

            //    var extendTwoOrder = orderInfos.FirstOrDefault(o => o.OrderId != order.OrderId && o.MemberId == order.MemberId && o.PayTime < threthholdTwoTime && o.PayTime > beginDate);
            //    if (extendTwoOrder != null)
            //    {
            //        order.IsNextTwoExtend = true;
            //        order.ExtendTwoOrderId = extendTwoOrder.OrderId;
            //        order.ExtendTwoTime = extendTwoOrder.PayTime;
            //    }
            //}
           
        }

        private static bool EachOrderEachStudent(ConcurrentDictionary<string, int> studentConsumcount, ConsumptionInfo item, OrderInfo order)
        {
            var key = GetStudentConsumeKey(item.MemberId, order.OrderId);

            var count = studentConsumcount.GetOrAdd(key, 0);

            count += item.Count;
            var result = false;
            if (count >= order.Count)
            {
                order.FinishYearMonth = item.ComsumDate.ToString("yyyy-MM");
                count -= order.Count;
                result = true;
            }
            studentConsumcount.AddOrUpdate(key, count, (old, newvalue) => count);
            return result;
        }

        private static string GetStudentConsumeKey(string memberId, string orderId)
        {
            return $"{memberId}-{orderId}";
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            MainWindow.SaveFile(newOrders);
        }
    }
}
