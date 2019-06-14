using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
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
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private List<ConsumptionInfo> consumInfos = null;
        private List<OrderInfo> orderInfos = null;
        private List<OrderInfo> newOrders = null;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == true)
            {
                XSSFWorkbook wk = new XSSFWorkbook(file.OpenFile());
                lbMessage.Content = "begin loadData";
                orderInfos = DealWithSheet1(wk);
                lvOrders.ItemsSource = orderInfos;
                consumInfos = DealWithSheet2(wk);
                //lvCon.ItemsSource = consumInfos;
                lbMessage.Content = "end loadData";
            }
        }
        
        public static List<ConsumptionInfo> DealWithSheet2(XSSFWorkbook wk)
        {
            try
            {
                var sheet = wk.GetSheet("Sheet2");
                var rows = sheet.GetRowEnumerator();
                var listOrderInfo = new List<ConsumptionInfo>();
                rows.MoveNext();
                while (rows.MoveNext())
                {
                    var row = (IRow)rows.Current;
                    var order = new ConsumptionInfo();
                    order.MemberId = row.GetCell(0).ToString().Trim();
                    order.Subject = row.GetCell(1).StringCellValue.Trim();
                    var yearMonth = row.GetCell(2).DateCellValue;
                    order.Year = yearMonth.Year;
                    order.Month = yearMonth.Month;
                    order.Day = yearMonth.Day;
                    order.ConsumCount = (int)row.GetCell(3).NumericCellValue;
                    order.Count = (int)row.GetCell(4).NumericCellValue;
                    //order.StudentName = row.GetCell(5).StringCellValue;
                    listOrderInfo.Add(order);
                }

                return listOrderInfo;
            }catch(Exception ex)
            {
                MessageBox.Show("获取Sheet2数据出问题");
                return new List<ConsumptionInfo>();
            }
        }

        public static List<OrderInfo> DealWithSheet1(XSSFWorkbook wk)
        {
            try
            {
                var sheet = wk.GetSheet("Sheet1");
                var rows = sheet.GetRowEnumerator();
                var listOrderInfo = new List<OrderInfo>();
                rows.MoveNext();
                while (rows.MoveNext())
                {
                    var row = (IRow)rows.Current;
                    var order = new OrderInfo();
                    order.OrderId = row.GetCell(0).ToString();
                    order.PayTime = row.GetCell(1).DateCellValue;
                    order.MemberId = row.GetCell(2).ToString();
                    order.Count = (int)row.GetCell(3).NumericCellValue;
                    //var yearMonth = row.GetCell(4).StringCellValue;
                    //order.OrderYear = int.Parse(yearMonth.Substring(0, 4));
                    //order.OrderMonth = int.Parse(yearMonth.Substring(5, yearMonth.Length - 5));
                    order.IsNew = row.GetCell(5).StringCellValue == "是";
                    listOrderInfo.Add(order);
                }
                //var test = listOrderInfo.Where(o => o.MemberId == -2147483648).ToList();
                return listOrderInfo;
            }catch(Exception ex)
            {
                MessageBox.Show("获取Sheet1数据出问题");
                return new List<OrderInfo>();
            }
        }

        private void btnEveryStudentEveryOrder_Click(object sender, RoutedEventArgs e)
        {
            var window = new btnEveryStudentEveryOrder()
            {
                Title = "每个学生每笔订单续费情况"
            };
            window.Show();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            lbMessage.Content = "begin analsys";
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
                if(order == null)
                {
                    continue;
                }

                if (!studentConsumcount.ContainsKey(item.MemberId))
                {
                    studentConsumcount.Add(item.MemberId, 0);
                }
                studentConsumcount[item.MemberId] += item.Count;
                if(studentConsumcount[item.MemberId] >= order.Count)
                {
                    studentDict.Add(item.MemberId, item.ComsumDate);
                    order.FinishYearMonth = item.ComsumDate.ToString("yyyy-MM");
                }
            }

            foreach (var order in newOrders.Where(_ => !string.IsNullOrWhiteSpace(_.FinishYearMonth)))
            {
                var yearMonth = order.FinishYearMonth.Split('-');
                var beginDate = new DateTime(int.Parse(yearMonth[0]), int.Parse(yearMonth[1]), 1);
                var threthholdTime = beginDate.AddMonths(1);
                var extendOrder = orderInfos.FirstOrDefault(o => o.OrderId != order.OrderId && o.MemberId == order.MemberId && o.PayTime < threthholdTime);
                if (extendOrder != null)
                {
                    order.IsExtend = true;
                    order.ExtendOrderId = extendOrder.OrderId;
                    order.ExtendTime = extendOrder.PayTime;
                }
                var threthholdTwoTime = beginDate.AddMonths(2);
                
                var extendTwoOrder = orderInfos.FirstOrDefault(o => o.OrderId != order.OrderId && o.MemberId == order.MemberId && o.PayTime < threthholdTwoTime && o.PayTime > beginDate);
                if(extendTwoOrder != null)
                {
                    order.IsNextTwoExtend = true;
                    order.ExtendTwoOrderId = extendTwoOrder.OrderId;
                    order.ExtendTwoTime = extendTwoOrder.PayTime;
                }
            }
            newOrder.ItemsSource = newOrders;
            lbMessage.Content = "analsys finish";

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            SaveFile(newOrders);
        }

        public static void SaveFile<T>(List<T> orders)
        {
            var saveFile = new SaveFileDialog();
            if (saveFile.ShowDialog() == true)
            {
                var wk = new XSSFWorkbook();
                var tb = wk.CreateSheet("WLGZS");
                var rowIndex = 1;
                IRow row = tb.CreateRow(rowIndex);
                rowIndex++;
                var properties = typeof(T).GetProperties();
                var index = 1;
                foreach (var p in properties)
                {
                    ICell cell = row.CreateCell(index);  //在第二行中创建单元格
                    cell.SetCellValue(p.Name);//循环往第二行的单元格中添加数据
                    index++;
                }

                foreach (var order in orders)
                {
                    row = tb.CreateRow(rowIndex);
                    rowIndex++;
                    index = 1;
                    foreach (var p in properties)
                    {
                        ICell cell = row.CreateCell(index);  //在第二行中创建单元格
                        var value = p.GetValue(order, null);
                        if (value == null)
                        {
                            value = string.Empty;
                        }
                        cell.SetCellValue(value.ToString());//循环往第二行的单元格中添加数据
                        index++;
                    }
                }

                using (FileStream fs = File.OpenWrite(saveFile.FileName)) //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
                {
                    wk.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。
                    MessageBox.Show("提示：创建成功！");
                }
            }
        }

        private void Button_Click_Switch(object sender, RoutedEventArgs e)
        {
            var window = new MonthlyAnalysis();
            window.Show();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var window = new WindowLeft40();
            window.Show();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            StudentWindow sw = new StudentWindow();
            sw.Show();
        }

        private void btnyingxushixu_Click(object sender, RoutedEventArgs e)
        {
            ShouldHaveWindow window = new ShouldHaveWindow();
            window.Show();
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            var window = new ExcelDivid();
            window.Show();
        }
    }

    public class OrderInfo
    {
        public string OrderId { get; set; }

        public DateTime PayTime { get; set; }

        public string MemberId { get; set; }

        public int Count { get; set; }

        public int OrderYear { get; set; }

        public int OrderMonth { get; set; }

        public bool IsNew { get; set; }

        public string FinishYearMonth { get; set; }

        /// <summary>
        /// 續費
        /// </summary>
        public bool IsExtend { get; set; }

        public string ExtendOrderId { get; set; }

        public DateTime? ExtendTime { get; set; }

        public bool IsNextTwoExtend { get; set; }

        public string ExtendTwoOrderId { get; set; }

        public DateTime? ExtendTwoTime { get; set; }

        public override string ToString()
        {
            return string.Format("Order:{0}-studentdent:{1}-count:{2}-finish:{3}", OrderId, MemberId, Count, FinishYearMonth);
        }
    }

    public class ConsumptionInfo
    {
        public string MemberId { get; set; }

        public string StudentName { get; set; }

        public string Subject { get; set; }

        public int Year { get; set; }

        public int Month { get; set; }

        public int? Day { get; set; }

        public DateTime ComsumDate
        {
            get
            {
                return new DateTime(Year, Month, Day ?? 1);
            }
        }

        /// <summary>
        /// 課次
        /// </summary>
        public int ConsumCount { get; set; }

        /// <summary>
        /// 課時
        /// </summary>
        public int Count { get; set; }

        public override string ToString()
        {
            return string.Format("member:{0}-Subject:{1}-consum:{2}-count:{3}", MemberId, Subject, ComsumDate.ToString("yyyy-MM"), ConsumCount); 
        }
    }
}
