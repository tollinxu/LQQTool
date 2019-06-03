using Microsoft.Win32;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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
    /// ShouldHaveWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ShouldHaveWindow : Window
    {
        public ShouldHaveWindow()
        {
            InitializeComponent();
        }

        private List<ConsumptionInfo>  consumInfos;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == true)
            {
                XSSFWorkbook wk = new XSSFWorkbook(file.OpenFile());
                lbMessage.Content = "begin loadData";
               
                consumInfos = MainWindow.DealWithSheet2(wk);
                dgLoad.ItemsSource = consumInfos;
                lbMessage.Content = "end loadData";
            }
        }
        const int threadhold = 40;
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var thread = new Thread(new ThreadStart(() =>
            {

                //System.Threading.Tasks.Task.Factory.StartNew(() =>
                //{
                this.Dispatcher.BeginInvoke(new Action(() => lbMessage.Content = "begin analysis"));

                var infos = consumInfos.Select(c => c.MemberId.Trim()).Distinct();
                var subjects = consumInfos.Select(c => c.Subject).Distinct();
                var subjectStudents = consumInfos.GroupBy(_ => _.Subject).ToDictionary(_ => _.Key, g => g.Select(_ => _.MemberId).Distinct());
                var min = consumInfos.Min(_ => _.ComsumDate);
                var endDate = consumInfos.Max(_ => _.ComsumDate).AddMonths(1);

                //while (min.Year < endDate.Year || (min.Month <= endDate.Month && min.Year == endDate.Year))
                //{

                //    min = min.AddMonths(1);
                //}

                results = new List<ResultClass>();

                System.Threading.Tasks.Parallel.ForEach(subjectStudents, sub =>
                //foreach (var sub in subjectStudents)
                {
                    //var t = new Thread(new ThreadStart(() =>
                    //     {
                    var subResult = new List<ResultClass>();
                    var accumulate = new ConcurrentDictionary<string, int>();
                    foreach (var student in sub.Value)
                    //System.Threading.Tasks.Parallel.ForEach(infos, student =>
                    {
                        //this.Dispatcher.BeginInvoke(new Action(() => lbMessage.Content = $"1---send:{student}"));
                        var studentCost = consumInfos.Where(_ => _.MemberId == student && _.Subject == sub.Key).ToList();
                        var current = studentCost.Min(_ => _.ComsumDate);
                        var studentMaxDate = studentCost.Max(_ => _.ComsumDate).AddMonths(1);


                        while (current.Year < studentMaxDate.Year || (current.Month <= studentMaxDate.Month && current.Year == studentMaxDate.Year))
                        {
                            //Total += 
                            var time = current.ToString("yyyy-MM");
                            var begin = DateTime.Now;
                            var beginTime = begin.ToString("HH:mm:ss");
                            this.Dispatcher.BeginInvoke(new Action(() => lbMessage.Content = $"1---send:{student} - {sub.Key} - {time}--begin:{beginTime}"));
                            var thisTime = current;
                            System.Threading.Tasks.Task.Factory.StartNew(Add);
                            CaculateEach(thisTime, accumulate, student, sub.Key, subResult, studentCost.Where(cc => cc.Month == current.Month && cc.Year == current.Year));
                            var during = (DateTime.Now - begin).TotalSeconds;
                            var endTime = DateTime.Now.ToString("HH:mm:ss");
                            this.Dispatcher.BeginInvoke(new Action(() => lbMessage.Content = $"1---send:{student} - {sub.Key} - {time}--end:{endTime} --- during:{during}s"));
                            current = current.AddMonths(1);
                        }
                    }
                    results.AddRange(subResult);

                    // }));
                    //  t.Start();
                });
                // }
                this.Dispatcher.BeginInvoke(new Action(() => lbMessage.Content = "计算完成"));
            }));
            thread.Start();
            dg_Result.ItemsSource = results;
            //});
        }

        static int Counter = 0;
        static int Total = 0;
        private static readonly object obj = new object();
        private void Add()
        {
            lock (obj)
            {
                Counter++;
                this.Dispatcher.BeginInvoke(new Action(() => lbMessage2.Content = "count:" + Counter + "/" + Total));
            }
        }
        
        private void CaculateEach(DateTime current, ConcurrentDictionary<string, int> accumulate, string student, string sub, List<ResultClass> subResult, IEnumerable<ConsumptionInfo> consumptions)
        {
            var consumDatas = consumptions.Where(c => c.MemberId == student && c.ComsumDate.Year == current.Year && c.ComsumDate.Month == current.Month && c.Subject == sub);

            var currentStd = subResult.FirstOrDefault(re => re.StudentId == student && re.YearMonthString == current.ToString("yyyy-MM") && re.Subject == sub);
            if (currentStd == null)
            {
                currentStd = new ResultClass()
                {
                    StudentId = student,
                    YearMonth = current,
                    Subject = sub
                };
                subResult.Add(currentStd);
            }

            var shouldbuyKey = ShouldBuyDict(student, sub);

            if (!accumulate.ContainsKey(shouldbuyKey))
            {
                accumulate.TryAdd(shouldbuyKey, 0);
            }
            
            if (consumDatas != null && consumDatas.Any())
            {
                currentStd.MonthConsume = consumDatas.Sum(c => c.Count);
                currentStd.StartDate = consumDatas.Min(_ => _.ComsumDate);
                currentStd.EndDate = consumDatas.Max(_ => _.ComsumDate);

                var thisIsShouldBuy = false;
                //把之前的记录改为停课
                for (int i = subResult.Count - 1; i >= 0; i--)
                {
                    var temp = subResult[i];

                    if(currentStd.Subject == temp.Subject && temp.StudentId == currentStd.StudentId && currentStd.YearMonthString == temp.YearMonthString)
                    {
                        continue;
                    }

                    if (temp.StudentId != student || temp.Subject != sub)
                    {
                        continue;
                    }
                    
                    if (temp.StudentCurrentStatus == StudentStatus.Reading)
                    {
                        break;
                    }
                    temp.StudentCurrentStatus = StudentStatus.Stop;
                    if (temp.ShouldBuy && !thisIsShouldBuy)
                    {
                        thisIsShouldBuy = true;
                    }
                    temp.ShouldBuy = false;
                }
                
                if (accumulate[shouldbuyKey] == 0)
                {
                    currentStd.IsBegin = true;
                }

                if (thisIsShouldBuy)
                {
                    currentStd.ShouldBuy = true;
                }

                accumulate[shouldbuyKey] += currentStd.MonthConsume;

                if (accumulate[shouldbuyKey] >= threadhold)
                {
                    

                    if (!CanShouldBuyList.Contains(shouldbuyKey))
                    {
                        CanShouldBuyList.Add(shouldbuyKey);
                    }
                    var next = new ResultClass()
                    {
                        StudentId = student,
                        YearMonth = current.AddMonths(1),
                        Subject = sub,
                        IsBegin = true,
                    };
                    currentStd.IsEnd = true;
                    accumulate[shouldbuyKey] -= threadhold;
                    next.ShouldBuy = true;
                    if (!hasShouldBuy.ContainsKey(shouldbuyKey))
                    {
                        //currentStd.ShouldBuy = true;
                        hasShouldBuy.Add(shouldbuyKey, next);
                    }
                    if (accumulate[shouldbuyKey] >= threadhold)
                    {
                        currentStd.IsBegin = true;
                        accumulate[shouldbuyKey] -= threadhold;
                    }

                    subResult.Add(next);
                }

                //有消费记录叫在读
                currentStd.StudentCurrentStatus = StudentStatus.Reading;
            }
            else
            {
                //没有消费记录叫流失
                currentStd.StudentCurrentStatus = StudentStatus.Lost;
                if (CanShouldBuyList.Contains(shouldbuyKey) && !hasShouldBuy.ContainsKey(shouldbuyKey))
                {
                    currentStd.ShouldBuy = true;
                    hasShouldBuy.Add(shouldbuyKey, currentStd);
                }
            }

            
        }

        string ShouldBuyDict(string studentId, string subject)
        {
            return $"{studentId}-{subject}".ToLower();
        }

        Dictionary<string, ResultClass> hasShouldBuy = new Dictionary<string, ResultClass>();
        private List<string> CanShouldBuyList = new List<string>();
        List<ResultClass> results;
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            MainWindow.SaveFile(results);
        }
    }

    public class ResultClass
    {
        public string StudentId { get; set; }

        public string Subject { get; set; }

        public DateTime?  YearMonth { get; set; }

        public string YearMonthString { get { return YearMonth.HasValue ? YearMonth.Value.ToString("yyyy-MM") : "0000-00";  } }

        public int MonthConsume { get; set; }

        public DateTime? StartDate { get; set; }

        public DateTime? EndDate { get; set; }

        public bool IsBegin { get; set; }

        public string BeginString { get { return IsBegin ? "是" : "否"; } }
        
        public bool IsEnd { get; set; }

        public string EndString { get { return IsEnd ? "是" : "否"; } }

        public StudentStatus StudentCurrentStatus { get; set; }

        public bool ShouldBuy { get; set; }
    }

    public enum StudentStatus
    {
        Reading = 1,

        Lost = 2,

        Stop = 3
    }

    public class StudentKey
    {
        public string StudentId { get; set; }

        public string Subject { get; set; }

        public override bool Equals(object obj)
        {
            var key = obj as StudentKey;
            if(key == null)
            {
                return false;
            }
            return this.ToString() == key.ToString();
        }

        public override string ToString()
        {
            return $"{this.StudentId}-{this.Subject}".ToLower();
        }
    }
}
