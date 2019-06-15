using Microsoft.Win32;
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
using System.Windows.Shapes;

namespace LQQStudents
{
    /// <summary>
    /// ExcelDivid.xaml 的交互逻辑
    /// </summary>
    public partial class ExcelDivid : Window
    {
        public ExcelDivid()
        {
            InitializeComponent();
        }
        private IEnumerable<string> Headers;
        private List<string[]> Datas;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == true)
            {
                XSSFWorkbook wk = new XSSFWorkbook(file.OpenFile());

                var sheet = wk.GetSheetAt(0);
                Headers = sheet.GetRow(0).Select(cell => cell.StringCellValue);
                cbCols.ItemsSource = Headers;
                Datas = new List<string[]>();
                var rows = sheet.GetRowEnumerator();
                rows.MoveNext();
                while (rows.MoveNext())
                {
                    var row = (IRow)rows.Current;
                    Datas.Add(row.Select(cell => GetCellValue(cell)).ToArray());
                }
            }
        }

        public static string GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return "错误格式";
                case CellType.Formula:
                    return cell.CellFormula;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.Unknown:
                case CellType.String:
                default:
                    return cell.StringCellValue;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (Headers == null || !Headers.Any() || Datas == null || !Datas.Any())
            {
                MessageBox.Show("大姐没有数据啊!");
                return;
            }

            if (cbCols.SelectedIndex == -1)
            {
                MessageBox.Show("大姐你选个列呗 :(!");
                return;
            }
            var index = cbCols.SelectedIndex;
            var groupData = Datas.GroupBy(_ => _[index]).ToDictionary(_ => _.Key, __ => __.ToList());
            using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = fbd.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    var dir = fbd.SelectedPath;
                    
                    foreach (var sheet in groupData)
                    {
                        var wk = new XSSFWorkbook();
                        var rowIndex = 0;
                        var tb = wk.CreateSheet(sheet.Key);
                        IRow row = tb.CreateRow(rowIndex);
                        var colIndex = 0;
                        foreach (var head in Headers)
                        {
                            ICell cell = row.CreateCell(colIndex);  //在第二行中创建单元格
                            cell.SetCellValue(head);//循环往第二行的单元格中添加数据
                            colIndex++;
                        }

                        foreach (var data in sheet.Value)
                        {
                            rowIndex++;
                            row = tb.CreateRow(rowIndex);
                            colIndex = 0;
                            foreach (var value in data)
                            {
                                ICell cell = row.CreateCell(colIndex);  //在第二行中创建单元格
                                cell.SetCellValue(value);//循环往第二行的单元格中添加数据
                                colIndex++;
                            }
                        }

                        var fileName = System.IO.Path.Combine(dir, sheet.Key + ".xlsx");
                        using (FileStream fs = File.OpenWrite(fileName)) 
                        {
                            wk.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。                    
                        }
                    }
                    System.Windows.MessageBox.Show("创建好了， 给个掌声呗！");
                }
            }
        }
    }
}
