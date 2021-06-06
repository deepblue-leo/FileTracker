using System;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Spire.Xls;
using Spire.Xls.Charts;
using System.Drawing;
using Microsoft.Win32;
using ExcelParser;

namespace FileTracker
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private IExcelParser _EP = new ExcelParser.ExcelParser();

        public MainWindow()
        {
            InitializeComponent();
            MasterFilePathTxt.Text = string.Empty;
            SlaveFilePathTxt.Text = string.Empty;

            MasterFileIndexFilterTxt.Text = string.Empty;
            SlaveFileIndexFilterTxt.Text = string.Empty;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            sheet.Range["A2"].Text = "出口前";
            sheet.Range["A5"].Text = "出口后";
            sheet.Range["B1"].Text = "年份";
            sheet.Range["B2"].Text = "2017年";
            sheet.Range["B6"].Text = "2018年";
            sheet.Range["C1"].Text = "季度";
            sheet.Range["C2"].Text = "1季度";
            sheet.Range["C3"].Text = "2季度";
            sheet.Range["C4"].Text = "3季度";
            sheet.Range["C5"].Text = "4季度";
            sheet.Range["C6"].Text = "1季度";
            sheet.Range["C7"].Text = "2季度";
            sheet.Range["C8"].Text = "3季度";
            sheet.Range["C9"].Text = "4季度";
            sheet.Range["D1"].Text = "季度产量\n（万吨）";
            sheet.Range["D2"].Value = "1.56";
            sheet.Range["D3"].Value = "2.3";
            sheet.Range["D4"].Value = "3.21";
            sheet.Range["D5"].Value = "3.5";
            sheet.Range["D6"].Value = "4.8";
            sheet.Range["D7"].Value = "5.2";
            sheet.Range["D8"].Value = "5.79";
            sheet.Range["D9"].Value = "5.58";

            sheet.Range["A2:A4"].Merge();
            sheet.Range["A5:A9"].Merge();
            sheet.Range["B2:B5"].Merge();
            sheet.Range["B6:B9"].Merge();
            sheet.Range["A1:D9"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["A1:D9"].Style.VerticalAlignment = VerticalAlignType.Center;

            //添加柱状图表
            Chart chart = sheet.Charts.Add(ExcelChartType.LineMarkers);
            chart.ChartTitle = "季度产量（万吨）";//设置图表标题     
            //chart.PlotArea.Fill.FillType = ShapeFillType.NoFill; //不填充绘图区域（默认填充灰色）
            chart.Legend.Delete();//删除图例

            //指定柱状图表在工作表中的位置及宽度
            chart.LeftColumn = 5;
            chart.TopRow = 1;
            chart.RightColumn = 14;

            //设置图表系列数据来源
            chart.DataRange = sheet.Range["D2:D9"];
            chart.SeriesDataFromRange = false;
            chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            chart.Series[0].Format.LineProperties.Color = System.Drawing.Color.BlueViolet;

            //设置系列分类标签数据来源
            ChartSerie serie = chart.Series[0];
            serie.CategoryLabels = sheet.Range["A2:C9"];

            chart.PrimaryCategoryAxis.MultiLevelLable = true;

            //保存文档
            wb.SaveToFile("output.xlsx", ExcelVersion.Version2013);
            System.Diagnostics.Process.Start("output.xlsx");
        }

        private void MasterFileOpenBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {
                Workbook wb = _EP.LoadExcel(ofd.FileName);

                MasterFilePathTxt.Text = ofd.FileName.Trim();

                //MasterFileHeader.ItemsSource = _EP.GetHeads(wb.ActiveSheet);
                //MasterFileHeader.DisplayMemberPath = 
                List<string> ll = _EP.GetHeads(wb.ActiveSheet);
                foreach (var item in ll)
                {
                    MasterFileIndexColumn.Items.Add(item);
                    MasterFileDisplayColumn.Items.Add(item);
                }
            }
        }

        private static OpenFileDialog OpenFileDialog()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "Excel文件(*.xlsx)|*.xlsx|Excel文件(*.xls)|*.xls|Csv文件(*.csv)|*.csv|所有文件(*.*)|*.*"; //设置“另存为文件类型”或“文件类型”框中出现的选择内容
            ofd.FilterIndex = 1; //设置默认显示文件类型为Csv文件(*.csv)|*.csv
            ofd.Title = "打开文件"; //获取或设置文件对话框标题
            ofd.RestoreDirectory = true;
            return ofd;
        }

        private void Demo1()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Range["A1"].Value = "公司部门";
            sheet.Range["A3"].Value = "综合部";
            sheet.Range["A4"].Value = "行政";
            sheet.Range["A5"].Value = "人事";
            sheet.Range["A6"].Value = "市场部";
            sheet.Range["A7"].Value = "业务部";
            sheet.Range["A8"].Value = "客服部";
            sheet.Range["A9"].Value = "技术部";
            sheet.Range["A10"].Value = "技术开发";
            sheet.Range["A11"].Value = "技术支持";
            sheet.Range["A12"].Value = "售前支持";
            sheet.Range["A13"].Value = "售后支持";

            sheet.PageSetup.IsSummaryRowBelow = false;

            //选择行进行一级分组
            sheet.GroupByRows(2, 13, false);
            //选择行进行二级分组
            sheet.GroupByRows(4, 5, false);
            sheet.GroupByRows(7, 8, false);
            sheet.GroupByRows(10, 13, false);
            //选择行进行三级分组
            sheet.GroupByRows(12, 13, true);


            CellStyle style = workbook.Styles.Add("style");
            style.Font.IsBold = true;
            style.Color = System.Drawing.Color.LawnGreen;
            sheet.Range["A1"].CellStyleName = style.Name;
            sheet.Range["A3"].CellStyleName = style.Name;
            sheet.Range["A6"].CellStyleName = style.Name;
            sheet.Range["A9"].CellStyleName = style.Name;

            sheet.Range["A4:A5"].BorderAround(LineStyleType.Thin);
            sheet.Range["A4:A5"].BorderInside(LineStyleType.Thin);
            sheet.Range["A7:A8"].BorderAround(LineStyleType.Thin);
            sheet.Range["A7:A8"].BorderInside(LineStyleType.Thin);
            sheet.Range["A10:A13"].BorderAround(LineStyleType.Thin);
            sheet.Range["A10:A13"].BorderInside(LineStyleType.Thin);

            workbook.SaveToFile("output.xlsx", ExcelVersion.Version2013);
        }

        private void MasterFileIndexColumn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string indexColumn = (string)MasterFileDisplayColumn.Items[MasterFileIndexColumn.SelectedIndex];
            int a = MasterFileDisplayColumn.SelectedItems.IndexOf(indexColumn);
            if (a < 0)
            {
                MasterFileDisplayColumn.SelectedItems.Add(indexColumn);
            }            
        }

        private void SlaveFileOpenBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {
                Workbook wb = _EP.LoadExcel(ofd.FileName);

                SlaveFilePathTxt.Text = ofd.FileName.Trim();
                //MasterFileHeader.ItemsSource = _EP.GetHeads(wb.ActiveSheet);
                //MasterFileHeader.DisplayMemberPath = 
                List<string> ll = _EP.GetHeads(wb.ActiveSheet);
                foreach (var item in ll)
                {
                    SlaveFileIndexColumn.Items.Add(item);
                    SlaveFileDisplayColumn.Items.Add(item);
                }
            }
        }

        private void SlaveFileIndexColumn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string indexColumn = (string)SlaveFileDisplayColumn.Items[SlaveFileIndexColumn.SelectedIndex];
            int a = SlaveFileDisplayColumn.SelectedItems.IndexOf(indexColumn);
            if (a < 0)
            {
                SlaveFileDisplayColumn.SelectedItems.Add(indexColumn);
            }
        }
    }
}
