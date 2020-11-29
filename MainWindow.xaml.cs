using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;
using static FireExit_FlightSimulation.BinDing;

namespace FireExit_FlightSimulation
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        
        Source Data_S_source = new Source();
        Source Data_R_source = new Source();
        string filepath;
        Excel_NPIO EXCEL=new Excel_NPIO();
        public MainWindow()
        {
            InitializeComponent();

            Bind(Data_S_source, Data_S, TextBox.TextProperty, "String");
            Bind(Data_R_source, Data_R, TextBlock.TextProperty, "String");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog fileDialog = new OpenFileDialog();
            //SaveFileDialog fileDialog = new SaveFileDialog();
            //FolderBrowserDialog fileDialog = new FolderBrowserDialog();
            //fileDialog.Multiselect = true;//允许多选
            fileDialog.RestoreDirectory = false;//使用上次打开目录
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "Excel(*.xlsx)|*.xlsx;*.xls";
            if (fileDialog.ShowDialog()==true)
            {
                filepath = fileDialog.FileName;
            }
            EXCEL.Open_Create_WorkBook(filepath);
            EXCEL.Open_WorkSheet(0);
            Console.WriteLine( EXCEL.Read_WorkSheet().Count);
            Console.WriteLine(EXCEL.Read_WorkSheet()[0].Count);
        }
    }
}
