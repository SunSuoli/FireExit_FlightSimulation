using Microsoft.Win32;
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
        
        static Source Data_S_source = new Source();
        static Source Data_R_source = new Source();
        string filepath;
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
            ExcelHelper EXCEL = new ExcelHelper();
            EXCEL.File_OpenorCreate(filepath);
            EXCEL.WorkSheet_Choose(1);
            EXCEL.Cloum_Add(2);
            EXCEL.Cloum_Delete(4);
            EXCEL.File_SaveAs(filepath);
            EXCEL.File_Close();
        }
    }

}
