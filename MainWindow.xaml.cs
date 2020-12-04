using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Threading;
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

       UDPHelper UDP_CONNECT=new UDPHelper();
        public MainWindow()
        {
            InitializeComponent();

            Bind(Data_S_source, Data_S, TextBox.TextProperty, "String");
            Bind(Data_R_source, Data_R, TextBlock.TextProperty, "String");

            ThreadStart childref = new ThreadStart(UDP_Receview);
            Thread childThread = new Thread(childref);
            childThread.Start();
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
            EXCEL.Range_Select(-1, 2,5,-1);
            string[,]  object_arry =EXCEL.Range_GetValue();
            Console.WriteLine(object_arry[0,0]);
            EXCEL.File_SaveAs(filepath);
            EXCEL.File_Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            UDP_CONNECT.UDP_Open("192.168.1.247", 51901);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            UDP_CONNECT.UDP_Write(Data_S_source.Data_String, "255.255.255.255", 51900);
        }

        private void UDP_Receview()
        {
            string data;
            string ip;
            int port;
            while (true)
            {
               UDP_CONNECT.UDP_Read( out data,out ip,out port);
                if (data != "")
                {
                    Data_R_source.Data_String += ip +":" +port.ToString() + "：" + data+"\n";
                }
                Thread.Sleep(100);
            }
        }
    }

}
