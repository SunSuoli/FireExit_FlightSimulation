using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Net.Sockets;
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
        Queue q = new Queue();//创建消息队列
        public class Element//含命令和数据的类
        {
            public string command;
            public object data;
        }

        static Source Data_S_source = new Source();
        static Source Data_R_source = new Source();
        string filepath;
        public MainWindow()
        {
            InitializeComponent();

            Bind(Data_S_source, Data_S, TextBox.TextProperty, "String");
            Bind(Data_R_source, Data_R, TextBlock.TextProperty, "String");

            ThreadStart childref = new ThreadStart(Machine);
            Thread childThread = new Thread(childref);
            childThread.Start();
        }

        private void Load_Click(object sender, RoutedEventArgs e)
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
        private void Open_Click(object sender, RoutedEventArgs e)
        {
           Element a=new Element();
            a.command = "打开通讯";
            a.data = null;
            q.Enqueue(a);
        }
        private void Send_Click(object sender, RoutedEventArgs e)
        {
            Element a = new Element();
            a.command = "写入数据";
            a.data = Data_S_source.Data_String;
            q.Enqueue(a);
        }
        private void Window_Close(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Element a = new Element();
            a.command = "退出";
            a.data = Data_S_source.Data_String;
            q.Enqueue(a);
        }
        private void Machine()
        {
            Client TCP_Client = new Client();
            HandleAndStream hs =null;
            bool run = true;
            while (run)
            {
                if (q.Count > 0)
                {
                    Element element = (Element)q.Dequeue();
                    switch (element.command)
                    {
                        case "打开通讯":
                            try
                            {
                                /*如果固定本地端口，服务器里已经保存此端口，重复开关客户端，会报错！
                                 如果使用不固定端口，则每次使用一个不同的端口号，不会与服务器保存的重复，不报错！*/
                                hs = TCP_Client.TCP_Connect("192.168.0.120", 51900, 0);
                                Data_R_source.Data_String += "通讯端口打开成功！" + "\n";
                            }
                            catch(Exception e)
                            {
                                Data_R_source.Data_String += "通讯端口打开失败！" + "\n"+e.Message+ "\n";
                            }
                            break;
                        case "写入数据":
                            TCP_Client.TCP_Write(hs, (string)element.data);
                            break;
                        case "退出":
                            TCP_Client.TCP_Close(hs, null);
                            run = false;
                            break;
                        default://默认读取数据
                            
                            break;
                    }
                }
                else
                {
                    string a = TCP_Client.TCP_Read(hs, 1024, 10);
                    if (a.Length > 0) 
                    {
                        Data_R_source.Data_String +=a+"\n";
                    }
                }
                Thread.Sleep(100);
            }
        }

       
    }

}
