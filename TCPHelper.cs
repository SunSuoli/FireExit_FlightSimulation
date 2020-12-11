using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace FireExit_FlightSimulation
{
    public class HandleAndStream//创建一个公共的TCP句柄和数据流类
    {
        public TcpClient handle;
        public NetworkStream stream;
    }
    class TCPHelper
    {
        public void TCP_Write(HandleAndStream hs, string message)//发送数据
        {
            Byte[] data_Send = Encoding.Default.GetBytes(message);
            int data_length = data_Send.Length;
            hs.stream.Write(data_Send,0, data_length);
        }
        public string TCP_Read(HandleAndStream hs, int lenght,int time_out)//接收数据
        {
            string message = "";
            int i = 0;
            bool read = true;
            while (read)
            {
                if (hs != null)
                {
                    //利用Available属性可以使阻塞式IO当做不阻塞使用
                    if (hs.handle.Available >= lenght)//如果数据长度满足则停止读取
                    {
                        Byte[] data = new Byte[lenght];
                        Int32 bytes = hs.stream.Read(data, 0, data.Length);
                        message = Encoding.Default.GetString(data, 0, bytes);
                        read = false;
                    }
                    else if (i >= time_out)//如果超时，则将现有的数据读出，停止读取
                    {
                        if (hs.handle.Available > 0)
                        {
                            Byte[] data = new Byte[lenght];
                            Int32 bytes = hs.stream.Read(data, 0, data.Length);
                            message = Encoding.Default.GetString(data, 0, bytes);
                        }
                        read = false;
                    }
                    else
                    {
                        message = "";
                    }
                }
                else
                {
                    message = "";
                    read = false;
                }
                i++;
                Thread.Sleep(1);//延时1毫秒
         }
            return message;
        }
        public void TCP_Close(HandleAndStream hs ,  TcpListener Listener_Handle)//关闭TCP句柄
        {
            if (hs != null)
            {
                hs.stream.Close();
                hs.handle.Close();
            }
            if (Listener_Handle != null)
            {
                Listener_Handle.Stop();
            }
        }
    }
    class Client: TCPHelper
    {
        public HandleAndStream TCP_Connect(string ip_remote, int port_remote, int port_local)//创建客户端连接
        {
            HandleAndStream hs = new HandleAndStream();

            IPEndPoint Clinet_EndPoint = new IPEndPoint(IPAddress.Any, port_local);//指定本地端口号
            TcpClient Connect_Handle = new TcpClient(Clinet_EndPoint);//重新实例化客户端
            Connect_Handle.Connect(ip_remote, port_remote);//绑定远程端口
            hs.handle = Connect_Handle;
            hs.stream = Connect_Handle.GetStream();
            return hs;
        }

    }
    class Server : TCPHelper
    {
        public void TCP_Listener_Create(string ip, int port)//创建TCP侦听器
        {
            TcpListener Listener_Handle = new TcpListener(IPAddress.Parse(ip), port);//绑定本地的IP（多个网卡中的某一个）和尝试连接的远程端口号
            Listener_Handle.Start();
        }
        public TcpClient TCP_Listener_Wait(TcpListener Listener_Handle, int time_out)
        {
            TcpClient client = null;
            int i = 0;
            bool wait = true;
            while (wait)
            {
                if (Listener_Handle.Pending())//侦听器正在挂起，无客户端接入
                {
                    if (i >= time_out)//等待已超时
                    {
                        wait = false;
                    }
                }
                else//有客户端接入
                {
                    client = Listener_Handle.AcceptTcpClient();
                    wait = false;
                }
                i++;
                Thread.Sleep(1);
            }
            return client;
        }//等待客户端接入
    }
}
