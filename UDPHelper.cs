using System;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace FireExit_FlightSimulation
{
    class UDPHelper
    {
        private static UdpClient UDPHandle=new UdpClient();
       
        public void UDP_Open(string ip,int port)
        {
            IPEndPoint UDPLocal = new IPEndPoint(IPAddress.Parse(ip), port);//绑定本地的IP（多个网卡中的某一个）和设置本地端口号
            UDPHandle = new UdpClient(UDPLocal);
            
        }
        public void UDP_Read(out string data,out string ip,out int port)//从任意远程目标监听数据
        {
            IPEndPoint UDPRemote = new IPEndPoint(IPAddress.Any, 0);//要监听的远程目标，这种写法表示监听任意目标
            if(UDPHandle.Available > 0)//利用Available属性可以使阻塞式IO当做不阻塞使用
            {
                data = Encoding.Default.GetString(UDPHandle.Receive(ref UDPRemote));//将字节数组转化成字符串
            }
            else
            {
                data = "";
            }
            ip = UDPRemote.Address.ToString();
            port = UDPRemote.Port;
        }
        public void UDP_Write(string data,string ip,int port)//发送数据到指定远程目标
        {
            UDPHandle.Connect(IPAddress.Parse(ip), port);//连接远程目标,ip为255.255.255.255时，数据进行广播。
            Byte[] data_Send = Encoding.Default.GetBytes(data);
            int data_length = data_Send.Length;
            UDPHandle.Send(data_Send, data_length);
        }
        public void Udp_Close()
        {
            UDPHandle.Close();
        }
    }
}
