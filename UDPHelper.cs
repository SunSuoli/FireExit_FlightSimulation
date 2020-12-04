using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace FireExit_FlightSimulation
{
    class UDPHelper
    {
        private static UdpClient UDPHandle;
        /*
         * ip:网络地址指定侦听的网络地址。 有多块网卡时，如需侦听特定地址上的网卡，应指定网卡的地址。
         * 如未指定网络地址，可侦听所有的网络地址。
         * 该函数仅在默认的网络地址上广播。 
         * port:要创建UDP套接字的本地端口。
         */
        public void UDP_Open(string ip,int port)
        {
            UDPHandle.Connect(IPAddress.Parse(ip), port);
        }
        public void UDP_Read(out string data,out string ip,out int port)//从任意远程目标监听数据
        {
            IPEndPoint UDPRemote = new IPEndPoint(IPAddress.Any, 0);//要监听的远程目标，这种写法表示监听任意目标
            if(UDPHandle.Available > 0)//利用Available属性可以使阻塞式IO当做不阻塞使用
            {
                data = Encoding.ASCII.GetString(UDPHandle.Receive(ref UDPRemote));//将字节数组转化成字符串
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
            IPEndPoint UDPRemote = new IPEndPoint(IPAddress.Parse(ip), port);//指定要发送的远程目标
            Byte[] data_Send = Encoding.ASCII.GetBytes(data);
            int data_length = data_Send.Length;
            UDPHandle.Send(data_Send, data_length,UDPRemote);
        }
        public void Udp_Close()
        {
            UDPHandle.Close();
        }
    }
}
