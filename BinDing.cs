using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace FireExit_FlightSimulation
{
    class BinDing
    {
        public class Source : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            //创建String类型数据源
            private string data_string;
            public string Data_String
            {
                get { return data_string; }
                set
                {
                    data_string = value;
                    if (PropertyChanged != null)
                    {
                        PropertyChanged(this, new PropertyChangedEventArgs("Data_String"));
                    }
                }
            }
            //创建Int类型数据源
            private string data_int;
            public string Data_Int
            {
                get { return data_int; }
                set
                {
                    data_int = value;
                    if (PropertyChanged != null)
                    {
                        PropertyChanged(this, new PropertyChangedEventArgs("Data_Int"));
                    }
                }
            }
        }

        public static void Bind(Source Data_Source, DependencyObject Control_Name, DependencyProperty Property,String Data_Type)
        {
            Binding Bind = new Binding();//实例化绑定
            Bind.Source = Data_Source;//设置绑定数据源
            Bind.Path= new PropertyPath("Data_"+Data_Type);//设置绑定路径
            Bind.Mode = BindingMode.TwoWay;//双向绑定
            Bind.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;//数据随控件改变
            BindingOperations.SetBinding(Control_Name, Property, Bind);
        }
    }
}
