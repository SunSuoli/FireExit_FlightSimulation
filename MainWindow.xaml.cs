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
        public MainWindow()
        {
            InitializeComponent();

            Bind(Data_S_source, Data_S, TextBox.TextProperty, "String");
            Bind(Data_R_source, Data_R, TextBlock.TextProperty, "String");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Data_S_source.Data_String += "12";
            Data_R_source.Data_String += "34";
        }
    }
}
