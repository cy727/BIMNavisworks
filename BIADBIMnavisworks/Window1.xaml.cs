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
using System.Windows.Shapes;

using Autodesk.Navisworks.Api.Controls;

namespace BIADBIMnavisworks
{
    /// <summary>
    /// Window1.xaml 的交互逻辑
    /// </summary>
    public partial class Window1 : Window
    {
        private Autodesk.Navisworks.Api.Controls.DocumentControl documentControlM;
        public Window1()
        {
            ApplicationControl.Initialize();
            InitializeComponent();

            this.documentControlM = new Autodesk.Navisworks.Api.Controls.DocumentControl();

            viewControl.DocumentControl = this.documentControlM;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            documentControlM.Document.TryOpenFile("E:\\chris\\idwork\\BIM\\BIM项目\\建院园区\\B座\\BIAD_B\\NWC模型（定期更新）\\B_建威大厦2#楼土建_F07.nwf");
            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.FullNavigationWheel;
        }
    }
}
