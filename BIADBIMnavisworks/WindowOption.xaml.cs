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
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace BIADBIMnavisworks
{
    /// <summary>
    /// WindowOption.xaml 的交互逻辑
    /// </summary>
    public partial class WindowOption : Window
    {
        private string dFileName = "";

        private System.Data.DataSet dSet = new DataSet();
        public WindowOption()
        {
            InitializeComponent();
            dFileName = Directory.GetCurrentDirectory() + "\\appcon.xml";
        }

        private void ButtonDirection_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialogD = new FolderBrowserDialog();
            folderBrowserDialogD.ShowNewFolderButton = false;
            folderBrowserDialogD.SelectedPath = TextBoxDirection.Text;


            if (folderBrowserDialogD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TextBoxDirection.Text = folderBrowserDialogD.SelectedPath;
            }
        }

        private void windowOption_Loaded(object sender, RoutedEventArgs e)
        {
            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
            }
            else  //建立文件
            {
                dSet.Tables.Add("映射目录");
                dSet.Tables["映射目录"].Columns.Add("目录地址", System.Type.GetType("System.String"));
                string[] strDRow0 = {""};
                dSet.Tables["映射目录"].Rows.Add(strDRow0);

                dSet.Tables.Add("图纸目录");
                dSet.Tables["图纸目录"].Columns.Add("目录地址", System.Type.GetType("System.String"));
                string[] strDRow01 = { "" };
                dSet.Tables["图纸目录"].Rows.Add(strDRow01);

                dSet.Tables.Add("动画目录");
                dSet.Tables["动画目录"].Columns.Add("目录地址", System.Type.GetType("System.String"));
                string[] strDRow02 = { "" };
                dSet.Tables["动画目录"].Rows.Add(strDRow02);

                dSet.Tables.Add("基本数据库信息");

                dSet.Tables["基本数据库信息"].Columns.Add("服务器地址", System.Type.GetType("System.String"));
                dSet.Tables["基本数据库信息"].Columns.Add("用户名", System.Type.GetType("System.String"));
                dSet.Tables["基本数据库信息"].Columns.Add("密码", System.Type.GetType("System.String"));
                dSet.Tables["基本数据库信息"].Columns.Add("数据库", System.Type.GetType("System.String"));

                string[] strDRow1 = { "", "", "", "" };
                dSet.Tables["基本数据库信息"].Rows.Add(strDRow1);

            }
            TextBoxDirection.Text = dSet.Tables["映射目录"].Rows[0][0].ToString();
            TextBoxDrawingsDirection.Text = dSet.Tables["图纸目录"].Rows[0][0].ToString();
            TextBoxMoviessDirection.Text = dSet.Tables["动画目录"].Rows[0][0].ToString();

        }

        private void buttonYes_Click(object sender, RoutedEventArgs e)
        {
            dSet.Tables["映射目录"].Rows[0][0] = TextBoxDirection.Text.Trim();
            dSet.Tables["图纸目录"].Rows[0][0] = TextBoxDrawingsDirection.Text.Trim();
            dSet.Tables["动画目录"].Rows[0][0] = TextBoxMoviessDirection.Text.Trim();
            dSet.WriteXml(dFileName);
            this.Close();
        }

        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
