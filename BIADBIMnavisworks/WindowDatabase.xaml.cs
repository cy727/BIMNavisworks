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
using System.Xml;
using System.IO;
using System.Data;


namespace BIADBIMnavisworks
{
    /// <summary>
    /// WindowDatabase.xaml 的交互逻辑
    /// </summary>
    public partial class WindowDatabase : Window
    {
        public string strConn = "";
        private string dFileName = "";
        public int intMode = 0;  //0=基本库，1=触点库，2=分析库
        public Boolean isComponent = true; //是否为构件数据库设置

        private System.Data.DataSet dSet = new DataSet();
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();

        public WindowDatabase()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dFileName = Directory.GetCurrentDirectory() + "\\dbcon.xml";



            sqlComm.Connection = sqlConn;

            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
            }
            else  //建立文件
            {
                dSet.Tables.Add("基本数据库配置");

                dSet.Tables["基本数据库配置"].Columns.Add("服务器地址", System.Type.GetType("System.String"));
                dSet.Tables["基本数据库配置"].Columns.Add("用户名", System.Type.GetType("System.String"));
                dSet.Tables["基本数据库配置"].Columns.Add("密码", System.Type.GetType("System.String"));
                dSet.Tables["基本数据库配置"].Columns.Add("数据库", System.Type.GetType("System.String"));

                string[] strDRow = { "", "", "", "" };
                dSet.Tables["基本数据库配置"].Rows.Add(strDRow);


                dSet.Tables.Add("监控数据库配置");

                dSet.Tables["监控数据库配置"].Columns.Add("服务器地址", System.Type.GetType("System.String"));
                dSet.Tables["监控数据库配置"].Columns.Add("用户名", System.Type.GetType("System.String"));
                dSet.Tables["监控数据库配置"].Columns.Add("密码", System.Type.GetType("System.String"));
                dSet.Tables["监控数据库配置"].Columns.Add("数据库", System.Type.GetType("System.String"));

                string[] strDRow1 = { "", "", "", "" };
                dSet.Tables["监控数据库配置"].Rows.Add(strDRow1);

                dSet.Tables.Add("分析数据库配置");

                dSet.Tables["分析数据库配置"].Columns.Add("服务器地址", System.Type.GetType("System.String"));
                dSet.Tables["分析数据库配置"].Columns.Add("用户名", System.Type.GetType("System.String"));
                dSet.Tables["分析数据库配置"].Columns.Add("密码", System.Type.GetType("System.String"));
                dSet.Tables["分析数据库配置"].Columns.Add("数据库", System.Type.GetType("System.String"));

                string[] strDRow2 = { "", "", "", "" };
                dSet.Tables["分析数据库配置"].Rows.Add(strDRow1);
            }

            switch (intMode)
            {
                case 0:
                    this.Title = "基本数据库配置";
                    TextBoxIP.Text = dSet.Tables["基本数据库配置"].Rows[0][0].ToString();
                    TextBoxUser.Text = dSet.Tables["基本数据库配置"].Rows[0][1].ToString();
                    TextBoxPassword.Password = dSet.Tables["基本数据库配置"].Rows[0][2].ToString();
                    TextBoxDB.Text = dSet.Tables["基本数据库配置"].Rows[0][3].ToString();
                    break;
                case 1:
                    this.Title = "监控数据库配置";
                    TextBoxIP.Text = dSet.Tables["监控数据库配置"].Rows[0][0].ToString();
                    TextBoxUser.Text = dSet.Tables["监控数据库配置"].Rows[0][1].ToString();
                    TextBoxPassword.Password = dSet.Tables["监控数据库配置"].Rows[0][2].ToString();
                    TextBoxDB.Text = dSet.Tables["监控数据库配置"].Rows[0][3].ToString();
                    break;
                case 2:
                    this.Title = "分析数据库配置";
                    TextBoxIP.Text = dSet.Tables["分析数据库配置"].Rows[0][0].ToString();
                    TextBoxUser.Text = dSet.Tables["分析数据库配置"].Rows[0][1].ToString();
                    TextBoxPassword.Password = dSet.Tables["分析数据库配置"].Rows[0][2].ToString();
                    TextBoxDB.Text = dSet.Tables["分析数据库配置"].Rows[0][3].ToString();
                    break;
            }



        }

        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void buttonTest_Click(object sender, RoutedEventArgs e)
        {
            strConn = "workstation id=CY;packet size=4096;user id=" + TextBoxUser.Text.Trim() + ";password=" + TextBoxPassword.Password.Trim() + ";data source=\"" + TextBoxIP.Text.Trim() + "\";;initial catalog=" + TextBoxDB.Text.Trim();

            sqlConn.ConnectionString = strConn;
            try
            {
                sqlConn.Open();
            }
            catch (System.Data.SqlClient.SqlException err)
            {
                MessageBox.Show("数据库连接错误，请与管理员联系");
                strConn = "";
                return;

            }

            MessageBox.Show("数据库连接正常");
            sqlConn.Close();


        }

        private void buttonYes_Click(object sender, RoutedEventArgs e)
        {
            switch (intMode)
            {
                case 0:
                    dSet.Tables["基本数据库配置"].Rows[0][0] = TextBoxIP.Text;
                    dSet.Tables["基本数据库配置"].Rows[0][1] = TextBoxUser.Text;
                    dSet.Tables["基本数据库配置"].Rows[0][2] = TextBoxPassword.Password;
                    dSet.Tables["基本数据库配置"].Rows[0][3] = TextBoxDB.Text;
                    break;
                case 1:
                    dSet.Tables["监控数据库配置"].Rows[0][0] = TextBoxIP.Text;
                    dSet.Tables["监控数据库配置"].Rows[0][1] = TextBoxUser.Text;
                    dSet.Tables["监控数据库配置"].Rows[0][2] = TextBoxPassword.Password;
                    dSet.Tables["监控数据库配置"].Rows[0][3] = TextBoxDB.Text;
                    break;
                case 2:
                    dSet.Tables["分析数据库配置"].Rows[0][0] = TextBoxIP.Text;
                    dSet.Tables["分析数据库配置"].Rows[0][1] = TextBoxUser.Text;
                    dSet.Tables["分析数据库配置"].Rows[0][2] = TextBoxPassword.Password;
                    dSet.Tables["分析数据库配置"].Rows[0][3] = TextBoxDB.Text;
                    break;
            }


            dSet.WriteXml(dFileName);
            MessageBox.Show("数据库配置保存完毕");
        }
    }
}
