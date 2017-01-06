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
using System.Drawing;
using System.Data;
using System.IO;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Timers;
using System.Windows.Threading;
//using System.Windows.Forms;

using System.Diagnostics;

using Visifire.Charts;

using System.Xml;
using System.Xml.Linq;

using Fluent;
using Xceed.Wpf.AvalonDock;
using BIADBIMnavisworks.ViewModel;

using Autodesk.Navisworks.Api.Controls;
using Autodesk.Navisworks.Api.ApplicationParts;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.ComApi;
//using Autodesk.Navisworks.Api.Interop.ComApi;



namespace BIADBIMnavisworks
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        private const double PI = 3.1415926;

        private Autodesk.Navisworks.Api.Controls.DocumentControl documentControlM;
        private Autodesk.Navisworks.Api.Controls.DocumentControl documentControlM2D;

        public static DockingManager DockingManager;

        //基本库
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();
        private System.Data.DataSet dSetOption = new DataSet();
        private System.Data.DataSet dSetDB = new DataSet();
        public string strConn = "";

        //监控库Monitor
        private System.Data.SqlClient.SqlConnection sqlConnM = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlCommM = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldrM;
        private System.Data.SqlClient.SqlDataAdapter sqlDAM = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSetM = new DataSet();
        public string strConnM = "";

        //分析库Analy
        private System.Data.SqlClient.SqlConnection sqlConnA = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlCommA = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldrA;
        private System.Data.SqlClient.SqlDataAdapter sqlDAA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSetA = new DataSet();
        public string strConnA = "";
        
        public string sSelectStyle = "";


        public string sModelName = "";
        public string sViewName = "";//主视图
        public string sViewNameAnimation = "";//主动态视图
        public string sViewNameTop = "";//顶视图
        public string sModelSite = "";//场地模型
        public string sModelSiteMEP = "";//小市政模型

        private string sDir = "";
        private string sFile = "";
        private string dFileName = "";

        private string sDirDrawings = "";
        private string sDirMovies = "";

        private UIStyle uiStyle = new UIStyle();

        private System.Data.DataTable dtDGMproperty = new DataTable();
        private System.Data.DataTable dtDGMmonitor = new DataTable();

        private System.Data.DataTable dtDWarnPool = new DataTable();

        private Timer timerMonitor = null;
        private Timer timerRoomMonitor = null;
        private Timer timerAnimation = null;

        //房间监控曲线
        private string sMonitorRoomCode = ""; //监控房间编号
        private string sMonitorRoomName = ""; //监控房间名称
        private List<ClassListEA> listChartWD = new List<ClassListEA>(); //温度监控
        private List<ClassListEA> listChartSD = new List<ClassListEA>(); //湿度监控
        private List<ClassListEA> listChartCO2 = new List<ClassListEA>(); //Co2监控
        private List<ClassListEA> listChartPM25 = new List<ClassListEA>(); //PM2.5监控
        private const int iMONITERNUMBER = 10; //房间监控数值数量
        private int ii = 1;

        private bool bMonitorRoomFirst = true; //重新更换了房间监控

        private delegate void TimerDispatcherDelegate();
        static private string sREPLACE = "^P1^";

        //缺省动画SavedViewpointAnimationCut
        private SavedViewpointAnimation svpacAnimationMain = null;
        //private static int iAnimationMain = 0;

        //U3d动画
        WindowsFormsControlLibraryU3D.UserControlU3dWebPlayer u3dPlayer;

        //搜索treeview用List
        List<TreeNodetagSelectTree> nodeList = new List<TreeNodetagSelectTree>();
        List<TreeNodetagSelectTree> nodeListSearch = new List<TreeNodetagSelectTree>();
        private int iSearchLocation = 0;//搜索当前位置

        private ClassAutodeskViewandDataService AutodeskVDS = new ClassAutodeskViewandDataService();

        //能源分析
        private ClassEnergyAnalysis classEA;

        //全部的范围，供ZOMMALL使用
        private BoundingBox3D boundingBox3DALL ;
        
        public MainWindow()
        {
            ApplicationControl.Initialize();
            Autodesk.Navisworks.Api.Controls.ApplicationControl.Initialize();
            
            InitializeComponent();

            string my_path = System.IO.Path.GetFullPath((new System.Uri(Assembly.GetExecutingAssembly().CodeBase)).LocalPath);
            Autodesk.Navisworks.Api.Application.Plugins.AddPluginAssembly(my_path);

            DockingManager = dockingManager;
            this.DataContext = new MainWindowViewModel();


            this.documentControlM = new Autodesk.Navisworks.Api.Controls.DocumentControl();
            viewControl.DocumentControl = this.documentControlM;
            documentControlM.SetAsMainDocument();

            this.documentControlM2D = new Autodesk.Navisworks.Api.Controls.DocumentControl();
            viewControl2D.DocumentControl = this.documentControlM2D;

            uiStyle.sSelectStyle = "自由";
            TextBlockSelectStyle.DataContext = uiStyle;

            
            this.checkboxSelectTree.DataContext = uiStyle;

            dtDGMproperty.Columns.Add("特性", System.Type.GetType("System.String"));//1
            dtDGMproperty.Columns.Add("值", System.Type.GetType("System.String"));//1
            dtDGMproperty.Columns.Add("控制", System.Type.GetType("System.String"));

            dtDGMmonitor.Columns.Add("特性", System.Type.GetType("System.String"));//1
            dtDGMmonitor.Columns.Add("值", System.Type.GetType("System.String"));//1
            dtDGMmonitor.Columns.Add("控制", System.Type.GetType("System.String"));


            dtDWarnPool.Columns.Add("楼层", System.Type.GetType("System.String"));//1
            dtDWarnPool.Columns.Add("房间", System.Type.GetType("System.String"));//1
            dtDWarnPool.Columns.Add("取值", System.Type.GetType("System.String"));

            timerMonitor = new Timer(1000);
            timerMonitor.Elapsed += new ElapsedEventHandler(OnTimedEventMonitor);
            timerMonitor.Interval = 1000;

            timerRoomMonitor = new Timer(1000);
            timerRoomMonitor.Elapsed += new ElapsedEventHandler(OnTimedEventRoomMonitor);
            timerRoomMonitor.Interval = 1000;

            timerAnimation = new Timer(500);
            timerAnimation.Elapsed += new ElapsedEventHandler(OnTimedEventAnimation);
            timerAnimation.Interval = 500;
            //aTimer.Enabled = true;

        }

        private void MainWindowNavisWorks_Loaded(object sender, RoutedEventArgs e)
        {
            
            dFileName = Directory.GetCurrentDirectory() + "\\appcon.xml";
            if (File.Exists(dFileName)) //存在文件
            {
                dSetOption.ReadXml(dFileName);
                sDir = dSetOption.Tables["映射目录"].Rows[0][0].ToString();
                sDirDrawings = dSetOption.Tables["图纸目录"].Rows[0][0].ToString();
                sDirMovies = dSetOption.Tables["动画目录"].Rows[0][0].ToString();
            }
            //var serializer = new Xceed.Wpf.AvalonDock.Layout.Serialization.XmlLayoutSerializer(this.dockingManager);
            //if (File.Exists(@".\AvalonDock.config"))
            //    serializer.Deserialize(@".\AvalonDock.config");

            dFileName = Directory.GetCurrentDirectory() + "\\dbcon.xml";
            if (!File.Exists(dFileName))
                return;
            
            dSetDB.ReadXml(dFileName);
            strConn = "workstation id=CY;packet size=4096;user id=" + dSetDB.Tables["基本数据库配置"].Rows[0][1].ToString() + ";password=" + dSetDB.Tables["基本数据库配置"].Rows[0][2].ToString() + ";data source=\"" + dSetDB.Tables["基本数据库配置"].Rows[0][0].ToString() + "\";;initial catalog=" + dSetDB.Tables["基本数据库配置"].Rows[0][3].ToString() + "";
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            strConnM = "workstation id=CY;packet size=4096;user id=" + dSetDB.Tables["监控数据库配置"].Rows[0][1].ToString() + ";password=" + dSetDB.Tables["监控数据库配置"].Rows[0][2].ToString() + ";data source=\"" + dSetDB.Tables["监控数据库配置"].Rows[0][0].ToString() + "\";;initial catalog=" + dSetDB.Tables["监控数据库配置"].Rows[0][3].ToString() + "";
            sqlConnM.ConnectionString = strConnM;
            sqlCommM.Connection = sqlConnM;
            sqlDAM.SelectCommand = sqlCommM;

            strConnA = "workstation id=CY;packet size=4096;user id=" + dSetDB.Tables["分析数据库配置"].Rows[0][1].ToString() + ";password=" + dSetDB.Tables["分析数据库配置"].Rows[0][2].ToString() + ";data source=\"" + dSetDB.Tables["分析数据库配置"].Rows[0][0].ToString() + "\";;initial catalog=" + dSetDB.Tables["分析数据库配置"].Rows[0][3].ToString() + "";
            sqlConnA.ConnectionString = strConnA;
            sqlCommA.Connection = sqlConnA;
            sqlDAA.SelectCommand = sqlCommA;

            //初始化数据集
            InitDateSet();
             

            //初始化选择树类型
            comoBoxSelectStyle.ItemsSource = null;
            //treeviewSelectTree.Items.Clear();
            List<tagSelectTreeStyle> itemListSelectTreeStyle = new List<tagSelectTreeStyle>();            tagSelectTreeStyle node1 = new tagSelectTreeStyle()
            {
                iStyle = 0,
                sStyle = "以空间为基准"
            };
            itemListSelectTreeStyle.Add(node1);

            tagSelectTreeStyle node2 = new tagSelectTreeStyle()
            {
                iStyle = 1,
                sStyle = "以系统为基准"
            };
            itemListSelectTreeStyle.Add(node2);
            comoBoxSelectStyle.ItemsSource = itemListSelectTreeStyle;
            comoBoxSelectStyle.DisplayMemberPath = "sStyle";
            
            
            comoBoxSelectStyle.SelectedValuePath = "iStyle";
            comoBoxSelectStyle.SelectedIndex = 0;

            //初始化选择树
            InitSelectTree(int.Parse(comoBoxSelectStyle.SelectedValue.ToString()));
            //InitSelectTree(0);

            int i;
            for (i = 0; i < dSet.Tables["项目参数表"].Rows.Count; i++)
            {
                sModelName = dSet.Tables["项目参数表"].Rows[i][6].ToString();
                sViewName = dSet.Tables["项目参数表"].Rows[i][7].ToString();
                sViewNameAnimation = dSet.Tables["项目参数表"].Rows[i][8].ToString();
                sViewNameTop = dSet.Tables["项目参数表"].Rows[i][9].ToString();
                sModelSite = dSet.Tables["项目参数表"].Rows[i][10].ToString();
                sModelSiteMEP = dSet.Tables["项目参数表"].Rows[i][11].ToString();

                sFile = sDir + "\\" + sModelName;
                break;
            }
            mainModelView.Title = dSet.Tables["项目参数表"].Rows[0][1].ToString();

            documentControlM.Document.TryOpenFile(sFile);
            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.FullNavigationWheel;
            uiStyle.sSelectStyle = "自由";
            

            //documentControlM2D.Document.TryOpenFile(sDir + "\\"+"t1.nwd");

            SavedItem siView;
            if(sViewNameAnimation!="")
                siView = getViewPoint(sViewNameAnimation);
            else
                siView = getViewPoint(sViewName);

            if (siView is SavedViewpointAnimation)
            {
                svpacAnimationMain = siView as SavedViewpointAnimation;
            }
            if (siView == null)
                return;
            //documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;
            documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = getViewPoint(sViewName);

            Autodesk.Navisworks.Api.Application.ActiveDocument.CurrentSelection.Changed +=  new EventHandler<EventArgs>(CurrentSelection_Changed);
            ApplicationControl.SelectionBehavior = SelectionBehavior.FirstObject;

            //加载U3d
            u3dPlayer = new WindowsFormsControlLibraryU3D.UserControlU3dWebPlayer();
            u3dPlayer.Dock = System.Windows.Forms.DockStyle.Fill;
            wfiU3dPlayer.Child = u3dPlayer;

            //能源分析
            classEA = new ClassEnergyAnalysis();
            classEA.strConnM = strConnM;

            classEA.Init();
            classEA.dSet = dSet;

            BoundingBox3D b3dTemp;
            double minX,minY,minZ;
            double maxX,maxY,maxZ;

            minX=documentControlM.Document.Models[0].RootItem.BoundingBox().Min.X;
            minY=documentControlM.Document.Models[0].RootItem.BoundingBox().Min.Y;
            minZ=documentControlM.Document.Models[0].RootItem.BoundingBox().Min.Z;

            maxX=documentControlM.Document.Models[0].RootItem.BoundingBox().Max.X;
            maxY=documentControlM.Document.Models[0].RootItem.BoundingBox().Max.Y;
            maxZ=documentControlM.Document.Models[0].RootItem.BoundingBox().Max.Z;

            //boundingBox3DALL初始化
            //boundingBox3DALL = documentControlM.Document.Models[0].RootItem.BoundingBox();
            for (i = 1; i < documentControlM.Document.Models.Count; i++)
            {
                b3dTemp = documentControlM.Document.Models[i].RootItem.BoundingBox();
                if (b3dTemp != null)
                {
                    if (b3dTemp.Min.X < minX)
                        minX = b3dTemp.Min.X;
                    if (b3dTemp.Min.Y < minY)
                        minY = b3dTemp.Min.Y;
                    if (b3dTemp.Min.Z < minZ)
                        minZ = b3dTemp.Min.Z;

                    if (b3dTemp.Max.X > maxX)
                        maxX = b3dTemp.Max.X;
                    if (b3dTemp.Max.Y > maxY)
                        maxY = b3dTemp.Max.Y;
                    if (b3dTemp.Max.Z > maxZ)
                        maxZ = b3dTemp.Max.Z;
                }
            }
            boundingBox3DALL = new BoundingBox3D(new Point3D(minX, minY, minZ), new Point3D(maxX, maxY, maxZ));

            
        }

        private void MainWindowNavisWorks_Unloaded(object sender, RoutedEventArgs e)
        {
            
            Autodesk.Navisworks.Api.Controls.ApplicationControl.Terminate();
            var serializer = new Xceed.Wpf.AvalonDock.Layout.Serialization.XmlLayoutSerializer(dockingManager);
            serializer.Serialize(@".\AvalonDock.config");
        }

        //初始化数据集
        private void InitDateSet()
        {
            sqlConn.Open();

            sqlComm.CommandText = "SELECT  ID, 项目名称, 项目说明, 项目编号, 客户名称, 项目状态, 主模型, 视图, 动态视图, 顶视图, 场地模型, 市政管线模型 FROM 项目参数表";
            if (dSet.Tables.Contains("项目参数表")) dSet.Tables.Remove("项目参数表");
            sqlDA.Fill(dSet, "项目参数表");

            sqlComm.CommandText = "SELECT 建筑ID, 建筑名称, 说明, 模型ID, 模型, 顶视图, 辅助模型, 基础模型, 二维图纸 FROM 建筑表";
            if (dSet.Tables.Contains("建筑表")) dSet.Tables.Remove("建筑表");
            sqlDA.Fill(dSet, "建筑表");

            sqlComm.CommandText = "SELECT 标高ID, 标高名称, 建筑ID, 说明, 模型ID, 模型, 视图, 土建模型组成, 机电模型组成, 标高排序, 图纸 , VR, 二维图纸 FROM 标高表 ORDER BY 标高排序";
            if (dSet.Tables.Contains("标高表")) dSet.Tables.Remove("标高表");
            sqlDA.Fill(dSet, "标高表");

            sqlComm.CommandText = "SELECT 房间ID, 房间名称, 房间编号, 标高ID, 说明, 模型ID, 模型, 视图, 二维图纸 FROM 房间表 ORDER BY 房间名称, 房间编号, 标高ID";
            if (dSet.Tables.Contains("房间表")) dSet.Tables.Remove("房间表");
            sqlDA.Fill(dSet, "房间表");

            //sqlComm.CommandText = "SELECT 系统表.系统ID, 系统表.设备系统, 系统表.数据源ID, 数据源表.数据源名称, 数据源表.数据源类型, 数据源表.数据源地址,数据源表.数据源 FROM 数据源表 INNER JOIN 系统表 ON 数据源表.ID = 系统表.数据源ID ORDER BY 系统表.排序";
            sqlComm.CommandText = "SELECT 系统ID, 设备系统, 数据源ID FROM 系统表 ORDER BY 排序";
            if (dSet.Tables.Contains("系统表")) dSet.Tables.Remove("系统表");
            sqlDA.Fill(dSet, "系统表");

            //sqlComm.CommandText = "SELECT 子系统ID, 子设备系统, 主系统ID, 主设备系统, 排序 FROM 子系统表 ORDER BY 主系统ID, 排序";
            sqlComm.CommandText = "SELECT 子系统表.子系统ID, 子系统表.子设备系统, 子系统表.主系统ID, 子系统表.主设备系统, 子系统表.排序,子系统表.数据源ID, 数据源表.数据源名称, 数据源表.数据源类型, 数据源表.数据源地址, 数据源表.数据源,数据源表.站点参数 FROM 子系统表 INNER JOIN 数据源表 ON 子系统表.数据源ID = 数据源表.ID ORDER BY 子系统表.主系统ID, 子系统表.排序";
            if (dSet.Tables.Contains("子系统表")) dSet.Tables.Remove("子系统表");
            sqlDA.Fill(dSet, "子系统表");

            sqlComm.CommandText = "SELECT 设备表.设备ID, 设备表.设备名称, 设备表.设备编号, 设备表.系统ID, 设备表.设备系统, 设备表.房间ID, 设备表.说明,设备表.模型ID, 设备表.模型, 设备表.视图, 房间表.房间名称, 房间表.房间编号, 房间表.模型 AS 房间说明, 房间表.视图 AS 房间视图, 标高表.标高名称, 标高表.模型 AS 标高模型, 标高表.视图 AS 标高视图, 建筑表.建筑名称, 建筑表.模型 AS 建筑模型, 建筑表.视图 AS 建筑视图, 建筑表.建筑ID, 标高表.标高ID, 子系统表.主系统ID, 子系统表.主设备系统, 设备表.OBIX站点, 设备表.二维图纸 FROM 设备表 INNER JOIN 房间表 ON 设备表.房间ID = 房间表.房间ID INNER JOIN 标高表 ON 房间表.标高ID = 标高表.标高ID INNER JOIN 建筑表 ON 标高表.建筑ID = 建筑表.建筑ID INNER JOIN 子系统表 ON 设备表.系统ID = 子系统表.子系统ID ORDER BY 子系统表.主系统ID, 设备表.系统ID, 设备表.设备编号";
            if (dSet.Tables.Contains("设备表")) dSet.Tables.Remove("设备表");
            sqlDA.Fill(dSet, "设备表");

            sqlComm.CommandText = "SELECT ID, 系统ID, 特性名称, 特性单位, 特性调整, 低阀值, 高阀值, 取值类型, OBIX取值, OBIX取值名, OBIX特征, OBIX首字清除,OBIX末字清除, OBIX显示关键字 FROM 系统监控特性表";
            if (dSet.Tables.Contains("系统监控特性表")) dSet.Tables.Remove("系统监控特性表");
            sqlDA.Fill(dSet, "系统监控特性表");

            sqlComm.CommandText = "SELECT 标高设备监控表.ID, 标高设备监控表.标高ID, 标高设备监控表.标高名称, 标高设备监控表.设备ID, 标高设备监控表.设备编号, 标高表.标高排序, 标高表.模型, 标高表.视图 FROM 标高设备监控表 INNER JOIN 标高表 ON 标高设备监控表.标高ID = 标高表.标高ID";
            if (dSet.Tables.Contains("标高设备监控表")) dSet.Tables.Remove("标高设备监控表");
            sqlDA.Fill(dSet, "标高设备监控表");

            sqlComm.CommandText = "SELECT   房间设备监控表.ID, 房间设备监控表.房间ID, 房间设备监控表.房间编号, 房间设备监控表.设备ID, 房间设备监控表.设备编号, 标高表.标高ID, 标高表.标高名称, 标高表.标高排序, 房间表.模型, 房间表.视图 FROM 房间设备监控表 INNER JOIN 房间表 ON 房间设备监控表.房间ID = 房间表.房间ID INNER JOIN 标高表 ON 房间表.标高ID = 标高表.标高ID INNER JOIN 设备表 ON 房间设备监控表.设备编号 = 设备表.设备编号 INNER JOIN 子系统表 ON 设备表.系统ID = 子系统表.子系统ID ORDER BY 子系统表.子系统ID";
            if (dSet.Tables.Contains("房间设备监控表")) dSet.Tables.Remove("房间设备监控表");
            sqlDA.Fill(dSet, "房间设备监控表");

            sqlComm.CommandText = "SELECT 设备控制监控表.ID, 设备控制监控表.设备ID, 设备控制监控表.设备编号, 设备控制监控表.控制设备ID,设备控制监控表.控制设备编号, 设备表.设备名称, 房间表.房间ID, 房间表.房间名称, 房间表.房间编号, 标高表.标高ID, 标高表.标高名称, 标高表.标高排序, 设备表.模型, 设备表.视图 FROM 设备控制监控表 INNER JOIN 设备表 ON 设备控制监控表.设备编号 = 设备表.设备编号 INNER JOIN 房间表 ON 设备表.房间ID = 房间表.房间ID INNER JOIN 标高表 ON 房间表.标高ID = 标高表.标高ID";
            if (dSet.Tables.Contains("设备控制监控表")) dSet.Tables.Remove("设备控制监控表");
            sqlDA.Fill(dSet, "设备控制监控表");


            sqlComm.CommandText = "SELECT ID, 设备ID, 设备编号, URN FROM 设备URN表";
            if (dSet.Tables.Contains("设备URN表")) dSet.Tables.Remove("设备URN表");
            sqlDA.Fill(dSet, "设备URN表");

            sqlConn.Close();
        }

        //----------------------------------------
        //初始化选择树，0，空间，1，系统
        //----------------------------------------
        private void InitSelectTree(int iStyle)
        {
            int i, j, k;
            string sModelT = "", sViewT = ""; //缺省模型视图
            string sDrawingT = "";//缺省二维图纸
            bool bTemp = true;

            treeviewSelectTree.ItemsSource = null;
            //treeviewSelectTree.Items.Clear();

            ObservableCollection<TreeNodetagSelectTree> itemList = new ObservableCollection<TreeNodetagSelectTree>();

            //根节点
            TreeNodetagSelectTree nodeR = new TreeNodetagSelectTree()
            {
                sDisplayName = dSet.Tables["项目参数表"].Rows[0][1].ToString(),
                sNote = dSet.Tables["项目参数表"].Rows[0][2].ToString(),
                sIcon = @"\Images\project.png",
                nodetype = nodeType.PROJECT,

                sModel = dSet.Tables["项目参数表"].Rows[0]["主模型"].ToString(),
                sView = dSet.Tables["项目参数表"].Rows[0]["视图"].ToString(),
                Parent=null,

            };
            sModelT = nodeR.sModel; sViewT = nodeR.sView;
            sDrawingT = "";

            for(i = 0; i < dSet.Tables["建筑表"].Rows.Count; i++) //建筑
            {
                //建筑节点
                TreeNodetagSelectTree nodeB = new TreeNodetagSelectTree()
                {
                    sDisplayName = dSet.Tables["建筑表"].Rows[i][1].ToString(),
                    sNote = dSet.Tables["建筑表"].Rows[i][2].ToString(),
                    sIcon = @"\Images\building.png",
                    nodetype = nodeType.BUILDING,
                    iBuildingID=int.Parse(dSet.Tables["建筑表"].Rows[i][0].ToString()),
                    sBuildingName = dSet.Tables["建筑表"].Rows[i][1].ToString(),

                    sModel = dSet.Tables["建筑表"].Rows[i][4].ToString(),
                    sView = dSet.Tables["建筑表"].Rows[i][5].ToString(),
                    sDrawing = dSet.Tables["建筑表"].Rows[i][8].ToString(),



                };
                if (nodeB.sModel == "" || nodeB.sModel == null)
                {
                    nodeB.sModel = sModelT;
                }
                else
                {
                    sModelT = nodeB.sModel;
                }

                if (nodeB.sView == "" || nodeB.sView == null)
                {
                    nodeB.sView = sViewT;
                }
                else
                {
                    sViewT = nodeB.sView;
                }

                if (nodeB.sDrawing == "" || nodeB.sDrawing == null)
                {
                    nodeB.sDrawing = sDrawingT;
                }
                else
                {
                    sDrawingT = nodeB.sDrawing;
                }

                
                switch(iStyle)
                {
                    #region 按空间定位
                    case 0: //按空间定位

                        //
                        var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                                     where (dtLevel.Field<int>("建筑ID") == int.Parse(dSet.Tables["建筑表"].Rows[i][0].ToString()))//条件
                                     select dtLevel;
                        foreach (var itemLevel in qLevel)//显示查询结果
                        {
                            //楼层节点
                            TreeNodetagSelectTree nodeL = new TreeNodetagSelectTree()
                            {
                                sDisplayName = itemLevel.Field<string>("标高名称"),
                                sNote = itemLevel.Field<string>("标高名称"),
                                sIcon = @"\Images\level.png",
                                nodetype = nodeType.LEVEL,
                                iBuildingID=nodeB.iBuildingID,
                                sBuildingName=nodeB.sSystemName,
                               
                                iLevelID = itemLevel.Field<int>("标高ID"),
                                sLevelName = itemLevel.Field<string>("标高名称"),

                                sLevelModelAS = itemLevel.Field<string>("土建模型组成"),
                                sLevelModelMEP = itemLevel.Field<string>("机电模型组成"),
                                fLevelOrder=itemLevel.Field<double>("标高排序"),
                                //fLevelOrder = 0,

                                sModel = itemLevel.Field<string>("模型"),
                                sView = itemLevel.Field<string>("视图"),
                                sDrawing = itemLevel.Field<string>("二维图纸"),
                            };
                            if (nodeL.sModel == "" || nodeL.sModel==null)
                            {
                                nodeL.sModel = sModelT;
                            }
                            else
                            {
                                sModelT = nodeL.sModel;
                            }

                            if (nodeL.sView == "" || nodeL.sView == null)
                            {
                                nodeL.sView = sViewT;
                            }
                            else
                            {
                                sViewT = nodeL.sView;
                            }

                            if (nodeL.sDrawing == "" || nodeL.sDrawing == null)
                            {
                                nodeL.sDrawing = sDrawingT;
                            }
                            else
                            {
                                sDrawingT = nodeL.sDrawing;
                            }

                            //筛选房间
                            var qRoom = from dtRoom in dSet.Tables["房间表"].AsEnumerable()//查询楼层
                                        where (dtRoom.Field<int>("标高ID") == nodeL.iLevelID)//条件
                                        select dtRoom;
                            foreach (var itemRoom in qRoom)//显示查询结果
                            {
                                //房间节点
                                TreeNodetagSelectTree nodeRo = new TreeNodetagSelectTree()
                                {
                                    sDisplayName = itemRoom.Field<string>("房间名称"),
                                    sNote = itemRoom.Field<string>("房间名称"),
                                    sIcon = @"\Images\room.png",
                                    nodetype = nodeType.ROOM,

                                    iBuildingID = nodeB.iBuildingID,
                                    sBuildingName = nodeB.sSystemName,
                                    iLevelID = nodeL.iLevelID,
                                    sLevelName = nodeL.sLevelName,

                                    sLevelModelAS=nodeL.sLevelModelAS,
                                    sLevelModelMEP = nodeL.sLevelModelMEP,
                                    fLevelOrder=nodeL.fLevelOrder,

                                    iRoomID = itemRoom.Field<int>("房间ID"),
                                    sRoomName = itemRoom.Field<string>("房间名称"),
                                    sRoomCode=itemRoom.Field<string>("房间编号"),

                                    sModel = itemRoom.Field<string>("模型"),
                                    sView = itemRoom.Field<string>("视图"),
                                    sDrawing = itemRoom.Field<string>("二维图纸"),
                                };
                                if (nodeRo.sModel == "" || nodeRo.sModel == null)
                                {
                                    nodeRo.sModel = sModelT;
                                }
                                else
                                {
                                    sModelT = nodeRo.sModel;
                                }

                                if (nodeRo.sView == "" || nodeRo.sView == null)
                                {
                                    nodeRo.sView = sViewT;
                                }
                                else
                                {
                                    sViewT = nodeRo.sView;
                                }
                                if (nodeRo.sDrawing == "" || nodeRo.sDrawing == null)
                                {
                                    nodeRo.sDrawing = sDrawingT;
                                }
                                else
                                {
                                    sDrawingT = nodeRo.sDrawing;
                                }

                                //筛选系统
                                var qSystem = from dtSystem in dSet.Tables["系统表"].AsEnumerable()//查询楼层
                                              //条件
                                              select dtSystem;
                                foreach (var itemSystem in qSystem)//显示查询结果
                                {
                                    //系统节点
                                    TreeNodetagSelectTree nodeSys = new TreeNodetagSelectTree()
                                    {
                                        sDisplayName = itemSystem.Field<string>("设备系统"),
                                        sNote = itemSystem.Field<string>("设备系统"),
                                        sIcon = @"",
                                        nodetype = nodeType.SYSTEM,

                                        iSystemID = itemSystem.Field<int>("系统ID"),
                                        sSystemName = itemSystem.Field<string>("设备系统"),

                                        sModel = "",
                                        sView = "",
                                        sDrawing="",
                                    };

                                    //筛选子系统
                                    var qSystemSub = from dtSystemSub in dSet.Tables["子系统表"].AsEnumerable()//查询楼层
                                                     where (dtSystemSub.Field<int>("主系统ID") == nodeSys.iSystemID)//条件
                                                     select dtSystemSub;
                                    foreach (var itemSystemSub in qSystemSub)//显示查询结果
                                    {
                                        //子系统节点
                                        TreeNodetagSelectTree nodeSysSub = new TreeNodetagSelectTree()
                                        {
                                            sDisplayName = itemSystemSub.Field<string>("子设备系统"),
                                            sNote = itemSystemSub.Field<string>("子设备系统"),
                                            sIcon = @"",
                                            nodetype = nodeType.SYSTEM,

                                            iSystemID = itemSystem.Field<int>("系统ID"),
                                            sSystemName = itemSystem.Field<string>("设备系统"),
                                            iSystemSubID = itemSystemSub.Field<int>("子系统ID"),
                                            sSystemSubName = itemSystemSub.Field<string>("子设备系统"),

                                            sModel = "",
                                            sView = "",
                                            sDrawing = "",
                                        };




                                        //筛选设备
                                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                                        where (dtFacility.Field<int>("房间ID") == nodeRo.iRoomID) && (dtFacility.Field<int>("系统ID") == nodeSysSub.iSystemSubID)//条件
                                                        select dtFacility;

                                        foreach (var itemFacility in qFacility)
                                        {
                                            //设备节点
                                            TreeNodetagSelectTree nodeFacility = new TreeNodetagSelectTree()
                                            {
                                                sDisplayName = itemFacility.Field<string>("设备名称") + "[" + itemFacility.Field<string>("设备编号") + "]",
                                                sNote = itemFacility.Field<string>("设备名称") + " 编号：" + itemFacility.Field<string>("设备编号") + "",
                                                sIcon = @"\Images\facility.png",
                                                nodetype = nodeType.FACILITY,

                                                iBuildingID = nodeB.iBuildingID,
                                                sBuildingName = nodeB.sBuildingName,
                                                iLevelID = nodeL.iLevelID,
                                                sLevelName = nodeL.sLevelName,
                                                sLevelModelAS = nodeL.sLevelModelAS,
                                                sLevelModelMEP = nodeL.sLevelModelMEP,
                                                fLevelOrder = nodeL.fLevelOrder,

                                                iSystemID = nodeSys.iSystemID,
                                                sSystemName = nodeSys.sSystemName,
                                                iSystemSubID = nodeSysSub.iSystemSubID,
                                                sSystemSubName = nodeSysSub.sSystemSubName,

                                                iRoomID = nodeRo.iRoomID,
                                                sRoomName = nodeRo.sRoomName,

                                                iFacilityID = itemFacility.Field<int>("设备ID"),
                                                sFacilityName = itemFacility.Field<string>("设备名称"),
                                                sFacilityCode = itemFacility.Field<string>("设备编号"),

                                                sModel = itemFacility.Field<string>("模型"),
                                                sView = itemFacility.Field<string>("视图"),
                                                sDrawing = itemFacility.Field<string>("二维图纸"),

                                            };

                                            //控制空间、设备赋值,楼层为主
                                            nodeFacility.sModel_controlled = sModelT;
                                            //筛选控制楼层
                                            var qLevelControlled = from dtLevelControlled in dSet.Tables["标高设备监控表"].AsEnumerable()//查询楼层
                                                                   where (dtLevelControlled.Field<string>("设备编号") == nodeFacility.sFacilityCode)//条件
                                                                   select dtLevelControlled;
                                            if (qLevelControlled.Count() > 0)
                                                nodeFacility.nodetype_controlled = nodeType.LEVEL;

                                            foreach (var itemLevelControlled in qLevelControlled)
                                            {
                                                nodeFacility.iLevelID_controlled = itemLevelControlled.Field<int>("标高ID");
                                                nodeFacility.sLevelName_controlled = itemLevelControlled.Field<string>("标高名称");
                                                nodeFacility.fLevelOrder_controlled = itemLevelControlled.Field<double>("标高排序");

                                                if (itemLevelControlled.Field<string>("模型") != null)
                                                    nodeFacility.sModel_controlled = itemLevelControlled.Field<string>("模型");
                                                nodeFacility.sView_controlled = itemLevelControlled.Field<string>("视图");
                                                break;
                                            }

                                            //筛选控制空间
                                            if (nodeFacility.nodetype_controlled == nodeType.UNKNOW)
                                            {
                                                var qRoomControlled = from dtRoomControlled in dSet.Tables["房间设备监控表"].AsEnumerable()//查询空间
                                                                      where (dtRoomControlled.Field<string>("设备编号") == nodeFacility.sFacilityCode)//条件
                                                                      select dtRoomControlled;
                                                if (qRoomControlled.Count() > 0)
                                                    nodeFacility.nodetype_controlled = nodeType.ROOM;

                                                foreach (var itemRoomControlled in qRoomControlled)
                                                {
                                                    nodeFacility.iRoomID_controlled = itemRoomControlled.Field<int>("房间ID");
                                                    nodeFacility.sRoomCode_controlled = itemRoomControlled.Field<string>("房间编号");
                                                    nodeFacility.iLevelID_controlled = itemRoomControlled.Field<int>("标高ID");
                                                    nodeFacility.sLevelName_controlled = itemRoomControlled.Field<string>("标高名称");
                                                    nodeFacility.fLevelOrder_controlled = itemRoomControlled.Field<double>("标高排序");

                                                    if (itemRoomControlled.Field<string>("模型") != null)
                                                        nodeFacility.sModel_controlled = itemRoomControlled.Field<string>("模型");
                                                    nodeFacility.sView_controlled = itemRoomControlled.Field<string>("视图");
                                                    break;
                                                }
                                            }

                                            //筛选控制设备
                                            var qFacilityControll = from dtFacilityControlled in dSet.Tables["设备控制监控表"].AsEnumerable()//查询控制设备
                                                                      where (dtFacilityControlled.Field<string>("设备编号") == nodeFacility.sFacilityCode)//条件
                                                                      select dtFacilityControlled;
                                            foreach (var itemFacilityControll in qFacilityControll)
                                            {
                                                nodeFacility.fiFacility_control.iFacilityID = itemFacilityControll.Field<int>("控制设备ID");
                                                nodeFacility.fiFacility_control.sFacilityCode = itemFacilityControll.Field<string>("控制设备编号");
                                                nodeFacility.fiFacility_control.sFacilityName = itemFacilityControll.Field<string>("设备名称");
                                                
                                                break;
                                            }


                                            //筛选受控设备
                                            var qFacilityControlled = from dtFacilityControlled in dSet.Tables["设备控制监控表"].AsEnumerable()//查询受控设备
                                                                      where (dtFacilityControlled.Field<string>("控制设备编号") == nodeFacility.sFacilityCode)//条件
                                                                        select dtFacilityControlled;
                                            foreach (var itemFacilityControlled in qFacilityControlled)
                                            {
                                                FacilityIdenti fiControlled = new FacilityIdenti();
                                                fiControlled.iFacilityID = itemFacilityControlled.Field<int>("设备ID");
                                                fiControlled.sFacilityCode = itemFacilityControlled.Field<string>("设备编号");
                                                fiControlled.sFacilityName = itemFacilityControlled.Field<string>("设备名称");

                                                nodeFacility.listFacility_controlled.Add(fiControlled);
                                            }
                                            

                                            if (nodeFacility.sModel == "" || nodeFacility.sModel == null)
                                            {
                                                nodeFacility.sModel = sModelT;
                                            }
                                            else
                                            {
                                                sModelT = nodeFacility.sModel;
                                            }

                                            if (nodeFacility.sView == "" || nodeFacility.sView == null)
                                            {
                                                nodeFacility.sView = sViewT;
                                            }
                                            else
                                            {
                                                sViewT = nodeFacility.sView;
                                            }

                                            if (nodeFacility.sDrawing == "" || nodeFacility.sDrawing == null)
                                            {
                                                nodeFacility.sDrawing = sDrawingT;
                                            }
                                            else
                                            {
                                                //sDrawingT = nodeFacility.sDrawing;
                                            }

                                            nodeFacility.Parent = nodeSysSub;
                                            nodeSysSub.Children.Add(nodeFacility);
                                        }//设备

                                        nodeSysSub.Parent = nodeSys;
                                        if (nodeSysSub.Children.Count > 0)
                                            nodeSys.Children.Add(nodeSysSub);
                                    
                                    }//子系统

                                    nodeSys.Parent = nodeRo;
                                    if(nodeSys.Children.Count>0)
                                        nodeRo.Children.Add(nodeSys);
                                }//系统


                                if (checkboxFacilityOnly.IsChecked.HasValue)
                                    bTemp = (bool)checkboxFacilityOnly.IsChecked;
                                else
                                    bTemp = false;

                                nodeRo.Parent = nodeL;
                                if (!bTemp || nodeRo.Children.Count>0) //只显示设备空间
                                    nodeL.Children.Add(nodeRo);
                            }//房间

                            nodeL.Parent = nodeB;
                            nodeB.Children.Add(nodeL);
                        }//标高

                        break;
                    #endregion

                    #region 按系统定位
                    case 1: //按系统定位

                        //筛选系统
                        var qSystem1 = from dtSystem in dSet.Tables["系统表"].AsEnumerable()//查询楼层
                                      //条件
                                  select dtSystem;

                        foreach (var itemSystem in qSystem1)//显示查询结果
                        {
                            //系统节点
                            TreeNodetagSelectTree nodeSys = new TreeNodetagSelectTree()
                            {
                                sDisplayName = itemSystem.Field<string>("设备系统"),
                                sNote = itemSystem.Field<string>("设备系统"),
                                sIcon = @"",
                                nodetype = nodeType.SYSTEM,

                                iSystemID = itemSystem.Field<int>("系统ID"),
                                sSystemName = itemSystem.Field<string>("设备系统"),

                                sModel = "",
                                sView = "",
                                sDrawing="",
                            };

                            //筛选子系统
                            var qSystemSub = from dtSystemSub in dSet.Tables["子系统表"].AsEnumerable()//查询楼层
                                                where (dtSystemSub.Field<int>("主系统ID") == nodeSys.iSystemID)//条件
                                                select dtSystemSub;
                            foreach (var itemSystemSub in qSystemSub)//显示查询结果
                            {
                                //子系统节点
                                TreeNodetagSelectTree nodeSysSub = new TreeNodetagSelectTree()
                                {
                                    sDisplayName = itemSystemSub.Field<string>("子设备系统"),
                                    sNote = itemSystemSub.Field<string>("子设备系统"),
                                    sIcon = @"",
                                    nodetype = nodeType.SYSTEM,

                                    iSystemID = itemSystem.Field<int>("系统ID"),
                                    sSystemName = itemSystem.Field<string>("设备系统"),
                                    iSystemSubID = itemSystemSub.Field<int>("子系统ID"),
                                    sSystemSubName = itemSystemSub.Field<string>("子设备系统"),

                                    sModel = "",
                                    sView = "",
                                    sDrawing="",
                                };


                                qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                                         where (dtLevel.Field<int>("建筑ID") == int.Parse(dSet.Tables["建筑表"].Rows[i][0].ToString()))//条件
                                         select dtLevel;
                                foreach (var itemLevel in qLevel)//显示查询结果
                                {
                                    //楼层节点
                                    TreeNodetagSelectTree nodeL = new TreeNodetagSelectTree()
                                    {
                                        sDisplayName = itemLevel.Field<string>("标高名称"),
                                        sNote = itemLevel.Field<string>("标高名称"),
                                        sIcon = @"\Images\level.png",
                                        nodetype = nodeType.LEVEL,
                                        iBuildingID = nodeB.iBuildingID,
                                        sBuildingName = nodeB.sSystemName,

                                        iLevelID = itemLevel.Field<int>("标高ID"),
                                        sLevelName = itemLevel.Field<string>("标高名称"),

                                        sLevelModelAS = itemLevel.Field<string>("土建模型组成"),
                                        sLevelModelMEP = itemLevel.Field<string>("机电模型组成"),
                                        fLevelOrder = itemLevel.Field<double>("标高排序"),

                                        sModel = itemLevel.Field<string>("模型"),
                                        sView = itemLevel.Field<string>("视图"),
                                        sDrawing = itemLevel.Field<string>("二维图纸"),
                                    };


                                    if (nodeL.sModel == "" || nodeL.sModel == null)
                                    {
                                        nodeL.sModel = sModelT;
                                    }
                                    else
                                    {
                                        sModelT = nodeL.sModel;
                                    }

                                    if (nodeL.sView == "" || nodeL.sView == null)
                                    {
                                        nodeL.sView = sViewT;
                                    }
                                    else
                                    {
                                        sViewT = nodeL.sView;
                                    }

                                    if (nodeL.sDrawing == "" || nodeL.sDrawing == null)
                                    {
                                        nodeL.sDrawing = sDrawingT;
                                    }
                                    else
                                    {
                                        sDrawingT = nodeL.sDrawing;
                                    }

                                    //筛选房间
                                    var qRoom = from dtRoom in dSet.Tables["房间表"].AsEnumerable()//查询楼层
                                                where (dtRoom.Field<int>("标高ID") == nodeL.iLevelID)//条件
                                                select dtRoom;
                                    foreach (var itemRoom in qRoom)//显示查询结果
                                    {
                                        //房间节点
                                        TreeNodetagSelectTree nodeRo = new TreeNodetagSelectTree()
                                        {
                                            sDisplayName = itemRoom.Field<string>("房间名称"),
                                            sNote = itemRoom.Field<string>("房间名称"),
                                            sIcon = @"\Images\room.png",
                                            nodetype = nodeType.ROOM,

                                            iBuildingID = nodeB.iBuildingID,
                                            sBuildingName = nodeB.sSystemName,
                                            iLevelID = nodeL.iLevelID,
                                            sLevelName = nodeL.sLevelName,
                                            sLevelModelAS = nodeL.sLevelModelAS,
                                            sLevelModelMEP = nodeL.sLevelModelMEP,
                                            fLevelOrder = nodeL.fLevelOrder,

                                            iRoomID = itemRoom.Field<int>("房间ID"),
                                            sRoomName = itemRoom.Field<string>("房间名称"),
                                            sRoomCode = itemRoom.Field<string>("房间编号"),

                                            sModel = itemRoom.Field<string>("模型"),
                                            sView = itemRoom.Field<string>("视图"),
                                            sDrawing = itemRoom.Field<string>("二维图纸"),
                                        };
                                        if (nodeRo.sModel == "" || nodeRo.sModel == null)
                                        {
                                            nodeRo.sModel = sModelT;
                                        }
                                        else
                                        {
                                            sModelT = nodeRo.sModel;
                                        }

                                        if (nodeRo.sView == "" || nodeRo.sView == null)
                                        {
                                            nodeRo.sView = sViewT;
                                        }
                                        else
                                        {
                                            sViewT = nodeRo.sView;
                                        }

                                        if (nodeRo.sDrawing == "" || nodeRo.sDrawing == null)
                                        {
                                            nodeRo.sDrawing = sDrawingT;
                                        }
                                        else
                                        {
                                            sDrawingT = nodeRo.sDrawing;
                                        }
                                        //筛选设备
                                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                                        where (dtFacility.Field<int>("房间ID") == nodeRo.iRoomID) && (dtFacility.Field<int>("系统ID") == nodeSysSub.iSystemSubID)//条件
                                                        select dtFacility;

                                        foreach (var itemFacility in qFacility)
                                        {
                                            //设备节点
                                            TreeNodetagSelectTree nodeFacility = new TreeNodetagSelectTree()
                                            {
                                                sDisplayName = itemFacility.Field<string>("设备名称") + "[" + itemFacility.Field<string>("设备编号") + "]",
                                                sNote = itemFacility.Field<string>("设备名称") + "(" + itemFacility.Field<string>("设备编号") + ")",
                                                sIcon = @"\Images\facility.png",
                                                nodetype = nodeType.FACILITY,

                                                iBuildingID = nodeB.iBuildingID,
                                                sBuildingName = nodeB.sBuildingName,
                                                iLevelID = nodeL.iLevelID,
                                                sLevelName = nodeL.sLevelName,
                                                sLevelModelAS = nodeL.sLevelModelAS,
                                                sLevelModelMEP = nodeL.sLevelModelMEP,
                                                fLevelOrder = nodeL.fLevelOrder,

                                                iRoomID = nodeRo.iRoomID,
                                                sRoomName = nodeRo.sRoomName,

                                                iSystemID = nodeSys.iSystemID,
                                                sSystemName = nodeSys.sSystemName,
                                                iSystemSubID = nodeSysSub.iSystemSubID,
                                                sSystemSubName = nodeSysSub.sSystemSubName,


                                                iFacilityID = itemFacility.Field<int>("设备ID"),
                                                sFacilityName = itemFacility.Field<string>("设备名称"),
                                                sFacilityCode = itemFacility.Field<string>("设备编号"),

                                                sModel = itemFacility.Field<string>("模型"),
                                                sView = itemFacility.Field<string>("视图"),
                                                sDrawing = itemFacility.Field<string>("二维图纸"),
                                            };

                                            //控制空间、设备赋值
                                            nodeFacility.sModel_controlled = sModelT;
                                            //筛选控制楼层
                                            var qLevelControlled = from dtLevelControlled in dSet.Tables["标高设备监控表"].AsEnumerable()//查询楼层
                                                                   where (dtLevelControlled.Field<string>("设备编号") == nodeFacility.sFacilityCode)//条件
                                                                   select dtLevelControlled;
                                            if (qLevelControlled.Count() > 0)
                                                nodeFacility.nodetype_controlled = nodeType.LEVEL;

                                            foreach (var itemLevelControlled in qLevelControlled)
                                            {
                                                nodeFacility.iLevelID_controlled = itemLevelControlled.Field<int>("标高ID");
                                                nodeFacility.sLevelName_controlled = itemLevelControlled.Field<string>("标高名称");
                                                nodeFacility.fLevelOrder_controlled = itemLevelControlled.Field<double>("标高排序");

                                                if (itemLevelControlled.Field<string>("模型") != null)
                                                    nodeFacility.sModel_controlled = itemLevelControlled.Field<string>("模型");
                                                nodeFacility.sView_controlled = itemLevelControlled.Field<string>("视图");
                                                break;
                                            }

                                            //筛选控制空间
                                            if (nodeFacility.nodetype_controlled == nodeType.UNKNOW)
                                            {
                                                var qRoomControlled = from dtRoomControlled in dSet.Tables["房间设备监控表"].AsEnumerable()//查询空间
                                                                      where (dtRoomControlled.Field<string>("设备编号") == nodeFacility.sFacilityCode)//条件
                                                                      select dtRoomControlled;
                                                if (qRoomControlled.Count() > 0)
                                                    nodeFacility.nodetype_controlled = nodeType.ROOM;


                                                foreach (var itemRoomControlled in qRoomControlled)
                                                {
                                                    nodeFacility.iRoomID_controlled = itemRoomControlled.Field<int>("房间ID");
                                                    nodeFacility.sRoomCode_controlled = itemRoomControlled.Field<string>("房间编号");
                                                    nodeFacility.iLevelID_controlled = itemRoomControlled.Field<int>("标高ID");
                                                    nodeFacility.sLevelName_controlled = itemRoomControlled.Field<string>("标高名称");
                                                    nodeFacility.fLevelOrder_controlled = itemRoomControlled.Field<double>("标高排序");

                                                    if (itemRoomControlled.Field<string>("模型") != null)
                                                        nodeFacility.sModel_controlled = itemRoomControlled.Field<string>("模型");
                                                    nodeFacility.sView_controlled = itemRoomControlled.Field<string>("视图");
                                                    break;
                                                }
                                            }

                                            //筛选控制设备
                                            var qFacilityControll = from dtFacilityControlled in dSet.Tables["设备控制监控表"].AsEnumerable()//查询控制设备
                                                                    where (dtFacilityControlled.Field<string>("设备编号") == nodeFacility.sFacilityCode)//条件
                                                                    select dtFacilityControlled;
                                            foreach (var itemFacilityControll in qFacilityControll)
                                            {
                                                nodeFacility.fiFacility_control.iFacilityID = itemFacilityControll.Field<int>("控制设备ID");
                                                nodeFacility.fiFacility_control.sFacilityCode = itemFacilityControll.Field<string>("控制设备编号");
                                                nodeFacility.fiFacility_control.sFacilityName = itemFacilityControll.Field<string>("设备名称");

                                                break;
                                            }


                                            //筛选受控设备
                                            var qFacilityControlled = from dtFacilityControlled in dSet.Tables["设备控制监控表"].AsEnumerable()//查询受控设备
                                                                      where (dtFacilityControlled.Field<string>("控制设备编号") == nodeFacility.sFacilityCode)//条件
                                                                      select dtFacilityControlled;
                                            foreach (var itemFacilityControlled in qFacilityControlled)
                                            {

                                                FacilityIdenti fiControlled = new FacilityIdenti();
                                                fiControlled.iFacilityID = itemFacilityControlled.Field<int>("设备ID");
                                                fiControlled.sFacilityCode = itemFacilityControlled.Field<string>("设备编号");
                                                fiControlled.sFacilityName = itemFacilityControlled.Field<string>("设备名称");

                                                nodeFacility.listFacility_controlled.Add(fiControlled);
                                            }
                                            //if (nodeFacility.nodetype_controlled == nodeType.UNKNOW)
                                            //{
                                            //    var qFacilityControlled = from dtFacilityControlled in dSet.Tables["房间设备监控表"].AsEnumerable()//查询空间
                                            //                              where (dtFacilityControlled.Field<string>("设备编号") == nodeFacility.sFacilityCode)//条件
                                            //                              select dtFacilityControlled;
                                            //    if (qFacilityControlled.Count() > 0)
                                            //        nodeFacility.nodetype_controlled = nodeType.FACILITY;

                                            //    foreach (var itemFacilityControlled in qFacilityControlled)
                                            //    {
                                            //        nodeFacility.iRoomID_controlled = itemFacilityControlled.Field<int>("房间ID");
                                            //        nodeFacility.sRoomCode_controlled = itemFacilityControlled.Field<string>("房间编号");
                                            //        nodeFacility.iLevelID_controlled = itemFacilityControlled.Field<int>("标高ID");
                                            //        nodeFacility.sLevelName_controlled = itemFacilityControlled.Field<string>("标高名称");
                                            //        nodeFacility.fLevelOrder_controlled = itemFacilityControlled.Field<double>("标高排序");

                                            //        FacilityIdenti fiControlled = new FacilityIdenti();
                                            //        fiControlled.iFacilityID = itemFacilityControlled.Field<int>("控制设备ID");
                                            //        fiControlled.sFacilityCode = itemFacilityControlled.Field<string>("控制设备编号");
                                            //        fiControlled.sFacilityName = itemFacilityControlled.Field<string>("设备名称");

                                            //        if (itemFacilityControlled.Field<string>("模型") != null)
                                            //            nodeFacility.sModel_controlled = itemFacilityControlled.Field<string>("模型");
                                            //        nodeFacility.sView_controlled = itemFacilityControlled.Field<string>("视图");

                                            //        nodeFacility.listFacility_controlled.Add(fiControlled);
                                            //    }
                                            //}

                                            if (nodeFacility.sModel == "" || nodeFacility.sModel == null)
                                            {
                                                nodeFacility.sModel = sModelT;
                                            }
                                            else
                                            {
                                                sModelT = nodeFacility.sModel;
                                            }

                                            if (nodeFacility.sDrawing == "" || nodeFacility.sDrawing == null)
                                            {
                                                nodeFacility.sDrawing = sDrawingT;
                                            }
                                            else
                                            {
                                                //sDrawingT = nodeFacility.sDrawing;
                                            }

                                            if (nodeFacility.sView == "" || nodeFacility.sView == null)
                                            {
                                                nodeFacility.sView = sViewT;
                                            }
                                            else
                                            {
                                                sViewT = nodeFacility.sView;
                                            }

                                            nodeFacility.Parent = nodeRo;
                                            nodeRo.Children.Add(nodeFacility);
                                        }//设备

                                        if (checkboxFacilityOnly.IsChecked.HasValue)
                                            bTemp = (bool)checkboxFacilityOnly.IsChecked;
                                        else
                                            bTemp = false;

                                        nodeRo.Parent = nodeL;
                                        if (!bTemp || nodeRo.Children.Count > 0) //只显示设备空间
                                            nodeL.Children.Add(nodeRo);
                                    }

                                    nodeL.Parent = nodeSysSub;
                                    if (nodeL.Children.Count > 0)
                                        nodeSysSub.Children.Add(nodeL);
                                } //楼层

                                nodeSysSub.Parent = nodeSys;
                                if (nodeSysSub.Children.Count > 0)
                                    nodeSys.Children.Add(nodeSysSub);
                            } //子系统

                            nodeSys.Parent = nodeB;
                            if (nodeSys.Children.Count > 0)
                                nodeB.Children.Add(nodeSys);
                        } //系统

                        break;
                    #endregion
                }

                nodeB.Parent = nodeR;
                nodeR.Children.Add(nodeB);
            }

            
            itemList.Add(nodeR);

            this.treeviewSelectTree.ItemsSource = itemList;
            nodeList.Clear();
            GetAllNodes(itemList);
        }

        /*
        private class tagSelectTree
        {
            public nodeType nodetype { get; set; }  //节点类型
            public string sIcon { get; set; } //节点图标
            //public string EditIcon { get; set; }
            public string sDisplayName { get; set; } //显示名称
            public string sNote { get; set; } //注释
            public int iBuildingID { get; set; } //建筑ID
            public string sBuildingName { get; set; } //建筑名称
            public int iLevelID { get; set; } //ID
            public string sLevelName { get; set; } //         
            public int iRoomID { get; set; } //ID
            public string sRoomName { get; set; } //   
            public int iSystemID { get; set; } //ID
            public string sSystemName { get; set; } //   
            public int iFacilityID { get; set; } //ID
            public string sFacilityName { get; set; } //   
            public string sFacilityCode { get; set; } //设备编号
            public string sModel { get; set; } //模型
            public string sView { get; set; } //模型视图
            public bool IsExpanded { set; get; }
            public List<tagSelectTree> Children { get; set; }


            public tagSelectTree()
            {
                Children = new List<tagSelectTree>();

                iBuildingID = 0; sBuildingName = "";
                iLevelID = 0; sLevelName = "";
                iRoomID = 0; sRoomName = "";
                iSystemID = 0; sSystemName = "";
                iFacilityID = 0; sFacilityName = ""; sFacilityCode = "";
                IsExpanded = false;
            }

        }
        */
        private class tagSelectTreeStyle
        {
            public int iStyle{ get; set; }
            public string sStyle { get; set; }

            public tagSelectTreeStyle()
            {
                iStyle = 0;
                sStyle = "";
            }
        }
        private void ButtonSelectTree_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            
            
            DockSelectTree.IsVisible = !DockSelectTree.IsVisible;
            if (DockSelectTree.IsVisible)
                fb.Foreground = new SolidColorBrush(Colors.Black);
            else
                fb.Foreground = new SolidColorBrush(Colors.LightGray);
        }
        private void ButtonFacilitystate_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;

            DockFacilityStatus.IsVisible = !DockFacilityStatus.IsVisible;
            if (DockFacilityStatus.IsVisible)
                fb.Foreground = new SolidColorBrush(Colors.Black);
            else
                fb.Foreground = new SolidColorBrush(Colors.LightGray);
        }

        private void ButtonMonitorShow_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;

            DockMonitor.IsVisible = !DockMonitor.IsVisible;
            if (DockMonitor.IsVisible)
                fb.Foreground = new SolidColorBrush(Colors.Black);
            else
                fb.Foreground = new SolidColorBrush(Colors.LightGray);
        }
        private void ButtonProperty_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;

            Dockproperty.IsVisible = !Dockproperty.IsVisible;
            if (Dockproperty.IsVisible)
                fb.Foreground = new SolidColorBrush(Colors.Black);
            else
                fb.Foreground = new SolidColorBrush(Colors.LightGray);
        }
        private void treeviewSelectTree_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            /*
            tagSelectTree t = treeviewSelectTree.SelectedItem as tagSelectTree;
            if (t == null)
                return;
            MessageBox.Show(t.sFacilityCode);
             */
        }

        private void ButtonSelectTreeRefresh_Click(object sender, RoutedEventArgs e)
        {
            //初始化选择树
            InitSelectTree(int.Parse(comoBoxSelectStyle.SelectedValue.ToString()));
        }

        private void comoBoxSelectStyle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            InitSelectTree(int.Parse(comoBoxSelectStyle.SelectedValue.ToString()));
        }

        // 辅助函数进行递归查询视点
        private SavedItem getViewPoint(string sViewName)
        {
            SavedItem si = null;
            if (sViewName.Trim() == "")
                return null;

            Document oDoc = documentControlM.Document;

            foreach (SavedItem oSVP in oDoc.SavedViewpoints.Value)
            {
                if (oSVP.DisplayName == sViewName)
                    return oSVP;

                if (oSVP.IsGroup)  // 如果是组或者动画剪辑组  
                {
                    si = recurseView(oSVP, sViewName);
                    if (si != null)
                        return si;
                }
                else
                {  // 访问保存视点信息，例如对应的视点  }
                    if (oSVP.DisplayName == sViewName)
                    {
                        return oSVP;
                    }
                }
            }
            return null;
        }
        private SavedItem recurseView(SavedItem oFolder, string sViewName)
        {
            SavedItem si = null;

            foreach (SavedItem oSItem in ((Autodesk.Navisworks.Api.GroupItem)oFolder).Children)
            {
                if (oSItem.DisplayName == sViewName)
                    return oSItem;

                if (oSItem.IsGroup)
                {
                    si=recurseView(oSItem, sViewName);
                    if (si != null)
                        return si;
                }

                if (oSItem is SavedViewpoint)
                {
                    SavedViewpoint oSVPt = oSItem as SavedViewpoint;
                    // 访问保存视点信息，例如对应的视点}
                    if (oSVPt.DisplayName == sViewName)
                        return oSItem;

                    if (oSItem is SavedViewpointAnimationCut)
                    {
                        //SavedViewpointAnimationCut oSVCut = oSItem as SavedViewpointAnimationCut;
                        // 访问保存视点信息，例如对应的视点} 
                    }
                }
            }
            return null;
        }

        private void ButtonSelect_Click(object sender, RoutedEventArgs e)
        {
            if (documentControlM.Document == null)
                return;

            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.Select;

            if (mainModelView.IsSelected)
                viewControl.Focus();
            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();

            if (documentControlM2D.Document == null)
                return;
            documentControlM2D.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.Select;

            if (mainModelView2D.IsSelected)
                viewControl2D.Focus();

        }

        private void ButtonSelectBox_Click(object sender, RoutedEventArgs e)
        {
            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.SelectBox;
            
            viewControl.Focus();
            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();
        }

        private void ButtonPan_Click(object sender, RoutedEventArgs e)
        {
            
            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigatePan;
            viewControl.Focus();
            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();

            if (documentControlM2D.Document == null)
                return;
            documentControlM2D.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigatePan;
        }

        private void ButtonZoom_Click(object sender, RoutedEventArgs e)
        {
            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigateZoom;
            viewControl.Focus();
            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();

            if (documentControlM2D.Document == null)
                return;
            documentControlM2D.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigateZoom;
        }

        private void ButtonWalk_Click(object sender, RoutedEventArgs e)
        {
            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigateWalk;
            viewControl.Focus();


            //ComApiBridge.State.CurrentView.ViewPoint.
            Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2 oV = (Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2)ComApiBridge.State.CurrentView.ViewPoint;

            if (oV.Paradigm == Autodesk.Navisworks.Api.Interop.ComApi.nwEParadigm.eParadigm_WALK)
            {
                oV.Viewer.AutoCrouch = true; //toogle auto crouch on

                if ((bool)checkboxCollisionDetection.IsChecked)
                    oV.Viewer.CollisionDetection = true;
                else
                    oV.Viewer.CollisionDetection = false;

                if ((bool)checkboxGravity.IsChecked)
                    oV.Viewer.Gravity = true;
                else
                    oV.Viewer.Gravity = false;
            }


            //Viewpoint v = documentControlM.Document.CurrentViewpoint;
            //MessageBox.Show(v.ViewerAvatar);
            //documentControlM.Document.CurrentViewpoint.CopyFrom(v);

            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();
        }

        private void ButtonLookAround_Click(object sender, RoutedEventArgs e)
        {
            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigateLookAround;
            viewControl.Focus();

            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();
        }

        private void ButtonOrbit_Click(object sender, RoutedEventArgs e)
        {
            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigateConstrainedOrbit;
            viewControl.Focus();

            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();
        }


        private void ButtonMeasure_Click(object sender, RoutedEventArgs e)
        {

            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.MeasurePointToPoint;
            viewControl.Focus();
            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();
        }

        private void ButtonFree_Click(object sender, RoutedEventArgs e)
        {
            if (documentControlM.Document == null)
                return;
            TimerStop();

            documentControlM.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigateFreeOrbit;
            viewControl.Focus();
            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();

            
        }

        private void ContextMenuSelectTree_Loaded(object sender, RoutedEventArgs e)
        {
            if(treeviewSelectTree.SelectedItem==null)
                return;

            TreeNodetagSelectTree tagST = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            switch(tagST.nodetype)
            {
                case nodeType.PROJECT:
                    miSelectLocation.IsEnabled = true;
                    miSelectDrawings.Visibility = Visibility.Collapsed;
                    break;
                case nodeType.BUILDING:
                    miSelectLocation.IsEnabled = true;
                    miSelectDrawings.Visibility = Visibility.Collapsed;
                    break;
                case nodeType.ROOM:
                    miSelectLocation.IsEnabled = true;
                    miSelectDrawings.Visibility = Visibility.Collapsed;
                    break;
                case nodeType.SYSTEM:
                    miSelectLocation.IsEnabled = false;
                    miSelectDrawings.Visibility = Visibility.Collapsed;
                    break;
                case nodeType.FACILITY:
                    miSelectLocation.IsEnabled = true;
                    miSelectDrawings.Visibility = Visibility.Collapsed;
                    break;
                case nodeType.LEVEL:
                    miSelectLocation.IsEnabled = true;
                    miSelectDrawings.Visibility = Visibility.Visible;
                    break;
                default:
                    break;
            }
        }

        private void miSelectExpand_Click(object sender, RoutedEventArgs e)
        {
            if (treeviewSelectTree.SelectedItem == null)
                return;

            ExpandInternal(treeviewSelectTree.SelectedItem as TreeNodetagSelectTree);

        }

        //展开
        private static void ExpandInternal(TreeNodetagSelectTree treeViewItem)
        {
            if (treeViewItem == null) return;
            treeViewItem.IsExpanded = true;

            foreach (TreeNodetagSelectTree treeItem in treeViewItem.Children)
            {
                if (treeItem == null) continue;
                if (treeItem.Children.Count<1) continue;

                treeItem.IsExpanded = true;
                ExpandInternal(treeItem);
            }

        }

        //选择图纸
        private void miSelectDrawings_Click(object sender, RoutedEventArgs e)
        {
            string sFileDrawing;

            if (treeviewSelectTree.SelectedItem == null)
                return;

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.nodetype != nodeType.LEVEL)
            {
                MessageBox.Show("请选择楼层");
                return;
            }

            //找到图纸
            var qLevelDrawings = from dtLevelDrawings in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                                 where (dtLevelDrawings.Field<int>("标高ID") == tnSelectTree.iLevelID)//条件
                                 select dtLevelDrawings;


            foreach (var itemLevelDrawings in qLevelDrawings)//显示查询结果
            {
                sFileDrawing = itemLevelDrawings.Field<string>("图纸");
                if (sFileDrawing == null || sFileDrawing == "")
                    break;

                mainDrawingView.Title = sFileDrawing;
                sFileDrawing = sDirDrawings + "/" + sFileDrawing;
                //MessageBox.Show(sFileDrawing);

                if (File.Exists(sFileDrawing))
                    //viewPDF.LoadPDFfile(sFileDrawing);
                {
                    string s = "file:///" + sFileDrawing;
                    webBrowerDrawing.Navigate("file:///" + sFileDrawing);
                }
                break;
            }

            
        }

        //二维图纸定位
        private void miSelect2DDrawing_Click(object sender, RoutedEventArgs e)
        {


            string sFileDrawing;

            if (treeviewSelectTree.SelectedItem == null)
                return;

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.sDrawing=="")
            {
                MessageBox.Show("没有找到对应二维图纸");
                return;
            }



            sFileDrawing = sDirDrawings + "/" + tnSelectTree.sDrawing;

            try
            {   
                if (!File.Exists(sFileDrawing)) //没有文件，返回
                //viewPDF.LoadPDFfile(sFileDrawing);
                {
                    return;
                }


                if(documentControlM2D.Document!=null) 
                {
                    if(documentControlM2D.Document.FileName.ToUpper()!=sFileDrawing) //不同文件，调入
                    {
                        documentControlM2D.Document.TryOpenFile(sFileDrawing);
                    }
                }
                else //调入新文件
                {
                    documentControlM2D.Document.TryOpenFile(sFileDrawing);
                }

                //查找
                //先定位模型
                List<int> listElementID;
                int iSelectNumber = modelLocation(out listElementID);

                if (iSelectNumber > 0) //有选择的
                {
                    documentControlM2D.Document.CurrentSelection.Clear();
                    ModelItemCollection items=new ModelItemCollection();
                    foreach (int iEid in listElementID)
                    {
                        Search search = new Search();
                        search.Selection.SelectAll();

                        SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");


                        oSearchCondition = oSearchCondition.DisplayStringContains(iEid.ToString());
                        search.SearchConditions.Add(oSearchCondition);
                        ModelItemCollection items1 = search.FindAll(documentControlM2D.Document, false);

                        items.AddRange(items1.ToArray());
                    }

                    documentControlM2D.Document.CurrentSelection.CopyFrom(items);

                    Viewpoint v = documentControlM2D.Document.CurrentViewpoint;
                    BoundingBox3D box = documentControlM2D.Document.CurrentSelection.SelectedItems.BoundingBox();

                    v.ZoomBox(box);

                    documentControlM2D.Document.CurrentViewpoint.CopyFrom(v);

                    mainModelView2D.IsSelected = true;
                }

                

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
            



        }


        private DependencyObject GetContainerFormObject(ItemsControl item, object obj)
        {
            if (item == null)
                return null;

            DependencyObject dObject = null;
            dObject = item.ItemContainerGenerator.ContainerFromItem(obj);

            if (dObject != null)
                return dObject;

            var query = from childItem in item.Items.Cast<object>()
                        let childControl = item.ItemContainerGenerator.ContainerFromItem(childItem) as ItemsControl
                        select GetContainerFormObject(childControl, obj);

            return query.FirstOrDefault(i => i != null);
        }


        //定位
        private void miSelectLocation_Click(object sender, RoutedEventArgs e)
        {
            //int i,j;

            //if (treeviewSelectTree.SelectedItem == null)
            //    return;

            //TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            //if (tnSelectTree.sModel.Trim() == "")
            //    return;

            ////隐藏，显示所需的楼层
            //switch (tnSelectTree.nodetype)
            //{
            //    case nodeType.PROJECT:
            //        showProject(true);
            //        break;
            //    case nodeType.BUILDING:
            //        showSelectedBuilding(tnSelectTree.iBuildingID,true);
            //        break;
            //    default:
            //        showSelectedLevel(tnSelectTree.sLevelModelAS, tnSelectTree.sLevelModelMEP, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);
            //        break;
            //}
            ////showSelectedLevel(tnSelectTree.sLevelModelAS, tnSelectTree.sLevelModelMEP, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);
            //if (sModelName != tnSelectTree.sModel)
            //{
            //    sModelName = tnSelectTree.sModel;
            //    if (sModelName.Trim()!="")
            //    {
            //        sFile = sDir + "\\" + sModelName;
            //        documentControlM.Document.TryOpenFile(sFile);
            //    }
                
            //}
            //sViewName = tnSelectTree.sView;
            //if (sViewName.Trim() == "")
            //    return;

            //SavedItem siView = getViewPoint(sViewName);
            //if (siView == null)
            //    return;
            //documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;

            //Search search = new Search();
            //search.Selection.SelectAll();

            //VariantData oData;


            ////ModelItemCollection hidden = new ModelItemCollection();
            //switch (tnSelectTree.nodetype)
            //{
            //    case nodeType.PROJECT:
            //        break;
            //    case nodeType.BUILDING:
            //        break;
            //    case nodeType.LEVEL:
            //        break;
            //    case nodeType.SYSTEM:
            //        break;
            //    case nodeType.ROOM:
            //        //showSelectedLevel(tnSelectTree.sLevelModel, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID,true,true);
            //        //选取
            //        oData = VariantData.FromDisplayString(tnSelectTree.sRoomCode);

            //        SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素", "编号");
            //        oSearchCondition = oSearchCondition.EqualValue(oData);
            //        search.SearchConditions.Add(oSearchCondition);

            //        ModelItemCollection items = search.FindAll(documentControlM.Document, false);
            //        //documentControlM.Document.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
            //        //documentControlM.Document.Models.OverridePermanentTransparency(items, 0.5);
            //        documentControlM.Document.CurrentSelection.Clear();
            //        documentControlM.Document.CurrentSelection.CopyFrom(items);

            //        //显示选定房间
            //        documentControlM.Document.Models.SetHidden(items, false); 


            //        break;
            //    case nodeType.FACILITY:
            //        //showSelectedLevel(tnSelectTree.sLevelModel, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);

            //        oData = VariantData.FromDisplayString(tnSelectTree.sFacilityCode);

            //        SearchCondition oSearchCondition1 = SearchCondition.HasPropertyByDisplayName("元素", "设备编号");
            //        oSearchCondition1 = oSearchCondition1.EqualValue(oData);
            //        search.SearchConditions.Add(oSearchCondition1);

            //        ModelItemCollection items1 = search.FindAll(documentControlM.Document, false);
            //        //documentControlM.Document.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
            //        //documentControlM.Document.Models.OverridePermanentTransparency(items, 0.5);
            //        documentControlM.Document.CurrentSelection.Clear();
            //        documentControlM.Document.CurrentSelection.CopyFrom(items1);

            //        //显示选定设备
            //        documentControlM.Document.Models.SetHidden(items1, false); 
            //        break;

            //}

            List<int> listElementID ;
            int iSelectNumber = modelLocation(out listElementID);

        }

        //定位
        private int modelLocation(out List<int> listElementID)
        {
            int iSelectItemNumber = 0;
            int i,j;

            listElementID = new List<int>();

            if (treeviewSelectTree.SelectedItem == null)
                return 0;

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.sModel.Trim() == "")
                return 0;

            //隐藏，显示所需的楼层
            switch (tnSelectTree.nodetype)
            {
                case nodeType.PROJECT:
                    showProject(true);
                    break;
                case nodeType.BUILDING:
                    showSelectedBuilding(tnSelectTree.iBuildingID,true);
                    break;
                default:
                    showSelectedLevel(tnSelectTree.sLevelModelAS, tnSelectTree.sLevelModelMEP, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);
                    break;
            }
            //showSelectedLevel(tnSelectTree.sLevelModelAS, tnSelectTree.sLevelModelMEP, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);
            if (sModelName != tnSelectTree.sModel)
            {
                sModelName = tnSelectTree.sModel;
                if (sModelName.Trim()!="")
                {
                    sFile = sDir + "\\" + sModelName;
                    documentControlM.Document.TryOpenFile(sFile);
                }
                
            }
            sViewName = tnSelectTree.sView;
            if (sViewName.Trim() == "")
                return 0;

            SavedItem siView = getViewPoint(sViewName);
            if (siView != null)
            {
                documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;
            }

            Search search = new Search();
            search.Selection.SelectAll();

            VariantData oData;


            //ModelItemCollection hidden = new ModelItemCollection();
            switch (tnSelectTree.nodetype)
            {
                case nodeType.PROJECT:
                    break;
                case nodeType.BUILDING:
                    break;
                case nodeType.LEVEL:
                    if (timerMonitor.Enabled) //全层监视跳过
                        break;



                    //showSelectedLevel(tnSelectTree.sLevelModel, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID,true,true);
                    //选取
                    oData = VariantData.FromDisplayString("文件");
                    SearchCondition oSearchConditionL = SearchCondition.HasPropertyByDisplayName("项目", "类型");
                    oSearchConditionL = oSearchConditionL.EqualValue(oData);
                    search.SearchConditions.Add(oSearchConditionL);

                    oData = VariantData.FromDisplayString(tnSelectTree.sLevelModelAS);
                    SearchCondition oSearchConditionL1 = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                    oSearchConditionL1 = oSearchConditionL1.EqualValue(oData);
                    search.SearchConditions.Add(oSearchConditionL1);

                    ModelItemCollection itemsL = search.FindAll(documentControlM.Document, false);


                    documentControlM.Document.CurrentSelection.Clear();
                    documentControlM.Document.CurrentSelection.CopyFrom(itemsL);

                    //显示选定楼层
                    documentControlM.Document.Models.SetHidden(itemsL, false);

                    //取得ElementID

                    break;

                case nodeType.SYSTEM:
                    break;
                case nodeType.ROOM:
                    //showSelectedLevel(tnSelectTree.sLevelModel, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID,true,true);
                    //选取
                    oData = VariantData.FromDisplayString(tnSelectTree.sRoomCode);

                    SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素", "编号");
                    oSearchCondition = oSearchCondition.EqualValue(oData);
                    search.SearchConditions.Add(oSearchCondition);

                    ModelItemCollection items = search.FindAll(documentControlM.Document, false);
                    //documentControlM.Document.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
                    //documentControlM.Document.Models.OverridePermanentTransparency(items, 0.5);

                    documentControlM.Document.CurrentSelection.Clear();
                    documentControlM.Document.CurrentSelection.CopyFrom(items);

                    //显示选定房间
                    documentControlM.Document.Models.SetHidden(items, false);

                    //取得ElementID
                    DataProperty oDP_DWGHandle;
                    foreach (ModelItem mi in items)
                    {
                        oDP_DWGHandle = mi.PropertyCategories.FindPropertyByDisplayName("元素 ID", "值");
                        if (oDP_DWGHandle!=null)
                        {
                            listElementID.Add(int.Parse(oDP_DWGHandle.Value.ToDisplayString()));
                        }
                    }
                    iSelectItemNumber = listElementID.Count;
                    break;
                case nodeType.FACILITY:
                    //showSelectedLevel(tnSelectTree.sLevelModel, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);

                    oData = VariantData.FromDisplayString(tnSelectTree.sFacilityCode);

                    SearchCondition oSearchCondition1 = SearchCondition.HasPropertyByDisplayName("元素", "设备编号");
                    oSearchCondition1 = oSearchCondition1.EqualValue(oData);
                    search.SearchConditions.Add(oSearchCondition1);

                    ModelItemCollection items1 = search.FindAll(documentControlM.Document, false);
                    //documentControlM.Document.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
                    //documentControlM.Document.Models.OverridePermanentTransparency(items, 0.5);
                    documentControlM.Document.CurrentSelection.Clear();
                    documentControlM.Document.CurrentSelection.CopyFrom(items1);

                    //显示选定设备
                    documentControlM.Document.Models.SetHidden(items1, false);

                    //取得ElementID
                    DataProperty oDP_DWGHandle1;
                    foreach (ModelItem mi in items1)
                    {
                        oDP_DWGHandle1 = mi.PropertyCategories.FindPropertyByDisplayName("元素 ID", "值");
                        if (oDP_DWGHandle1 != null)
                        {
                            listElementID.Add(int.Parse(oDP_DWGHandle1.Value.ToDisplayString()));
                        }
                    }
                    iSelectItemNumber = listElementID.Count;
                    break;

            }



            return iSelectItemNumber;
        }

        //监控定位
        private void miSelectControlLocation_Click(object sender, RoutedEventArgs e)
        {
            int i, j;

            if (treeviewSelectTree.SelectedItem == null)
                return;

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.sModel.Trim() == "")
                return;

            if (tnSelectTree.nodetype_controlled == nodeType.UNKNOW)
            {
                MessageBox.Show("没有控制记录");
                return;
            }

            //隐藏，显示所需的楼层
            switch (tnSelectTree.nodetype_controlled)
            {
                case nodeType.PROJECT:
                    break;
                case nodeType.BUILDING:
                    showSelectedBuilding(tnSelectTree.iBuildingID, true);
                    break;
                default:
                    showSelectedLevel(tnSelectTree.sLevelModelAS, tnSelectTree.sLevelModelMEP, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);
                    break;
            }
            //showSelectedLevel(tnSelectTree.sLevelModelAS, tnSelectTree.sLevelModelMEP, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);
            if (sModelName != tnSelectTree.sModel_controlled)
            {
                sModelName = tnSelectTree.sModel_controlled;
                if (sModelName.Trim() != "")
                {
                    sFile = sDir + "\\" + sModelName;
                    documentControlM.Document.TryOpenFile(sFile);
                }

            }
            sViewName = tnSelectTree.sView_controlled;
            if (sViewName.Trim() == "")
                return;

            SavedItem siView = getViewPoint(sViewName);
            if (siView == null)
                return;
            documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;

            Search search = new Search();
            search.Selection.SelectAll();

            VariantData oData;


            //ModelItemCollection hidden = new ModelItemCollection();
            switch (tnSelectTree.nodetype_controlled)
            {
                case nodeType.PROJECT:
                    break;
                case nodeType.BUILDING:
                    break;
                case nodeType.LEVEL:
                    //选取土建模型展示

                    oData = VariantData.FromDisplayString(tnSelectTree.sLevelModelAS);

                    SearchCondition oSearchCondition1 = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                    oSearchCondition1 = oSearchCondition1.EqualValue(oData);
                    search.SearchConditions.Add(oSearchCondition1);

                    ModelItemCollection items1 = search.FindAll(documentControlM.Document, false);
                    //documentControlM.Document.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
                    //documentControlM.Document.Models.OverridePermanentTransparency(items, 0.5);
                    documentControlM.Document.CurrentSelection.Clear();
                    documentControlM.Document.CurrentSelection.CopyFrom(items1);

                    //显示选定楼层
                    documentControlM.Document.Models.SetHidden(items1, false);

                    refreshLevelMoniter(tnSelectTree.iLevelID_controlled);

                    break;
                case nodeType.SYSTEM:
                    break;
                case nodeType.ROOM:
                    //showSelectedLevel(tnSelectTree.sLevelModel, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID,true,true);
                    //选取
                    oData = VariantData.FromDisplayString(tnSelectTree.sRoomCode_controlled);

                    SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素", "编号");
                    oSearchCondition = oSearchCondition.EqualValue(oData);
                    search.SearchConditions.Add(oSearchCondition);

                    ModelItemCollection items = search.FindAll(documentControlM.Document, false);
                    //documentControlM.Document.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
                    //documentControlM.Document.Models.OverridePermanentTransparency(items, 0.5);
                    documentControlM.Document.CurrentSelection.Clear();
                    documentControlM.Document.CurrentSelection.CopyFrom(items);

                    //显示选定房间
                    documentControlM.Document.Models.SetHidden(items, false);


                    break;
                //case nodeType.FACILITY:
                //    //showSelectedLevel(tnSelectTree.sLevelModel, tnSelectTree.fLevelOrder, tnSelectTree.iBuildingID, true, true);

                //    oData = VariantData.FromDisplayString(tnSelectTree.sFacilityCode);

                //    SearchCondition oSearchCondition1 = SearchCondition.HasPropertyByDisplayName("元素", "设备编号");
                //    oSearchCondition1 = oSearchCondition1.EqualValue(oData);
                //    search.SearchConditions.Add(oSearchCondition1);

                //    ModelItemCollection items1 = search.FindAll(documentControlM.Document, false);
                //    //documentControlM.Document.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
                //    //documentControlM.Document.Models.OverridePermanentTransparency(items, 0.5);
                //    documentControlM.Document.CurrentSelection.Clear();
                //    documentControlM.Document.CurrentSelection.CopyFrom(items1);

                //    //显示选定设备
                //    documentControlM.Document.Models.SetHidden(items1, false);
                    //break;

            }

        }


        private void ButtonOption_Click(object sender, RoutedEventArgs e)
        {
            WindowOption wOption = new WindowOption();
            wOption.ShowDialog();
        }

        private void CurrentSelection_Changed(object sender, EventArgs e)
        {
            showFeature();

        }

        private void OnTimedEventMonitor(object sender, EventArgs e)
        {
             this.Dispatcher.Invoke(DispatcherPriority.Normal,
                 new TimerDispatcherDelegate(updateUIMonitor1));
        }

        private void OnTimedEventRoomMonitor(object sender, EventArgs e)
        {
            this.Dispatcher.Invoke(DispatcherPriority.Normal,
                new TimerDispatcherDelegate(updateUIRoomMonitor));
        }

        //监控过滤
        private void updateUIMonitor1()
        {
            if (!timerMonitor.Enabled)
                return;

            updateUIMonitor();
        }

        //监控设备参数
         private void updateUIMonitor()
         {
             //if (!timerMonitor.Enabled)
             //    return;

             int i, j;
             string sCode;

             dtDGMmonitor.Clear();
             // Respond to event...
             if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
                 return;


             ModelItem oSelectedItem = documentControlM.Document.CurrentSelection.SelectedItems.ElementAt<ModelItem>(0);
             switch(selectItemType(oSelectedItem,out sCode))
             {
                 case nodeType.ROOM: //房间监控
                     refreshRoomMoniter(sCode);
                     break;

                 case nodeType.FACILITY: //设备监控
                     refreshFacilityMoniter();

                     break;

                 default:
                     break;


             }

         }

         //监控房间动态曲线
         private void updateUIRoomMonitor()
         {
             if (!timerRoomMonitor.Enabled)
                 return;

             int i, j;
             string sCode;

             if (sMonitorRoomCode == "")
                 return;

             refreshRoomMoniterCurve(sMonitorRoomCode); //
             bMonitorRoomFirst = false;

             //得到现选的房间编号
             if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
             {

             }
             else
             {
                 ModelItem oSelectedItem = documentControlM.Document.CurrentSelection.SelectedItems.ElementAt<ModelItem>(0);
                 if (selectItemType(oSelectedItem, out sCode) != nodeType.ROOM)
                 {
                     if (sCode != sMonitorRoomCode) //更换了房间
                     {
                         sMonitorRoomCode = sCode;
                         bMonitorRoomFirst = true;
                     }

                 }
             }


             
 
         }

        //***************************************************
        //刷新设备监控属性
        /*
        private void refreshFacilityMoniter(string sCode)
        {
            int i,j;
            int iParaCount = 0;

            SolidColorBrush scbBackGround = new SolidColorBrush(Colors.LightGreen);
            DataGridRow dgRow;

            if (sCode == null || sCode == "")
                return;

            string sDataSourceLocation, sDataSource,sFacilityName,sSystemName;
            int iStyleNumber;
            string sValue3D="",sFloor3D="",sCeil3D="",sNormal="1";
            string sPara = "";

            DataSourceType dsType = selectFacilityDataSource(sCode, out sDataSourceLocation, out sDataSource, out iStyleNumber, out sFacilityName, out sSystemName);

            if (dsType == DataSourceType.NONE)
                return;

            
            //得到参数列表和调整值
            var qPara = from dtPara in dSet.Tables["系统监控特性表"].AsEnumerable()//查询特性
                where (dtPara.Field<int>("系统ID") == iStyleNumber) //条件
                select dtPara;
            iParaCount = qPara.Count();

            object[] oTemp = new object[3];
            oTemp[0] = "编号"; oTemp[1] = sCode; oTemp[2] = "0";
            dtDGMmonitor.Rows.Add(oTemp);

            oTemp[0] = "名称"; oTemp[1] = sFacilityName;
            dtDGMmonitor.Rows.Add(oTemp);

            //得到数据源
            switch (dsType)
            {
                #region DATABASE DATASOURCE
                case DataSourceType.DATABASE: //数据库

                    //现在先考虑单一数据库操作
                    try
                    {
                        sqlConnM.Open();
                        sqlCommM.CommandText = sDataSource.Replace(sREPLACE, sCode);

                        if (dSetM.Tables.Contains("设备监控表")) dSetM.Tables.Remove("设备监控表");
                        sqlDAM.Fill(dSetM, "设备监控表");
                        sqlConnM.Close();

                        for (i = 0; i < dSetM.Tables["设备监控表"].Rows.Count; i++)
                        {
                            for (j = 0; j < dSetM.Tables["设备监控表"].Columns.Count; j++)
                            {
                                if (j < iParaCount) //采用定义参数名，并调整参数
                                {
                                    if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                        oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                                    else
                                        oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                                    oTemp[1] = (decimal.Parse(dSetM.Tables["设备监控表"].Rows[i][j].ToString()) * qPara.ElementAt(j).Field<decimal>("特性调整")).ToString();
                                }
                                else
                                {
                                    oTemp[0] = dSetM.Tables["设备监控表"].Columns[j].ColumnName; oTemp[1] = dSetM.Tables["设备监控表"].Rows[i][j].ToString();
                                }
                                sValue3D = oTemp[1].ToString();
                                dtDGMmonitor.Rows.Add(oTemp);
                            }
                        }

                        DGMmonitor.ItemsSource = dtDGMmonitor.AsDataView();

                        if (!Window.GetWindow(DGMmonitor).IsVisible)
                        {
                            Window.GetWindow(DGMmonitor).Show();
                        }
                        DGMmonitor.UpdateLayout();

                        for (i = 2; i < DGMmonitor.ItemContainerGenerator.Items.Count; i++)
                        {
                            dgRow = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(i);
                            //
                            if (i - 2 < iParaCount) //取值小于属性定义，判断取值空间
                            {
                                if (qPara.ElementAt(i - 2).Field<decimal?>("低阀值") != null) //有低值
                                {
                                    sFloor3D = qPara.ElementAt(i - 2).Field<decimal>("低阀值").ToString();
                                    if (decimal.Parse(dtDGMmonitor.Rows[i][1].ToString()) < qPara.ElementAt(i - 2).Field<decimal>("低阀值"))
                                    {
                                        scbBackGround = new SolidColorBrush(Colors.LightPink);
                                        sNormal = "0";
                                    }
                                }
                                if (qPara.ElementAt(i - 2).Field<decimal?>("高阀值") != null) //有高值
                                {
                                    sCeil3D = qPara.ElementAt(i - 2).Field<decimal>("高阀值").ToString();
                                    if (decimal.Parse(dtDGMmonitor.Rows[i][1].ToString()) > qPara.ElementAt(i - 2).Field<decimal>("高阀值"))
                                    {
                                        scbBackGround = new SolidColorBrush(Colors.Red);
                                        sNormal = "0";
                                    }
                                }
                            }
                            dgRow.Background = scbBackGround;
                        }
                    }
                    catch
                    {
                        oTemp[0] = "错误";
                        oTemp[1] = "数据库读取失败";
                        dtDGMmonitor.Rows.Add(oTemp);
                        DGMmonitor.ItemsSource = dtDGMmonitor.AsDataView();
                        DGMmonitor.UpdateLayout();
                        dgRow = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(2);
                        dgRow.Background = new SolidColorBrush(Colors.Red); 
                    }

                    break;
                #endregion



                #region OBIX DATASOURCE
                case DataSourceType.OBIX://Niagara
                    HttpWebResponseUtility.tagUrl = sDataSourceLocation + sCode + "/";


                    string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponse();
                    if (responseFromServer == null || responseFromServer == "")
                        break;

                    XmlDocument xmlDoc = new XmlDocument();
                    XDocument xDoc=new XDocument();
                    XmlReader xmlReader = xDoc.CreateReader();

                    xmlDoc.LoadXml(responseFromServer);
                    XElement xEle = XElement.Parse(responseFromServer);
                    xmlDoc.PreserveWhitespace = false;


                    //查询每个参数
                    for(i=0;i<iParaCount;i++)
                    {
                        sPara = qPara.ElementAt(i).Field<string>("OBIX取值");

                        var text = from t in xEle.Elements()//定位到节点 
                                    .Where(w => w.Attribute("name").Value.Equals(sPara))   //若要筛选就用上这个语句 
                                   select new
                                   {
                                       disp = t.Attribute("display").Value   //注意此处用到 attribute 
                                   };

                        foreach (var sValue in text)
                        {
                            if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                            else
                                oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                            switch (qPara.ElementAt(0).Field<string>("取值类型"))
                            {
                                case "INT":
                                    oTemp[1] = (decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp)) * qPara.ElementAt(i).Field<decimal>("特性调整")).ToString();
                                    break;
                                default:
                                    oTemp[1] = HttpWebResponseUtility.GetResultObix(sValue.disp);
                                    break;
                            }

                            sValue3D = oTemp[1].ToString(); //？
                            dtDGMmonitor.Rows.Add(oTemp);
                            break; //只查找第一个参数
                        }

                    }
                   
                    DGMmonitor.ItemsSource = dtDGMmonitor.AsDataView();
                    if (!Window.GetWindow(DGMmonitor).IsVisible)
                    {
                        Window.GetWindow(DGMmonitor).Show();
                    }
                    DGMmonitor.UpdateLayout();


                    for (i = 2; i < DGMmonitor.ItemContainerGenerator.Items.Count; i++)
                    {
                        //判断是否取到该项取值
                        if (!dtDGMmonitor.Rows[i][0].ToString().StartsWith(qPara.ElementAt(i - 2).Field<string>("特性名称")))
                        {
                            continue;
                        }
                        dgRow = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(i);
                        //
                        if (i - 2 < iParaCount) //取值小于属性定义，判断取值空间
                        {
                            if (qPara.ElementAt(i - 2).Field<decimal?>("低阀值") != null) //有低值
                            {
                                sFloor3D = qPara.ElementAt(i - 2).Field<decimal>("低阀值").ToString();
                                if (decimal.Parse(dtDGMmonitor.Rows[i][1].ToString()) < qPara.ElementAt(i - 2).Field<decimal>("低阀值"))
                                {
                                    scbBackGround = new SolidColorBrush(Colors.LightPink);
                                    sNormal = "0";
                                }
                            }
                            if (qPara.ElementAt(i - 2).Field<decimal?>("高阀值") != null) //有高值
                            {
                                sCeil3D = qPara.ElementAt(i - 2).Field<decimal>("高阀值").ToString();
                                if (decimal.Parse(dtDGMmonitor.Rows[i][1].ToString()) > qPara.ElementAt(i - 2).Field<decimal>("高阀值"))
                                {
                                    scbBackGround = new SolidColorBrush(Colors.Red);
                                    sNormal = "0";
                                }
                            }
                        }
                        dgRow.Background = scbBackGround;
                    }
                    
         

                    //DataRowView drv = DGMmonitor.Items[2] as DataRowView;
                    //DataGridRow row = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(2);
                    //if (row != null)
                    //{
                    //    int it = 0;
                    //    it=int.Parse(dSetM.Tables["设备监控表"].Rows[0][0].ToString());
                    //    if (it<10)
                    //    row.Background = new SolidColorBrush(Colors.LightGreen);
                    //    if (it >= 10 && it < 20)
                    //        row.Background = new SolidColorBrush(Colors.LightPink);
                    //    if (it >= 20)
                    //        row.Background = new SolidColorBrush(Colors.Red);

                    //}

                    break;
                #endregion
                default:
                    break;
            }


            //U3d显示
            //showFeatureU3D(sSystemName,sCode,sValue3D,sFloor3D,sCeil3D);
            showFeatureU3D(sSystemName, sCode, sValue3D, sNormal);
            
        }
         */

        //***************************************************
        //刷新所有设备属性
        private void refreshFacilityMoniter()
        {


            string sFacilityCode = "", sFacilityCodeStore = "",sValueU3D = "",sNormalU3D = "1";
            int i, j;
            int iParaCount = 0;


            string sDataSourceLocation, sDataSource, sFacilityName, sSystemName="";
            int iStyleNumber;
            string sFloor3D = "", sCeil3D = "";
            string sPara = "";
            string sParaName = "",sParaAttr="",sParaDisp="";
            string sT1 = "",sT2="";

            DataSourceType dsType;
            object[] oTemp = new object[3]; //3颜色控制: 0 正常 1 过高 -1 过低 100 设备名称 200 房间名称

            string sCode="";
            //bool bfirstMoniter = true;

            foreach (ModelItem oSelectedItem in documentControlM.Document.CurrentSelection.SelectedItems)//显示查询结果
            {
                //sFacilityCode = "";
                selectItemType(oSelectedItem, out sFacilityCodeStore);

                if (sFacilityCodeStore == "")
                    continue;

                if (sFacilityCodeStore == sFacilityCode) //相同的编号跳过
                    continue;

                sFacilityCode = sFacilityCodeStore;

                dsType = selectFacilityDataSource(sFacilityCode, out sDataSourceLocation, out sDataSource, out iStyleNumber, out sFacilityName, out sSystemName);
                if (dsType == DataSourceType.NONE)
                    continue;

                //得到参数列表和调整值
                var qPara = from dtPara in dSet.Tables["系统监控特性表"].AsEnumerable()//查询特性
                            where (dtPara.Field<int>("系统ID") == iStyleNumber) //条件
                            select dtPara;
                iParaCount = qPara.Count();

                oTemp[0] = "设备"; oTemp[1] = sFacilityName + "(" + sFacilityCode + ")"; oTemp[2] = "100";
                dtDGMmonitor.Rows.Add(oTemp);

                //得到数据源
                switch (dsType)
                {
                    #region DATABASE DATASOURCE
                    case DataSourceType.DATABASE: //数据库

                        //现在先考虑单一数据库操作
                        try
                        {
                            sqlConnM.Open();
                            sqlCommM.CommandText = sDataSource.Replace(sREPLACE, sFacilityCode);

                            if (dSetM.Tables.Contains("设备监控表")) dSetM.Tables.Remove("设备监控表");
                            sqlDAM.Fill(dSetM, "设备监控表");
                            sqlConnM.Close();

                            sValueU3D = ""; sNormalU3D = "1";

                            for (i = 0; i < dSetM.Tables["设备监控表"].Rows.Count; i++)
                            {
                                for (j = 0; j < dSetM.Tables["设备监控表"].Columns.Count; j++)
                                {
                                    if (j < iParaCount) //采用定义参数名，并调整参数
                                    {
                                        if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                            oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                                        else
                                            oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                                        oTemp[1] = (decimal.Parse(dSetM.Tables["设备监控表"].Rows[i][j].ToString()) * qPara.ElementAt(j).Field<decimal>("特性调整")).ToString();

                                        //判断取值
                                        oTemp[2] = "0";

                                        sValueU3D += oTemp[0].ToString() + ":" + oTemp[1].ToString()+"\n";
                                        sNormalU3D = "1";
                                    }
                                    else
                                    {
                                        oTemp[0] = dSetM.Tables["设备监控表"].Columns[j].ColumnName; oTemp[1] = dSetM.Tables["设备监控表"].Rows[i][j].ToString();
                                        oTemp[2] = "0";

                                        sValueU3D += oTemp[0].ToString() + ":" + oTemp[1].ToString() + "\n";
                                        sNormalU3D = "1";
                                    }

                                    //oTemp[1]="0";
                                    dtDGMmonitor.Rows.Add(oTemp);
                                }
                            }


                        }
                        catch
                        {
                            oTemp[0] = "错误";
                            oTemp[1] = "数据库读取失败"; oTemp[2] = "-100";
                            dtDGMmonitor.Rows.Add(oTemp);
                        }

                        break;
                    #endregion



                    #region OBIX DATASOURCE
                    case DataSourceType.OBIX://Niagara
                        HttpWebResponseUtility.tagUrl = sDataSourceLocation + sFacilityCode + "/";


                        //string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponse();
                        string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponseTransmit();

                        if (responseFromServer == null || responseFromServer == "")
                            break;

                        XmlDocument xmlDoc = new XmlDocument();
                        XDocument xDoc = new XDocument();
                        XmlReader xmlReader = xDoc.CreateReader();

                        xmlDoc.LoadXml(responseFromServer);
                        XElement xEle = XElement.Parse(responseFromServer);
                        xmlDoc.PreserveWhitespace = false;

                         sValueU3D = ""; sNormalU3D = "1";
                        //查询每个参数
                        for (i = 0; i < iParaCount; i++)
                        {
                            sPara = qPara.ElementAt(i).Field<string>("OBIX取值");
                            sParaName = qPara.ElementAt(i).Field<string>("OBIX特征");
                            sParaAttr = qPara.ElementAt(i).Field<string>("OBIX取值名");
                            sParaDisp = qPara.ElementAt(i).Field<string>("OBIX显示关键字");

                            var text = from t in xEle.Elements()//定位到节点 
                                        .Where(w => (w.Attribute(sParaAttr).Value.Equals(sPara) && w.Name.LocalName.Equals(sParaName)))   //若要筛选就用上这个语句 
                                       // .Where(w => (w.Attribute(sParaAttr).Value.Equals(sPara)))   //若要筛选就用上这个语句 
                                       select new
                                       {
                                           disp = t.Attribute(sParaDisp).Value   //注意此处用到 attribute 
                                       };

                            foreach (var sValue in text)
                            {
                                if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                    oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                                else
                                    oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                                switch (qPara.ElementAt(0).Field<string>("取值类型"))
                                {
                                    case "INT": //整数
                                    case "DOUBLE": //浮点双精度

                                        oTemp[1] = (decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp)) * qPara.ElementAt(i).Field<decimal>("特性调整")).ToString();

                                        //判断取值
                                        oTemp[2] = "0";
                                        if (qPara.ElementAt(i).Field<decimal?>("低阀值") != null) //有低值
                                        {
                                            //sFloor3D = qPara.ElementAt(i).Field<decimal>("低阀值").ToString();
                                            if (decimal.Parse(oTemp[1].ToString()) < qPara.ElementAt(i).Field<decimal>("低阀值"))
                                            {
                                                oTemp[2] = "-1";
                                            }
                                        }
                                        if (qPara.ElementAt(i).Field<decimal?>("高阀值") != null) //有高值
                                        {
                                            //sCeil3D = qPara.ElementAt(i).Field<decimal>("高阀值").ToString();
                                            if (decimal.Parse(oTemp[1].ToString()) > qPara.ElementAt(i).Field<decimal>("高阀值"))
                                            {
                                                oTemp[2] = "1";
                                            }
                                        }

                                        sValueU3D += oTemp[0].ToString() + ":" + oTemp[1].ToString()+"\n";
                                        if (sNormalU3D == "1")
                                        {
                                            if (oTemp[2].ToString()!="0")
                                                sNormalU3D = "0";
                                        }

                                        break;
                                    //case "DOUBLE": //浮点双精度

                                    //    break;
                                    default: //文字
                                        oTemp[1] = HttpWebResponseUtility.GetResultObix(sValue.disp);
                                        oTemp[2] = "0";

                                        sValueU3D += oTemp[0].ToString() + ":" + oTemp[1].ToString() + "\n";
                                        break;
                                }

                                //oTemp{2]="0";
                                dtDGMmonitor.Rows.Add(oTemp);
                                break; //只查找第一个参数
                            }

                        }


                        break;
                    #endregion

                    #region SOAP DATASOURCE
                    case DataSourceType.SOAP://SOAP WEB SEVICE 
                        ClassWebResponseSoap.tagURL = sDataSourceLocation;
                        ClassWebResponseSoap.SOAPAction = sDataSource;


                        string responseFromServerSoap = ClassWebResponseSoap.CreateGetHttpResponseSoap(sFacilityCode);
                        if (responseFromServerSoap == null || responseFromServerSoap == "")
                            break;

                        //XmlDocument xmlDocSoap = new XmlDocument();
                        //XDocument xDocSoap = new XDocument();
                        //XmlReader xmlReaderSoap = xDocSoap.CreateReader();
                        //xmlDocSoap.LoadXml(responseFromServerSoap);
                        //string s = xmlDocSoap.DocumentElement["env:Header"]["NotifySOAPHeader"].InnerXml;
                        //xmlDocSoap.PreserveWhitespace = false;
                        //responseFromServerSoap="<?xml version=\"1.0\" encoding=\"GB2312\" ?><getJCIInfoResult xmlns=\"http://tempuri.org/\"><sFacilityNo>CH/B2-1</sFacilityNo><dtReadingTime>2016/5/19 18:52:47</dtReadingTime></getJCIInfoResult>";

                        XElement xEleSoap = XElement.Parse("<?xml version=\"1.0\" encoding=\"GB2312\" ?>"+responseFromServerSoap);


                        sValueU3D = ""; sNormalU3D = "1";
                        var textsoap = from t in xEleSoap.Elements().Last().Elements()//定位到节点 
                                       where t.Name.LocalName.Equals("dParameter") //确定节点
                                       select t;


                        foreach (var sValue in textsoap)
                        {
                            oTemp[2] = "0";
                            foreach (var sValueget in sValue.Elements())
                            {
                                switch (sValueget.Name.LocalName)
                                {
                                    case "Para_Name":
                                        oTemp[0] = sValueget.Value;
                                        break;
                                    case "Para_Value":
                                        oTemp[1] = sValueget.Value;
                                        break;
                                    case "Para_Unit":
                                        oTemp[0] = oTemp[0] + "(" + sValueget.Value + ")";
                                        break;
                                    case "bWarn":
                                        if (sValueget.Value == "true")
                                        {
                                            oTemp[2] = "1";
                                        };
                                        break;
                                }
                            }
                            sValueU3D += oTemp[0].ToString() + ":" + oTemp[1].ToString() + "\n";
                            dtDGMmonitor.Rows.Add(oTemp);
                        }
                        break;
                    #endregion
                    default:
                        break;
                }


            }

            DGMmonitor.ItemsSource = dtDGMmonitor.AsDataView();

            if (!Window.GetWindow(DGMmonitor).IsVisible)
            {
                Window.GetWindow(DGMmonitor).Show();
            }
            DGMmonitor.UpdateLayout();

            DataGridRow dgRow;
            SolidColorBrush scbBackGround = null;
            for (i = 0; i < DGMmonitor.ItemContainerGenerator.Items.Count; i++)
            {
                dgRow = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(i);
                switch (dtDGMmonitor.Rows[i][2].ToString())
                {
                    case "1":
                        scbBackGround = new SolidColorBrush(Colors.Red);
                        break;
                    case "-1":
                        scbBackGround = new SolidColorBrush(Colors.Red);
                        break;
                    case "0":
                        scbBackGround = new SolidColorBrush(Colors.LightGreen);
                        break;
                    case "100":
                        scbBackGround = new SolidColorBrush(Colors.LightGray);
                        break;
                    case "200":
                        scbBackGround = new SolidColorBrush(Colors.Gray);
                        break;
                }
                //判断是否取到该项取值
                if (scbBackGround != null)
                    dgRow.Background = scbBackGround;
            }

            //U3d show
            //wfiU3dPlayer.Focus();
            if (sSystemName!="")
                showFeatureU3D(sSystemName, sFacilityCodeStore, sValueU3D, sNormalU3D);
            else
                showFeatureU3D("", "", "", "");
            
        }


        //判断设备的数据源
        //参数：sCode 设备编号
        //返回：数据源类型 sDataSourceLocation 数据源地址（DB，OBIX）数据源（数据库，Folder）系统编号 iStyleNumber
        private DataSourceType selectFacilityDataSource(string sCode, out string sDataSourceLocation, out string sDataSource,out int iStyleNumber,out string sFacilityName,out string sStsyemName)
        {
            string sTemp = "";
            sDataSourceLocation = "";
            sDataSource = "";
            iStyleNumber=0;
            sFacilityName = "";
            sStsyemName = "";

            DataSourceType dsType=DataSourceType.NONE;

            //得到设备类型
            var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询设备编号
                where (dtFacility.Field<string>("设备编号") == sCode)//条件
                select dtFacility;
            foreach (var itemFaclity in qFacility)//显示查询结果
            {

                sFacilityName = itemFaclity.Field<string>("设备名称");
                //得到数据源
                var qDS = from dtDS in dSet.Tables["子系统表"].AsEnumerable()//查询系统
                          where (dtDS.Field<int>("子系统ID") == itemFaclity.Field<int>("系统ID"))//条件
                    select dtDS;
                foreach (var itemDS in qDS)
                {
                    switch(itemDS.Field<string>("数据源类型").ToUpper().Trim())
                    {
                        case "DB":
                            dsType = DataSourceType.DATABASE;
                            break;
                        case "OBIX":
                            dsType = DataSourceType.OBIX;
                            break;
                        case "SOAP":
                            dsType = DataSourceType.SOAP;
                            break;
                        default:
                            return dsType;
                    }
                    sDataSourceLocation = itemDS.Field<string>("数据源地址").Trim();
                    //OBIX数据源替换站点

                    if (dsType == DataSourceType.OBIX)
                    {
                        if (itemFaclity.Field<string>("OBIX站点") != null)
                        {
                            sDataSourceLocation=sDataSourceLocation.Replace(sREPLACE, itemFaclity.Field<string>("OBIX站点"));
                        }
                    }

                    if (itemDS.Field<string>("数据源")==null)
                        sDataSource = "";
                    else
                        sDataSource = itemDS.Field<string>("数据源").Trim();

                    iStyleNumber=itemFaclity.Field<int>("系统ID");
                    sStsyemName = itemFaclity.Field<string>("设备系统");
                    break;//只取一个
                }
                break; //只取一个
            }



            return dsType;
        }

        //判断所选项目类型,
        //参数：oSelectedItem 所选项目 
         //返回：类型 sCode 编码（设备，房间）
        private nodeType selectItemType(ModelItem oSelectedItem,out string sCode)
         {
             sCode = "";

             if (oSelectedItem == null)
                 return nodeType.UNKNOW;

             //是否有设备编号,判断为设备
             DataProperty oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素", "设备编号");
             DataProperty oDP_DWGHandle1 = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("项目", "名称");

             if (oDP_DWGHandle != null)
             {
                 if (oDP_DWGHandle.Value.ToDisplayString().Trim() != "")
                 {
                     sCode = oDP_DWGHandle.Value.ToDisplayString();
                 }
                 return nodeType.FACILITY;
             }
             else //所选不为设备
             {
                 //是否为房间
                 oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("项目", "类型");
                 if (oDP_DWGHandle != null)
                 {
                     if (oDP_DWGHandle.Value.ToDisplayString().Trim() == "房间") //找到房间
                     {
                         oDP_DWGHandle1 = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素", "编号");
                         //oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素", "名称");

                         if (oDP_DWGHandle1 != null)
                         {
                             if (oDP_DWGHandle1.Value.ToDisplayString().Trim() != "")
                             {
                                 sCode = oDP_DWGHandle1.Value.ToDisplayString();
                             }
                         }
                         return nodeType.ROOM;
                     }
                 }

             }

             return nodeType.UNKNOW;
         }

        //得到现在建筑的基础模型、楼层体量模型(建筑名，基础模型，体量楼层)
        private void getModelNameAssist(string sBuildingname, out string sModelFoundation, out string sModelBodyMass)
        {
            sModelFoundation = ""; sModelBodyMass = "";

            //找到图纸
            var qBuildingModels = from dtBuildingModels in dSet.Tables["建筑表"].AsEnumerable()//查询
                                  where (dtBuildingModels.Field<string>("建筑名称") == sBuildingname)//条件
                                  select dtBuildingModels;

            foreach (var itemBuildingModels in qBuildingModels)//显示查询结果
            {
                if (itemBuildingModels.Field<string>("基础模型") != null)
                    sModelFoundation = itemBuildingModels.Field<string>("基础模型");

                if (itemBuildingModels.Field<string>("辅助模型") != null)
                    sModelBodyMass = itemBuildingModels.Field<string>("辅助模型");
            }
        }


        //刷新房间监控,房间编号
        private void refreshRoomMoniter(string sCode)
        { 
            string sFacilityCode="";
            int i,j;
            int iParaCount = 0;
            string sParaName = "", sParaAttr = "", sParaDisp = "";   

            if (sCode == null || sCode == "")
                return;

            string sDataSourceLocation, sDataSource,sFacilityName,sSystemName;
            int iStyleNumber;
            string sFloor3D="",sCeil3D="";
            string sPara = "";
            DataSourceType dsType;
            object[] oTemp = new object[3]; //3颜色控制: 0 正常 1 过高 -1 过低 100 设备名称 200 房间名称

            //得到和房间名称
            var qRoom = from dtRoom in dSet.Tables["房间表"].AsEnumerable()//查询楼层
                where (dtRoom.Field<string>("房间编号") == sCode)//条件
                select dtRoom;

            foreach (var itemRoom in qRoom)//显示查询结果
            {
                oTemp[0] = "名称"; oTemp[1] = itemRoom.Field<string>("房间名称"); oTemp[2] = "200";
                dtDGMmonitor.Rows.Add(oTemp);

                oTemp[0] = "编号"; oTemp[1] = sCode; oTemp[2] = "200";
                dtDGMmonitor.Rows.Add(oTemp);

                break;
            }

            //得到和房间相关的所有设备
            var qFacility = from dtFacility in dSet.Tables["房间设备监控表"].AsEnumerable()//查询楼层
                where (dtFacility.Field<string>("房间编号") == sCode)//条件
                select dtFacility;

            foreach (var itemFacility in qFacility)//显示查询结果
            {
                sFacilityCode=itemFacility.Field<string>("设备编号");

                dsType = selectFacilityDataSource(sFacilityCode, out sDataSourceLocation, out sDataSource, out iStyleNumber, out sFacilityName, out sSystemName);
                if (dsType == DataSourceType.NONE)
                    return;

                //得到参数列表和调整值
                var qPara = from dtPara in dSet.Tables["系统监控特性表"].AsEnumerable()//查询特性
                    where (dtPara.Field<int>("系统ID") == iStyleNumber) //条件
                    select dtPara;
                iParaCount = qPara.Count();

                oTemp[0] = "设备"; oTemp[1] = sFacilityName + "(" + sFacilityCode + ")"; oTemp[2] = "100";
                dtDGMmonitor.Rows.Add(oTemp);

                //得到数据源
                switch (dsType)
                {
                    #region DATABASE DATASOURCE
                    case DataSourceType.DATABASE: //数据库

                        //现在先考虑单一数据库操作
                        try
                        {
                            sqlConnM.Open();
                            sqlCommM.CommandText = sDataSource.Replace(sREPLACE, sFacilityCode);

                            if (dSetM.Tables.Contains("设备监控表")) dSetM.Tables.Remove("设备监控表");
                            sqlDAM.Fill(dSetM, "设备监控表");
                            sqlConnM.Close();

                            for (i = 0; i < dSetM.Tables["设备监控表"].Rows.Count; i++)
                            {
                                for (j = 0; j < dSetM.Tables["设备监控表"].Columns.Count; j++)
                                {
                                    if (j < iParaCount) //采用定义参数名，并调整参数
                                    {
                                        if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                            oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                                        else
                                            oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                                        oTemp[1] = (decimal.Parse(dSetM.Tables["设备监控表"].Rows[i][j].ToString()) * qPara.ElementAt(j).Field<decimal>("特性调整")).ToString();

                                        //判断取值
                                        oTemp[2] = 0;
                                       
                                    }
                                    else
                                    {
                                        oTemp[0] = dSetM.Tables["设备监控表"].Columns[j].ColumnName; oTemp[1] = dSetM.Tables["设备监控表"].Rows[i][j].ToString();
                                        oTemp[2] = 0;
                                    }

                                    //oTemp[1]="0";
                                    dtDGMmonitor.Rows.Add(oTemp);
                                }
                            }


                        }
                        catch
                        {
                            oTemp[0] = "错误";
                            oTemp[1] = "数据库读取失败";oTemp[2]="-100";
                            dtDGMmonitor.Rows.Add(oTemp);
                        }

                        break;
                    #endregion



                    #region OBIX DATASOURCE
                    case DataSourceType.OBIX://Niagara
                        HttpWebResponseUtility.tagUrl = sDataSourceLocation + sFacilityCode + "/";


                        //string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponse();
                        string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponseTransmit();
                        if (responseFromServer == null || responseFromServer == "")
                            break;

                        XmlDocument xmlDoc = new XmlDocument();
                        XDocument xDoc = new XDocument();
                        XmlReader xmlReader = xDoc.CreateReader();

                        xmlDoc.LoadXml(responseFromServer);
                        XElement xEle = XElement.Parse(responseFromServer);
                        xmlDoc.PreserveWhitespace = false;


                        //查询每个参数
                        for (i = 0; i < iParaCount; i++)
                        {
                            sPara = qPara.ElementAt(i).Field<string>("OBIX取值");
                            sParaName = qPara.ElementAt(i).Field<string>("OBIX特征");
                            sParaAttr = qPara.ElementAt(i).Field<string>("OBIX取值名");
                            sParaDisp = qPara.ElementAt(i).Field<string>("OBIX显示关键字");


                            var text = from t in xEle.Elements()//定位到节点 
                                        .Where(w => (w.Attribute(sParaAttr).Value.Equals(sPara) && w.Name.LocalName.Equals(sParaName)))   //若要筛选就用上这个语句 
                                       select new
                                       {
                                           disp = t.Attribute(sParaDisp).Value   //注意此处用到 attribute 
                                       };

                            foreach (var sValue in text)
                            {
                                if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                    oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                                else
                                    oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                                switch (qPara.ElementAt(0).Field<string>("取值类型"))
                                {
                                    case "INT":
                                    case "DOUBLE":
                                        oTemp[1] = (decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp)) * qPara.ElementAt(i).Field<decimal>("特性调整")).ToString();

                                        //判断取值
                                        oTemp[2] = "0";
                                        if (qPara.ElementAt(i).Field<decimal?>("低阀值") != null) //有低值
                                        {
                                            //sFloor3D = qPara.ElementAt(i).Field<decimal>("低阀值").ToString();
                                            if (decimal.Parse(oTemp[1].ToString()) < qPara.ElementAt(i).Field<decimal>("低阀值"))
                                            {
                                                oTemp[2]= "-1";
                                            }
                                        }
                                        if (qPara.ElementAt(i).Field<decimal?>("高阀值") != null) //有高值
                                        {
                                            //sCeil3D = qPara.ElementAt(i).Field<decimal>("高阀值").ToString();
                                            if (decimal.Parse(oTemp[1].ToString()) > qPara.ElementAt(i).Field<decimal>("高阀值"))
                                            {
                                                oTemp[2] = "1";
                                            }
                                        }

                                        break;
                                    default:
                                        oTemp[1] = HttpWebResponseUtility.GetResultObix(sValue.disp);
                                        oTemp[2]="0";
                                        break;
                                }

                                //oTemp{2]="0";
                                dtDGMmonitor.Rows.Add(oTemp);
                                break; //只查找第一个参数
                            }

                        }


                        break;
                    #endregion
                    default:
                        break;
                }



            }



            DGMmonitor.ItemsSource = dtDGMmonitor.AsDataView();
            
            if (!Window.GetWindow(DGMmonitor).IsVisible)
            {
                Window.GetWindow(DGMmonitor).Show();
            }
            DGMmonitor.UpdateLayout();

            DataGridRow dgRow;
            SolidColorBrush scbBackGround=null;
            for (i = 0; i < DGMmonitor.ItemContainerGenerator.Items.Count; i++)
            {
                dgRow = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(i);
                switch(dtDGMmonitor.Rows[i][2].ToString())
                {
                    case "1":
                        scbBackGround = new SolidColorBrush(Colors.Red);
                        break;
                    case "-1":
                        scbBackGround = new SolidColorBrush(Colors.Red);
                        break;
                    case "0":
                        scbBackGround = new SolidColorBrush(Colors.LightGreen);
                        break;
                    case "100":
                        scbBackGround = new SolidColorBrush(Colors.LightGray);
                        break;
                    case "200":
                        scbBackGround = new SolidColorBrush(Colors.Gray);
                        break;
                }
                //判断是否取到该项取值
                if (scbBackGround!=null)
                    dgRow.Background = scbBackGround;
            }


            //U3D
            showFeatureU3D("", "", "", "");

        }

        //刷新楼层监控,标高ID
        private void refreshLevelMoniter(int iLevelID)
        {
            string sFacilityCode = "";
            int i, j;
            int iParaCount = 0;
            string sParaName = "", sParaAttr = "", sParaDisp = ""; 

            if (iLevelID == null || iLevelID == 0)
                return;

            string sDataSourceLocation, sDataSource, sFacilityName, sSystemName;
            int iStyleNumber;
            string sFloor3D = "", sCeil3D = "";
            string sPara = "";
            DataSourceType dsType;
            object[] oTemp = new object[3]; //3颜色控制: 0 正常 1 过高 -1 过低 100 设备名称 200 房间名称

            dtDGMmonitor.Clear();
            
            //得到和房间名称
            var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                         where (dtLevel.Field<int>("标高ID") == iLevelID)//条件
                         select dtLevel;

            foreach (var itemLevel in qLevel)//显示查询结果
            {
                oTemp[0] = "名称"; oTemp[1] = itemLevel.Field<string>("标高名称"); oTemp[2] = "200";
                dtDGMmonitor.Rows.Add(oTemp);

                //oTemp[0] = "编号"; oTemp[1] = sCode; oTemp[2] = "200";
                //dtDGMmonitor.Rows.Add(oTemp);

                break;
            }

            
            //得到和楼层相关的所有设备
            var qFacility = from dtFacility in dSet.Tables["标高设备监控表"].AsEnumerable()//查询楼层
                            where (dtFacility.Field<int>("标高ID") == iLevelID)//条件
                            select dtFacility;

            foreach (var itemFacility in qFacility)//显示查询结果
            {
                sFacilityCode = itemFacility.Field<string>("设备编号");

                dsType = selectFacilityDataSource(sFacilityCode, out sDataSourceLocation, out sDataSource, out iStyleNumber, out sFacilityName, out sSystemName);
                if (dsType == DataSourceType.NONE)
                    return;

                //得到参数列表和调整值
                var qPara = from dtPara in dSet.Tables["系统监控特性表"].AsEnumerable()//查询特性
                            where (dtPara.Field<int>("系统ID") == iStyleNumber) //条件
                            select dtPara;
                iParaCount = qPara.Count();

                oTemp[0] = "设备"; oTemp[1] = sFacilityName + "(" + sFacilityCode + ")"; oTemp[2] = "100";
                dtDGMmonitor.Rows.Add(oTemp);

                //得到数据源
                switch (dsType)
                {
                    #region DATABASE DATASOURCE
                    case DataSourceType.DATABASE: //数据库

                        //现在先考虑单一数据库操作
                        try
                        {
                            sqlConnM.Open();
                            sqlCommM.CommandText = sDataSource.Replace(sREPLACE, sFacilityCode);

                            if (dSetM.Tables.Contains("设备监控表")) dSetM.Tables.Remove("设备监控表");
                            sqlDAM.Fill(dSetM, "设备监控表");
                            sqlConnM.Close();

                            for (i = 0; i < dSetM.Tables["设备监控表"].Rows.Count; i++)
                            {
                                for (j = 0; j < dSetM.Tables["设备监控表"].Columns.Count; j++)
                                {
                                    if (j < iParaCount) //采用定义参数名，并调整参数
                                    {
                                        if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                            oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                                        else
                                            oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                                        oTemp[1] = (decimal.Parse(dSetM.Tables["设备监控表"].Rows[i][j].ToString()) * qPara.ElementAt(j).Field<decimal>("特性调整")).ToString();

                                        //判断取值

                                    }
                                    else
                                    {
                                        oTemp[0] = dSetM.Tables["设备监控表"].Columns[j].ColumnName; oTemp[1] = dSetM.Tables["设备监控表"].Rows[i][j].ToString();
                                    }

                                    //oTemp[1]="0";
                                    dtDGMmonitor.Rows.Add(oTemp);
                                }
                            }


                        }
                        catch
                        {
                            oTemp[0] = "错误";
                            oTemp[1] = "数据库读取失败"; oTemp[2] = "-100";
                            dtDGMmonitor.Rows.Add(oTemp);
                        }

                        break;
                    #endregion



                    #region OBIX DATASOURCE
                    case DataSourceType.OBIX://Niagara
                        HttpWebResponseUtility.tagUrl = sDataSourceLocation + sFacilityCode + "/";


                        //string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponse();
                        string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponseTransmit();
                        if (responseFromServer == null || responseFromServer == "")
                            break;

                        XmlDocument xmlDoc = new XmlDocument();
                        XDocument xDoc = new XDocument();
                        XmlReader xmlReader = xDoc.CreateReader();

                        xmlDoc.LoadXml(responseFromServer);
                        XElement xEle = XElement.Parse(responseFromServer);
                        xmlDoc.PreserveWhitespace = false;


                        //查询每个参数
                        for (i = 0; i < iParaCount; i++)
                        {
                            sPara = qPara.ElementAt(i).Field<string>("OBIX取值");
                            sParaName = qPara.ElementAt(i).Field<string>("OBIX特征");
                            sParaAttr = qPara.ElementAt(i).Field<string>("OBIX取值名");
                            sParaDisp = qPara.ElementAt(i).Field<string>("OBIX显示关键字");


                            var text = from t in xEle.Elements()//定位到节点 
                                        .Where(w => (w.Attribute(sParaAttr).Value.Equals(sPara) && w.Name.LocalName.Equals(sParaName)))   //若要筛选就用上这个语句 
                                       select new
                                       {
                                           disp = t.Attribute(sParaDisp).Value   //注意此处用到 attribute 
                                       };

                            foreach (var sValue in text)
                            {
                                if (qPara.ElementAt(i).Field<string>("特性单位") != null)
                                    oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称") + "(" + qPara.ElementAt(i).Field<string>("特性单位") + ")";
                                else
                                    oTemp[0] = qPara.ElementAt(i).Field<string>("特性名称");

                                switch (qPara.ElementAt(0).Field<string>("取值类型"))
                                {
                                    case "INT":
                                    case "DOUBLE": //浮点双精度

                                        oTemp[1] = (decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp)) * qPara.ElementAt(i).Field<decimal>("特性调整")).ToString();

                                        //判断取值
                                        oTemp[2] = "0";
                                        if (qPara.ElementAt(i).Field<decimal?>("低阀值") != null) //有低值
                                        {
                                            //sFloor3D = qPara.ElementAt(i).Field<decimal>("低阀值").ToString();
                                            if (decimal.Parse(oTemp[1].ToString()) < qPara.ElementAt(i).Field<decimal>("低阀值"))
                                            {
                                                oTemp[2] = "-1";
                                            }
                                        }
                                        if (qPara.ElementAt(i).Field<decimal?>("高阀值") != null) //有高值
                                        {
                                            //sCeil3D = qPara.ElementAt(i).Field<decimal>("高阀值").ToString();
                                            if (decimal.Parse(oTemp[1].ToString()) > qPara.ElementAt(i).Field<decimal>("高阀值"))
                                            {
                                                oTemp[2] = "1";
                                            }
                                        }

                                        break;
                                    default:
                                        oTemp[1] = HttpWebResponseUtility.GetResultObix(sValue.disp);
                                        oTemp[2] = "0";
                                        break;
                                }

                                //oTemp{2]="0";
                                dtDGMmonitor.Rows.Add(oTemp);
                                break; //只查找第一个参数
                            }

                        }


                        break;
                    #endregion
                    default:
                        break;
                }


            }

            DGMmonitor.ItemsSource = dtDGMmonitor.AsDataView();

            if (!Window.GetWindow(DGMmonitor).IsVisible)
            {
                Window.GetWindow(DGMmonitor).Show();
            }
            DGMmonitor.UpdateLayout();

            DataGridRow dgRow;
            SolidColorBrush scbBackGround = null;
            for (i = 0; i < DGMmonitor.ItemContainerGenerator.Items.Count; i++)
            {

                dgRow = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(i);

                //if (dgRow == null)
                //{
                //    DGMmonitor.UpdateLayout();
                //    DGMmonitor.ScrollIntoView(DGMmonitor.Items[i]);
                //    dgRow = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(i);
                //}

                DataGridRow d1 = (DataGridRow)this.DGMmonitor.ItemContainerGenerator.ContainerFromIndex(13);
                switch (dtDGMmonitor.Rows[i][2].ToString())
                {
                    case "1":
                        scbBackGround = new SolidColorBrush(Colors.Red);
                        break;
                    case "-1":
                        scbBackGround = new SolidColorBrush(Colors.Red);
                        break;
                    case "0":
                        scbBackGround = new SolidColorBrush(Colors.LightGreen);
                        break;
                    case "100":
                        scbBackGround = new SolidColorBrush(Colors.LightGray);
                        break;
                    case "200":
                        scbBackGround = new SolidColorBrush(Colors.Gray);
                        break;
                }
                //判断是否取到该项取值
                if (scbBackGround != null && dgRow != null)
                    dgRow.Background = scbBackGround;
            }


            //qqq

        }

        //刷新房间环境曲线
        private void refreshRoomMoniterCurve(string sCode)
        {
            int i;

            if (RoomMoniter.Children.Count < 1)
                return;

            decimal dWD, dSD, dCO2, dPM25;
            if (getRoomMoniterPara(sMonitorRoomCode, out dWD, out dSD, out dCO2))
            {
                DateTime dtNow = DateTime.Now;
                if (listChartWD.Count > 0)
                {
                    listChartWD.RemoveAt(0);
                    ClassListEA cleaWD = new ClassListEA();
                    cleaWD.sName = dtNow.ToString();
                    cleaWD.sValue = (dWD).ToString();
                    listChartWD.Add(cleaWD);
                }

                if (listChartSD.Count > 0)
                {
                    listChartSD.RemoveAt(0);
                    ClassListEA cleaSD = new ClassListEA();
                    cleaSD.sName = dtNow.ToString();
                    cleaSD.sValue = (dSD).ToString();
                    listChartSD.Add(cleaSD);
                }

                if (listChartCO2.Count > 0)
                {
                    listChartCO2.RemoveAt(0);
                    ClassListEA cleaCO2 = new ClassListEA();
                    cleaCO2.sName = dtNow.ToString();
                    cleaCO2.sValue = (dCO2).ToString();
                    listChartCO2.Add(cleaCO2);
                }

                //ii++;

            }

            if (getRoomMoniterParaPM25(sMonitorRoomCode, out dPM25))
            {
                DateTime dtNow = DateTime.Now;
                if (listChartPM25.Count > 0)
                {
                    listChartPM25.RemoveAt(0);
                    ClassListEA cleaPM25 = new ClassListEA();
                    cleaPM25.sName = dtNow.ToString();
                    cleaPM25.sValue = (dPM25).ToString();
                    listChartPM25.Add(cleaPM25);
                }

            }



            Grid G_WD = RoomMoniter.Children[0] as Grid;
            Chart chartSplineWD = G_WD.Children[0] as Chart;
            Grid G_SD = RoomMoniter.Children[1] as Grid;
            Chart chartSplineSD = G_SD.Children[0] as Chart;
            Grid G_CO2 = RoomMoniter.Children[2] as Grid;
            Chart chartSplineCO2 = G_CO2.Children[0] as Chart;
            Grid G_PM25 = RoomMoniter.Children[3] as Grid;
            Chart chartSplinePM25 = G_PM25.Children[0] as Chart;
            //首次创建，创建一个标题的对象
            if (bMonitorRoomFirst)
            {
                Title titlechartSpline = chartSplineWD.Titles[0];
                //设置标题的名称
                titlechartSpline.Text = sMonitorRoomName+":温度";
            }


            //初始化一个新的Axis
            Axis xaxisSplineWD = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplineWD.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplineWD.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplineWD.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis         
            chartSplineWD.AxesX.Clear();
            chartSplineWD.AxesX.Add(xaxisSplineWD);

            DataPoint dataPointSplineWD;
            DataSeries dataSeriesSplineWD = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplineWD.LegendText = "";

            dataSeriesSplineWD.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplineWD.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartWD.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplineWD = new DataPoint();
                // 设置X轴点                    
                dataPointSplineWD.XValue = DateTime.Parse(listChartWD[i].sName);
                //设置Y轴点                   
                dataPointSplineWD.YValue = double.Parse(listChartWD[i].sValue);
                dataPointSplineWD.MarkerSize = listChartWD.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplineWD.Color = new SolidColorBrush(Colors.LightGray);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplineWD.DataPoints.Add(dataPointSplineWD);
            }

            // 添加数据线到数据序列。 
            dataSeriesSplineWD.Color = new SolidColorBrush(Colors.LightPink);
            chartSplineWD.Series.Clear();
            chartSplineWD.Series.Add(dataSeriesSplineWD);


            //初始化一个新的Axis
            Axis xaxisSplineSD = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplineSD.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplineSD.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplineSD.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis         
            chartSplineSD.AxesX.Clear();
            chartSplineSD.AxesX.Add(xaxisSplineSD);

            DataPoint dataPointSplineSD;
            DataSeries dataSeriesSplineSD = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplineSD.LegendText = "";

            dataSeriesSplineSD.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplineSD.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartSD.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplineSD = new DataPoint();
                // 设置X轴点                    
                dataPointSplineSD.XValue = DateTime.Parse(listChartSD[i].sName);
                //设置Y轴点                   
                dataPointSplineSD.YValue = double.Parse(listChartSD[i].sValue);
                dataPointSplineSD.MarkerSize = listChartSD.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplineSD.Color = new SolidColorBrush(Colors.LightPink);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplineSD.DataPoints.Add(dataPointSplineSD);
            }

            // 添加数据线到数据序列。 
            dataSeriesSplineSD.Color = new SolidColorBrush(Colors.Orange);
            chartSplineSD.Series.Clear();
            chartSplineSD.Series.Add(dataSeriesSplineSD);

            //初始化一个新的Axis
            Axis xaxisSplineCO2 = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplineCO2.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplineCO2.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplineCO2.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis         
            chartSplineCO2.AxesX.Clear();
            chartSplineCO2.AxesX.Add(xaxisSplineCO2);


            DataPoint dataPointSplineCO2;
            DataSeries dataSeriesSplineCO2 = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplineCO2.LegendText = "";

            dataSeriesSplineCO2.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplineCO2.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartCO2.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplineCO2 = new DataPoint();
                // 设置X轴点                    
                dataPointSplineCO2.XValue = DateTime.Parse(listChartCO2[i].sName);
                //设置Y轴点                   
                dataPointSplineCO2.YValue = double.Parse(listChartCO2[i].sValue);
                dataPointSplineCO2.MarkerSize = listChartCO2.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplineCO2.Color = new SolidColorBrush(Colors.LightGray);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplineCO2.DataPoints.Add(dataPointSplineCO2);
            }

            // 添加数据线到数据序列。 
            dataSeriesSplineCO2.Color = new SolidColorBrush(Colors.LightBlue);
            chartSplineCO2.Series.Clear();
            chartSplineCO2.Series.Add(dataSeriesSplineCO2);

            //初始化一个新的Axis
            Axis xaxisSplinePM25 = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplinePM25.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplinePM25.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplinePM25.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis         
            chartSplinePM25.AxesX.Clear();
            chartSplinePM25.AxesX.Add(xaxisSplinePM25);

            DataPoint dataPointSplinePM25;
            DataSeries dataSeriesSplinePM25 = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplinePM25.LegendText = "";

            dataSeriesSplinePM25.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplinePM25.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartPM25.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplinePM25 = new DataPoint();
                // 设置X轴点                    
                dataPointSplinePM25.XValue = DateTime.Parse(listChartPM25[i].sName);
                //设置Y轴点                   
                dataPointSplinePM25.YValue = double.Parse(listChartPM25[i].sValue);
                dataPointSplinePM25.MarkerSize = listChartPM25.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplinePM25.Color = new SolidColorBrush(Colors.LightGray);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplinePM25.DataPoints.Add(dataPointSplinePM25);
            }

            // 添加数据线到数据序列。 
            dataSeriesSplinePM25.Color = new SolidColorBrush(Colors.Blue);
            chartSplinePM25.Series.Clear();
            chartSplinePM25.Series.Add(dataSeriesSplinePM25);


        }



   


        //////////////////////////////////
        //private void changeCameraViewDir_Way2()
        //{

        //    Document oDoc =
        //             documentControlM.Document;

        //    // make a copy of current viewpoint
        //    Viewpoint oCurrVCopy = oDoc.CurrentViewpoint.CreateCopy();

        //    // Focal Distance
        //    double oFocal =
        //        oCurrVCopy.FocalDistance;

        //    // new target is the center of the model
        //    Point3D oNewTarget = oDoc.Models[0].RootItem.BoundingBox().Center;

        //    //new direction is X- >> RIGHT
        //    Vector3D oNewViewDir = new Vector3D(-1, 0, 0);
        //    //new direction is X >> LEFT
        //    //Vector3D oNewViewDir = new Vector3D(1, 0, 0);

        //    //calculate the new position by the target and focal distance
        //    Point3D oNewPos = new Point3D(oNewTarget.X - oNewViewDir.X * oFocal,
        //        oNewTarget.Y - oNewViewDir.Y * oFocal,
        //        oNewTarget.Z - oNewViewDir.Z * oFocal);

        //    //set the position
        //    oCurrVCopy.Position = oNewPos;
        //    //set the target
        //    oCurrVCopy.PointAt(oNewTarget);
        //    //set view direction
        //    oCurrVCopy.AlignDirection(oNewViewDir);
        //    // set which direction is up: in this case it is Z+
        //    oCurrVCopy.AlignUp(new Vector3D(0, 0, 1));

        //    // update current viewpoint
        //    oDoc.CurrentViewpoint.CopyFrom(oCurrVCopy);
        //}

        //private void moveCameraAlongViewDir()
        //{
        //    Document oDoc =
        //        Autodesk.Navisworks.Api.Application.ActiveDocument;
        //    // make a copy of current viewpoint
        //    Viewpoint oCurrVCopy = oDoc.CurrentViewpoint.CreateCopy();
        //    // get view direction
        //    Vector3D oViewDir = getViewDir(oCurrVCopy);
        //    //step to move
        //    double step = 2;
        //    //  create the new position
        //    Point3D newPos =
        //        new Point3D(oCurrVCopy.Position.X + oViewDir.X * step,
        //                    oCurrVCopy.Position.Y + oViewDir.Y * step,
        //                    oCurrVCopy.Position.Z + oViewDir.Z * step);
        //    oCurrVCopy.Position = newPos;

        //    // update current viewpoint
        //    oDoc.CurrentViewpoint.CopyFrom(oCurrVCopy);

        //}

        private Vector3D getViewDir(Viewpoint oVP)
        {
            Rotation3D oRot = oVP.Rotation;

            // calculate view direction
            Rotation3D oNegtiveZ =
                new Rotation3D(0, 0, -1, 0);

            Rotation3D otempRot =
                MultiplyRotation3D(oNegtiveZ, oRot.Invert());

            Rotation3D oViewDirRot =
                MultiplyRotation3D(oRot, otempRot);

            // get view direction
            Vector3D oViewDir =
                new Vector3D(oViewDirRot.A,
                            oViewDirRot.B,
                            oViewDirRot.C);

            oViewDir.Normalize();

            return new Vector3D(oViewDir.X,
                        oViewDir.Y,
                        oViewDir.Z);

        }

        private Rotation3D MultiplyRotation3D(Rotation3D r2, Rotation3D r1)
        {

            Rotation3D oRot =
                new Rotation3D(r2.D * r1.A + r2.A * r1.D +
                                    r2.B * r1.C - r2.C * r1.B,
                                r2.D * r1.B + r2.B * r1.D +
                                    r2.C * r1.A - r2.A * r1.C,
                                r2.D * r1.C + r2.C * r1.D +
                                    r2.A * r1.B - r2.B * r1.A,
                                r2.D * r1.D - r2.A * r1.A -
                                    r2.B * r1.B - r2.C * r1.C);

            oRot.Normalize();

            return oRot;

        }
        
        //旋转视图
        private void ButtonVPMove_Click(object sender, RoutedEventArgs e)
        {
            //dumpViewPoint();
            //updateUIAnimation();

            //if (svpacAnimationMain == null)
            //    return;
            Fluent.Button fb = sender as Fluent.Button;

            timerAnimation.Interval = 100;
            //iAnimationMain = 0;

            if (timerAnimation.Enabled)
            {
                timerAnimation.Stop();
                fb.Foreground = new SolidColorBrush(Colors.Black);
            }
            else
            {
                timerAnimation.Start();
                fb.Foreground = new SolidColorBrush(Colors.LightGray);
            }

        }

        private void OnTimedEventAnimation(object sender, EventArgs e)
        {
            this.Dispatcher.Invoke(DispatcherPriority.Normal,
                new TimerDispatcherDelegate(updateUIAnimation));
        }

        //动态视图：旋转
        private void updateUIAnimation()
        {
            /*
            int i;

            if (iAnimationMain >= svpacAnimationMain.Children.Count)
                iAnimationMain = 0;

            SavedItem oSItem = svpacAnimationMain.Children[iAnimationMain];
            iAnimationMain++;

            if (oSItem == null)
                return;
            
            //if (oSItem is SavedViewpointAnimationCut)
            //{
            //    SavedViewpointAnimationCut s = oSItem as SavedViewpointAnimationCut;
            //}
            
            documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = oSItem;
            */
            double deta = 720;

            //iAnimationMain++;
            //if (iAnimationMain >= 360) //360等分
            //    iAnimationMain = 0;

            //if(iAnimationMain==0)
            //    return;

            try
            {
                Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();

                Point3D oPos = oCurrVCopy.Position;
                Vector3D oViewDir = getViewDir(oCurrVCopy);
                double oFocal = oCurrVCopy.FocalDistance;
                //得到目标点
                Point3D oTarget = new Point3D(oPos.X + oViewDir.X * oFocal, oPos.Y + oViewDir.Y * oFocal, oPos.Z + oViewDir.Z * oFocal);

                //旋转相机
                //  设置需要旋转的轴为（Ｚ：０,０,１）
                UnitVector3D odeltaA = new UnitVector3D(0, 0, 1);
                // 设置旋转的增量四元素: 轴是 Z, 角度45度
                Rotation3D delta = new Rotation3D(odeltaA, PI / deta);
                //以原四元素乘以增量得到新四元素
                oCurrVCopy.Rotation = Multiply(oCurrVCopy.Rotation, delta);
                oViewDir = getViewDir(oCurrVCopy);

                //计算新的相机位置
                Point3D newPos = new Point3D(oTarget.X - oViewDir.X * oFocal, oTarget.Y - oViewDir.Y * oFocal, oTarget.Z - oViewDir.Z * oFocal);

                oCurrVCopy.Position = newPos;
                documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);
            }
            catch
            {
            }

        }

        public static Rotation3D Multiply(Rotation3D r1, Rotation3D r2)
        {
            Rotation3D res = new Rotation3D(r2.D * r1.A + r2.A * r1.D + r2.B * r1.C - r2.C * r1.B,
                                            r2.D * r1.B + r2.B * r1.D + r2.C * r1.A - r2.A * r1.C,
                                            r2.D * r1.C + r2.C * r1.D + r2.A * r1.B - r2.B * r1.A,
                                            r2.D * r1.D - r2.A * r1.A - r2.B * r1.B - r2.C * r1.C);

            return res;
        }

        /// <summary>
        /// ///////////////////////////////////////////////////////
        /// 测试读取函数,临时
        /// //////////////////////////////////////////////////////
        /// </summary>
        private void dumpViewPoint()
        {
            Document oDoc =
                Autodesk.Navisworks.Api.Application.ActiveDocument;
            // get current viewpoint
            Viewpoint oCurVP = oDoc.CurrentViewpoint;
 
            // common properties
            Debug.Print("Far Plane Distance: " +
                    oCurVP.FarPlaneDistance);
            Debug.Print("Near Plane Distance: " +
                    oCurVP.NearPlaneDistance);
            Debug.Print("Projection: " +
                    oCurVP.Projection);
            Debug.Print("Aspect Ratio: " +
                    oCurVP.AspectRatio);
 
 
            if (oCurVP.HasLinearSpeed)
                Debug.Print("Linear Speed: " +
                    oCurVP.LinearSpeed);
            else
                Debug.Print("Linear Speed: <None>");
            if (oCurVP.HasAngularSpeed)
                Debug.Print("Angular Speed: " +
                    oCurVP.AngularSpeed);
            else
                Debug.Print("Angular Speed: <None>");
            if (oCurVP.HasLighting)
                Debug.Print("Lighting: " +
                    oCurVP.Lighting);
            else
                Debug.Print("Lighting: <None>");
 
            if (oCurVP.HasRenderStyle)
                Debug.Print("RenderStyle: " +
                    oCurVP.RenderStyle);
            else
                Debug.Print("RenderStyle:<None> " +
                    oCurVP.RenderStyle);
 
            if (oCurVP.HasTool)
                Debug.Print("Tool: " +
                    oCurVP.Tool);
            else
                Debug.Print("Tool:<None> " +
                    oCurVP.Tool);
 
 
            // camera properties
            double oFocal =
                oCurVP.FocalDistance;
            Debug.Print("Focal Distance: " +
                oFocal);
            // Rotation
            Rotation3D oRot = oCurVP.Rotation;
            Debug.Print("Quaternion: <A (x),B(y),C(z),D(w)> = <{0},{1},{2},{3}>",oRot.A,oRot.B,oRot.C,oRot.D);
 
            // calculate view direction
            Rotation3D oNegtiveZ =
                new Rotation3D(0, 0, -1, 0);
            Rotation3D otempRot =
                MultiplyRotation3D(oNegtiveZ, oRot.Invert());
            Rotation3D oViewDirRot =
                MultiplyRotation3D(oRot, otempRot);
            // get view direction
            Vector3D oViewDir =
                new Vector3D(oViewDirRot.A,
                            oViewDirRot.B,
                            oViewDirRot.C);
            oViewDir.Normalize();
            Debug.Print("View Direction:<{0},{1},{2}> ",
                        oViewDir.X,
                        oViewDir.Y,
                        oViewDir.Z);
            // position
            Point3D oPos = oCurVP.Position;
            Debug.Print("Position:<{0},{1},{2}> ",
                        oPos.X,
                        oPos.Y,
                        oPos.Z);
            // target
            Point3D oTarget =
                new Point3D(oPos.X + oViewDir.X * oFocal,
                            oPos.Y + oViewDir.Y * oFocal,
                            oPos.Z + oViewDir.Z * oFocal);
            Debug.Print("Target:<{0},{1},{2}> ",
                            oTarget.X,
                            oTarget.Y,
                            oTarget.Z);
 
            // rotation information
            AxisAndAngleResult oAR =
                oRot.ToAxisAndAngle();
            Debug.Print("Rotation Axis:<{0},{1},{2}> ",
                        oAR.Axis.X,
                        oAR.Axis.Y,
                        oAR.Axis.Z);
 
            Debug.Print("Rotation Angle: " +
                oAR.Angle);




 
        }


        private void ButtonFeature_Click(object sender, RoutedEventArgs e)
        {

            //ApplicationControl.SelectionBehavior = SelectionBehavior.FirstObject;
            TimerStop();

            viewControl.Focus();
            //set the active View (and the ActiveDocument to DocumentControl.Document)
            viewControl.SetActiveView();


            //viewControl.DocumentControl.Document.Tool.Value = Autodesk.Navisworks.Api.Tool.NavigateFreeOrbit; 

            ToolPluginRecord toolPluginRecord =
                      (ToolPluginRecord)Autodesk.Navisworks.Api.Application.Plugins.FindPlugin("ToolPluginSelect.BIADBIM");
            viewControl.DocumentControl.Document.Tool.SetCustomToolPlugin(toolPluginRecord.LoadPlugin());

            uiStyle.sSelectStyle = ((Fluent.Button)sender).Header.ToString();
        }

        private void ButtonBIMDB_Click(object sender, RoutedEventArgs e)
        {
            WindowDatabase wDatabase = new WindowDatabase();
            wDatabase.intMode = 0;
            wDatabase.ShowDialog();
        }

        private void ButtonMonitorDB_Click(object sender, RoutedEventArgs e)
        {
            WindowDatabase wDatabase = new WindowDatabase();
            wDatabase.intMode = 1;
            wDatabase.ShowDialog();
        }

        private void ButtonAnalysisDB_Click(object sender, RoutedEventArgs e)
        {
            WindowDatabase wDatabase = new WindowDatabase();
            wDatabase.intMode = 2;
            wDatabase.ShowDialog();
        }



        private void ButtonMonitor_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            if (timerMonitor.Enabled)
            {
                timerMonitor.Stop();
                fb.Foreground = new SolidColorBrush(Colors.Black);
            }
            else
            {
                timerMonitor.Start();
                fb.Foreground = new SolidColorBrush(Colors.LightGray);
            }

            //timerMonitor.Start();
        }

        private void MainWindowNavisWorks_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            timerMonitor.Stop();
            TimerStop();
            //timerAnimation.Stop();
        }

        //停止所有时钟动画
        private void TimerStop()
        {
            timerAnimation.Stop();
            ButtonVPMove.Foreground = new SolidColorBrush(Colors.Black);

            timerRoomMonitor.Stop();
            ButtonRoomMonitor.Foreground = new SolidColorBrush(Colors.Black);
        }

        private void ButtonHidden_Click(object sender, RoutedEventArgs e)
        {
            ModelItemCollection hidden = new ModelItemCollection();
            hidden.AddRange(documentControlM.Document.CurrentSelection.SelectedItems);
            documentControlM.Document.Models.SetHidden(hidden, true);
        }

        private void ButtonUnHidden_Click(object sender, RoutedEventArgs e)
        {
            // 需要隐藏的集合
            ModelItemCollection hidden = new ModelItemCollection();
            // 可见的集合
            ModelItemCollection visible = new ModelItemCollection();

            foreach (ModelItem item in documentControlM.Document.CurrentSelection.SelectedItems)
            {
                //该条目的祖先				
                if (item.AncestorsAndSelf != null)
                    visible.AddRange(item.AncestorsAndSelf);
                //该条目的子孙
                if (item.Descendants != null)
                    visible.AddRange(item.Descendants);
            }

            //mark as invisible all the siblings of the visible items as well as the visible items
            foreach (ModelItem toShow in visible)
            {
                if (toShow.Parent != null)
                {
                    hidden.AddRange(toShow.Parent.Children);
                }
            }
            //remove the visible items from the collection
            foreach (ModelItem toShow in visible)
            {
                hidden.Remove(toShow);
            }
            //hide the remaining items
            documentControlM.Document.Models.SetHidden(hidden, true);


        }

        private void ButtonShowAll_Click(object sender, RoutedEventArgs e)
        {
            IEnumerable<ModelItem> items = Autodesk.Navisworks.Api.Application.ActiveDocument.Models.RootItemDescendantsAndSelf.
    Where(x => x.IsHidden == true);
            //Application.ActiveDocument.CurrentSelection.CopyFrom(items);

            ModelItemCollection hidden = new ModelItemCollection();
            hidden.AddRange(items);
            documentControlM.Document.Models.SetHidden(hidden, false);

            //隐藏辅助模型
            Search search = new Search();
            VariantData oData;

            search.Clear();
            search.Selection.SelectAll();
            var qBuildingModels = from dtBuildingModels in dSet.Tables["建筑表"].AsEnumerable()//查询
                              select dtBuildingModels;
            foreach (var itemBuildingModels in qBuildingModels)//显示查询结果
            {
                if (itemBuildingModels.Field<string>("辅助模型") != null)
                {
                    if (itemBuildingModels.Field<string>("辅助模型") != "")
                    {
                        oData = VariantData.FromDisplayString(itemBuildingModels.Field<string>("辅助模型"));

                        SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                        oSearchCondition = oSearchCondition.EqualValue(oData);
                        search.SearchConditions.Add(oSearchCondition);
                    }
                }
            }

            ModelItemCollection items1;
            items1 = search.FindAll(documentControlM.Document, false);
            //显示辅助
            documentControlM.Document.Models.SetHidden(items1.DescendantsAndSelf, true);


        }

        private void ButtonVPRoom_Click(object sender, RoutedEventArgs e)
        {
            //TimerStop();
            // 需要隐藏的集合
            ModelItemCollection hidden = new ModelItemCollection();
            // 可见的集合
            ModelItemCollection visible = new ModelItemCollection();

            Search search = new Search();
            search.Selection.SelectAll();
            VariantData oData = VariantData.FromDisplayString("房间");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection itemsRoom = search.FindAll(documentControlM.Document, false);

            foreach (ModelItem item in itemsRoom)
            {
                //该条目的祖先				
                if (item.AncestorsAndSelf != null)
                    visible.AddRange(item.AncestorsAndSelf);
                //该条目的子孙
                if (item.Descendants != null)
                    visible.AddRange(item.Descendants);
            }

            //mark as invisible all the siblings of the visible items as well as the visible items
            foreach (ModelItem toShow in visible)
            {
                if (toShow.Parent != null)
                {
                    hidden.AddRange(toShow.Parent.Children);
                }
            }
            //remove the visible items from the collection
            foreach (ModelItem toShow in visible)
            {
                hidden.Remove(toShow);
            }
            //hide the remaining items
            documentControlM.Document.Models.SetHidden(hidden, true);
        }

        //显示所选择的标高层，避免遮挡视线
        private void showSelectedLevel(string sLevelModelAS,string sLevelModelMEP,double fLevelOrder,int iBuildings,bool bShowLevelDown,bool bHidden)
        {
            string sTemp;
            double dTemp = 0;

            // 需要隐藏的集合
            ModelItemCollection hidden = new ModelItemCollection();
            // 可见的集合
            ModelItemCollection visible = new ModelItemCollection();

            if (iBuildings == 0)
                return;

            if (sLevelModelAS == "" && sLevelModelMEP == "")
                return;

            //找到所有文件
            Search search = new Search();
            search.Selection.SelectAll();
            VariantData oData = VariantData.FromDisplayString("文件");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection itemsFiles = search.FindAll(documentControlM.Document, false);

            foreach (ModelItem item in itemsFiles)
            {
                sTemp = ""; 
                sTemp=item.PropertyCategories.FindPropertyByDisplayName("项目", "名称").Value.ToDisplayString();
                if (sTemp == "")
                    continue;

                if (sTemp.ToLower() == sLevelModelAS.ToLower() || sTemp.ToLower() == sLevelModelMEP.ToLower()) //显示本层
                {
                    visible.Add(item);
                }
                else //不是本层
                {
                    if (!bShowLevelDown) //不显示以下层
                    {
                        hidden.Add(item);
                        continue;
                    }
                    else //显示以下层
                    {
                        var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                                     where (dtLevel.Field<int>("建筑ID") == iBuildings) && ((dtLevel.Field<string>("土建模型组成") == sTemp) || (dtLevel.Field<string>("机电模型组成") == sTemp))//条件
                                 select dtLevel;
                        
                        foreach (var itemLevel in qLevel)//显示查询结果
                        {
                            dTemp = itemLevel.Field<double>("标高排序");
                            if (dTemp > fLevelOrder)
                            {
                                hidden.Add(item);
                            }
                            else
                            {
                                visible.Add(item);
                            }
                            break;
                        }
                    }
                }

                

            }
            documentControlM.Document.Models.SetHidden(hidden, true);
            documentControlM.Document.Models.SetHidden(visible, false);

        }

        //显示所选建筑
        private void showSelectedBuilding(int iBuildings, bool bHidden)
        {
            string sTemp;

            // 需要隐藏的集合
            ModelItemCollection hidden = new ModelItemCollection();
            // 可见的集合
            ModelItemCollection visible = new ModelItemCollection();

            if (iBuildings == 0)
                return;


            //找到所有文件
            Search search = new Search();
            search.Selection.SelectAll();
            VariantData oData = VariantData.FromDisplayString("文件");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection itemsFiles = search.FindAll(documentControlM.Document, false);

            foreach (ModelItem item in itemsFiles)
            {

                sTemp = "";
                sTemp = item.PropertyCategories.FindPropertyByDisplayName("项目", "名称").Value.ToDisplayString();
                if (sTemp == "")
                    continue;

                var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                             where (dtLevel.Field<int>("建筑ID") == iBuildings) && ((dtLevel.Field<string>("土建模型组成") == sTemp) || (dtLevel.Field<string>("机电模型组成") == sTemp))//条件
                             select dtLevel;

                if (qLevel.Count() < 1)
                {
                    hidden.Add(item);
                    continue;
                }

                foreach (var itemLevel in qLevel)//显示查询结果
                {
                    visible.Add(item);
                }

            }
            documentControlM.Document.Models.SetHidden(hidden, true);
            documentControlM.Document.Models.SetHidden(visible, false);
        }

        //显示项目
        private void showProject(bool bHidden)
        {
            string sTemp;

            // 需要隐藏的集合
            ModelItemCollection hidden = new ModelItemCollection();
            // 可见的集合
            ModelItemCollection visible = new ModelItemCollection();


            //找到所有文件
            Search search = new Search();
            search.Selection.SelectAll();
            VariantData oData = VariantData.FromDisplayString("文件");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection itemsFiles = search.FindAll(documentControlM.Document, false);

            foreach (ModelItem item in itemsFiles)
            {

                sTemp = "";
                sTemp = item.PropertyCategories.FindPropertyByDisplayName("项目", "名称").Value.ToDisplayString();
                if (sTemp == "")
                    continue;

                var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                             where ((dtLevel.Field<string>("土建模型组成") == sTemp) || (dtLevel.Field<string>("机电模型组成") == sTemp))//条件
                             select dtLevel;

                if (qLevel.Count() < 1)
                {
                    hidden.Add(item);
                    continue;
                }

                foreach (var itemLevel in qLevel)//显示查询结果
                {
                    visible.Add(item);
                }

            }
            documentControlM.Document.Models.SetHidden(hidden, true);
            documentControlM.Document.Models.SetHidden(visible, false);

            //调缺省视图
            SavedItem siView = getViewPoint(sViewName);
            if (siView == null)
                return;
            documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;


        }

        //显示静态特性
        public void showFeature()
        {
            if (documentControlM.Document.CurrentSelection.IsEmpty)
                return;

            //MessageBox.Show("1111");

            foreach (ModelItem item in documentControlM.Document.CurrentSelection.SelectedItems)
            {
                //StatusBarItemSelect.Content = item.PropertyCategories.FindPropertyByDisplayName("项目", "类型").Value.ToDisplayString() + "：" + item.PropertyCategories.FindPropertyByDisplayName("项目", "名称").Value.ToDisplayString(); ;
                DataProperty oDP = item.PropertyCategories.FindPropertyByDisplayName("项目", "类型");
                DataProperty oDP1 = item.PropertyCategories.FindPropertyByDisplayName("项目", "名称");
                if (oDP != null && oDP1 != null)
                    StatusBarItemSelect.Content = oDP.Value.ToDisplayString() + " -> " + oDP1.Value.ToDisplayString();
                else
                    if (oDP != null)
                        StatusBarItemSelect.Content = oDP.Value.ToDisplayString();
                    else
                        StatusBarItemSelect.Content = "";
                break; //第一个



            }


            int i, j;
            string sFCode = ""; //设备编码
            bool bfirstSelect = true;

            dtDGMproperty.Clear();

            //是否有设备编号
            //ModelItem oSelectedItem = documentControlM.Document.CurrentSelection.SelectedItems.ElementAt<ModelItem>(0);
            foreach (ModelItem oSelectedItem in documentControlM.Document.CurrentSelection.SelectedItems)
            {
                DataProperty oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素", "设备编号");
                DataProperty oDP_DWGHandle1 = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("项目", "名称");
                DataProperty oDP_DWGHandle2 = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素", "设备铭牌号");

                sFCode = "";
                if (oDP_DWGHandle != null)
                {
                    sFCode = oDP_DWGHandle.Value.ToDisplayString().Trim();
                }
                

                if (oDP_DWGHandle2 != null)
                {
                    if (oDP_DWGHandle2.Value.ToDisplayString().Trim() != "")
                    {

                        //MessageBox.Show(oDP_DWGHandle.Value.ToDisplayString());
                        //初始化DGMproperty
                        sqlConn.Open();

                        sqlComm.CommandText = "SELECT 类型, 厂商信息, 采购日期, 维护日期, 维护公司, 维护电话, 支持连接 FROM 设备特性表 WHERE (设备铭牌号 = N'" + oDP_DWGHandle2.Value.ToDisplayString() + "')";
                        if (dSet.Tables.Contains("设备特性表")) dSet.Tables.Remove("设备特性表");
                        sqlDA.Fill(dSet, "设备特性表");

                        sqlConn.Close();



                        //DGMproperty
                        object[] oTemp = new object[3];
                        oTemp[0] = "设备名称"; oTemp[1] = ""; oTemp[2] = "100";
                        if (oDP_DWGHandle1 != null)
                        {
                            oTemp[1] = oDP_DWGHandle1.Value.ToDisplayString();
                        }
                        dtDGMproperty.Rows.Add(oTemp);

                        oTemp[0] = "控制编号"; oTemp[1] = sFCode; oTemp[2] = "0";
                        dtDGMproperty.Rows.Add(oTemp);

                        oTemp[0] = "设备铭牌号"; oTemp[1] = oDP_DWGHandle2.Value.ToDisplayString(); oTemp[2] = "0";
                        dtDGMproperty.Rows.Add(oTemp);


                        for (i = 0; i < dSet.Tables["设备特性表"].Rows.Count; i++)
                        {
                            for (j = 0; j < dSet.Tables["设备特性表"].Columns.Count; j++)
                            {
                                oTemp[0] = dSet.Tables["设备特性表"].Columns[j].ColumnName; oTemp[1] = dSet.Tables["设备特性表"].Rows[i][j].ToString(); oTemp[2] = "1";
                                dtDGMproperty.Rows.Add(oTemp);
                            }
                        }

                        //反选选择树
                        if (uiStyle.bSelectStyle && bfirstSelect) 
                        {
                            selectTreeNode(sFCode, nodeType.FACILITY);
                        }
                    }


                }
                else //所选不为设备
                {
                    //是否有设备编号
                    oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("项目", "类型");
                    if (oDP_DWGHandle != null)
                    {
                        if (oDP_DWGHandle.Value.ToDisplayString().Trim() == "房间") //找到房间
                        {
                            oDP_DWGHandle1 = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素", "编号");
                            oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素", "名称");

                            if (oDP_DWGHandle1 != null)
                            {
                                if (oDP_DWGHandle1.Value.ToDisplayString().Trim() != "")
                                {
                                    //MessageBox.Show(oDP_DWGHandle.Value.ToDisplayString());
                                    //初始化DGMproperty
                                    sqlConn.Open();

                                    sqlComm.CommandText = "SELECT 所属单位, 责任人, 管理单位, 责任电话 FROM 房间特性表 WHERE (房间编号 = N'" + oDP_DWGHandle1.Value.ToDisplayString() + "')";
                                    if (dSet.Tables.Contains("房间特性表")) dSet.Tables.Remove("房间特性表");
                                    sqlDA.Fill(dSet, "房间特性表");

                                    sqlConn.Close();



                                    //DGMproperty
                                    object[] oTemp = new object[3];
                                    oTemp[0] = "房间名称"; oTemp[1] = ""; oTemp[2] = "100";
                                    if (oDP_DWGHandle != null)
                                    {
                                        oTemp[1] = oDP_DWGHandle.Value.ToDisplayString();
                                    }

                                    dtDGMproperty.Rows.Add(oTemp);
                                    oTemp[0] = "编号"; oTemp[1] = oDP_DWGHandle1.Value.ToDisplayString(); oTemp[2] = "0";
                                    dtDGMproperty.Rows.Add(oTemp);
                                    for (i = 0; i < dSet.Tables["房间特性表"].Rows.Count; i++)
                                    {
                                        for (j = 0; j < dSet.Tables["房间特性表"].Columns.Count; j++)
                                        {
                                            oTemp[0] = dSet.Tables["房间特性表"].Columns[j].ColumnName; oTemp[1] = dSet.Tables["房间特性表"].Rows[i][j].ToString(); oTemp[2] = "1";
                                            dtDGMproperty.Rows.Add(oTemp);
                                        }
                                    }

                                    //反选选择树
                                    if (uiStyle.bSelectStyle && bfirstSelect)
                                    {
                                        selectTreeNode(oDP_DWGHandle1.Value.ToDisplayString(), nodeType.ROOM);
                                    }

                                }
                            }
                        }
                    }

                }

                //判断是否为第一个
                if (bfirstSelect)
                    bfirstSelect = false;
            }

            DGMproperty.ItemsSource = dtDGMproperty.AsDataView();

            if (!Window.GetWindow(DGMproperty).IsVisible)
            {
                Window.GetWindow(DGMproperty).Show();
            }
            DGMproperty.UpdateLayout();

            DataGridRow dgRow;
            SolidColorBrush scbBackGround = null;
            for (i = 0; i < DGMproperty.ItemContainerGenerator.Items.Count; i++)
            {
                dgRow = (DataGridRow)this.DGMproperty.ItemContainerGenerator.ContainerFromIndex(i);
                switch (dtDGMproperty.Rows[i][2].ToString())
                {
                    case "1":
                        scbBackGround = new SolidColorBrush(Colors.LightSeaGreen);
                        break;
                    case "-1":
                        scbBackGround = new SolidColorBrush(Colors.Red);
                        break;
                    case "0":
                        scbBackGround = new SolidColorBrush(Colors.LightGreen);
                        break;
                    case "100":
                        scbBackGround = new SolidColorBrush(Colors.LightGray);
                        break;
                    case "200":
                        scbBackGround = new SolidColorBrush(Colors.Gray);
                        break;
                }
                //判断是否取到该项取值
                if (scbBackGround != null)
                    dgRow.Background = scbBackGround;
            }

            

        }

        //搜索模型选择树 sCode:编号 ntype:类型
        private void selectTreeNode(string sCode,nodeType ntype)
        {
            nodeListSearch.Clear();
            List<TreeNodetagSelectTree> listSelectTreeViewNode=new List<TreeNodetagSelectTree>();

            switch (ntype)
            {
                case nodeType.FACILITY:
                    var qSelectTreeViewNode = from SelectTreeViewNode in nodeList   //查询空间
                                              where SelectTreeViewNode.nodetype.Equals(ntype) && SelectTreeViewNode.sFacilityCode == sCode
                                              select SelectTreeViewNode;
                    if (qSelectTreeViewNode.Count() <= 0)
                    {
                        //MessageBox.Show("没找到搜索内容", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    listSelectTreeViewNode = qSelectTreeViewNode.ToList();
                    break;
                case nodeType.ROOM:
                    qSelectTreeViewNode = from SelectTreeViewNode in nodeList   //查询空间
                                              where SelectTreeViewNode.nodetype.Equals(ntype) && SelectTreeViewNode.sRoomCode == sCode
                                              select SelectTreeViewNode;
                    if (qSelectTreeViewNode.Count() <= 0)
                    {
                        //MessageBox.Show("没找到搜索内容", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    listSelectTreeViewNode = qSelectTreeViewNode.ToList();
                    break;
                default:
                    return;
            }





            foreach (TreeNodetagSelectTree tnNode in listSelectTreeViewNode)
            {
                nodeListSearch.Add(tnNode);
            }

            iSearchLocation = 0;
            nodeListSearch[iSearchLocation].IsSelected = true;

            //节点显示
            showSelectTreeNode(nodeListSearch[iSearchLocation]); ;

        }


        private void checkboxRoomSelect_Checked(object sender, RoutedEventArgs e)
        {
            checkboxRoomSelectChange();
        }

        private void checkboxRoomSelect_Unchecked(object sender, RoutedEventArgs e)
        {
            checkboxRoomSelectChange();
        }

        private void checkboxRoomSelectChange()
        {
            bool bTemp = false;
            if (checkboxRoomSelect.IsChecked.HasValue)
                bTemp = (bool)checkboxRoomSelect.IsChecked;
            else
                bTemp = false;

            Search search = new Search();
            search.Selection.SelectAll();
            VariantData oData = VariantData.FromDisplayString("房间");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = search.FindAll(documentControlM.Document, false);

            //documentControlM.Document.CurrentSelection.CopyFrom(items);

            ModelItemCollection hidden = new ModelItemCollection();
            //hidden.AddRange(documentControlM.Document.CurrentSelection.SelectedItems);
            hidden.AddRange(items);

            if (bTemp) //显示所有房间
            {
                documentControlM.Document.Models.SetHidden(hidden, false);
            }
            else //隐藏所有房间
            {
                documentControlM.Document.Models.SetHidden(hidden, true);
            }
        }

        //U3d
        private void Buttonu3d_Click(object sender, RoutedEventArgs e)
        {
            //u3dPlayer.Focus();

            //string oString = "电子除尘装置|AB01|100||";
            ////string oString = "";
            

            //u3dPlayer.u3dPlayer.SendMessage("GameCenture", "Change", oString);

            updateUIMonitor();

        }

        //U3d场景调整
        //private void showFeatureU3D(string sSystem, string sCode, string sValue, string sFloor, string sCeil)
        //{
        //    string oString;

        //     //u3dPlayer.Focus();
        //     if(sSystem=="")
        //         oString="";
        //     else
        //         oString = sSystem+"|"+sCode+"|"+sValue.ToString()+"|"+sFloor.ToString()+"|"+sCeil.ToString();
        //         //oString = "s|" + sCode + "|" + sValue.ToString() + "|" + sFloor.ToString() + "|" + sCeil.ToString();

        //     //oString = "s|024|1000|3000|2000";
        //     u3dPlayer.u3dPlayer.SendMessage("GameCenture", "Change", oString);
             
        //}

        //U3D场景新参数
        private void showFeatureU3D(string sSystem, string sCode, string sValue, string sNormal)
        {
            string oString;

            //u3dPlayer.Focus();
            if (sSystem == "")
                oString = "";
            else
                oString = sSystem + "|" + sCode + "|" + sValue.ToString() + "|" + sNormal;
            //oString = "s|" + sCode + "|" + sValue.ToString() + "|" + sNormal;

            //oString = "s|024|1000|3000|2000";
            u3dPlayer.u3dPlayer.SendMessage("GameCenture", "Change", oString);

        }

        //顶视图
        private void ButtonVPTOP_Click(object sender, RoutedEventArgs e)
        {
            SavedItem siView = getViewPoint(sViewNameTop);
            if (siView == null)
                return;
            documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;
        }

        //视图充满
        private void ButtonVPFIT_Click(object sender, RoutedEventArgs e)
        {
            //documentControlM.Document.SetPlainBackground(Autodesk.Navisworks.Api.Color.Green);
            Viewpoint v = documentControlM.Document.CurrentViewpoint;
            BoundingBox3D box = documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox();

            v.ZoomBox(box);

            documentControlM.Document.CurrentViewpoint.CopyFrom(v);
        }

        //楼板显示控制
        private void ButtonShowFloor_Click(object sender, RoutedEventArgs e)
        {
            bool bHidden=false;
            Fluent.Button fb = sender as Fluent.Button;
            //if(fb.Foreground == SolidColorBrush.ColorProperty.SolidColorBrush(Colors.LightGray))
            if (fb.Header.ToString() == "隐藏天花")
            {
                fb.Foreground = new SolidColorBrush(Colors.Blue);
                fb.Header = "显示天花";
                bHidden = true;
            }
            else
            {
                fb.Foreground = new SolidColorBrush(Colors.Black);
                fb.Header = "隐藏天花";
                bHidden = false;
            }


            Search search = new Search();
            search.Selection.SelectAll();
            VariantData oData = VariantData.FromDisplayString("天花板");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = search.FindAll(documentControlM.Document, false);

            documentControlM.Document.Models.SetHidden(items,bHidden);
        }

        //房间楼板显示控制
        private void ButtonShowRoom_Click(object sender, RoutedEventArgs e)
        {
            bool bHidden = false;
            Fluent.Button fb = sender as Fluent.Button;
            //if(fb.Foreground == SolidColorBrush.ColorProperty.SolidColorBrush(Colors.LightGray))
            if (fb.Header.ToString() == "隐藏房间")
            {
                fb.Foreground = new SolidColorBrush(Colors.Blue);
                fb.Header = "显示房间";
                bHidden = true;
            }
            else
            {
                fb.Foreground = new SolidColorBrush(Colors.Black);
                fb.Header = "隐藏房间";
                bHidden = false;
            }


            Search search = new Search();
            search.Selection.SelectAll();
            VariantData oData = VariantData.FromDisplayString("房间");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = search.FindAll(documentControlM.Document, false);

            documentControlM.Document.Models.SetHidden(items, bHidden);
        }

        //机电显示控制
        private void ButtonShowMEP_Click(object sender, RoutedEventArgs e)
        {
            bool bHidden = false;
            Fluent.Button fb = sender as Fluent.Button;
            //if(fb.Foreground == SolidColorBrush.ColorProperty.SolidColorBrush(Colors.LightGray))
            if (fb.Header.ToString() == "隐藏机电")
            {
                fb.Foreground = new SolidColorBrush(Colors.Blue);
                fb.Header = "显示机电";
                bHidden = true;
            }
            else
            {
                fb.Foreground = new SolidColorBrush(Colors.Black);
                fb.Header = "隐藏机电";
                bHidden = false;
            }
            
            //找到机电模型
            var qLevelModels = from dtLevelDrawings in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                                 //where (dtLevelDrawings.Field<int>("标高ID") == tnSelectTree.iLevelID)//条件
                                 select dtLevelDrawings;


            foreach (var itemLevelModel in qLevelModels)//显示查询结果
            {
                Search search = new Search();
                search.Selection.SelectAll();
                VariantData oData = VariantData.FromDisplayString(itemLevelModel.Field<string>("机电模型组成"));

                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                oSearchCondition = oSearchCondition.EqualValue(oData);
                search.SearchConditions.Add(oSearchCondition);
                ModelItemCollection items = search.FindAll(documentControlM.Document, false);

                documentControlM.Document.Models.SetHidden(items, bHidden);
            }
        }

        //能源分析,分层电量
        private void ButtonEleLevelAnalysis_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            LayoutEnergyAnalysis.IsVisible=true;
            //LayoutEnergyAnalysis.Float();

            classEA.getLevelEnergyAnalysis_Ele();
            EnergyAnalysis.Children.Clear();//清除分析图表
            showEnergyAnalysisStackedColumn("分层电量统计","度",classEA.listEA);
            showEnergyAnalysisStackedPie("分层电量比", "度", classEA.listEA);
        }

        
        //能源分析,部门电量
        private void ButtonEleDepAnalysis_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            LayoutEnergyAnalysis.IsVisible = true;

            classEA.getDepartmentEnergyAnalysis_Ele();
            EnergyAnalysis.Children.Clear();//清除分析图表
            showEnergyAnalysisStackedColumn("部门电量统计", "度", classEA.listEA);
            showEnergyAnalysisStackedPie("部门电量比", "度", classEA.listEA);
        }

        //能源分析,分层水量
        private void ButtonWaterLevelAnalysis_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            LayoutEnergyAnalysis.IsVisible = true;
            //LayoutEnergyAnalysis.Float();

            classEA.getLevelEnergyAnalysis_1("水表");
            EnergyAnalysis.Children.Clear();//清除分析图表
            showEnergyAnalysisStackedColumn("分层水量统计", "吨", classEA.listEA);
            showEnergyAnalysisStackedPie("分层量比", "吨", classEA.listEA);
        }

        private void ButtonWaterDepAnalysis_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            LayoutEnergyAnalysis.IsVisible = true;

            classEA.getDepartmentEnergyAnalysis_1("水表");
            EnergyAnalysis.Children.Clear();//清除分析图表
            showEnergyAnalysisStackedColumn("部门水量统计", "吨", classEA.listEA);
            showEnergyAnalysisStackedPie("部门水量比", "吨", classEA.listEA);
        }

        private void ButtonHeatLevelAnalysis_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            LayoutEnergyAnalysis.IsVisible = true;
            //LayoutEnergyAnalysis.Float();

            classEA.getLevelEnergyAnalysis_1("热计量");
            EnergyAnalysis.Children.Clear();//清除分析图表
            showEnergyAnalysisStackedColumn("分层热计量统计", "千卡", classEA.listEA);
            showEnergyAnalysisStackedPie("分层量比", "千卡", classEA.listEA);
        }

        private void ButtonHeatDepAnalysis_Click(object sender, RoutedEventArgs e)
        {
            Fluent.Button fb = sender as Fluent.Button;
            LayoutEnergyAnalysis.IsVisible = true;

            classEA.getDepartmentEnergyAnalysis_1("热计量");
            EnergyAnalysis.Children.Clear();//清除分析图表
            showEnergyAnalysisStackedColumn("部门热量统计", "千卡", classEA.listEA);
            showEnergyAnalysisStackedPie("部门热量比", "千卡", classEA.listEA);
        }

        //能源分析,部门月度电量
        private void ButtonEleMonthAnalysis_Click(object sender, RoutedEventArgs e)
        {
            //Fluent.Button fb = sender as Fluent.Button;
            LayoutEnergyAnalysis.IsVisible = true;

            EnergyAnalysis.Children.Clear();//清除分析图表
            classEA.getDepartmentEnergyNowMonthAnalysis_Ele();
            showEnergyAnalysisStackedColumn("本月电量统计", "度", classEA.listEA);

            //classEA.getDepartmentEnergyLastMonthAnalysis_Ele();
            //showEnergyAnalysisStackedColumn("上月电量比", "度", classEA.listEA);

            classEA.getDepartmentEnergyDaysAnalysis_Ele(5);
            showEnergyAnalysisStackedSPline("五日电量走势", "度", "B座", classEA.listEA);
        }

        //显示能源分析数据柱状图,标题，单位，数据
        private void showEnergyAnalysisStackedColumn(string chartTitle, string Suffix,List<ClassListEA> listChart)
        {
            int i, j;

            #region 柱状图
            Chart chartStackedColumn = new Chart();

            //设置图标的宽度和高度
            chartStackedColumn.Width = 600;
            chartStackedColumn.Height = 300;
            chartStackedColumn.Margin = new Thickness(5, 5, 10, 5);
            //是否启用打印和保持图片
            chartStackedColumn.ToolBarEnabled = false;

            //设置图标的属性
            chartStackedColumn.ScrollingEnabled = false;//是否启用或禁用滚动
            chartStackedColumn.View3D = true;//3D效果显示

            //创建一个标题的对象
            Title titlechartStackedColumn = new Title();
            //设置标题的名称
            titlechartStackedColumn.Text = chartTitle;
            titlechartStackedColumn.Padding = new Thickness(0, 10, 5, 0);
            //向图标添加标题
            chartStackedColumn.Titles.Add(titlechartStackedColumn);

            Axis yAxischartStackedColumn = new Axis();
            //设置图标中Y轴的最小值永远为0           
            yAxischartStackedColumn.AxisMinimum = 0;
            //设置图表中Y轴的后缀          
            yAxischartStackedColumn.Suffix = Suffix;
            chartStackedColumn.AxesY.Add(yAxischartStackedColumn);

            // 创建一个新的数据线。               
            DataSeries dataSeriesStackedColumn = new DataSeries();
            // 设置数据线的格式
            dataSeriesStackedColumn.RenderAs = RenderAs.StackedColumn;//柱状Stacked

            //List<string> valuexStackedColumn = new List<string>() { "第一天", "第二天", "第三天", "第四天", "第五天" };
            //List<string> valueyStackedColumn = new List<string>() { "130", "750", "600", "380", "970" };

            // 设置数据点              
            DataPoint dataPointStackedColumn;
            for (i = 0; i < listChart.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointStackedColumn = new DataPoint();
                // 设置X轴点                    
                dataPointStackedColumn.AxisXLabel = listChart[i].sName;
                //设置Y轴点                   
                dataPointStackedColumn.YValue = double.Parse(listChart[i].sValue);
                //添加一个点击事件        
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesStackedColumn.DataPoints.Add(dataPointStackedColumn);
            }

            // 添加数据线到数据序列。                
            chartStackedColumn.Series.Add(dataSeriesStackedColumn);

            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            Grid grStackedColumn = new Grid();
            grStackedColumn.Children.Add(chartStackedColumn);
            EnergyAnalysis.Children.Add(grStackedColumn);
            //Grid.SetRow(chartStackedColumn, 0);
            grStackedColumn.VerticalAlignment = VerticalAlignment.Top;
            #endregion
        }

        //显示能源分析数据饼图,标题，单位，数据
        private void showEnergyAnalysisStackedPie(string chartTitle, string Suffix, List<ClassListEA> listChart)
        {
            int i, j;

            #region 饼图
            //创建一个图标
            Chart chartPie = new Chart();

            //设置图标的宽度和高度
            chartPie.Width = 600;
            chartPie.Height = 300;
            chartPie.Margin = new Thickness(5, 5, 10, 5);
            //是否启用打印和保持图片
            chartPie.ToolBarEnabled = false;

            //设置图标的属性
            chartPie.ScrollingEnabled = false;//是否启用或禁用滚动
            chartPie.View3D = true;//3D效果显示

            //创建一个标题的对象
            Title titlePie = new Title();

            //设置标题的名称
            titlePie.Text = chartTitle;
            titlePie.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chartPie.Titles.Add(titlePie);

            // 创建一个新的数据线。               
            DataSeries dataSeriesPie = new DataSeries();

            // 设置数据线的格式
            dataSeriesPie.RenderAs = RenderAs.Pie;//柱状Stacked
            //dataSeriesPie.PercentageFormatString

            // 设置数据点              
            DataPoint dataPointPie;
            for (i = 0; i < listChart.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointPie = new DataPoint();
                // 设置X轴点                    
                dataPointPie.AxisXLabel = listChart[i].sName;

                dataPointPie.LegendText = "##" + listChart[i].sName;
                //设置Y轴点                   
                dataPointPie.YValue = double.Parse(listChart[i].sValue);
                //添加一个点击事件        
                //dataPointPie.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesPie.DataPoints.Add(dataPointPie);
            }

            // 添加数据线到数据序列。                
            chartPie.Series.Add(dataSeriesPie);

            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            Grid grStackedPie = new Grid();
            grStackedPie.Children.Add(chartPie);
            EnergyAnalysis.Children.Add(grStackedPie);
            grStackedPie.VerticalAlignment = VerticalAlignment.Top;
            #endregion
        }

        //显示能源分析数据曲线图,标题，单位，数据
        private void showEnergyAnalysisStackedSPline(string chartTitle, string Suffix, string LegendText,List<ClassListEA> listChart)
        {
            int i, j;

            #region 曲线图
            Chart chartSpline = new Chart();

            //设置图标的宽度和高度
            chartSpline.Width = 600;
            chartSpline.Height = 300;

            chartSpline.Margin = new Thickness(5, 5, 10, 5);
            //是否启用打印和保持图片
            chartSpline.ToolBarEnabled = false;

            //设置图标的属性
            chartSpline.ScrollingEnabled = false;//是否启用或禁用滚动
            chartSpline.View3D = true;//3D效果显示

            //创建一个标题的对象
            Title titlechartSpline = new Title();

            //设置标题的名称
            titlechartSpline.Text = chartTitle;
            titlechartSpline.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chartSpline.Titles.Add(titlechartSpline);

            //初始化一个新的Axis
            Axis xaxisSpline = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSpline.IntervalType = IntervalTypes.Days;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSpline.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSpline.ValueFormatString = "yyyy-MM-dd";
            //给图标添加Axis            
            chartSpline.AxesX.Add(xaxisSpline);

            Axis yAxisSpline = new Axis();
            //设置图标中Y轴的最小值永远为0           
            yAxisSpline.AxisMinimum = 0;
            //设置图表中Y轴的后缀          
            yAxisSpline.Suffix = Suffix;
            chartSpline.AxesY.Add(yAxisSpline);


            // 创建一个新的数据线。               
            DataSeries dataSeriesSpline = new DataSeries();
            // 设置数据线的格式。               
            dataSeriesSpline.LegendText = LegendText;

            dataSeriesSpline.RenderAs = RenderAs.Spline;//折线图

            dataSeriesSpline.XValueType = ChartValueTypes.DateTime;
            // 设置数据点              
            DataPoint dataPointSpline;

            //List<DateTime> LsTimeSpline = new List<DateTime>()
            //    { 
            //       DateTime.Now.AddDays(-4),
            //       DateTime.Now.AddDays(-3),
            //       DateTime.Now.AddDays(-2),
            //       DateTime.Now.AddDays(-1),
            //       DateTime.Now,
            //    };

            //List<string> listYValueSpline = new List<string>() { "33", "75", "60", "98", "67" };

            for (i = 0; i < listChart.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSpline = new DataPoint();
                // 设置X轴点                    
                dataPointSpline.XValue = DateTime.Parse(listChart[i].sName);
                //设置Y轴点                   
                dataPointSpline.YValue = double.Parse(listChart[i].sValue);
                dataPointSpline.MarkerSize = listChart.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                // dataPoint.Color = new SolidColorBrush(Colors.LightGray);                   
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSpline.DataPoints.Add(dataPointSpline);
            }

            // 添加数据线到数据序列。                
            chartSpline.Series.Add(dataSeriesSpline);



            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            Grid grStackedSpline = new Grid();
            grStackedSpline.Children.Add(chartSpline);
            EnergyAnalysis.Children.Add(grStackedSpline);
            grStackedSpline.VerticalAlignment = VerticalAlignment.Top;
            //grStackedSpline.SetValue(DockPanel.DockProperty, Dock.Top);

            //Grid.SetRow(chartSpline, 0);


            #endregion

        }

        //显示能源分析数据
        //private void showEnergyAnalysis1()
        //{

        //    int i, j;

        //    #region 柱状图
        //    EnergyAnalysis.Children.Clear();
        //    Chart chartStackedColumn = new Chart();

        //    //设置图标的宽度和高度
        //    chartStackedColumn.Width = 300;
        //    chartStackedColumn.Height = 200;
        //    chartStackedColumn.Margin = new Thickness(5, 5, 10, 5);
        //    //是否启用打印和保持图片
        //    chartStackedColumn.ToolBarEnabled = false;

        //    //设置图标的属性
        //    chartStackedColumn.ScrollingEnabled = false;//是否启用或禁用滚动
        //    chartStackedColumn.View3D = true;//3D效果显示

        //    //创建一个标题的对象
        //    Title titlechartStackedColumn = new Title();
        //    //设置标题的名称
        //    titlechartStackedColumn.Text = "5日内用电量";
        //    titlechartStackedColumn.Padding = new Thickness(0, 10, 5, 0);
        //    //向图标添加标题
        //    chartStackedColumn.Titles.Add(titlechartStackedColumn);

        //    Axis yAxischartStackedColumn = new Axis();
        //    //设置图标中Y轴的最小值永远为0           
        //    yAxischartStackedColumn.AxisMinimum = 0;
        //    //设置图表中Y轴的后缀          
        //    yAxischartStackedColumn.Suffix = "度";
        //    chartStackedColumn.AxesY.Add(yAxischartStackedColumn);

        //    // 创建一个新的数据线。               
        //    DataSeries dataSeriesStackedColumn = new DataSeries();
        //    // 设置数据线的格式
        //    dataSeriesStackedColumn.RenderAs = RenderAs.StackedColumn;//柱状Stacked

        //    List<string> valuexStackedColumn = new List<string>() { "职能", "一所", "二所", "三所", "四所" };
        //    List<string> valueyStackedColumn = new List<string>() { "130", "750", "600", "380", "970" };

        //    // 设置数据点              
        //    DataPoint dataPointStackedColumn;
        //    for (i = 0; i < valuexStackedColumn.Count; i++)
        //    {
        //        // 创建一个数据点的实例。                   
        //        dataPointStackedColumn = new DataPoint();
        //        // 设置X轴点                    
        //        dataPointStackedColumn.AxisXLabel = valuexStackedColumn[i];
        //        //设置Y轴点                   
        //        dataPointStackedColumn.YValue = double.Parse(valueyStackedColumn[i]);
        //        //添加一个点击事件        
        //        //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
        //        //添加数据点                   
        //        dataSeriesStackedColumn.DataPoints.Add(dataPointStackedColumn);
        //    }

        //    // 添加数据线到数据序列。                
        //    chartStackedColumn.Series.Add(dataSeriesStackedColumn);

        //    //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
        //    Grid grStackedColumn = new Grid();
        //    grStackedColumn.Children.Add(chartStackedColumn);
        //    EnergyAnalysis.Children.Add(grStackedColumn);
        //    //Grid.SetRow(chartStackedColumn, 0);
        //    grStackedColumn.VerticalAlignment = VerticalAlignment.Top;
        //    #endregion

        //    #region 曲线图
        //    Chart chartSpline = new Chart();

        //    //设置图标的宽度和高度
        //    chartSpline.Width = 300;
        //    chartSpline.Height = 200;

        //    chartSpline.Margin = new Thickness(5, 5, 10, 5);
        //    //是否启用打印和保持图片
        //    chartSpline.ToolBarEnabled = false;

        //    //设置图标的属性
        //    chartSpline.ScrollingEnabled = false;//是否启用或禁用滚动
        //    chartSpline.View3D = true;//3D效果显示

        //    //创建一个标题的对象
        //    Title titlechartSpline = new Title();

        //    //设置标题的名称
        //    titlechartSpline.Text = "五日用水量";
        //    titlechartSpline.Padding = new Thickness(0, 10, 5, 0);

        //    //向图标添加标题
        //    chartSpline.Titles.Add(titlechartSpline);

        //    //初始化一个新的Axis
        //    Axis xaxisSpline = new Axis();
        //    //设置Axis的属性
        //    //图表的X轴坐标按什么来分类，如时分秒
        //    xaxisSpline.IntervalType = IntervalTypes.Days;
        //    //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
        //    xaxisSpline.Interval = 1;
        //    //设置X轴的时间显示格式为7-10 11：20           
        //    xaxisSpline.ValueFormatString = "yyyy-MM-dd";
        //    //给图标添加Axis            
        //    chartSpline.AxesX.Add(xaxisSpline);

        //    Axis yAxisSpline = new Axis();
        //    //设置图标中Y轴的最小值永远为0           
        //    yAxisSpline.AxisMinimum = 0;
        //    //设置图表中Y轴的后缀          
        //    yAxisSpline.Suffix = "吨";
        //    chartSpline.AxesY.Add(yAxisSpline);


        //    // 创建一个新的数据线。               
        //    DataSeries dataSeriesSpline = new DataSeries();
        //    // 设置数据线的格式。               
        //    dataSeriesSpline.LegendText = "B座";

        //    dataSeriesSpline.RenderAs = RenderAs.Spline;//折线图

        //    dataSeriesSpline.XValueType = ChartValueTypes.DateTime;
        //    // 设置数据点              
        //    DataPoint dataPointSpline;

        //    List<DateTime> LsTimeSpline = new List<DateTime>()
        //    { 
        //       DateTime.Now.AddDays(-4),
        //       DateTime.Now.AddDays(-3),
        //       DateTime.Now.AddDays(-2),
        //       DateTime.Now.AddDays(-1),
        //       DateTime.Now,
        //    };

        //    List<string> listYValueSpline = new List<string>() { "33", "75", "60", "98", "67" };

        //    for (i = 0; i < LsTimeSpline.Count; i++)
        //    {
        //        // 创建一个数据点的实例。                   
        //        dataPointSpline = new DataPoint();
        //        // 设置X轴点                    
        //        dataPointSpline.XValue = LsTimeSpline[i];
        //        //设置Y轴点                   
        //        dataPointSpline.YValue = double.Parse(listYValueSpline[i]);
        //        dataPointSpline.MarkerSize = 5;
        //        //dataPoint.Tag = tableName.Split('(')[0];
        //        //设置数据点颜色                  
        //        // dataPoint.Color = new SolidColorBrush(Colors.LightGray);                   
        //        //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
        //        //添加数据点                   
        //        dataSeriesSpline.DataPoints.Add(dataPointSpline);
        //    }

        //    // 添加数据线到数据序列。                
        //    chartSpline.Series.Add(dataSeriesSpline);



        //    //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
        //    Grid grStackedSpline = new Grid();
        //    grStackedSpline.Children.Add(chartSpline);
        //    EnergyAnalysis.Children.Add(grStackedSpline);
        //    grStackedSpline.VerticalAlignment = VerticalAlignment.Top;
        //    //grStackedSpline.SetValue(DockPanel.DockProperty, Dock.Top);

        //    //Grid.SetRow(chartSpline, 0);


        //    #endregion

        //    #region 饼图
        //    //创建一个图标
        //    Chart chartPie = new Chart();

        //    //设置图标的宽度和高度
        //    chartPie.Width = 300;
        //    chartPie.Height = 200;
        //    chartPie.Margin = new Thickness(5, 5, 10, 5);
        //    //是否启用打印和保持图片
        //    chartPie.ToolBarEnabled = false;

        //    //设置图标的属性
        //    chartPie.ScrollingEnabled = false;//是否启用或禁用滚动
        //    chartPie.View3D = true;//3D效果显示

        //    //创建一个标题的对象
        //    Title titlePie = new Title();

        //    //设置标题的名称
        //    titlePie.Text = "能耗比";
        //    titlePie.Padding = new Thickness(0, 10, 5, 0);

        //    //向图标添加标题
        //    chartPie.Titles.Add(titlePie);

        //    // 创建一个新的数据线。               
        //    DataSeries dataSeriesPie = new DataSeries();

        //    // 设置数据线的格式
        //    dataSeriesPie.RenderAs = RenderAs.Pie;//柱状Stacked

        //    List<string> valuexPie = new List<string>() { "职能", "一院", "二院", "三院", "四院" };
        //    List<string> valueyPie = new List<string>() { "13", "75", "38", "60", "97" };

        //    // 设置数据点              
        //    DataPoint dataPointPie;
        //    for (i = 0; i < valuexPie.Count; i++)
        //    {
        //        // 创建一个数据点的实例。                   
        //        dataPointPie = new DataPoint();
        //        // 设置X轴点                    
        //        dataPointPie.AxisXLabel = valuexPie[i];

        //        dataPointPie.LegendText = "##" + valuexPie[i];
        //        //设置Y轴点                   
        //        dataPointPie.YValue = double.Parse(valueyPie[i]);
        //        //添加一个点击事件        
        //        //dataPointPie.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
        //        //添加数据点                   
        //        dataSeriesPie.DataPoints.Add(dataPointPie);
        //    }

        //    // 添加数据线到数据序列。                
        //    chartPie.Series.Add(dataSeriesPie);

        //    //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
        //    Grid grStackedPie = new Grid();
        //    grStackedPie.Children.Add(chartPie);
        //    EnergyAnalysis.Children.Add(grStackedPie);
        //    grStackedPie.VerticalAlignment = VerticalAlignment.Top;
        //    #endregion


        //}


        //搜索
        
        public void GetAllNodes(ObservableCollection<TreeNodetagSelectTree> itemList)//ItemCollection nodeCollection, List<TreeNodetagSelectTree> nodeList)
        {

            //nodeList.Clear();
            foreach (TreeNodetagSelectTree itemNode in itemList)
            {
                nodeList.Add(itemNode);
                GetAllNodes(itemNode.Children);
            }

        }

        //搜索
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if(textboxFacilitySearch.Text.Trim()=="")
            {
                MessageBox.Show("请输入搜索关键字","提示",MessageBoxButton.OK,MessageBoxImage.Information);
                return;
            }

            //ObservableCollection<TreeNodetagSelectTree> itemList = treeviewSelectTree.ItemsSource as ObservableCollection<TreeNodetagSelectTree>;

            //foreach (TreeNodetagSelectTree itemNode in itemList)
            //{
            //    itemNode.IsSelected=true;
            //}
            nodeListSearch.Clear();

            var qSelectTreeViewNode = from SelectTreeViewNode in nodeList   //查询空间
                                      where SelectTreeViewNode.sDisplayName.Contains(textboxFacilitySearch.Text.Trim())
                                      select SelectTreeViewNode;

            if (qSelectTreeViewNode.Count() <= 0)
            {
                MessageBox.Show("没找到搜索内容", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            foreach (TreeNodetagSelectTree tnNode in qSelectTreeViewNode)
            {
                nodeListSearch.Add(tnNode);
            }

            iSearchLocation = 0;
            FindNodeAllUpDown(0);

        }

        //定位下一个查询。iMoveCount：+1 下一个查询 -1：上一个查询
        private void FindNodeAllUpDown(int iMoveCount)
        {
            if (nodeListSearch.Count <= 0)
                return;

            if (iSearchLocation + iMoveCount < 0)
            {
                MessageBox.Show("已到搜索起始点", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (iSearchLocation + iMoveCount >= nodeListSearch.Count)
            {
                MessageBox.Show("已到搜索终点", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            iSearchLocation += iMoveCount;

            nodeListSearch[iSearchLocation].IsSelected = true;
            

            //节点显示
            showSelectTreeNode(nodeListSearch[iSearchLocation]); ;
        }

        //节点显示
        private void showSelectTreeNode(TreeNodetagSelectTree tnNode)
        {
            if (tnNode.Parent == null)
                return;


            tnNode.Parent.IsExpanded = true;
            showSelectTreeNode(tnNode.Parent);
        }

        private void btnSearchNext_Click(object sender, RoutedEventArgs e)
        {
            FindNodeAllUpDown(1);
        }

        private void btnSearchLast_Click(object sender, RoutedEventArgs e)
        {
            FindNodeAllUpDown(-1);
        }

        private void ButtonuVR_Click(object sender, RoutedEventArgs e)
        {
            string sFileVR;

            if (treeviewSelectTree.SelectedItem == null)
            {
                MessageBox.Show("请选择楼层");
                return;
            }

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.nodetype != nodeType.LEVEL)
            {
                MessageBox.Show("请选择楼层");
                return;
            }

            //找到VR
            var qLevelVR = from dtLevelVRs in dSet.Tables["标高表"].AsEnumerable()//查询楼层
                                 where (dtLevelVRs.Field<int>("标高ID") == tnSelectTree.iLevelID)//条件
                                 select dtLevelVRs;


            foreach (var itemLevelVR in qLevelVR)//显示查询结果
            {
                sFileVR = itemLevelVR.Field<string>("VR");
                if (sFileVR == null || sFileVR == "")
                {
                    MessageBox.Show("此层没有VR周游场景");
                    break;
                }

                mainDrawingView.Title = sFileVR;
                sFileVR = sDirDrawings + "/VR/" + sFileVR; //目录暂定为图纸子目录
                //MessageBox.Show(sFileDrawing);

                if (File.Exists(sFileVR))
                //viewPDF.LoadPDFfile(sFileDrawing);
                {
                    Process process = new Process();
                    //process.StartInfo.UseShellExecute = false;
                    process.StartInfo.FileName = sFileVR;
                    //process.StartInfo.CreateNoWindow = true;
                    process.Start();
                }
                break;
            }//VR
            
            

        }

        //云端浏览
        private void Buttonlmv_Click(object sender, RoutedEventArgs e)
        {
            //
            int i, j;

            if (treeviewSelectTree.SelectedItem == null)
            {
                MessageBox.Show("请选择设备");
                return;
            }

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.sModel.Trim() == "")
                return;

            if (tnSelectTree.nodetype != nodeType.FACILITY)
            {
                MessageBox.Show("请选择设备");
                return;
            }

            //寻找设备
                      
            var qFacility = from dtFacility in dSet.Tables["设备URN表"].AsEnumerable()//查询设备编号
                where (dtFacility.Field<string>("设备编号") == tnSelectTree.sFacilityCode)//条件
                select dtFacility;
            if (qFacility.Count() < 1)
            {
                MessageBox.Show("设备在云端没有数据");
                return;
            }
            foreach (var itemFaclity in qFacility)//显示查询结果
            {
                string _fileUrn = itemFaclity.Field<string>("URN");
                if (_fileUrn != "")
                {
                    //得到云端权限
                    if(AutodeskVDS.authentication())
                    {
                        String _token = AutodeskVDS._token;
                        string url = string.Format("http://viewer.autodesk.io/node/view-helper?urn={0}&token={1}", _fileUrn, _token);
                        Process.Start(url);
                    }
                    
                }
                break;
            }
        }

        private void ButtonRenderFull_Click(object sender, RoutedEventArgs e)
        {
            Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();
            oCurrVCopy.RenderStyle = ViewpointRenderStyle.FullRender;
            documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);

            StatusBarRender.Content = (sender as Fluent.Button).Header;
        }

        private void ButtonRenderShaded_Click(object sender, RoutedEventArgs e)
        {
            Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();
            oCurrVCopy.RenderStyle = ViewpointRenderStyle.Shaded;
            documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);

            StatusBarRender.Content = (sender as Fluent.Button).Header;
        }



        private void checkboxThirdPerson_Checked(object sender, RoutedEventArgs e)
        {
            Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2 oV = (Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2)ComApiBridge.State.CurrentView.ViewPoint;
            oV.Viewer.CameraMode = Autodesk.Navisworks.Api.Interop.ComApi.nwECameraMode.eCameraMode_ThirdPerson;
            
        }

        private void checkboxThirdPerson_Unchecked(object sender, RoutedEventArgs e)
        {
            Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2 oV = (Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2)ComApiBridge.State.CurrentView.ViewPoint;

            oV.Viewer.CameraMode = Autodesk.Navisworks.Api.Interop.ComApi.nwECameraMode.eCameraMode_FirstPerson;
        }

        private void checkboxGravity_Checked(object sender, RoutedEventArgs e)
        {
            checkboxCollisionDetection.IsChecked = true;

            Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2 oV = (Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2)ComApiBridge.State.CurrentView.ViewPoint;

            //if (oV.Paradigm == Autodesk.Navisworks.Api.Interop.ComApi.nwEParadigm.eParadigm_WALK)
            //{
                // gravity, coollison or crouch can only take effect in Walk mode.
                oV.Viewer.Gravity = true;//toogle gravity on
                 
            //}
        }

        private void checkboxCollisionDetection_Unchecked(object sender, RoutedEventArgs e)
        {
            checkboxGravity.IsChecked = false;

            Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2 oV = (Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2)ComApiBridge.State.CurrentView.ViewPoint;

            //if (oV.Paradigm == Autodesk.Navisworks.Api.Interop.ComApi.nwEParadigm.eParadigm_WALK)
            //{
                // gravity, coollison or crouch can only take effect in Walk mode.
                oV.Viewer.CollisionDetection = false;//toogle gravity on
            //}
        }

        private void checkboxGravity_Unchecked(object sender, RoutedEventArgs e)
        {
            Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2 oV = (Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2)ComApiBridge.State.CurrentView.ViewPoint;

            //if (oV.Paradigm == Autodesk.Navisworks.Api.Interop.ComApi.nwEParadigm.eParadigm_WALK)
            //{
                // gravity, coollison or crouch can only take effect in Walk mode.
                oV.Viewer.Gravity = false;//toogle gravity on
            //}
        }

        private void checkboxCollisionDetection_Checked(object sender, RoutedEventArgs e)
        {
            Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2 oV = (Autodesk.Navisworks.Api.Interop.ComApi.InwNvViewPoint2)ComApiBridge.State.CurrentView.ViewPoint;

            //if (oV.Paradigm == Autodesk.Navisworks.Api.Interop.ComApi.nwEParadigm.eParadigm_WALK)
            //{
                // gravity, coollison or crouch can only take effect in Walk mode.
                oV.Viewer.CollisionDetection = true;//toogle gravity on
            //}
        }

        private void ButtonFullLights_Click(object sender, RoutedEventArgs e)
        {
            Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();
            oCurrVCopy.Lighting=ViewpointLighting.FullLights;
            documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);

            StatusBarLighting.Content = (sender as Fluent.Button).Header;

            DropDownButtonLighting.Focus();
        }

        private void ButtonHeadlight_Click(object sender, RoutedEventArgs e)
        {
            Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();
            oCurrVCopy.Lighting = ViewpointLighting.Headlight;
            documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);

            StatusBarLighting.Content = (sender as Fluent.Button).Header;
        }

        private void ButtonSceneLights_Click(object sender, RoutedEventArgs e)
        {
            Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();
            oCurrVCopy.Lighting = ViewpointLighting.SceneLights;
            documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);

            StatusBarLighting.Content = (sender as Fluent.Button).Header;
        }

        private void ButtonNoneLights_Click(object sender, RoutedEventArgs e)
        {
            Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();
            oCurrVCopy.Lighting = ViewpointLighting.None;
            documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);

            StatusBarLighting.Content = (sender as Fluent.Button).Header;
        }

        //系统视图，暖通水
        private void ButtonSystemM_NTS_Click(object sender, RoutedEventArgs e)
        {
            viewSystem("暖通-暖通水");
            StatusBarSystem.Content = (sender as Fluent.Button).Header;
        }
        private void ButtonSystemM_NTF_Click(object sender, RoutedEventArgs e)
        {
            viewSystem("暖通-暖通风");
            StatusBarSystem.Content = (sender as Fluent.Button).Header;
        }
        private void ButtonSystemP_XF_Click(object sender, RoutedEventArgs e)
        {
            viewSystem("给排水-消防");
            StatusBarSystem.Content = (sender as Fluent.Button).Header;
        }

        private void ButtonSystemP_GPS_Click(object sender, RoutedEventArgs e)
        {
            viewSystem("给排水-给排水");
            StatusBarSystem.Content = (sender as Fluent.Button).Header;
        }

        private void ButtonSystemE_QD_Click(object sender, RoutedEventArgs e)
        {
            viewSystem("电气-强电");
            StatusBarSystem.Content = (sender as Fluent.Button).Header;
        }

        private void ButtonSystemE_RD_Click(object sender, RoutedEventArgs e)
        {
            viewSystem("电气-弱电");
            StatusBarSystem.Content = (sender as Fluent.Button).Header;
        }

        private void ButtonSystem_TLLC_Click(object sender, RoutedEventArgs e)
        {
            viewSystem("");
            StatusBarSystem.Content = (sender as Fluent.Button).Header;
        }

        //单系统显示，系统名称，""只显示体量
        private void viewSystem(string sSystemname)
        {
            
            Search search = new Search();
            VariantData oData;
            ModelItemCollection items;


            //隐藏所有土建模型
            var qBuildingModels = from dtBuildingModels in dSet.Tables["标高表"].AsEnumerable()//查询
                              select dtBuildingModels;
            //显示系统
            if (sSystemname == "") //不显示系统
            {
                search.Clear();
                search.Selection.SelectAll();
                qBuildingModels = from dtBuildingModels in dSet.Tables["标高表"].AsEnumerable()//查询
                                  select dtBuildingModels;
                foreach (var itemBuildingModels in qBuildingModels)//显示查询结果
                {
                    if (itemBuildingModels.Field<string>("土建模型组成") != null)
                    {
                        if (itemBuildingModels.Field<string>("土建模型组成") != "")
                        {
                            System.Collections.Generic.List<SearchCondition> oG = new System.Collections.Generic.List<SearchCondition>();
                            oData = VariantData.FromDisplayString(itemBuildingModels.Field<string>("土建模型组成"));


                            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                            oSearchCondition = oSearchCondition.EqualValue(oData);
                            oG.Add(oSearchCondition);
                            search.SearchConditions.AddGroup(oG);
                        }
                    }
                }
                items = search.FindAll(documentControlM.Document, false);
                //不显示土建
                documentControlM.Document.Models.SetHidden(items, true);

                search.Clear();
                search.Selection.SelectAll();
                foreach (var itemBuildingModels in qBuildingModels)//显示查询结果
                {
                    if (itemBuildingModels.Field<string>("机电模型组成") != null)
                    {
                        if (itemBuildingModels.Field<string>("机电模型组成") != "")
                        {
                            System.Collections.Generic.List<SearchCondition> oG = new System.Collections.Generic.List<SearchCondition>();
                            oData = VariantData.FromDisplayString(itemBuildingModels.Field<string>("机电模型组成"));


                            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                            oSearchCondition = oSearchCondition.EqualValue(oData);
                            oG.Add(oSearchCondition);
                            search.SearchConditions.AddGroup(oG);
                        }
                    }
                }
                items = search.FindAll(documentControlM.Document, false);
                //不显示机电
                documentControlM.Document.Models.SetHidden(items, true);
            }
            else //显示机电系统
            {
                search.Clear();
                search.Selection.SelectAll();
                foreach (var itemBuildingModels in qBuildingModels)//显示查询结果
                {
                    if (itemBuildingModels.Field<string>("机电模型组成") != null)
                    {
                        if (itemBuildingModels.Field<string>("机电模型组成") != "")
                        {
                            System.Collections.Generic.List<SearchCondition> oG = new System.Collections.Generic.List<SearchCondition>();
                            oData = VariantData.FromDisplayString(itemBuildingModels.Field<string>("机电模型组成"));
                            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                            oSearchCondition = oSearchCondition.EqualValue(oData);
                            oG.Add(oSearchCondition);
                            search.SearchConditions.AddGroup(oG);
                        }
                    }

                    if (itemBuildingModels.Field<string>("土建模型组成") != null)
                    {
                        if (itemBuildingModels.Field<string>("土建模型组成") != "")
                        {
                            System.Collections.Generic.List<SearchCondition> oG = new System.Collections.Generic.List<SearchCondition>();
                            oData = VariantData.FromDisplayString(itemBuildingModels.Field<string>("土建模型组成"));
                            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                            oSearchCondition = oSearchCondition.EqualValue(oData);
                            oG.Add(oSearchCondition);
                            search.SearchConditions.AddGroup(oG);
                        }
                    }
                }
                items = search.FindAll(documentControlM.Document, false);

                //ModelItemCollection miHiddenCollection = new ModelItemCollection();
                //ModelItemCollection miUnHiddenCollection = new ModelItemCollection();
                //显示模型
                documentControlM.Document.Models.SetHidden(items, false);

                //状态
                ProgressBarStatus.Maximum = items.Count*2+2;
                ProgressBarStatus.Value = 0;
                Dispatcher.Invoke(new Action(() =>
                {
                   
                }), System.Windows.Threading.DispatcherPriority.Background);

                //显示工作集
                foreach (ModelItem modelitem in items)
                {
                    //显示文件
                    documentControlM.Document.Models.SetHidden(modelitem.Descendants, false);
                    ProgressBarStatus.Value++;
                    Dispatcher.Invoke(new Action(() =>
                    {
                    }), System.Windows.Threading.DispatcherPriority.Background);

                    //隐藏非指定工作集
                    IEnumerable<ModelItem> items2 = from item2 in modelitem.Descendants
                                                    where (getPropertyValueString(item2, "元素", "工作集") != sSystemname) && (item2.PropertyCategories.FindPropertyByDisplayName("元素", "工作集") != null)
                                                    select item2;
                    documentControlM.Document.Models.SetHidden(items2, true);
                    ProgressBarStatus.Value++;
                    Dispatcher.Invoke(new Action(() =>
                    {
                    }), System.Windows.Threading.DispatcherPriority.Background);
                }


            }

            search.Clear();
            search.Selection.SelectAll();
            //隐藏所有土建基础模型
            qBuildingModels = from dtBuildingModels in dSet.Tables["建筑表"].AsEnumerable()//查询
                                  select dtBuildingModels;
            foreach (var itemBuildingModels in qBuildingModels)//显示查询结果
            {
                if (itemBuildingModels.Field<string>("基础模型") != null)
                {
                    if (itemBuildingModels.Field<string>("基础模型") != "")
                    {
                        oData = VariantData.FromDisplayString(itemBuildingModels.Field<string>("基础模型"));

                        SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                        oSearchCondition = oSearchCondition.EqualValue(oData);
                        search.SearchConditions.Add(oSearchCondition);
                    }
                }
            }
            items = search.FindAll(documentControlM.Document, false);
            ////隐藏基础
            documentControlM.Document.Models.SetHidden(items, true);
            ProgressBarStatus.Value++;
            Dispatcher.Invoke(new Action(() =>
            {
            }), System.Windows.Threading.DispatcherPriority.Background);

            //显示所有土建辅助模型
            search.Clear();
            search.Selection.SelectAll();
            foreach (var itemBuildingModels in qBuildingModels)//显示查询结果
            {
                if (itemBuildingModels.Field<string>("辅助模型") != null)
                {
                    if (itemBuildingModels.Field<string>("辅助模型") != "")
                    {
                        oData = VariantData.FromDisplayString(itemBuildingModels.Field<string>("辅助模型"));

                        SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "名称");
                        oSearchCondition = oSearchCondition.EqualValue(oData);
                        search.SearchConditions.Add(oSearchCondition);
                    }
                }
            }
            items = search.FindAll(documentControlM.Document, false);
            //显示辅助
            documentControlM.Document.Models.SetHidden(items.DescendantsAndSelf, false);



            
        }

        private string getPropertyValueString(ModelItem item, string sCategories, string sProperty)
        {
            string sValue="";

            if (item.PropertyCategories.FindPropertyByDisplayName(sCategories, sProperty) != null)
                sValue = item.PropertyCategories.FindPropertyByDisplayName(sCategories, sProperty).Value.ToDisplayString();

            return sValue;
        }

        //动态刷新房间属性
        private void ButtonRoomMonitor_Click(object sender, RoutedEventArgs e)
        {
            LayoutRoomMoniter.IsVisible = true;

            Fluent.Button fb = sender as Fluent.Button;

            if (timerRoomMonitor.Enabled)
            {
                timerRoomMonitor.Stop();
                sMonitorRoomCode = ""; sMonitorRoomName = "";
                RoomMoniter.Children.Clear();
                bMonitorRoomFirst = true;
                fb.Foreground = new SolidColorBrush(Colors.Black);
            }
            else
            {
                if(InitRoomCurve())
                  fb.Foreground = new SolidColorBrush(Colors.LightGray);
            }

        }

        //动态曲线初始化
        private bool InitRoomCurve()
        {
            string sCode = "";
            int i = 0;

            if (timerRoomMonitor.Enabled)
                timerRoomMonitor.Stop();

            listChartWD.Clear(); listChartSD.Clear(); listChartCO2.Clear(); listChartPM25.Clear();

            if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
            {
                MessageBox.Show("请选择需要显示的房间", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            ModelItem oSelectedItem = documentControlM.Document.CurrentSelection.SelectedItems.ElementAt<ModelItem>(0);
            if (selectItemType(oSelectedItem, out sCode) != nodeType.ROOM)
            {
                MessageBox.Show("请选择显示的房间", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            sMonitorRoomCode = sCode;
            //找到房间名
            //得到和房间名称
            var qRoom = from dtRoom in dSet.Tables["房间表"].AsEnumerable()//查询楼层
                        where (dtRoom.Field<string>("房间编号") == sCode)//条件
                        select dtRoom;

            foreach (var itemRoom in qRoom)//显示查询结果
            {
                sMonitorRoomName = itemRoom.Field<string>("房间名称");
                break;
            }

            bMonitorRoomFirst = true;

            #region 曲线图
            decimal dWD=0, dSD=0, dCO2=0, dPM25;
            
            if (getRoomMoniterPara(sMonitorRoomCode, out dWD, out dSD, out dCO2))
            {
                DateTime dtNow = DateTime.Now;
                for (i = 0; i < iMONITERNUMBER; i++)
                {
                    ClassListEA cleaWD = new ClassListEA();
                    cleaWD.sName = dtNow.AddSeconds(-1 * (iMONITERNUMBER - 1 - i)).ToString();
                    cleaWD.sValue = dWD.ToString();
                    listChartWD.Add(cleaWD);

                    ClassListEA cleaSD = new ClassListEA();
                    cleaSD.sName = cleaWD.sName;
                    cleaSD.sValue = dSD.ToString();
                    listChartSD.Add(cleaSD);

                    ClassListEA cleaCO2 = new ClassListEA();
                    cleaCO2.sName = cleaWD.sName;
                    cleaCO2.sValue = dCO2.ToString();
                    listChartCO2.Add(cleaCO2);
                }
            }

            if (getRoomMoniterParaPM25(sMonitorRoomCode, out dPM25))
            {
                DateTime dtNow = DateTime.Now;
                for (i = 0; i < iMONITERNUMBER; i++)
                {
                    ClassListEA cleaMP25 = new ClassListEA();
                    cleaMP25.sName = dtNow.AddSeconds(-1 * (iMONITERNUMBER - 1 - i)).ToString();
                    cleaMP25.sValue = dPM25.ToString();
                    listChartPM25.Add(cleaMP25);
                }
            }


            //加入温度曲线图 
            Chart chartSplineWD = new Chart();
            //设置图标的宽度和高度
            chartSplineWD.Width = 460;
            chartSplineWD.Height = 210;

            chartSplineWD.Margin = new Thickness(1, 1, 1, 1);
            //是否启用打印和保持图片
            chartSplineWD.ToolBarEnabled = false;

            //设置图标的属性
            chartSplineWD.ScrollingEnabled = false;//是否启用或禁用滚动
            chartSplineWD.View3D = false;//3D效果显示

            //创建一个标题的对象
            Title titlechartSplineWD = new Title();

            //设置标题的名称
            titlechartSplineWD.Text = sMonitorRoomName + ":温度";
            titlechartSplineWD.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chartSplineWD.Titles.Add(titlechartSplineWD);

            //初始化一个新的Axis
            Axis xaxisSplineWD = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplineWD.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplineWD.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplineWD.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis            
            chartSplineWD.AxesX.Add(xaxisSplineWD);

            Axis yAxisSplineWD = new Axis();
            //设置图标中Y轴的最小值永远为0           
            //yAxisSplineWD.AxisMinimum = -10;
            //yAxisSplineWD.AxisMaximum = 23.5;
            //yAxisSplineWD.AxisMinimum = 22.5;
            //yAxisSplineWD.Interval = 0.5;
            yAxisSplineWD.AxisMaximum = (double)dWD + 1;
            yAxisSplineWD.AxisMinimum = (double)dWD - 1;
            yAxisSplineWD.Interval = 0.5;
            //设置图表中Y轴的后缀          
            yAxisSplineWD.Suffix = "度";
            chartSplineWD.AxesY.Add(yAxisSplineWD);

            DataPoint dataPointSplineWD;
            DataSeries dataSeriesSplineWD = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplineWD.LegendText = "";

            dataSeriesSplineWD.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplineWD.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartWD.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplineWD = new DataPoint();
                // 设置X轴点                    
                dataPointSplineWD.XValue = DateTime.Parse(listChartWD[i].sName);
                //设置Y轴点                   
                dataPointSplineWD.YValue = double.Parse(listChartWD[i].sValue);
                dataPointSplineWD.MarkerSize = listChartWD.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplineWD.Color = new SolidColorBrush(Colors.LightGray);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplineWD.DataPoints.Add(dataPointSplineWD);
            }

            // 添加数据线到数据序列。 
            dataSeriesSplineWD.Color = new SolidColorBrush(Colors.LightGray);
            chartSplineWD.Series.Add(dataSeriesSplineWD);

            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            Grid grStackedSplineWD = new Grid();
            grStackedSplineWD.Children.Add(chartSplineWD);
            RoomMoniter.Children.Add(grStackedSplineWD);
            grStackedSplineWD.VerticalAlignment = VerticalAlignment.Top;


            //加入湿度曲线图 
            Chart chartSplineSD = new Chart();
            //设置图标的宽度和高度
            chartSplineSD.Width = 460;
            chartSplineSD.Height = 210;

            chartSplineSD.Margin = new Thickness(1, 1, 1, 1);
            //是否启用打印和保持图片
            chartSplineSD.ToolBarEnabled = false;

            //设置图标的属性
            chartSplineSD.ScrollingEnabled = false;//是否启用或禁用滚动
            chartSplineSD.View3D = false;//3D效果显示

            //创建一个标题的对象
            Title titlechartSplineSD = new Title();

            //设置标题的名shi
            titlechartSplineSD.Text = sMonitorRoomName + ":湿度";
            titlechartSplineSD.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chartSplineSD.Titles.Add(titlechartSplineSD);

            //初始化一个新的Axis
            Axis xaxisSplineSD = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplineSD.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplineSD.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplineSD.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis            
            chartSplineSD.AxesX.Add(xaxisSplineSD);

            Axis yAxisSplineSD = new Axis();
            //设置图标中Y轴的最小值永远为0           
            //yAxisSplineSD.AxisMinimum = 43;
            yAxisSplineSD.AxisMinimum = (double)dSD-2.0;
            yAxisSplineSD.AxisMaximum = (double)dSD + 2.0;
            //yAxisSplineSD.AxisMaximum = 44;
            yAxisSplineSD.Interval = 1;
            //设置图表中Y轴的后缀          
            yAxisSplineSD.Suffix = "";
            chartSplineSD.AxesY.Add(yAxisSplineSD);

            DataPoint dataPointSplineSD;
            DataSeries dataSeriesSplineSD = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplineSD.LegendText = "";

            dataSeriesSplineSD.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplineSD.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartSD.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplineSD = new DataPoint();
                // 设置X轴点                    
                dataPointSplineSD.XValue = DateTime.Parse(listChartSD[i].sName);
                //设置Y轴点                   
                dataPointSplineSD.YValue = double.Parse(listChartSD[i].sValue);
                dataPointSplineSD.MarkerSize = listChartSD.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplineSD.Color = new SolidColorBrush(Colors.LightGray);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplineSD.DataPoints.Add(dataPointSplineSD);
            }

            // 添加数据线到数据序列。           
            dataSeriesSplineSD.Color = new SolidColorBrush(Colors.LightGray);
            chartSplineSD.Series.Add(dataSeriesSplineSD);

            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            Grid grStackedSplineSD = new Grid();
            grStackedSplineSD.Children.Add(chartSplineSD);
            RoomMoniter.Children.Add(grStackedSplineSD);
            grStackedSplineSD.VerticalAlignment = VerticalAlignment.Top;


            //加入co2线图 
            Chart chartSplineCO2 = new Chart();

            //设置图标的宽度和高度
            chartSplineCO2.Width = 460;
            chartSplineCO2.Height = 210;

            chartSplineCO2.Margin = new Thickness(1, 1, 1, 1);
            //是否启用打印和保持图片
            chartSplineCO2.ToolBarEnabled = false;

            //设置图标的属性
            chartSplineCO2.ScrollingEnabled = false;//是否启用或禁用滚动
            chartSplineCO2.View3D = false;//3D效果显示

            //创建一个标题的对象
            Title titlechartSplineCO2 = new Title();

            //设置标题的名称
            titlechartSplineCO2.Text = sMonitorRoomName + ":CO2";
            titlechartSplineCO2.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chartSplineCO2.Titles.Add(titlechartSplineCO2);

            //初始化一个新的Axis
            Axis xaxisSplineCO2 = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplineCO2.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplineCO2.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplineCO2.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis            
            chartSplineCO2.AxesX.Add(xaxisSplineCO2);

            Axis yAxisSplineCO2 = new Axis();
            //设置图标中Y轴的最小值永远为0           
            //yAxisSplineCO2.AxisMinimum = 560;
            yAxisSplineCO2.AxisMinimum = (double)dCO2-20;
            yAxisSplineCO2.AxisMaximum = (double)dCO2 + 10;
            //yAxisSplineCO2.AxisMaximum = 600;
            yAxisSplineCO2.Interval = 10;
            //设置图表中Y轴的后缀          
            yAxisSplineCO2.Suffix = "";
            chartSplineCO2.AxesY.Add(yAxisSplineCO2);

            DataPoint dataPointSplineCO2;
            DataSeries dataSeriesSplineCO2 = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplineCO2.LegendText = "";

            dataSeriesSplineCO2.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplineCO2.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartCO2.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplineCO2 = new DataPoint();
                // 设置X轴点                    
                dataPointSplineCO2.XValue = DateTime.Parse(listChartCO2[i].sName);
                //设置Y轴点                   
                dataPointSplineCO2.YValue = double.Parse(listChartCO2[i].sValue);
                dataPointSplineCO2.MarkerSize = listChartCO2.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplineCO2.Color = new SolidColorBrush(Colors.LightGray);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplineCO2.DataPoints.Add(dataPointSplineCO2);
            }

            // 添加数据线到数据序列。    
            dataSeriesSplineCO2.Color = new SolidColorBrush(Colors.LightGray);
            chartSplineCO2.Series.Add(dataSeriesSplineCO2);

            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            Grid grStackedSplineCO2 = new Grid();
            grStackedSplineCO2.Children.Add(chartSplineCO2);
            RoomMoniter.Children.Add(grStackedSplineCO2);
            grStackedSplineCO2.VerticalAlignment = VerticalAlignment.Top;



            //加入PM25线图 
            Chart chartSplinePM25 = new Chart();

            //设置图标的宽度和高度
            chartSplinePM25.Width = 460;
            chartSplinePM25.Height = 210;

            chartSplinePM25.Margin = new Thickness(1, 1, 1, 1);
            //是否启用打印和保持图片
            chartSplinePM25.ToolBarEnabled = false;

            //设置图标的属性
            chartSplinePM25.ScrollingEnabled = false;//是否启用或禁用滚动
            chartSplinePM25.View3D = false;//3D效果显示

            //创建一个标题的对象
            Title titlechartSplinePM25 = new Title();

            //设置标题的名称
            titlechartSplinePM25.Text = sMonitorRoomName + ":PM2.5";
            titlechartSplinePM25.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chartSplinePM25.Titles.Add(titlechartSplinePM25);

            //初始化一个新的Axis
            Axis xaxisSplinePM25 = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒
            xaxisSplinePM25.IntervalType = IntervalTypes.Seconds;
            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxisSplinePM25.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20           
            xaxisSplinePM25.ValueFormatString = "MM-dd hh:mm:ss";
            //给图标添加Axis            
            chartSplinePM25.AxesX.Add(xaxisSplinePM25);

            Axis yAxisSplinePM25 = new Axis();
            //设置图标中Y轴的最小值永远为0           
            //yAxisSplinePM25.AxisMinimum = 0;
            yAxisSplinePM25.AxisMinimum = (double)dPM25-50;
            yAxisSplinePM25.AxisMaximum = (double)dPM25 + 50;
            yAxisSplinePM25.Interval = 25;

            //设置图表中Y轴的后缀          
            yAxisSplinePM25.Suffix = "μg/M3";
            chartSplinePM25.AxesY.Add(yAxisSplinePM25);

            DataPoint dataPointSplinePM25;
            DataSeries dataSeriesSplinePM25 = new DataSeries();
            // 创建一个新的数据线。               
            dataSeriesSplinePM25.LegendText = "";

            dataSeriesSplinePM25.RenderAs = RenderAs.Spline;//折线图
            dataSeriesSplinePM25.XValueType = ChartValueTypes.DateTime;

            for (i = 0; i < listChartPM25.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPointSplinePM25 = new DataPoint();
                // 设置X轴点                    
                dataPointSplinePM25.XValue = DateTime.Parse(listChartPM25[i].sName);
                //设置Y轴点                   
                dataPointSplinePM25.YValue = double.Parse(listChartPM25[i].sValue);
                dataPointSplinePM25.MarkerSize = listChartPM25.Count;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPointSplinePM25.Color = new SolidColorBrush(Colors.LightGray);
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeriesSplinePM25.DataPoints.Add(dataPointSplinePM25);
            }

            // 添加数据线到数据序列。
            dataSeriesSplinePM25.Color = new SolidColorBrush(Colors.LightGray);
            chartSplinePM25.Series.Add(dataSeriesSplinePM25);

            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            Grid grStackedSplinePM25 = new Grid();
            grStackedSplinePM25.Children.Add(chartSplinePM25);
            RoomMoniter.Children.Add(grStackedSplinePM25);
            grStackedSplinePM25.VerticalAlignment = VerticalAlignment.Top;


            Dispatcher.Invoke(new Action(() =>
            {
            }), System.Windows.Threading.DispatcherPriority.Background);
            #endregion

            //updateUIRoomMonitor();
            timerRoomMonitor.Start();

            return true;
        }

        //临时，得到房间一体监控参数。房间编号，温度，湿度，CO2，返回： 成功 true 失败 false
        private bool getRoomMoniterPara(string sRoomCode, out decimal dWD, out decimal dSD, out decimal dCO2)
        {
            dWD = 0; dSD = 0; dCO2 = 0;

            string sFacilityCode = "";
            //得到一体机编号
            var qFacilityControlled = from dtFacilityControlled in dSet.Tables["房间设备监控表"].AsEnumerable()//查询空间
                                      where (dtFacilityControlled.Field<string>("房间编号") == sRoomCode) && (dtFacilityControlled.Field<string>("设备编号").StartsWith("THC"))//条件
                                      select dtFacilityControlled;
            if (qFacilityControlled.Count() < 1)
                return false;

            foreach (var itemFacilityControlled in qFacilityControlled)
            {
                sFacilityCode = itemFacilityControlled.Field<string>("设备编号");
                break;
            }

            //得到设备
            var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询
                            where (dtFacility.Field<string>("设备编号") == sFacilityCode) //条件
                            select dtFacility;

            int iSystem = 0;//系统ID
            string sOBIX = "";
            foreach (var itemFacility in qFacility)
            {
                iSystem = itemFacility.Field<int>("系统ID");
                sOBIX = itemFacility.Field<string>("OBIX站点");
                break;
            }

            //得到子系统数据源
            string sDataSourceLocation = "";
            //筛选子系统
            var qSystemSub = from dtSystemSub in dSet.Tables["子系统表"].AsEnumerable()//查询
                             where (dtSystemSub.Field<int>("子系统ID") == iSystem)//条件
                             select dtSystemSub;
            foreach (var itemSystemSub in qSystemSub)//显示查询结果
            {
                sDataSourceLocation = itemSystemSub.Field<string>("数据源地址");
                break;
            }
            sDataSourceLocation = sDataSourceLocation.Replace(sREPLACE, sOBIX);
            HttpWebResponseUtility.tagUrl = sDataSourceLocation + sFacilityCode + "/";

            //string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponse();
            string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponseTransmit();
            if (responseFromServer == null || responseFromServer == "")
                return false;

            XmlDocument xmlDoc = new XmlDocument();
            XDocument xDoc = new XDocument();
            XmlReader xmlReader = xDoc.CreateReader();

            xmlDoc.LoadXml(responseFromServer);
            XElement xEle = XElement.Parse(responseFromServer);
            xmlDoc.PreserveWhitespace = false;

            var text = from t in xEle.Elements()//定位到节点 
                        .Where(w => w.Attribute("name").Value.Equals("RetAirTemp"))   //若要筛选就用上这个语句 
                       select new
                       {
                           disp = t.Attribute("display").Value   //注意此处用到 attribute 
                       };
            foreach (var sValue in text)
            {
                dWD = decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp));
                break;
            }

            text = from t in xEle.Elements()//定位到节点 
            .Where(w => w.Attribute("name").Value.Equals("RetAirHum"))   //若要筛选就用上这个语句 
                   select new
                   {
                       disp = t.Attribute("display").Value   //注意此处用到 attribute 
                   };
            foreach (var sValue in text)
            {
                dSD = decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp));
                break;
            }

            text = from t in xEle.Elements()//定位到节点 
            .Where(w => w.Attribute("name").Value.Equals("CO2"))   //若要筛选就用上这个语句 
                   select new
                   {
                       disp = t.Attribute("display").Value   //注意此处用到 attribute 
                   };
            foreach (var sValue in text)
            {
                dCO2 = decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp));
                break;
            }
            return true;
        }

        //临时，得到房间一体监控参数。房间编号，PM25，返回： 成功 true 失败 false
        private bool getRoomMoniterParaPM25(string sRoomCode, out decimal dPM25)
        {
            dPM25 = 0;

            string sFacilityCode = "";
            //得到一体机编号
            var qFacilityControlled = from dtFacilityControlled in dSet.Tables["房间设备监控表"].AsEnumerable()//查询空间
                                      where (dtFacilityControlled.Field<string>("房间编号") == sRoomCode) && (dtFacilityControlled.Field<string>("设备编号").StartsWith("PM25"))//条件
                                      select dtFacilityControlled;
            if (qFacilityControlled.Count() < 1)
                return false;

            foreach (var itemFacilityControlled in qFacilityControlled)
            {
                sFacilityCode = itemFacilityControlled.Field<string>("设备编号");
                break;
            }

            //得到设备
            var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询
                            where (dtFacility.Field<string>("设备编号") == sFacilityCode) //条件
                            select dtFacility;

            int iSystem = 0;//系统ID
            string sOBIX = "";
            foreach (var itemFacility in qFacility)
            {
                iSystem = itemFacility.Field<int>("系统ID");
                sOBIX = itemFacility.Field<string>("OBIX站点");
                break;
            }

            //得到子系统数据源
            string sDataSourceLocation = "";
            //筛选子系统
            var qSystemSub = from dtSystemSub in dSet.Tables["子系统表"].AsEnumerable()//查询
                             where (dtSystemSub.Field<int>("子系统ID") == iSystem)//条件
                             select dtSystemSub;
            foreach (var itemSystemSub in qSystemSub)//显示查询结果
            {
                sDataSourceLocation = itemSystemSub.Field<string>("数据源地址");
                break;
            }
            sDataSourceLocation = sDataSourceLocation.Replace(sREPLACE, sOBIX);
            HttpWebResponseUtility.tagUrl = sDataSourceLocation + sFacilityCode + "/";

            //string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponse();
            string responseFromServer = HttpWebResponseUtility.CreateGetHttpResponseTransmit();
            if (responseFromServer == null || responseFromServer == "")
                return false;

            XmlDocument xmlDoc = new XmlDocument();
            XDocument xDoc = new XDocument();
            XmlReader xmlReader = xDoc.CreateReader();

            xmlDoc.LoadXml(responseFromServer);
            XElement xEle = XElement.Parse(responseFromServer);
            xmlDoc.PreserveWhitespace = false;

            var text = from t in xEle.Elements()//定位到节点 
                        .Where(w => w.Attribute("name").Value.Equals("PM25Value"))   //若要筛选就用上这个语句 
                       select new
                       {
                           disp = t.Attribute("display").Value   //注意此处用到 attribute 
                       };
            foreach (var sValue in text)
            {
                dPM25 = decimal.Parse(HttpWebResponseUtility.GetResultObix(sValue.disp));
                break;
            }

            return true;
        }

        //方向：正面（90，0，0）
        private void ButtonViewCubeFront_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();

                Point3D oPos = oCurrVCopy.Position;
                Vector3D oViewDir = getViewDir(oCurrVCopy);

                //AxisAndAngleResult r = oCurrVCopy.Rotation.ToAxisAndAngle();

                UnitVector3D UV3d1 = new UnitVector3D(1, 0, 0);
                Rotation3D r1 = new Rotation3D(UV3d1, PI / 2);

                //UnitVector3D UV3d2 = new UnitVector3D(0, 0, 1);
                //Rotation3D r2 = new Rotation3D(UV3d2, PI);

                //Rotation3D mlrResult = MultiplyRotation3D(r1, r2);
                oCurrVCopy.Rotation = r1;


                if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
                    oCurrVCopy.ZoomBox(boundingBox3DALL);
                else
                    oCurrVCopy.ZoomBox(documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox());

                documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);
            }
            catch
            {
            }
        }

        //方向：背面（-90，0，180）
        private void ButtonViewCubeBack_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();

                Point3D oPos = oCurrVCopy.Position;
                Vector3D oViewDir = getViewDir(oCurrVCopy);

                //AxisAndAngleResult r = oCurrVCopy.Rotation.ToAxisAndAngle();

                UnitVector3D UV3d1 = new UnitVector3D(1, 0, 0);
                Rotation3D r1 = new Rotation3D(UV3d1, -PI / 2);

                UnitVector3D UV3d2 = new UnitVector3D(0, 0, 1);
                Rotation3D r2 = new Rotation3D(UV3d2, PI);

                Rotation3D mlrResult = MultiplyRotation3D(r1, r2);
                oCurrVCopy.Rotation = mlrResult;

                
                if (documentControlM.Document.CurrentSelection.SelectedItems.Count<1)
                    oCurrVCopy.ZoomBox(boundingBox3DALL);
                else
                    oCurrVCopy.ZoomBox(documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox());

                documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);

            }
            catch
            {
            }
        }

        //方向：左面（90，-90，0）
        private void ButtonViewCubeLeft_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();

                Point3D oPos = oCurrVCopy.Position;
                Vector3D oViewDir = getViewDir(oCurrVCopy);

                //AxisAndAngleResult r = oCurrVCopy.Rotation.ToAxisAndAngle();

                UnitVector3D UV3d1 = new UnitVector3D(1, 0, 0);
                Rotation3D r1 = new Rotation3D(UV3d1, PI / 2);

                UnitVector3D UV3d2 = new UnitVector3D(0,1, 0);
                Rotation3D r2 = new Rotation3D(UV3d2, -PI/2);

                Rotation3D mlrResult = MultiplyRotation3D(r1, r2);
                oCurrVCopy.Rotation = mlrResult;


                if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
                    oCurrVCopy.ZoomBox(boundingBox3DALL);
                else
                    oCurrVCopy.ZoomBox(documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox());

                documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);
            }
            catch
            {
            }
        }

        //方向：右面（90，90，0）
        private void ButtonViewCubeRight_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();

                Point3D oPos = oCurrVCopy.Position;
                Vector3D oViewDir = getViewDir(oCurrVCopy);

                //AxisAndAngleResult r = oCurrVCopy.Rotation.ToAxisAndAngle();

                UnitVector3D UV3d1 = new UnitVector3D(1, 0, 0);
                Rotation3D r1 = new Rotation3D(UV3d1, PI / 2);

                UnitVector3D UV3d2 = new UnitVector3D(0, 1, 0);
                Rotation3D r2 = new Rotation3D(UV3d2, PI / 2);

                Rotation3D mlrResult = MultiplyRotation3D(r1, r2);
                oCurrVCopy.Rotation = mlrResult;


                if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
                    oCurrVCopy.ZoomBox(boundingBox3DALL);
                else
                    oCurrVCopy.ZoomBox(documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox());

                documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);
            }
            catch
            {
            }

        }

        //方向：顶面（0，0，0）
        private void ButtonViewCubeTop_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();

                Point3D oPos = oCurrVCopy.Position;
                Vector3D oViewDir = getViewDir(oCurrVCopy);

                //AxisAndAngleResult r = oCurrVCopy.Rotation.ToAxisAndAngle();

                UnitVector3D UV3d1 = new UnitVector3D(0, 0, 1);
                Rotation3D r1 = new Rotation3D(UV3d1, 0);

                //UnitVector3D UV3d2 = new UnitVector3D(0, 1, 0);
                //Rotation3D r2 = new Rotation3D(UV3d2, PI / 2);

                //Rotation3D mlrResult = MultiplyRotation3D(r1, r2);
                oCurrVCopy.Rotation = r1;


                if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
                    oCurrVCopy.ZoomBox(boundingBox3DALL);
                else
                    oCurrVCopy.ZoomBox(documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox());

                documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);
            }
            catch
            {
            }
        }

        //方向：底面（180，0，180）
        private void ButtonViewCubeBottom_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Viewpoint oCurrVCopy = documentControlM.Document.CurrentViewpoint.CreateCopy();

                Point3D oPos = oCurrVCopy.Position;
                Vector3D oViewDir = getViewDir(oCurrVCopy);

                //AxisAndAngleResult r = oCurrVCopy.Rotation.ToAxisAndAngle();

                UnitVector3D UV3d1 = new UnitVector3D(1, 0, 0);
                Rotation3D r1 = new Rotation3D(UV3d1, PI);

                UnitVector3D UV3d2 = new UnitVector3D(0, 0, 1);
                Rotation3D r2 = new Rotation3D(UV3d2, PI);

                Rotation3D mlrResult = MultiplyRotation3D(r1, r2);
                oCurrVCopy.Rotation = mlrResult;


                if (documentControlM.Document.CurrentSelection.SelectedItems.Count < 1)
                    oCurrVCopy.ZoomBox(boundingBox3DALL);
                else
                    oCurrVCopy.ZoomBox(documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox());

                documentControlM.Document.CurrentViewpoint.CopyFrom(oCurrVCopy);
            }
            catch
            {
            }
        }


        private void B1_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show(uiStyle.bSelectStyle.ToString());

            Search search = new Search();
            search.Selection.SelectAll();
          
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");


            oSearchCondition = oSearchCondition.DisplayStringContains("313037");
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = search.FindAll(documentControlM2D.Document, false);

            documentControlM2D.Document.CurrentSelection.Clear();
            documentControlM2D.Document.CurrentSelection.CopyFrom(items);

            Viewpoint v = documentControlM2D.Document.CurrentViewpoint;
            BoundingBox3D box = documentControlM2D.Document.CurrentSelection.SelectedItems.BoundingBox();

            v.ZoomBox(box);

            documentControlM2D.Document.CurrentViewpoint.CopyFrom(v);

            mainModelView2D.IsSelected = true;


            

           
        }

        private void view2Ddrawing_Click(object sender, RoutedEventArgs e)
        {
            string sCode;
            string sFile="";

            //获得选择的构件
            if (documentControlM.Document.CurrentSelection.IsEmpty)
            {
                MessageBox.Show("没有选择");
                return;
            }

            foreach (ModelItem oSelectedItem in documentControlM.Document.CurrentSelection.SelectedItems)
            {
                //获取图纸文件
                switch(selectItemType(oSelectedItem, out sCode))
                {
                    case nodeType.FACILITY:
                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询设备
                                     where (dtFacility.Field<string>("设备编号") == sCode)//条件
                                     select dtFacility;

                        foreach (var itemFacility in qFacility)//显示查询结果
                        {
                            sFile = itemFacility.Field<string>("二维图纸");
                            break;
                        }

                        break;
                    default:
                        sFile="";
                        break;
                }

                //打开文件
                if(sFile!="")
                {
                    sFile=sDirDrawings + "/" + sFile;

                    if (File.Exists(sFile)) //有文件
                    //viewPDF.LoadPDFfile(sFileDrawing);
                    {
                        if (documentControlM2D.Document != null)
                        {
                            if (documentControlM2D.Document.FileName.ToUpper() != sFile) //不同文件，调入
                            {
                                documentControlM2D.Document.TryOpenFile(sFile);
                            }
                        }
                        else //调入新文件
                        {
                            documentControlM2D.Document.TryOpenFile(sFile);
                        }
                    }

                }

                //得到ElementID
                DataProperty oDP_DWGHandle;
                int iEID=0;

                oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("元素 ID", "值");
                if(oDP_DWGHandle!=null)
                    iEID = int.Parse(oDP_DWGHandle.Value.ToDisplayString());

                //查询对应
                if (documentControlM2D.Document != null && iEID != 0)
                {
                    documentControlM2D.Document.CurrentSelection.Clear();
                    ModelItemCollection items = new ModelItemCollection();
                    Search search = new Search();
                    search.Selection.SelectAll();

                    SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
                    oSearchCondition = oSearchCondition.DisplayStringContains(iEID.ToString());
                    search.SearchConditions.Add(oSearchCondition);
                    ModelItemCollection items1 = search.FindAll(documentControlM2D.Document, false);

                    documentControlM2D.Document.CurrentSelection.CopyFrom(items1);

                    Viewpoint v = documentControlM2D.Document.CurrentViewpoint;
                    BoundingBox3D box = documentControlM2D.Document.CurrentSelection.SelectedItems.BoundingBox();

                    v.ZoomBox(box);

                    documentControlM2D.Document.CurrentViewpoint.CopyFrom(v);

                    mainModelView2D.IsSelected = true;
                }

                break; //只取一个
            }
        }

        private void view3DModel_Click(object sender, RoutedEventArgs e)
        {
            string sT = "";
            int i1, i2;
            int eId = 0;

            if (documentControlM2D.Document == null)
                return;

            //获得选择的构件
            if (documentControlM2D.Document.CurrentSelection.IsEmpty)
            {
                MessageBox.Show("没有选择");
                return;
            }

            DataProperty oDP_DWGHandle;
            ModelItemCollection items = new ModelItemCollection();
            Search search = new Search();

            foreach (ModelItem oSelectedItem in documentControlM2D.Document.CurrentSelection.SelectedItems)
            {
                eId = 0;
                oDP_DWGHandle = oSelectedItem.PropertyCategories.FindPropertyByDisplayName("项目", "类型");
                if (oDP_DWGHandle != null)
                {
                    sT = oDP_DWGHandle.Value.ToDisplayString();

                    //得到ID号
                    i1 = sT.IndexOf('[');
                    i2 = sT.IndexOf(']');

                    if (i1 != -1 && i2 != -1)
                    {
                        sT = sT.Substring(i1+1, i2 - i1-1);
                        if (!int.TryParse(sT, out eId))
                            eId = 0;
                    }

                    //查找模型中的对应ID
                    if (eId != 0)
                    {
                        search.Selection.SelectAll();
                        SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素 ID", "值");

                        oSearchCondition = oSearchCondition.EqualValue(VariantData.FromDisplayString(eId.ToString()));
                        search.SearchConditions.Add(oSearchCondition);
                        ModelItemCollection items1 = search.FindAll(documentControlM.Document, false);

                        items.AddRange(items1.ToArray());
                    }



                }
                
                break; //只取一个
            }

            if (items.Count() > 0)
            {
                documentControlM.Document.CurrentSelection.Clear();
                documentControlM.Document.CurrentSelection.CopyFrom(items);

                Viewpoint v = documentControlM.Document.CurrentViewpoint;
                BoundingBox3D box = documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox();

                v.ZoomBox(box);

                documentControlM.Document.CurrentViewpoint.CopyFrom(v);

                mainModelView.IsSelected = true;
            }


        }


        //poufen
        private void ButtonSectioning_Click(object sender, RoutedEventArgs e)
        {
            SavedItem siView = getViewPoint("111");
            if (siView == null)
                return ;
            documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;


        }

        //安选择项
        private void ButtonSectioningCreate_Click(object sender, RoutedEventArgs e)
        {
            if(documentControlM.Document.CurrentSelection.IsEmpty)
            {
                MessageBox.Show("没有选择项");
                return;
            }

            Autodesk.Navisworks.Api.Interop.ComApi.InwOpState10 state;
            state = ComApiBridge.State;

            BoundingBox3D Bbox = documentControlM.Document.CurrentSelection.SelectedItems.BoundingBox(true);

            // create a geometry vector as the normal of section plane
            Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f sectionPlaneNormalTop = (Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f)state.ObjectFactory(
                Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLUnitVec3f,null,null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f sectionPlaneNormalBottom = (Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f)state.ObjectFactory(
                Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLUnitVec3f,null,null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f sectionPlaneNormalLeft = (Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f)state.ObjectFactory(
                Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLUnitVec3f, null, null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f sectionPlaneNormalRight = (Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f)state.ObjectFactory(
                Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLUnitVec3f, null, null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f sectionPlaneNormalFront = (Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f)state.ObjectFactory(
                Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLUnitVec3f, null, null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f sectionPlaneNormalBack = (Autodesk.Navisworks.Api.Interop.ComApi.InwLUnitVec3f)state.ObjectFactory(
                Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLUnitVec3f, null, null);
            sectionPlaneNormalTop.SetValue(0, 0, -1);
            sectionPlaneNormalBottom.SetValue(0, 0, 1);
            sectionPlaneNormalLeft.SetValue(1, 0, 0);
            sectionPlaneNormalRight.SetValue(-1, 0, 0);
            sectionPlaneNormalFront.SetValue(0, 1, 0);
            sectionPlaneNormalBack.SetValue(0, -1, 0);

            // create a geometry plane
            Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f sectionPlaneTop = (Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f)state.ObjectFactory(Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLPlane3f,null,null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f sectionPlaneBottom = (Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f)state.ObjectFactory(Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLPlane3f,null,null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f sectionPlaneLeft = (Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f)state.ObjectFactory(Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLPlane3f, null, null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f sectionPlaneRight = (Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f)state.ObjectFactory(Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLPlane3f, null, null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f sectionPlaneFront = (Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f)state.ObjectFactory(Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLPlane3f, null, null);
            Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f sectionPlaneBack = (Autodesk.Navisworks.Api.Interop.ComApi.InwLPlane3f)state.ObjectFactory(Autodesk.Navisworks.Api.Interop.ComApi.nwEObjectType.eObjectType_nwLPlane3f, null, null);

            //get collection of sectioning planes
            Autodesk.Navisworks.Api.Interop.ComApi.InwClippingPlaneColl2 clipColl =
                (Autodesk.Navisworks.Api.Interop.ComApi.InwClippingPlaneColl2)state.CurrentView.ClippingPlanes();
            //clear plans
            clipColl.Clear();

            // get the count of current sectioning planes
            //int planeCount = clipColl.Count + 1;
            int planeCount = 1;

            // create a new sectioning plane
            // it forces creation of planes up to this index.
            clipColl.CreatePlane(1);
            // get the last sectioning plane which are what we created
            Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlaneTop =
                (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane)state.CurrentView.ClippingPlanes().Last();
            
            clipColl.CreatePlane(2);
            Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlaneBottom =
                (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane)state.CurrentView.ClippingPlanes().Last();

            clipColl.CreatePlane(3);
            Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlaneLeft =
                (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane)state.CurrentView.ClippingPlanes().Last();

            clipColl.CreatePlane(4);
            Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlaneRight =
                (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane)state.CurrentView.ClippingPlanes().Last();

            clipColl.CreatePlane(5);
            Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlaneFront =
                (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane)state.CurrentView.ClippingPlanes().Last();

            clipColl.CreatePlane(6);
            Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlaneBack =
                (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane)state.CurrentView.ClippingPlanes().Last();

            //assign the geometry vector with the plane

            sectionPlaneTop.SetValue(sectionPlaneNormalTop, -1*Bbox.Max.Z);
            sectionPlaneBottom.SetValue(sectionPlaneNormalBottom, Bbox.Min.Z);
            sectionPlaneLeft.SetValue(sectionPlaneNormalLeft, Bbox.Min.X);
            sectionPlaneRight.SetValue(sectionPlaneNormalRight, -1*Bbox.Max.X);
            sectionPlaneFront.SetValue(sectionPlaneNormalFront, Bbox.Min.Y);
            sectionPlaneBack.SetValue(sectionPlaneNormalBack, -1*Bbox.Max.Y);

            // ask the sectioning plane uses the new geometry plane
            cliPlaneTop.Plane = sectionPlaneTop;
            cliPlaneBottom.Plane = sectionPlaneBottom;
            cliPlaneLeft.Plane = sectionPlaneLeft;
            cliPlaneRight.Plane = sectionPlaneRight;
            cliPlaneFront.Plane = sectionPlaneFront;
            cliPlaneBack.Plane = sectionPlaneBack;

            // enable this sectioning plane
            cliPlaneTop.Enabled = true;
            cliPlaneBottom.Enabled = true;
            cliPlaneLeft.Enabled = true;
            cliPlaneRight.Enabled = true;
            cliPlaneFront.Enabled = true;
            cliPlaneBack.Enabled = true;

        }

        //清除剖分
        private void ButtonSectioningClear_Click(object sender, RoutedEventArgs e)
        {
            Autodesk.Navisworks.Api.Interop.ComApi.InwOpState10 state;
            state = ComApiBridge.State;

            Autodesk.Navisworks.Api.Interop.ComApi.InwClippingPlaneColl2 clipColl = (Autodesk.Navisworks.Api.Interop.ComApi.InwClippingPlaneColl2)state.CurrentView.ClippingPlanes();
            //clipColl.Clear();

            foreach (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlane in state.CurrentView.ClippingPlanes())
            {
                cliPlane.Enabled = false;
            }

        }

        private void ButtonSectioningEnable_Click(object sender, RoutedEventArgs e)
        {
            Autodesk.Navisworks.Api.Interop.ComApi.InwOpState10 state;
            state = ComApiBridge.State;

            Autodesk.Navisworks.Api.Interop.ComApi.InwClippingPlaneColl2 clipColl = (Autodesk.Navisworks.Api.Interop.ComApi.InwClippingPlaneColl2)state.CurrentView.ClippingPlanes();
            //clipColl.Clear();

            foreach (Autodesk.Navisworks.Api.Interop.ComApi.InwOaClipPlane cliPlane in state.CurrentView.ClippingPlanes())
            {
                cliPlane.Enabled = true;
            }
        }

        //控制设备
        private void miSelectControlFacility_Click(object sender, RoutedEventArgs e)
        {
            //int i, j;
            //FacilityIdenti fiControll = new FacilityIdenti();

            //if (treeviewSelectTree.SelectedItem == null)
            //    return;

            //TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            //if (tnSelectTree.sModel.Trim() == "")
            //    return;

            //if (tnSelectTree.sFacilityCode.Trim() == "")
            //    return;


            ////筛选控制设备
            //var qFacilityControll = from dtFacilityControlled in dSet.Tables["设备控制监控表"].AsEnumerable()//查询控制设备
            //                        where (dtFacilityControlled.Field<string>("设备编号") == tnSelectTree.sFacilityCode)//条件
            //                        select dtFacilityControlled;
            //if (qFacilityControll.Count()==0)
            //{
            //    MessageBox.Show("没有控制记录");
            //    return;
            //}

            //foreach (var itemFacilityControll in qFacilityControll) //只选择一个
            //{
            //    fiControll.iFacilityID = itemFacilityControll.Field<int>("控制设备ID");
            //    fiControll.sFacilityCode = itemFacilityControll.Field<string>("控制设备编号");

            //    break;
            //}

            ////得到设备
            //var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询控制设备
            //                where (dtFacility.Field<string>("设备编号") == fiControll.sFacilityCode)//条件
            //                select dtFacility;
            //if (qFacility.Count() == 0)
            //    return;

            //string sViewName = "";

            ////隐藏，显示所需的楼层
            //foreach (var itemFacility in qFacility) //只选择一个
            //{
            //    int iLevel=itemFacility.Field<int>("标高ID");
            //    if(iLevel==0)
            //        return;

            //    var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询标高
            //                 where (dtLevel.Field<int>("标高ID") == iLevel)//条件
            //                select dtLevel;
            //    if(qLevel.Count()==0)
            //        return;
            //    foreach (var itemLevel in qLevel) //只选择一个
            //    {
            //        showSelectedLevel(itemLevel.Field<string>("土建模型组成"), itemLevel.Field<string>("机电模型组成"), itemLevel.Field<double>("标高排序"), itemLevel.Field<int>("建筑ID"), true,true);                     
            //        break;
            //    }

            //    sViewName = itemFacility.Field<string>("视图");
            //    break;
            //}


            //if (sViewName.Trim() == "")
            //    return;

            //SavedItem siView = getViewPoint(sViewName);
            //if (siView == null)
            //    return;
            //documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;

            //Search search = new Search();
            //search.Selection.SelectAll();

            //VariantData oData;

            //oData = VariantData.FromDisplayString(fiControll.sFacilityCode);

            //SearchCondition oSearchCondition1 = SearchCondition.HasPropertyByDisplayName("元素", "设备编号");
            //oSearchCondition1 = oSearchCondition1.EqualValue(oData);
            //search.SearchConditions.Add(oSearchCondition1);

            //ModelItemCollection items1 = search.FindAll(documentControlM.Document, false);
            //documentControlM.Document.CurrentSelection.Clear();
            //documentControlM.Document.CurrentSelection.CopyFrom(items1);


            //多个控制设备，刷新代码
            int i, j;
            List<FacilityIdenti> liControlled = new List<FacilityIdenti>();


            if (treeviewSelectTree.SelectedItem == null)
                return;

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.sModel.Trim() == "")
                return;

            if (tnSelectTree.sFacilityCode.Trim() == "")
                return;


            //筛选控制设备
            var qFacilityControll = from dtFacilityControlled in dSet.Tables["设备控制监控表"].AsEnumerable()//查询控制设备
                                    where (dtFacilityControlled.Field<string>("设备编号") == tnSelectTree.sFacilityCode)//条件
                                    select dtFacilityControlled;
            if (qFacilityControll.Count() == 0)
            {
                MessageBox.Show("没有控制记录");
                return;
            }

            foreach (var itemFacilityControll in qFacilityControll) //选择
            {
                FacilityIdenti fiControll = new FacilityIdenti();
                fiControll.iFacilityID = itemFacilityControll.Field<int>("控制设备ID");
                fiControll.sFacilityCode = itemFacilityControll.Field<string>("控制设备编号");

                liControlled.Add(fiControll);

            }

            //得到设备楼层视图，选第一个
            var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询控制设备
                            where (dtFacility.Field<string>("设备编号") == liControlled[0].sFacilityCode)//条件
                            select dtFacility;
            if (qFacility.Count() == 0)
                return;

            string sViewName = "";

            //隐藏，显示所需的楼层
            foreach (var itemFacility in qFacility) //只选择一个
            {
                int iLevel = itemFacility.Field<int>("标高ID");
                if (iLevel == 0)
                    return;

                var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询标高
                             where (dtLevel.Field<int>("标高ID") == iLevel)//条件
                             select dtLevel;
                if (qLevel.Count() == 0)
                    return;
                foreach (var itemLevel in qLevel) //只选择一个
                {
                    showSelectedLevel(itemLevel.Field<string>("土建模型组成"), itemLevel.Field<string>("机电模型组成"), itemLevel.Field<double>("标高排序"), itemLevel.Field<int>("建筑ID"), true, true);
                    break;
                }

                sViewName = itemFacility.Field<string>("视图");
                break;
            }


            if (sViewName.Trim() == "")
                return;

            SavedItem siView = getViewPoint(sViewName);
            if (siView != null)
                //return;
                documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;


            Search search = new Search();
            search.Selection.SelectAll();

            VariantData oData;
            for (i = 0; i < liControlled.Count; i++)
            {
                if (liControlled[i].sFacilityCode == "") continue;

                System.Collections.Generic.List<SearchCondition> oG = new System.Collections.Generic.List<SearchCondition>();
                oData = VariantData.FromDisplayString(liControlled[i].sFacilityCode);


                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素", "设备编号");
                oSearchCondition = oSearchCondition.EqualValue(oData);
                oG.Add(oSearchCondition);
                search.SearchConditions.AddGroup(oG);
            }
            ModelItemCollection items = search.FindAll(documentControlM.Document, false);
            //不显示土建
            documentControlM.Document.Models.SetHidden(items, false);
            documentControlM.Document.CurrentSelection.Clear();
            documentControlM.Document.CurrentSelection.CopyFrom(items);




        }

        //受控制设备
        private void miSelectControledFacility_Click(object sender, RoutedEventArgs e)
        {

            int i, j;
            List<FacilityIdenti> liControlled = new List<FacilityIdenti>();
            

            if (treeviewSelectTree.SelectedItem == null)
                return;

            TreeNodetagSelectTree tnSelectTree = treeviewSelectTree.SelectedItem as TreeNodetagSelectTree;

            if (tnSelectTree.sModel.Trim() == "")
                return;

            if (tnSelectTree.sFacilityCode.Trim() == "")
                return;


            //筛选控制设备
            var qFacilityControll = from dtFacilityControlled in dSet.Tables["设备控制监控表"].AsEnumerable()//查询控制设备
                                    where (dtFacilityControlled.Field<string>("控制设备编号") == tnSelectTree.sFacilityCode)//条件
                                    select dtFacilityControlled;
            if (qFacilityControll.Count() == 0)
            {
                MessageBox.Show("没有控制记录");
                return;
            }

            foreach (var itemFacilityControll in qFacilityControll) //选择
            {
                FacilityIdenti fiControll = new FacilityIdenti();
                fiControll.iFacilityID = itemFacilityControll.Field<int>("设备ID");
                fiControll.sFacilityCode = itemFacilityControll.Field<string>("设备编号");

                liControlled.Add(fiControll);

            }

            //得到设备楼层视图，选第一个
            var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询控制设备
                            where (dtFacility.Field<string>("设备编号") == liControlled[0].sFacilityCode)//条件
                            select dtFacility;
            if (qFacility.Count() == 0)
                return;

            string sViewName = "";

            //隐藏，显示所需的楼层
            foreach (var itemFacility in qFacility) //只选择一个
            {
                int iLevel = itemFacility.Field<int>("标高ID");
                if (iLevel == 0)
                    return;

                var qLevel = from dtLevel in dSet.Tables["标高表"].AsEnumerable()//查询标高
                             where (dtLevel.Field<int>("标高ID") == iLevel)//条件
                             select dtLevel;
                if (qLevel.Count() == 0)
                    return;
                foreach (var itemLevel in qLevel) //只选择一个
                {
                    showSelectedLevel(itemLevel.Field<string>("土建模型组成"), itemLevel.Field<string>("机电模型组成"), itemLevel.Field<double>("标高排序"), itemLevel.Field<int>("建筑ID"), true, true);
                    break;
                }

                sViewName = itemFacility.Field<string>("视图");
                break;
            }


            if (sViewName.Trim() == "")
                return;

            SavedItem siView = getViewPoint(sViewName);
            if (siView != null)
                //return;
                documentControlM.Document.SavedViewpoints.CurrentSavedViewpoint = siView;


            Search search = new Search();
            search.Selection.SelectAll();

            VariantData oData;
            for (i = 0; i < liControlled.Count; i++)
            {
                if (liControlled[i].sFacilityCode == "") continue;

                System.Collections.Generic.List<SearchCondition> oG = new System.Collections.Generic.List<SearchCondition>();
                oData = VariantData.FromDisplayString(liControlled[i].sFacilityCode);


                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素", "设备编号");
                oSearchCondition = oSearchCondition.EqualValue(oData);
                oG.Add(oSearchCondition);
                search.SearchConditions.AddGroup(oG);
            }
            ModelItemCollection items = search.FindAll(documentControlM.Document, false);
            //不显示土建
            documentControlM.Document.Models.SetHidden(items, false);
            documentControlM.Document.CurrentSelection.Clear();
            documentControlM.Document.CurrentSelection.CopyFrom(items);


        }

        private void ButtonRefreshAll_Click(object sender, RoutedEventArgs e)
        {
            //int i,j;
            //string sT = "";
            //sqlConn.Open();

            //sqlComm.CommandText = "SELECT ID, 设备名称, 设备编号, 设备铭牌号, 类型, 厂商信息, 采购日期, 维护日期, 维护公司, 维护电话, 支持连接 FROM 设备特性表";
            //if (dSet.Tables.Contains("设备特性表1")) dSet.Tables.Remove("设备特性表1");
            //sqlDA.Fill(dSet, "设备特性表1");

            //for (i = 0; i < dSet.Tables["设备特性表1"].Rows.Count; i++)
            //{
            //    //查找设备表
            //    sqlComm.CommandText = "SELECT 设备ID, 设备名称, 设备编号, 原始ID FROM 设备表 WHERE (设备编号 = '" + dSet.Tables["设备特性表1"].Rows[i][2].ToString()+ "')";

            //    sqldr = sqlComm.ExecuteReader();
            //    if (!sqldr.HasRows)
            //    {
            //        sqldr.Close();
            //        continue;
            //    }

            //    sqldr.Read();
            //    sT=sqldr.GetValue(3).ToString(); //原始ID

            //    if (sT == "")
            //    {
            //        sqldr.Close();
            //        continue;
            //    }

            //    j = dSet.Tables["设备特性表1"].Rows[i][3].ToString().IndexOf(" ");
            //    sT = dSet.Tables["设备特性表1"].Rows[i][3].ToString().Substring(0, j) + " " + sT;

            //    sqldr.Close();
            //    sqlComm.CommandText = "UPDATE 设备特性表 SET 设备铭牌号 = N'" + sT + "' WHERE (设备编号 = N'" + dSet.Tables["设备特性表1"].Rows[i][2].ToString() + "')";
            //    sqlComm.ExecuteNonQuery();



            //}


            //sqlConn.Close();
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            //String url = "http://172.18.11.111/jci/WebServiceBIADBIMJCI.asmx";
            //String soapAction = "http://tempuri.org/HelloWorld1";

            //var soapClient = new SoapClient(url, soapAction);
            //soapClient.Arguments.Add(new SoapParameter("requestXml", "{'Head': { 'MethodCode': 'M1001', 'Security': { 'Token': ''}},'Body': { 'FlowID': 'ca0a9a91-bb13-4717-8590-d9258f5c292f'}}"));
            //Object ob = soapClient.GetResult();

            ClassWebResponseSoap.tagURL = "http://172.18.11.111/jci/WebServiceBIADBIMJCI.asmx";
            ClassWebResponseSoap.SOAPAction = "http://tempuri.org/HelloWorld";

            string responseFromServer = ClassWebResponseSoap.CreateGetHttpResponseSoap("1");
        }

        private void MainWindowNavisWorks_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                //S ystem.Windows.Forms.MessageBox.Show("111");
                System.Windows.Forms.Help.ShowHelp(null, System.Windows.Forms.Application.StartupPath+"\\bimhelp.chm");  
            }
        }

        private void WarnPOOLRefresh_Click(object sender, RoutedEventArgs e)
        {
            
            LabelWarnPOOL.Content = System.DateTime.Now.ToShortDateString() + "  " + System.DateTime.Now.ToShortTimeString();

            dtDWarnPool.Clear();
            // Respond to event...

            string responseFromServerSoap = ClassWebResponseSoap.CreateGetHttpResponseSoap_byPara("http://123.127.216.6/airmoniter/WebServicePM25WarnPOOL.asmx", "PM25WarnPool", "http://tempuri.org/PM25WarnPool");
            if (responseFromServerSoap == null || responseFromServerSoap == "")
                return;

            //XmlDocument xmlDocSoap = new XmlDocument();
            //XDocument xDocSoap = new XDocument();
            //XmlReader xmlReaderSoap = xDocSoap.CreateReader();
            //xmlDocSoap.LoadXml(responseFromServerSoap);
            //string s = xmlDocSoap.DocumentElement["env:Header"]["NotifySOAPHeader"].InnerXml;
            //xmlDocSoap.PreserveWhitespace = false;
            //responseFromServerSoap="<?xml version=\"1.0\" encoding=\"GB2312\" ?><getJCIInfoResult xmlns=\"http://tempuri.org/\"><sFacilityNo>CH/B2-1</sFacilityNo><dtReadingTime>2016/5/19 18:52:47</dtReadingTime></getJCIInfoResult>";

            XElement xEleSoap = XElement.Parse("<?xml version=\"1.0\" encoding=\"GB2312\" ?>" + responseFromServerSoap);


            var textsoap = from t in xEleSoap.Elements().Last().Elements()//定位到节点 
                           where t.Name.LocalName.Equals("Warn") //确定节点
                           select t;

            object[] oTemp = new object[3]; 
            foreach (var sValue in textsoap)
            {
                oTemp[2] = "0";
                foreach (var sValueget in sValue.Elements())
                {
                    switch (sValueget.Name.LocalName)
                    {
                        case "LevelName":
                            oTemp[0] = sValueget.Value;
                            break;
                        case "RoomNo":
                            oTemp[1] = sValueget.Value;
                            break;
                        case "Value":
                            oTemp[2] = sValueget.Value;
                            break;

                    }
                }
                dtDWarnPool.Rows.Add(oTemp);
            }


            DGWarnPOOL.ItemsSource = dtDWarnPool.AsDataView();

            if (!Window.GetWindow(DGWarnPOOL).IsVisible)
            {
                Window.GetWindow(DGWarnPOOL).Show();
            }
            DGWarnPOOL.UpdateLayout();

            //DataGridRow dgRow;
            //SolidColorBrush scbBackGround = null;
            //int i;
            //for (i = 0; i < DGWarnPOOL.ItemContainerGenerator.Items.Count; i++)
            //{
            //    dgRow = (DataGridRow)this.DGWarnPOOL.ItemContainerGenerator.ContainerFromIndex(i);
            //    switch (dtDGMmonitor.Rows[i][2].ToString())
            //    {
            //        case "1":
            //            scbBackGround = new SolidColorBrush(Colors.Red);
            //            break;
            //        case "-1":
            //            scbBackGround = new SolidColorBrush(Colors.Red);
            //            break;
            //        case "0":
            //            scbBackGround = new SolidColorBrush(Colors.LightGreen);
            //            break;
            //        case "100":
            //            scbBackGround = new SolidColorBrush(Colors.LightGray);
            //            break;
            //        case "200":
            //            scbBackGround = new SolidColorBrush(Colors.Gray);
            //            break;
            //    }
            //    //判断是否取到该项取值
            //    if (scbBackGround != null)
            //        dgRow.Background = scbBackGround;
            //}

        }

        private void ButtonRoomPM25_Click(object sender, RoutedEventArgs e)
        {
            LayoutPM25WarnPool.IsVisible = true;
            LayoutPM25WarnPool.Show();

            if (!Window.GetWindow(DGWarnPOOL).IsVisible)
            {
                Window.GetWindow(DGWarnPOOL).Show();
            }
            WarnPOOLRefresh_Click(null, null);
        }






    }


    [Plugin("ToolPluginSelect", "BIADBIM")]
    public class ToolPluginTest : ToolPlugin
    {
        public override bool MouseDown(
            //the current view when mouse down 
             Autodesk.Navisworks.Api.View view,
            //Enumerates key modifiers used in input: 
            //None, Ctrl,Alt,Shift 
             KeyModifiers modifiers,
            //left mouse button:1, 
            //middle mouse button:2,
            //right mouse button:3
             ushort button,
            //screen coordinate x
             int x,
            //screen coordinate y
             int y,
            // not clear to me :-(  
             double timeOffset)
        {
            // key modifiers used in input
            //left/middle mouse
            //timeOffset

            //ApplicationControl.SelectionBehavior = SelectionBehavior.FirstObject;
            //string str = modifiers.ToString() + "\n" +
            //              button.ToString() + "\n" +
            //              timeOffset.ToString();

            //System.Windows.Forms.MessageBox.Show(str);

            // get info of selecting
            
            PickItemResult itemResult = view.PickItemFromPoint(x, y);

            if (itemResult != null)
            {
                //selected point in WCS
                //string oStr = string.Format("拾取点的坐标 {0},{1},{2}", itemResult.Point.X,
                //                             itemResult.Point.Y,
                //                             itemResult.Point.Z);

                //System.Windows.Forms.MessageBox.Show(oStr);

                //selected object
                ModelItem modelItem = itemResult.ModelItem;
                view.Document.CurrentSelection.Clear();
                view.Document.CurrentSelection.Add(modelItem);

                Window mainwin = System.Windows.Application.Current.MainWindow;
                //MessageBox.Show(mainwin.Title, mainwin.Title);
                MainWindow mwindow = mainwin as MainWindow;
                mwindow.showFeature();
                //System.Windows.Forms.MessageBox.Show("拾取的模型ClassDisplayName:" + modelItem.ClassDisplayName);
            }
            

            return false;
        }


    } 





}
