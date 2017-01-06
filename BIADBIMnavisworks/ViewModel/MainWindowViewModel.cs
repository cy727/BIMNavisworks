using System;
using System.ComponentModel;
using System.Windows;
using System.Collections.ObjectModel;
using System.Collections.Generic;

using Microsoft.Practices.Prism.Commands;
using Microsoft.Practices.Prism.ViewModel;

using Xceed.Wpf.AvalonDock;
using Xceed.Wpf.AvalonDock.Layout;

namespace BIADBIMnavisworks.ViewModel
{
    class MainWindowViewModel : NotificationObject
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private DockingManager _dockingManager;

        public DelegateCommand<object> SetTextCommand { get; set; }
        public DelegateCommand ExitSystemCommand { get; set; }
        public DelegateCommand SetDockSelectTreeIsVisibleCommand { get; set; }

        public MainWindowViewModel()
        {
            _dockingManager = MainWindow.DockingManager;

            SetTextCommand = new DelegateCommand<object>(SetText);
            ExitSystemCommand = new DelegateCommand(this.OnExit);
            SetDockSelectTreeIsVisibleCommand = new DelegateCommand(this.SetDockSelectTreeIsVisible);
        }
        private void SetText(object obj)
        {
            System.Windows.MessageBox.Show("1111");
        }

        private void SetDockSelectTreeIsVisible()
        {
            if (_isVisible)
                _isVisible = false;
            else
                _isVisible = true;
        }


        #region IsVisible

        private bool _isVisible = true;
        public bool IsVisible
        {
            get { return _isVisible; }
            set
            {
                if (_isVisible != value)
                {
                    _isVisible = value;
                    RaisePropertyChanged("IsVisible");
                }
            }
        }

        #endregion


        #region Exit
        private void OnExit()
        {
            //MessageBoxResult result = MessageBox.Show("确定要退出系统吗?", "确认消息", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (MessageBox.Show("确定要退出系统吗?", "确认消息", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.OK)
            {
                Application.Current.Shutdown();
                var serializer = new Xceed.Wpf.AvalonDock.Layout.Serialization.XmlLayoutSerializer(_dockingManager);
                serializer.Serialize(@".\AvalonDock.config");
            }
        }


        #endregion

    }

    public enum nodeType : short　　//显示指定枚举的底层数据类型
    {
        UNKNOW = -1, //未知
        PROJECT = 0, //项目
        BUILDING = 1,　　//建筑
        LEVEL = 2,　　//标高
        ROOM = 3,   //房间
        FACILITY = 4,   //设备
        SYSTEM = 5,     //系统
    };

    public enum DataSourceType : short　　//显示指定枚举的数据源数据类型
    {
        NONE = -1,  //未知
        DATABASE = 0, //数据库
        OBIX = 1,　　//OBIX
        OPC = 2,　　//OPC
        SOAP = 3,   //Web SOAP
    };

    public class FacilityIdenti //设备标示
    {
        public int iFacilityID{ get; set; } //受控设备ID
        public string sFacilityName{ get; set; } //受控设备名称   
        public string sFacilityCode{ get; set; } //受控设备编号


        public FacilityIdenti()
        {
            iFacilityID = 0; sFacilityName = ""; sFacilityCode = "";
        }
    }

    public class TreeNodetagSelectTree : INotifyPropertyChanged
    {
        public nodeType nodetype { get; set; }  //节点类型
        public string sIcon { get; set; } //节点图标
        //public string EditIcon { get; set; }
        public string sDisplayName { get; set; } //显示名称
        public string sNote { get; set; } //注释
        public int iBuildingID { get; set; } //建筑ID
        public string sBuildingName { get; set; } //建筑名称
        public int iLevelID { get; set; } //标高ID
        public string sLevelName { get; set; } //标高名称
        public int iRoomID { get; set; } //房间ID
        public string sRoomName { get; set; } //房间名称   
        public string sRoomCode { get; set; } //房间编码
        public int iSystemID { get; set; } //系统ID
        public string sSystemName { get; set; } //系统名称
        public int iSystemSubID { get; set; } //子系统ID
        public string sSystemSubName { get; set; } //子系统名称
        public int iFacilityID { get; set; } //设备ID
        public string sFacilityName { get; set; } //设备名称   
        public string sFacilityCode { get; set; } //设备编号
        public string sModel { get; set; } //模型
        public string sView { get; set; } //模型视图
        public string sLevelModelAS { get; set; }//标高模型土建
        public string sLevelModelMEP { get; set; }//标高模型机电
        public double fLevelOrder { get; set; }//标高排序
        public int ElementID { get; set; }//revit Elelemnt ID
        public string sDrawing { get; set; } //图纸



        public nodeType nodetype_controlled{ get; set; }//受控类型
        public int iLevelID_controlled { get; set; } //受控标高ID
        public string sLevelName_controlled { get; set; } //受控标高名称
        public double fLevelOrder_controlled { get; set; }//受控标高排序
        public int iRoomID_controlled { get; set; } //受控房间ID
        public string sRoomName_controlled { get; set; } //受控房间名称   
        public string sRoomCode_controlled { get; set; } //受控房间编码
        public List<FacilityIdenti> listFacility_controlled { get; set; } //受控设备ID,多个
        public FacilityIdenti fiFacility_control { get; set; } //控制设备ID,一个
        public string sModel_controlled { get; set; } //模型
        public string sView_controlled { get; set; } //模型视图

        public ObservableCollection<TreeNodetagSelectTree> Children { get; set; }
        public TreeNodetagSelectTree Parent { get; set; }
        bool _IsSel = false;
        bool _IsExpanded = false;

        /// <summary>
        /// 这个属性的名称要与属性值的名称一模一样
        /// </summary>
        public bool IsSelected
        {
            get { return this._IsSel; }
            set
            {
                if (value != _IsSel)
                {

                    this._IsSel = value;
                    OnPropertyChanged("IsSelected");
                }
            }
        }
        /// <summary>
        /// 这个属性的名称要与属性值的名称一模一样
        /// </summary>
        public bool IsExpanded
        {
            get { return this._IsExpanded; }
            set
            {
                if (value != _IsExpanded)
                {

                    this._IsExpanded = value;
                    OnPropertyChanged("IsExpanded");
                }
            }
        }

        public TreeNodetagSelectTree()
        {
            Children = new ObservableCollection<TreeNodetagSelectTree>();
            Parent = null;

            nodetype = nodeType.UNKNOW;
            nodetype_controlled = nodeType.UNKNOW;

            iBuildingID = 0; sBuildingName = "";
            iLevelID = 0; sLevelName = ""; sLevelModelAS = ""; sLevelModelMEP = ""; fLevelOrder = 0;
            iRoomID = 0; sRoomName = ""; sRoomCode = "";
            iSystemID = 0; sSystemName = ""; iSystemSubID = 0; sSystemSubName = "";
            iFacilityID = 0; sFacilityName = ""; sFacilityCode = "";
            sDrawing = "";

            //受控
            iLevelID_controlled = 0; sLevelName_controlled = ""; fLevelOrder_controlled = 0;
            iRoomID_controlled = 0; sRoomName_controlled = ""; sRoomCode_controlled = "";
            listFacility_controlled=new List<FacilityIdenti>();
            fiFacility_control = new FacilityIdenti();
            
            IsExpanded = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }


    public class UIStyle : ViewModelBase
    {
        private string _sSelectStyle;
        private bool _bSelectStyle;
        public string sSelectStyle
        {
            get { return _sSelectStyle; }
            set { _sSelectStyle = value; RaisePropertyChanged("sSelectStyle"); }
        }

        public bool bSelectStyle
        {
            get { return _bSelectStyle; }
            set { _bSelectStyle = value; RaisePropertyChanged("bSelectStyle"); }
        }

    }
}