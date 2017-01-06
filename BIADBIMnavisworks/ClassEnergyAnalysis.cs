using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace BIADBIMnavisworks
{
    class ClassEnergyAnalysis
    {
        //监控库Monitor
        private System.Data.SqlClient.SqlConnection sqlConnM = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlCommM = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldrM;
        private System.Data.SqlClient.SqlDataAdapter sqlDAM = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSetM = new DataSet();
        public string strConnM = "";

        public List<ClassListEA> listEA = new List<ClassListEA>();

        private List<string> listLevel = new List<string>() { "B2", "B1", "B1夹", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11" };
        //private List<string> listLevel = new List<string>() { "B2", "B1", "B1夹", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10"};
        private List<string> listDepa = new List<string>() { "职能", "一院", "二院", "三院", "四院", "其他（公用）"};
        public List<ClassListLevelDapa> listLD = new List<ClassListLevelDapa>();

        //设备表
        public System.Data.DataSet dSet;


        //初始化楼层对应
        public void Init()
        {
            listLD.Clear();

            ClassListLevelDapa ld1 = new ClassListLevelDapa("B2", "其他（公用）");
            listLD.Add(ld1);

            ClassListLevelDapa ld2 = new ClassListLevelDapa("B1", "其他（公用）");
            listLD.Add(ld2);

            ClassListLevelDapa ld3 = new ClassListLevelDapa("B1夹", "其他（公用）");
            listLD.Add(ld3);

            ClassListLevelDapa ld4 = new ClassListLevelDapa("F1", "职能");
            listLD.Add(ld4);

            ClassListLevelDapa ld5 = new ClassListLevelDapa("F2", "职能");
            listLD.Add(ld5);

            ClassListLevelDapa ld6 = new ClassListLevelDapa("F3", "四院");
            listLD.Add(ld6);

            ClassListLevelDapa ld7 = new ClassListLevelDapa("F4", "四院");
            listLD.Add(ld7);

            ClassListLevelDapa ld8 = new ClassListLevelDapa("F5", "三院");
            listLD.Add(ld8);

            ClassListLevelDapa ld9 = new ClassListLevelDapa("F6", "三院");
            listLD.Add(ld9);

            ClassListLevelDapa ld10 = new ClassListLevelDapa("F7", "二院");
            listLD.Add(ld10);

            ClassListLevelDapa ld11 = new ClassListLevelDapa("F8", "二院");
            listLD.Add(ld11);

            ClassListLevelDapa ld12 = new ClassListLevelDapa("F9", "一院");
            listLD.Add(ld12);

            ClassListLevelDapa ld13 = new ClassListLevelDapa("F10", "一院");
            listLD.Add(ld13);

            ClassListLevelDapa ld14 = new ClassListLevelDapa("F11", "其他（公用）");
            listLD.Add(ld14);

            sqlConnM.ConnectionString = strConnM;
            sqlCommM.Connection = sqlConnM;
            sqlDAM.SelectCommand = sqlCommM;
        }


        //电量按层取值
        public void getLevelEnergyAnalysis_Ele()
        {
            double dSum=0;
            //清除
            listEA.Clear();

            try
            {
                foreach (string sLevel in listLevel)
                {
                    dSum = 0;
                    //取得此层电表
                    var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                    where (dtFacility.Field<string>("设备系统") == "电表") && (dtFacility.Field<string>("标高名称") == sLevel)//条件
                                    select dtFacility;

                    if (qFacility.Count() < 1) //没有电表，继续
                        continue;

                    foreach (var itemFacility in qFacility)
                    {
                        sqlConnM.Open();
                        //取得电表的值
                        sqlCommM.CommandText = "SELECT TOP(1) value FROM " + itemFacility.Field<string>("设备编号") + "span ORDER BY time DESC";
                        sqldrM = sqlCommM.ExecuteReader();
                        if (sqldrM.HasRows)
                        {
                            sqldrM.Read();
                            dSum += double.Parse(sqldrM.GetValue(0).ToString())/10;
                        }

                        if (!sqldrM.IsClosed)
                            sqldrM.Close();
                        sqlConnM.Close();
                    }

                    ////////////////////////////////////数值错误
                    if (sLevel == "F9" || sLevel == "F10")
                        dSum *= 10;

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = sLevel; classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);
                }

            }
            catch(Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:"+e.Message,"错误提示",System.Windows.MessageBoxButton.OK,System.Windows.MessageBoxImage.Error);
            }
            finally
            {

            }

        }

        //电量按层取值
        public void getLevelEnergyAnalysis_1(string sSystem)
        {
            double dSum = 0;
            //清除
            listEA.Clear();

            try
            {
                foreach (string sLevel in listLevel)
                {
                    dSum = 0;
                    //取得此层电表
                    var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                    where (dtFacility.Field<string>("设备系统") == sSystem) && (dtFacility.Field<string>("标高名称") == sLevel)//条件
                                    select dtFacility;

                    if (qFacility.Count() < 1) //没有电表，继续
                        continue;

                    foreach (var itemFacility in qFacility)
                    {
                        sqlConnM.Open();
                        //取得电表的值
                        sqlCommM.CommandText = "SELECT MAX(value) FROM " + itemFacility.Field<string>("设备编号") + "span";
                        sqldrM = sqlCommM.ExecuteReader();
                        if (sqldrM.HasRows)
                        {
                            sqldrM.Read();
                            dSum += double.Parse(sqldrM.GetValue(0).ToString()) / 10;
                        }

                        if (!sqldrM.IsClosed)
                            sqldrM.Close();
                        sqlConnM.Close();
                    }

                    ////////////////////////////////////数值错误
                    //if (sLevel == "F9" || sLevel == "F10")
                    //    dSum *= 10;

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = sLevel; classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);
                }

            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:" + e.Message, "错误提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {

            }

        }

        //电量按部门取值
        public void getDepartmentEnergyAnalysis_Ele()
        {
            double dSum = 0;
            //清除
            listEA.Clear();

            try
            {
                foreach (string sDepartment in listDepa)
                {
                    dSum = 0;
                    //找到部门的所有楼层
                    var qLevel = from dtLevel in listLD//查询楼层
                                 where (dtLevel.sDepartment == sDepartment)//条件
                                 select dtLevel;


                    foreach (ClassListLevelDapa sLevel in qLevel)
                    {

                        //取得此层电表
                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                        where (dtFacility.Field<string>("设备系统") == "电表") && (dtFacility.Field<string>("标高名称") == sLevel.sLevel)//条件
                                        select dtFacility;

                        if (qFacility.Count() < 1) //没有电表，继续
                            continue;

                        foreach (var itemFacility in qFacility)
                        {
                            sqlConnM.Open();
                            //取得电表的值
                            sqlCommM.CommandText = "SELECT TOP(1) value FROM " + itemFacility.Field<string>("设备编号") + "span ORDER BY time DESC";
                            sqldrM = sqlCommM.ExecuteReader();
                            if (sqldrM.HasRows)
                            {
                                sqldrM.Read();
                                if (itemFacility.Field<string>("标高名称") == "F9" || itemFacility.Field<string>("标高名称") == "F10")
                                    ////////////////////////////////////数值错误
                                    dSum += double.Parse(sqldrM.GetValue(0).ToString())*10/10;
                                else
                                    dSum += double.Parse(sqldrM.GetValue(0).ToString())/10;
                            }
                            if (!sqldrM.IsClosed)
                                sqldrM.Close();
                            sqlConnM.Close();



                        }

                    }

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = sDepartment; classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);
                }
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:" + e.Message, "错误提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {

            }
        }
        public void getDepartmentEnergyAnalysis_1(string sSystem)
        {
            double dSum = 0;
            //清除
            listEA.Clear();

            try
            {
                foreach (string sDepartment in listDepa)
                {
                    dSum = 0;
                    //找到部门的所有楼层
                    var qLevel = from dtLevel in listLD//查询楼层
                                 where (dtLevel.sDepartment == sDepartment)//条件
                                 select dtLevel;


                    foreach (ClassListLevelDapa sLevel in qLevel)
                    {

                        //取得此层电表
                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                        where (dtFacility.Field<string>("设备系统") == sSystem) && (dtFacility.Field<string>("标高名称") == sLevel.sLevel)//条件
                                        select dtFacility;

                        if (qFacility.Count() < 1) //没有电表，继续
                            continue;

                        foreach (var itemFacility in qFacility)
                        {
                            sqlConnM.Open();
                            //取得电表的值
                            sqlCommM.CommandText = "SELECT MAX(value) FROM " + itemFacility.Field<string>("设备编号") + "span";
                            sqldrM = sqlCommM.ExecuteReader();
                            if (sqldrM.HasRows)
                            {
                                sqldrM.Read();
                                //if (itemFacility.Field<string>("标高名称") == "F9" || itemFacility.Field<string>("标高名称") == "F10")
                                    ////////////////////////////////////数值错误
                                    dSum += double.Parse(sqldrM.GetValue(0).ToString()) * 10 / 10;
                                //else
                                //    dSum += double.Parse(sqldrM.GetValue(0).ToString()) / 10;
                            }
                            if (!sqldrM.IsClosed)
                                sqldrM.Close();
                            sqlConnM.Close();



                        }

                    }

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = sDepartment; classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);
                }
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:" + e.Message, "错误提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {

            }
        }

        //电量按部门取值
        public void getDepartmentEnergyAnalysis_Ele(string sSystem)
        {
            double dSum = 0;
            //清除
            listEA.Clear();

            try
            {
                foreach (string sDepartment in listDepa)
                {
                    dSum = 0;
                    //找到部门的所有楼层
                    var qLevel = from dtLevel in listLD//查询楼层
                                 where (dtLevel.sDepartment == sDepartment)//条件
                                 select dtLevel;


                    foreach (ClassListLevelDapa sLevel in qLevel)
                    {

                        //取得此层电表
                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                        where (dtFacility.Field<string>("设备系统") == sSystem) && (dtFacility.Field<string>("标高名称") == sLevel.sLevel)//条件
                                        select dtFacility;

                        if (qFacility.Count() < 1) //没有电表，继续
                            continue;

                        foreach (var itemFacility in qFacility)
                        {
                            sqlConnM.Open();
                            //取得电表的值
                            sqlCommM.CommandText = "SELECT TOP(1) value FROM " + itemFacility.Field<string>("设备编号") + "span ORDER BY time DESC";
                            sqldrM = sqlCommM.ExecuteReader();
                            if (sqldrM.HasRows)
                            {
                                sqldrM.Read();
                                if (itemFacility.Field<string>("标高名称") == "F9" || itemFacility.Field<string>("标高名称") == "F10")
                                    ////////////////////////////////////数值错误
                                    dSum += double.Parse(sqldrM.GetValue(0).ToString()) * 10 / 10;
                                else
                                    dSum += double.Parse(sqldrM.GetValue(0).ToString()) / 10;
                            }
                            if (!sqldrM.IsClosed)
                                sqldrM.Close();
                            sqlConnM.Close();



                        }

                    }

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = sDepartment; classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);
                }
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:" + e.Message, "错误提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {

            }
        }

        //电量当前月度统计
        public void getDepartmentEnergyNowMonthAnalysis_Ele()
        {
            double dSum = 0,dTemp2=0,dTemp1=0;
            DateTime dtNow=DateTime.Now,dtLastMonth,dtLastMonth_1;

            try
            {
                //得到日期
                sqlConnM.Open();
                sqlCommM.CommandText = "SELECT GETDATE() AS Date";
                sqldrM = sqlCommM.ExecuteReader();
                if (sqldrM.HasRows)
                {
                    sqldrM.Read();
                    dtNow = DateTime.Parse(sqldrM.GetValue(0).ToString());
                }
                sqldrM.Close();
                sqlConnM.Close();
                dtLastMonth = dtNow.AddMonths(-1);

                //清除
                listEA.Clear();
                //sqlConnM.Open();
                foreach (string sDepartment in listDepa)
                {
                    dSum = 0;
                    //找到部门的所有楼层
                    var qLevel = from dtLevel in listLD//查询楼层
                                 where (dtLevel.sDepartment == sDepartment)//条件
                                 select dtLevel;

                    foreach (ClassListLevelDapa sLevel in qLevel)
                    {
                        //取得此层电表
                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                        where (dtFacility.Field<string>("设备系统") == "电表") && (dtFacility.Field<string>("标高名称") == sLevel.sLevel)//条件
                                        select dtFacility;

                        if (qFacility.Count() < 1) //没有电表，继续
                            continue;

                        foreach (var itemFacility in qFacility)
                        {
                            dTemp1 = 0; dTemp2 = 0;
                            sqlConnM.Open();
                            //取得电表的当月大值
                            sqlCommM.CommandText = "SELECT TOP (1) value, time FROM " + itemFacility.Field<string>("设备编号") + "span WHERE (time LIKE N'%" + dtNow.Year.ToString() + "-" + dtNow.Month.ToString() + "%') ORDER BY time DESC";
                            sqldrM = sqlCommM.ExecuteReader();
                            if (sqldrM.HasRows)
                            {
                                sqldrM.Read();
                                dTemp1 = double.Parse(sqldrM.GetValue(0).ToString());
                            }
                            sqldrM.Close();

                            //取得电表的当月小值
                            sqlCommM.CommandText = "SELECT TOP (1) value, time FROM " + itemFacility.Field<string>("设备编号") + "span WHERE (time LIKE N'%" + dtNow.Year.ToString() + "-" + dtNow.Month.ToString() + "%') ORDER BY time";
                            sqldrM = sqlCommM.ExecuteReader();
                            if (sqldrM.HasRows)
                            {
                                sqldrM.Read();
                                dTemp2 = double.Parse(sqldrM.GetValue(0).ToString());
                            }
                            sqldrM.Close();
                            sqlConnM.Close();

                            dTemp1-=dTemp2;
                            if (dTemp1 < 0)
                                dTemp1 = 0;

                            dSum += dTemp1;


                        }

                    }

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = sDepartment; classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);
                }
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:" + e.Message, "错误提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                if (sqlConnM.State == System.Data.ConnectionState.Open)
                {
                    sqlConnM.Close();
                }
            }
        }

        //电量上月度统计
        public void getDepartmentEnergyLastMonthAnalysis_Ele()
        {
            double dSum = 0, dTemp2 = 0, dTemp1 = 0;
            DateTime dtNow = DateTime.Now, dtLastMonth;

            try
            {
                //得到日期
                sqlConnM.Open();
                sqlCommM.CommandText = "SELECT GETDATE() AS Date";
                sqldrM = sqlCommM.ExecuteReader();
                if (sqldrM.HasRows)
                {
                    sqldrM.Read();
                    dtNow = DateTime.Parse(sqldrM.GetValue(0).ToString());
                }
                sqldrM.Close();
                sqlConnM.Close();
                dtLastMonth = dtNow.AddMonths(-1);

                //清除
                listEA.Clear();
                //sqlConnM.Open();
                foreach (string sDepartment in listDepa)
                {
                    dSum = 0;
                    //找到部门的所有楼层
                    var qLevel = from dtLevel in listLD//查询楼层
                                 where (dtLevel.sDepartment == sDepartment)//条件
                                 select dtLevel;

                    foreach (ClassListLevelDapa sLevel in qLevel)
                    {
                        //取得此层电表
                        var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询楼层
                                        where (dtFacility.Field<string>("设备系统") == "电表") && (dtFacility.Field<string>("标高名称") == sLevel.sLevel)//条件
                                        select dtFacility;

                        if (qFacility.Count() < 1) //没有电表，继续
                            continue;

                        foreach (var itemFacility in qFacility)
                        {
                            dTemp1 = 0; dTemp2 = 0;
                            sqlConnM.Open();
                            //取得电表的当月大值
                            sqlCommM.CommandText = "SELECT TOP (1) value, time FROM " + itemFacility.Field<string>("设备编号") + "span WHERE (time LIKE N'%" + dtLastMonth.Year.ToString() + "-" + dtLastMonth.Month.ToString().PadLeft(2, '0') + "%') ORDER BY time DESC";
                            sqldrM = sqlCommM.ExecuteReader();
                            if (sqldrM.HasRows)
                            {
                                sqldrM.Read();
                                dTemp1 = double.Parse(sqldrM.GetValue(0).ToString());
                            }
                            sqldrM.Close();

                            //取得电表的当月小值
                            sqlCommM.CommandText = "SELECT TOP (1) value, time FROM " + itemFacility.Field<string>("设备编号") + "span WHERE (time LIKE N'%" + dtLastMonth.Year.ToString() + "-" + dtLastMonth.Month.ToString().PadLeft(2, '0') + "%') ORDER BY time";
                            sqldrM = sqlCommM.ExecuteReader();
                            if (sqldrM.HasRows)
                            {
                                sqldrM.Read();
                                dTemp2 = double.Parse(sqldrM.GetValue(0).ToString());
                            }
                            sqldrM.Close();
                            sqlConnM.Close();

                            dTemp1 -= dTemp2;
                            if (dTemp1 < 0)
                                dTemp1 = 0;

                            dSum += dTemp1;


                        }

                    }

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = sDepartment; classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);
                }
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:" + e.Message, "错误提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                if (sqlConnM.State == System.Data.ConnectionState.Open)
                {
                    sqlConnM.Close();
                }
            }
        }

        //电量五日走势统计
        public void getDepartmentEnergyDaysAnalysis_Ele(int iDays)
        {
            double dSum = 0, dTemp2 = 0, dTemp1 = 0;
            DateTime dtNow = DateTime.Now;
            int i;

            try
            {
                //得到日期
                sqlConnM.Open();
                sqlCommM.CommandText = "SELECT GETDATE() AS Date";
                sqldrM = sqlCommM.ExecuteReader();
                if (sqldrM.HasRows)
                {
                    sqldrM.Read();
                    dtNow = DateTime.Parse(sqldrM.GetValue(0).ToString());
                }
                dtNow = dtNow.AddDays(-1); //不取当天值
                sqldrM.Close();
                sqlConnM.Close();

                //清除
                listEA.Clear();
                
                //取得电表
                for (i = 0; i < iDays; i++)
                {
                    var qFacility = from dtFacility in dSet.Tables["设备表"].AsEnumerable()//查询
                                    where (dtFacility.Field<string>("设备系统") == "电表")//条件
                                    select dtFacility;

                    if (qFacility.Count() < 1) //没有电表，返回
                        return;
                    dSum = 0;
                    foreach (var itemFacility in qFacility)
                    {
                        dTemp1 = 0; dTemp2 = 0;

                        sqlConnM.Open();
                        //取得电表的当天大值
                        sqlCommM.CommandText = "SELECT TOP (1) value, time FROM " + itemFacility.Field<string>("设备编号") + "span WHERE (time LIKE N'%" + dtNow.Year.ToString() + "-" + dtNow.Month.ToString().PadLeft(2, '0') + "-" + dtNow.Day.ToString().PadLeft(2, '0') + "%') ORDER BY time DESC";
                        sqldrM = sqlCommM.ExecuteReader();
                        if (sqldrM.HasRows)
                        {
                            sqldrM.Read();
                            dTemp1 = double.Parse(sqldrM.GetValue(0).ToString());
                        }
                        sqldrM.Close();

                        //取得电表的当月小值
                        sqlCommM.CommandText = "SELECT TOP (1) value, time FROM " + itemFacility.Field<string>("设备编号") + "span WHERE (time LIKE N'%" + dtNow.Year.ToString() + "-" + dtNow.Month.ToString().PadLeft(2, '0') + "-" + dtNow.Day.ToString().PadLeft(2, '0') + "%') ORDER BY time";
                        sqldrM = sqlCommM.ExecuteReader();
                        if (sqldrM.HasRows)
                        {
                            sqldrM.Read();
                            dTemp2 = double.Parse(sqldrM.GetValue(0).ToString());
                        }
                        sqldrM.Close();
                        sqlConnM.Close();

                        dTemp1 -= dTemp2;
                        if (dTemp1 < 0)
                            dTemp1 = 0;

                        dSum += dTemp1;
                        
                    }

                    ClassListEA classlea = new ClassListEA();
                    classlea.sName = dtNow.ToShortDateString(); classlea.sValue = dSum.ToString();
                    listEA.Add(classlea);

                    dtNow = dtNow.AddDays(-1);
                } //for
                
            } //try
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("数据读取错误:" + e.Message, "错误提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                if (sqlConnM.State == System.Data.ConnectionState.Open)
                {
                    sqlConnM.Close();
                }
            }
        }


    }

    //返回值，返回项名称，返回项取值
    class ClassListEA
    {
        public string sName="";
        public string sValue="";

        public ClassListEA()
        {
            sName = ""; sValue = "";
        }
    }

    //部门与楼层对应
    class ClassListLevelDapa
    {
        public string sLevel = "";
        public string sDepartment = "";

        public ClassListLevelDapa()
        {
            sLevel = ""; sDepartment = "";
        }

        public ClassListLevelDapa(string Level, string Department)
        {
            sLevel = Level; sDepartment = Department;
        }
    }
}
