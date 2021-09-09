using System;
using System.Globalization;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Net;
using System.Runtime.InteropServices;
using System.IO.Ports;
using System.Threading;
using System.Collections;
using System.Diagnostics;
using System.Reflection;
using System.Data.SqlClient;

namespace AutoInfoSystem
{
    public partial class FrmMain : Form
    {
        public List<Communication> totalCommunication;
        public List<string> totalMachine;
        public int totalCount;
        public bool isInit;
        //建立哈希表
        public Hashtable configParaHashtable;

        public Thread UpdateInfoToInternetThread = null;

        private SqlConnection mycon;
        private SqlCommand SqlCmd;

        MyServiceReference.WebServiceSoapClient dd;

        public FrmMain()
        {
            InitializeComponent();
            isInit = true;
            try
            {
                //totalCommunication.ForEach(ItemCheckEventArgs);
                //导出系统配置参数文件
                string configFilePath;
                configFilePath = AppDomain.CurrentDomain.BaseDirectory + "\\AutoInfoConfig.xml";//没有用@，就要用\\
                configParaHashtable = ImportSystemConfig(configFilePath);
                totalMachine = new List<string>();
                totalCommunication = new List<Communication>();
                // 参数初始化
                totalCount = int.Parse(configParaHashtable["totalCount"].ToString());
                for (int i = 0; i < totalCount; i++)
                {
                    string temp = configParaHashtable["machine" + i.ToString()].ToString();
                    totalMachine.Add(temp);
                }

                
            }
            catch
            {
                MessageBox.Show("配置文件出错，请修改配置文件后重新打开软件！");
                isInit = false;
            }

        }
        private void FrmMain_Load(object sender, EventArgs e)
        {
            if (!isInit)
            {
                this.Close();
                this.Dispose();
            }
            else
            {
                try
                {
                    for (int i = 0; i < totalCount; i++)
                    {
                        Communication one = new Communication(totalMachine[i]);
                        totalCommunication.Add(one);
                    }
                    //mycon = new SqlConnection("server=(local);uid=sa;database=hebei");
                    mycon = new SqlConnection("Data Source=HE-PC\\MSSQL2008;Initial Catalog=hebei;Integrated Security=True");
                    mycon.Open();


                    UpdateInfoToInternetThread = new Thread(new ThreadStart(UpdateInfoToInternetThreadFun));

                    UpdateInfoToInternetThread.Name = "UpdateInfoToInternetThread";

                    UpdateInfoToInternetThread.Start(); 


                }
                catch
                {
                    MessageBox.Show("FrmMain_Load Error");
                }

            }
        }
        /// <summary>
        ///更新数据到外网
        /// </summary>
        private void UpdateInfoToInternetThreadFun()
        {
            dd = new AutoInfoSystem.MyServiceReference.WebServiceSoapClient();
            while(true) 
            {
                //如果网络不通
                if (!UpdateAlarmInfoToInternet())
                {
                    Thread.Sleep(20000);
                    dd = new AutoInfoSystem.MyServiceReference.WebServiceSoapClient();
                }
                //如果网络不通
                if (!UpdateRunStatusInfoToInternet())
                {
                    Thread.Sleep(20000);
                    dd = new AutoInfoSystem.MyServiceReference.WebServiceSoapClient();
                }
                //如果网络不通
                if (!UpdatePartInfoToInternet())
                {
                    Thread.Sleep(20000);
                    dd = new AutoInfoSystem.MyServiceReference.WebServiceSoapClient();
                }
                //如果网络不通
                if (!UpdateMaterialInfoToInternet())
                {
                    Thread.Sleep(20000);
                    dd = new AutoInfoSystem.MyServiceReference.WebServiceSoapClient();
                }
                ////如果返回为FALSE,sleep20秒，这个暂时不用传外网
                //if (!UpdateNowStayusInfoToInternet())
                //{
                //    Thread.Sleep(20000);
                //}
                //网络通,间隔两秒循环一次
                Thread.Sleep(2000);
            }
        }
        private Hashtable ImportSystemConfig(string configFilePath)//导入系统配置
        {
            Hashtable parameterTable = new Hashtable();
            try
            {
                XMLOperator xmlOp = new XMLOperator();
                parameterTable = xmlOp.LoadFile(configFilePath);
            }
            catch
            {
                MessageBox.Show("配置文件不存在或错误，请检查！！");
            }
            return parameterTable;
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("确定要关闭该信息统计软件？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                DateTime endDt = DateTime.Now;
                for (int i = 0; i < totalCommunication.Count; i++)
                {
                    //确认PA是否真的运行
                    if (totalCommunication[i].fmsKeyMonitorThread != null)
                    {
                        totalCommunication[i].UpdateMachineStatus(totalCommunication[i].machineName, totalCommunication[i].lastMachineStatus.machineStatus, endDt, true);
                    }
                }
                try
                {
                    //写机床的实时状态
                    string strStatus = "update 机床实时状态 set 运行状态=255";
                    SqlCmd = new SqlCommand(strStatus, mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
                catch { }
                Process p = Process.GetCurrentProcess();
                if (p != null)
                {
                    p.Kill();
                }
            }
            else
            {
                e.Cancel = true;
                return;
            }
        }
        /// </summary>
        /// 更新报警信息到外网
        /// </summary>
        private bool UpdateAlarmInfoToInternet()
        {
            try
            {
                string total = null;
                string totalLocal = null;
                string str = "select * from 报警信息 where 上传网络='0'";
                SqlCmd = new SqlCommand(str, mycon);
                SqlDataReader DataReader = SqlCmd.ExecuteReader();
                while (DataReader.Read())
                {
                    //上传信息字符串
                    string info = "insert into 报警信息(机床名,报警号,报警时间,报警类型) values('" + DataReader[0].ToString() + "','" + DataReader[1].ToString() + "','" + DataReader[2].ToString() + "','" + DataReader[3].ToString() + "')";
                    total = total + info + ";";
                    //跟新本地信息字符串
                    string rrr = "update 报警信息 set 上传网络=1 where 机床名='" + DataReader[0].ToString() + "'and 报警号='" + DataReader[1].ToString() + "' and 报警时间='" + DataReader[2].ToString() + "' and 报警类型='" + DataReader[3].ToString() + "'";
                    totalLocal = totalLocal + rrr + ";";
                }
                DataReader.Close();
                DataReader.Dispose();
                if (dd.ExecSqlCmd(total) > 0 || total !=null)
                {
                    UpdateLocalSql(totalLocal);
                }
                return true;
            }
            catch 
            {
                return false;
            }

        }
        /// <summary>
        /// 更新运行状态到外网
        /// </summary>
        private bool UpdateRunStatusInfoToInternet()
        {
            try
            {
                string total = null;
                string totalLocal = null;
                string str = "select * from 机床运行状态 where 上传网络='0'";
                SqlCmd = new SqlCommand(str, mycon);
                SqlDataReader DataReader = SqlCmd.ExecuteReader();
                while (DataReader.Read())
                {
                    //上传信息字符串
                    string info = "insert into 机床运行状态(机床名,机床状态,起始时间,结束时间,持续时间) values('" + DataReader[0].ToString() + "','" + DataReader[1].ToString() + "','" + DataReader[2].ToString() + "','" + DataReader[3].ToString() + "','" + DataReader[4].ToString() + "')";
                    total = total + info + ";";
                    //跟新本地信息字符串
                    string rrr = "update 机床运行状态 set 上传网络=1 where 机床名='" + DataReader[0].ToString() + "'and 机床状态='" + DataReader[1].ToString() + "' and 起始时间='" + DataReader[2].ToString() + "' and 结束时间='" + DataReader[3].ToString() + "'and 持续时间='" + DataReader[4].ToString() + "'";
                    totalLocal = totalLocal + rrr + ";";
                }
                DataReader.Close();
                DataReader.Dispose();
                if (dd.ExecSqlCmd(total) > 0 || total!=null)
                {
                    UpdateLocalSql(totalLocal);
                }
                return false;
            }
            catch { return false; }
        }
        /// <summary>
        /// 更新运工件加工信息到外网
        /// </summary>
        private bool UpdatePartInfoToInternet()
        {
            try
            {
                string total = null;
                string totalLocal = null;
                string str = "select * from 工件加工信息 where 上传网络='0'";
                SqlCmd = new SqlCommand(str, mycon);
                SqlDataReader DataReader = SqlCmd.ExecuteReader();
                while (DataReader.Read())
                {
                    //上传信息字符串
                    string info = "insert into 工件加工信息(工件名称,材料,厚度,数量,加工时间,加工机床,NC名称) values('" + DataReader[0].ToString() + "','" + DataReader[1].ToString() + "','" + DataReader[2].ToString() + "','" + DataReader[3].ToString() + "','" + DataReader[4].ToString() + "','" + DataReader[5].ToString() + "','" + DataReader[6].ToString() + "')";
                    total = total + info + ";";
                    //跟新本地信息字符串
                    string rrr = "update 工件加工信息 set 上传网络=1 where 工件名称='" + DataReader[0].ToString() + "'and 材料='" + DataReader[1].ToString() + "' and 厚度='" + DataReader[2].ToString() + "' and 数量='" + DataReader[3].ToString() + "'and 加工时间='" + DataReader[4].ToString() + "'and 加工机床='" + DataReader[5].ToString() + "'and NC名称='" + DataReader[6].ToString() + "'";
                    totalLocal = totalLocal + rrr + ";";
                }
                DataReader.Close();
                DataReader.Dispose();

                if (dd.ExecSqlCmd(total) >0 || total!=null)
                {
                    UpdateLocalSql(totalLocal);

                }
                return true;
            }
            catch { return false; }
        }
        /// <summary>
        /// 更新运行状态到外网
        /// </summary>
        private bool UpdateMaterialInfoToInternet()
        {
            try
            {
                string total = null;
                string totalLocal = null;
                string str = "select * from 板材信息表 where 上传网络='0'";
                SqlCmd = new SqlCommand(str, mycon);
                SqlDataReader DataReader = SqlCmd.ExecuteReader();
                while (DataReader.Read())
                {
                    //上传信息字符串
                    string info = "insert into 板材信息表(材料,长度,宽度,厚度,加工时间,加工机床,NC名称) values('" + DataReader[0].ToString() + "','" + DataReader[1].ToString() + "','" + DataReader[2].ToString() + "','" + DataReader[3].ToString() + "','" + DataReader[4].ToString() + "','" + DataReader[5].ToString() + "','" + DataReader[6].ToString() + "')";
                    total = total + info + ";";
                    //跟新本地信息字符串
                    string rrr = "update 板材信息表 set 上传网络=1 where 材料='" + DataReader[0].ToString() + "'and 长度='" + DataReader[1].ToString() + "' and 宽度='" + DataReader[2].ToString() + "' and 厚度='" + DataReader[3].ToString() + "'and 加工时间='" + DataReader[4].ToString() + "'and 加工机床='" + DataReader[5].ToString() + "'and NC名称='" + DataReader[6].ToString() + "'";
                    totalLocal = totalLocal + rrr + ";";
                }
                DataReader.Close();
                DataReader.Dispose();

                if (dd.ExecSqlCmd(total) >0 || total!=null)
                {
                    UpdateLocalSql(totalLocal);
                }
                return true;
            }
            catch { return false; }
        }
        /// <summary>
        /// 更新机床实时状态到外网
        /// </summary>
        private void UpdateNowStayusInfoToInternet()
        {
            //try
            //{
            //    string str = "select * from 机床实时状态";
            //    SqlCmd = new SqlCommand(str, mycon);
            //    SqlDataReader DataReader = SqlCmd.ExecuteReader();
            //    if (true)
            //    {
            //        while (DataReader.Read())
            //        {
            //            //如果上传成功，更新上传网络1
            //            string dd = "insert into 机床实时状态(机床名,运行状态) values('" + DataReader[0].ToString() + "','" + DataReader[1].ToString() + "')";
            //            //if (上传网络成功)
            //            //{
            //            //这个比较特殊

            //            //}
            //            //else
            //            //{
            //            break;
            //            //}
            //            //网络不通直接break返回
            //        }
            //        DataReader.Close();
            //        DataReader.Dispose();
            //    }
            //}
            //catch { }
        }
        /// <summary>
        /// 跟新本地的数据库
        /// </summary>
        /// <param name="str"></param>
        private void UpdateLocalSql(string str)
        {
            try
            {
                SqlCmd = new SqlCommand(str, mycon);
                SqlCmd.ExecuteNonQuery();
                SqlCmd.Dispose();
            }
            catch
            {
                //MessageBox.Show("UpdateLocalSql error!");
                mycon = new SqlConnection("Data Source=HE-PC\\MSSQL2008;Initial Catalog=hebei;Integrated Security=True");

                //mycon = new SqlConnection("server=(local);uid=sa;database=hebei");
                mycon.Open();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string total = null;
            string totalLocal = null;
            string str = "select * from 报警信息 where 上传网络='0'";
            SqlCmd = new SqlCommand(str, mycon);
            SqlDataReader DataReader = SqlCmd.ExecuteReader();
            while (DataReader.Read())
            {
                //上传信息字符串
                string info = "insert into 报警信息(机床名,报警号,报警时间,报警类型) values('" + DataReader[0].ToString() + "','" + DataReader[1].ToString() + "','" + DataReader[2].ToString() + "','" + DataReader[3].ToString() + "')";
                total = total + info + ";";
                //跟新本地信息字符串
                string rrr = "update 报警信息 set 上传网络=1 where 机床名='" + DataReader[0].ToString() + "'and 报警号='" + DataReader[1].ToString() + "' and 报警时间='" + DataReader[2].ToString() + "' and 报警类型='" + DataReader[3].ToString() + "'";
                totalLocal = totalLocal + rrr + ";";
            }
            DataReader.Close();
            DataReader.Dispose();
            if (dd.ExecSqlCmd(total) >0)
            {
                UpdateLocalSql(totalLocal);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                dd = new AutoInfoSystem.MyServiceReference.WebServiceSoapClient();
                string total = null;
                string totalLocal = null;
                string str = "select * from 机床运行状态 where 上传网络='0'";
                SqlCmd = new SqlCommand(str, mycon);
                SqlDataReader DataReader = SqlCmd.ExecuteReader();
                while (DataReader.Read())
                {
                    //上传信息字符串
                    string info = "insert into 机床运行状态(机床名,机床状态,起始时间,结束时间,持续时间) values('" + DataReader[0].ToString() + "','" + DataReader[1].ToString() + "','" + DataReader[2].ToString() + "','" + DataReader[3].ToString() + "','" + DataReader[4].ToString() + "')";
                    total = total + info + ";";
                    //跟新本地信息字符串
                    string rrr = "update 机床运行状态 set 上传网络=1 where 机床名='" + DataReader[0].ToString() + "'and 机床状态='" + DataReader[1].ToString() + "' and 起始时间='" + DataReader[2].ToString() + "' and 结束时间='" + DataReader[3].ToString() + "'and 持续时间='" + DataReader[4].ToString() + "'";
                    totalLocal = totalLocal + rrr + ";";
                }
                DataReader.Close();
                DataReader.Dispose();
                if (dd.ExecSqlCmd(total) > 0 || total != null)
                {
                    UpdateLocalSql(totalLocal);
                }
            }
            catch { }
        }
    }
}
