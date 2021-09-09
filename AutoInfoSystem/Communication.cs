using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace AutoInfoSystem
{
    public class Communication
    {
        public Thread fmsKeyMonitorThread = null;
        public Thread fmsPlcServerThread = null;
        public Thread fmsFileServerThread = null;

        private System.Timers.Timer GetMachineStatusTimer;
        /// <summary>
        /// 与CNC服务器Socket通讯
        /// </summary>
        private Socket clientCncSocket;
        private Socket SynclientCncSocket;
        private Socket clientPlcSocket;
        private Socket clientFileSocket;


        private SqlConnection mycon;
        private SqlCommand SqlCmd;

        private ManualResetEvent connectDone;
        private ManualResetEvent sendDone;
        private ManualResetEvent receiveDone;

        private ManualResetEvent connectDonePlc;
        private ManualResetEvent sendDonePlc;
        private ManualResetEvent receiveDonePlc;

        private ManualResetEvent connectDoneFile;
        private ManualResetEvent sendDoneFile;
        private ManualResetEvent receiveDoneFile;

        public lastMachineStatusInfo lastMachineStatus;


        private bool isActiveNCName;
        //private bool isOpenLight;
        private string activeNCName;
        private string NCPath;

        private string tempStr;

        public string machineIP;
        public string machineName;


        public List<partInfo> PartInfoList;
        public information NCInfo;

        private string NCDir;

        //初始化函数
        public Communication(string computerIP)
        {
            //CNC
            connectDone = new ManualResetEvent(false);
            sendDone = new ManualResetEvent(false);
            receiveDone = new ManualResetEvent(false);
            //PLC
            connectDonePlc = new ManualResetEvent(false);
            sendDonePlc = new ManualResetEvent(false);
            receiveDonePlc = new ManualResetEvent(false);

            //文件服务器
            connectDoneFile = new ManualResetEvent(false);
            sendDoneFile = new ManualResetEvent(false);
            receiveDoneFile = new ManualResetEvent(false);


            isActiveNCName = false;
            activeNCName = null;
            lastMachineStatus = new lastMachineStatusInfo();
            lastMachineStatus.machineStatus = null;
            this.machineIP = computerIP;
            try
            {
                string con = "Data Source=HE-PC\\MSSQL2008;Initial Catalog=hebei;Integrated Security=True";
                mycon = new SqlConnection(con);
                //Thread.Sleep(3000);//防止数据库服务器还没有启动
                if (mycon.State != ConnectionState.Open)
                {
                    mycon.Open();
                }
                SqlCmd = new SqlCommand("SELECT COUNT(*) FROM sys.databases where name='hebei'", mycon);
                SqlCmd.ExecuteNonQuery();
                if (0 == Convert.ToInt32(SqlCmd.ExecuteScalar()))
                {
                    string create = "create database hebei";
                    SqlCmd = new SqlCommand(create, mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                    mycon.Dispose();
                }
                else
                {
                    SqlCmd.Dispose();
                    mycon.Dispose();
                }

                mycon = new SqlConnection(con);
                mycon.Open();
                //判断数据库表是不是存在，不存在创建
                InitiTable();
            }
            catch
            {
                MessageBox.Show("连接数据库失败，请重新手动打开软件");
                return;
            }
            //首次启动监控CNC服务器线程
            fmsKeyMonitorThread = new Thread(new ThreadStart(fmsKeyMonitorThreadFun));
            fmsKeyMonitorThread.Name = "fmsKeyMonitorThread";
            fmsKeyMonitorThread.Start();

            // 首次启动监控PLC服务器线程
            fmsPlcServerThread = new Thread(new ThreadStart(fmsPlcServerThreadFun));
            fmsPlcServerThread.Name = "fmsPlcServerThread";
            fmsPlcServerThread.Start();

            // 首次启动监控PLC服务器线程
            fmsFileServerThread = new Thread(new ThreadStart(fmsFileServerThreadFun));
            fmsFileServerThread.Name = "fmsFileServerThread";
            fmsFileServerThread.Start();

            //定时连接到PA
            GetMachineStatusTimer = new System.Timers.Timer(2000);
            GetMachineStatusTimer.Elapsed += new System.Timers.ElapsedEventHandler(GetMachineStatusTimer_Elapsed);
            GetMachineStatusTimer.AutoReset = true;
            GetMachineStatusTimer.Enabled = true;

            NCDir = @"E:\";
        }
        public class information
        {
            public string fileName;        //NC程序名
            public string material;        //材料
            public string thickness;       //厚度
            public string length;          //长度
            public string width;           //宽度
            public string needCount;       //需求数    
        }
        public class NCCode
        {
            public int G; //G代码种类
            public float x;
            public float y;
            public float I;
            public float J;
        }
        public class partInfo
        {
            public string partName;
            public int partCount;
        }
        public class lastMachineStatusInfo
        {
            public string machineStatus;
            public string lastTime;
        }
        // State object for receiving data from remote device.
        public class StateObject
        {
            // Client socket.
            public Socket workSocket = null;
            // Size of receive buffer.
            public const int BufferSize = 256;
            // Receive buffer.
            public byte[] buffer = new byte[BufferSize];
            // Received data string.
            public StringBuilder sb = new StringBuilder();
        }
        /// <summary>
        /// 文件的socket发送信息
        /// </summary>
        /// <param name="client"></param>
        /// <param name="data"></param>
        private void FileSend(Socket client, String data)
        {
            try
            {
                // Convert the string data to byte data using ASCII encoding.
                byte[] byteData = Encoding.ASCII.GetBytes(data);

                // Begin sending the data to the remote device.
                client.BeginSend(byteData, 0, byteData.Length, 0,
                    new AsyncCallback(FileSendCallback), client);
            }
            catch
            {
                fmsFileServerThread = null;
            }
        }
        /// <summary>
        /// PLC的socket发送信息
        /// </summary>
        /// <param name="client"></param>
        /// <param name="data"></param>
        private void PlcSend(Socket client, String data)
        {
            try
            {
                // Convert the string data to byte data using ASCII encoding.
                byte[] byteData = Encoding.ASCII.GetBytes(data);

                // Begin sending the data to the remote device.
                client.BeginSend(byteData, 0, byteData.Length, 0,
                    new AsyncCallback(PlcSendCallback), client);
            }
            catch
            {
                fmsPlcServerThread = null;
            }
        }
        /// <summary>
        /// CNC的socket发送信息
        /// </summary>
        /// <param name="client"></param>
        /// <param name="data"></param>
        private void Send(Socket client, String data)
        {
            try
            {
                // Convert the string data to byte data using ASCII encoding.
                byte[] byteData = Encoding.ASCII.GetBytes(data);

                // Begin sending the data to the remote device.
                client.BeginSend(byteData, 0, byteData.Length, 0,
                    new AsyncCallback(SendCallback), client);
            }
            catch
            {
                fmsKeyMonitorThread = null;
            }
        }
        /// <summary>
        /// 文件异步发送信息回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void FileSendCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;

                // Complete sending the data to the remote device.
                int bytesSent = client.EndSend(ar);

                // Signal that all bytes have been sent.
                sendDoneFile.Set();
            }
            catch
            {
                fmsFileServerThread = null;
            }
        }
        /// <summary>
        /// PLC异步发送信息回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void PlcSendCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;

                // Complete sending the data to the remote device.
                int bytesSent = client.EndSend(ar);

                // Signal that all bytes have been sent.
                sendDonePlc.Set();
            }
            catch
            {
                fmsPlcServerThread = null;
            }
        }
        /// <summary>
        /// CNC异步发送信息回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void SendCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;

                // Complete sending the data to the remote device.
                int bytesSent = client.EndSend(ar);

                // Signal that all bytes have been sent.
                sendDone.Set();
            }
            catch
            {
                fmsKeyMonitorThread = null;
            }
        }
        /// <summary>
        /// CNC异步连接回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void ConnectCallback(IAsyncResult ar)
        {
            try
            {
                
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;
                // Complete the connection.
                client.EndConnect(ar);
                // Signal that the connection has been made.
                connectDone.Set();
            }
            catch
            {
                MessageBox.Show("PA连接失败，请检查网络！");
                fmsKeyMonitorThread = null;
            }
        }
        /// <summary>
        /// 文件服务器异步连接回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void FileConnectCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;
                // Complete the connection.
                client.EndConnect(ar);
                // Signal that the connection has been made.
                connectDoneFile.Set();
            }
            catch
            {
                fmsFileServerThread = null;
            }
        }
        /// <summary>
        /// 文件服务器异步接收的回调函数
        /// </summary>
        /// <param name="ar"></param>
        public void FileReceiveCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the state object and the client socket 
                // from the asynchronous state object.
                StateObject state = (StateObject)ar.AsyncState;
                Socket client = state.workSocket;

                // Read data from the remote device.
                string readStr;
                int bytesRead = client.EndReceive(ar);
                if (bytesRead > 0)
                {
                    readStr = Encoding.ASCII.GetString(state.buffer, 0, bytesRead);
                    if (readStr.Contains("<fload>"))
                    {
                        //向PA发送指令
                        FileSend(clientFileSocket, "<fload></fload>\n");

                        sendDoneFile.WaitOne();

                        StateObject stateFile = new StateObject();
                        stateFile.workSocket = clientFileSocket;
                        // Begin receiving the data from the remote device.
                        clientFileSocket.BeginReceive(stateFile.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(FileReceiveCallback), stateFile);
                        receiveDoneFile.WaitOne();

                    }
                    else
                    {
                        readStr = Encoding.ASCII.GetString(state.buffer, 0, bytesRead);
                        tempStr = tempStr + readStr;
                        if (readStr.Contains("M30"))
                        {
                            FileStream fs1 = new FileStream(NCDir + activeNCName, FileMode.Create, FileAccess.Write);
                            StreamWriter sw1 = new StreamWriter(fs1);
                            sw1.Write(tempStr);
                            tempStr = null;
                            sw1.Close();
                            fs1.Close();
                            sw1.Dispose();
                            fs1.Dispose();
                        }

                    }
                    //获取剩下的数据
                    client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(FileReceiveCallback), state);
                }
                else
                {

                    receiveDoneFile.Set();
                    fmsFileServerThread = null;
                    clientFileSocket = null;
                }
            }
            catch (Exception ex)
            {
                fmsFileServerThread = null;
                clientFileSocket = null;
            }
        }
        /// <summary>
        /// PLC异步连接回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void PlcConnectCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;
                // Complete the connection.
                client.EndConnect(ar);
                // Signal that the connection has been made.
                connectDonePlc.Set();
            }
            catch
            {
                fmsPlcServerThread = null;
            }
        }
        /// <summary>
        /// PLC异步接收的回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void PlcReceiveCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the state object and the client socket 
                // from the asynchronous state object.
                StateObject state = (StateObject)ar.AsyncState;
                Socket client = state.workSocket;

                // Read data from the remote device.
                string readStr;
                string tempStr;
                int bytesRead = client.EndReceive(ar);
                if (bytesRead > 0)
                {
                    readStr = Encoding.ASCII.GetString(state.buffer, 0, bytesRead);
					 //if (readStr.Contains("<FMS_CONTROL_Prg.Parr[583]"))
                     if (readStr.Contains("<.FMS>"))
                    {
                        //tempStr = readStr.Substring(readStr.IndexOf("<FMS_CONTROL_Prg.Parr[583]>") + 27);
                        tempStr = readStr.Substring(readStr.IndexOf("<.FMS>") + 6);
                        if (GetNumFromStr(tempStr) == "1")
                        {

                            //读取NC文件工件信息,要考虑路径的问题
                            if (GetNCPartInfo(NCDir + activeNCName))
                            {
                                //工件信息写入数
                                WritePartInfoSql();
                                //可以增加板材的信息，写到数据库
                                WriteMaterialSql();
                            }
                        }
                    }
                    //获取剩下的数据
                    client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(PlcReceiveCallback), state);
                }
                else
                {
                    receiveDonePlc.Set();
                    fmsPlcServerThread = null;
                    clientPlcSocket = null;
                }
            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                fmsPlcServerThread = null;
                clientPlcSocket = null;
            }
        }
        /// <summary>
        /// 激光，气体的消耗表
        /// </summary>
        public void WriteLightUsed()
        {

        }
        /// <summary>
        /// 更新工件信息数据库
        /// </summary>
        private void WritePartInfoSql()
        {
            if (mycon.State == ConnectionState.Open)
            {
                DateTime dt = DateTime.Now;
                string dtStr = dt.Year.ToString() + "-" + dt.Month.ToString() + "-" + dt.Day.ToString() + " " +
                   dt.Hour.ToString() + ":" + dt.Minute.ToString() + ":" + dt.Second.ToString();
                for (int k = 0; k < PartInfoList.Count; k++)
                {
                    string str = "insert into 工件加工信息(工件名称,材料,厚度,数量,加工时间,加工机床,NC名称,上传网络)" +
                                          " values('" + PartInfoList[k].partName + "'," +
                                                  "'" + NCInfo.material + "'," +
                                                  "'" + NCInfo.thickness + "'," +
                                                  "'" + PartInfoList[k].partCount + "'," +
                                                  "'" + dtStr + "'," +
                                                  "'" + machineName + "'," +
                                                  "'" + activeNCName + "'," +
                                                  "'0')";
                    SqlCmd = new SqlCommand(str, mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
            }
        }
        /// <summary>
        /// 写板材数据库
        /// </summary>
        private void WriteMaterialSql()
        {
            try
            {
                DateTime dt = DateTime.Now;
                string dtStr = dt.Year.ToString() + "-" + dt.Month.ToString() + "-" + dt.Day.ToString() + " " +
                   dt.Hour.ToString() + ":" + dt.Minute.ToString() + ":" + dt.Second.ToString();
                string str = "";
                if (mycon.State == ConnectionState.Open)
                {
                    str = "insert into 板材信息表(材料,长度,宽度,厚度,加工时间,加工机床,NC名称,上传网络)" +
                          " values('" + NCInfo.material + "'," +
                                  "'" + NCInfo.length + "'," +
                                  "'" + NCInfo.width + "'," +
                                  "'" + NCInfo.thickness + "'," +
                                  "'" + dtStr + "'," +
                                  "'" + machineName + "'," +
                                  "'" + activeNCName + "'," +
                                  "'0')";
                    SqlCmd = new SqlCommand(str, mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
            }
            catch { }
        }
        /// <summary>
        /// CNC异步接收的回调函数
        /// </summary>
        /// <param name="ar"></param>
        private void ReceiveCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the state object and the client socket 
                // from the asynchronous state object.
                StateObject state = (StateObject)ar.AsyncState;
                Socket client = state.workSocket;

                // Read data from the remote device.
                string readStr;
                string tempStr;
                int bytesRead = client.EndReceive(ar);
                if (bytesRead > 0)
                {
                    readStr = Encoding.ASCII.GetString(state.buffer, 0, bytesRead);
                    if (readStr.Contains("<mode>"))
                    {
                        tempStr = readStr.Substring(readStr.IndexOf("<mode>") + 6);
                        if (GetNumFromStr(tempStr) == "0")//自动模式
                        {
                            isActiveNCName = true;
                        }
                    }
                    if (readStr.Contains("<no>"))
                    {
                        tempStr = readStr.Substring(readStr.IndexOf("<no>") + 4);
                        string status = GetNumFromStr(tempStr);
                        string alarmStytle = readStr.Substring(readStr.IndexOf("<st>") + 4, readStr.IndexOf("</st>") - readStr.IndexOf("<st>") - 4);
                        if (alarmStytle != "plc" && alarmStytle != "run")
                        {
                            alarmStytle = "cnc";
                        }
                        //更新报警信息
                        UpdateAlarmInfo(machineName, status, alarmStytle);
                    }
                    if (readStr.Contains("<status>"))
                    {
                        DateTime dt = DateTime.Now;
                        tempStr = readStr.Substring(readStr.IndexOf("<status>") + 8);
                        string status = GetNumFromStr(tempStr);
                        //写机床的实时状态
                        string strStatus = "update 机床实时状态 set 运行状态='" + status + "',机床名='" + machineName + "' where 机床IP='" + machineIP + "'";
                        SqlCmd = new SqlCommand(strStatus, mycon);
                        SqlCmd.ExecuteNonQuery();
                        SqlCmd.Dispose();
                        if (status == "6")
                        {
                            //暂时注释掉
                            status = "RUN";
                            isActiveNCName = true;
                        }
                        if (status == "4")
                        {
                            status = "STOP";
                        }
                        if (status == "5")
                        {
                            status = "WAIT";
                        }
                        if (status == "2" || status == "3")
                        {
                            status = "RUN";
                        }
                        if (status == "1")
                        {
                            status = "ALARM";
                        }
                        //更新机床状态函数
                        UpdateMachineStatus(machineName, status, dt, false);
                    }

                    // Get the rest of the data.
                    client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReceiveCallback), state);
                }
                else
                {
                    receiveDone.Set();
                    //更新最后一个状态的结束时间
                    DateTime endDt = DateTime.Now;
                    UpdateMachineStatus(machineName, lastMachineStatus.machineStatus, endDt,true);
                    //写机床的实时状态
                    string strStatus = "update 机床实时状态 set 运行状态=255,机床名='" + machineName + "' where 机床IP='" + machineIP + "'";
                    SqlCmd = new SqlCommand(strStatus, mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                    fmsKeyMonitorThread = null;
                    clientCncSocket = null;
                    lastMachineStatus.machineStatus = null;
                    lastMachineStatus.lastTime = null;
                }
            }
            catch
            {
                fmsKeyMonitorThread = null;
            }
        }
        /// <summary>
        /// 更新机床状态数据库表
        /// </summary>
        /// <param name="machineName"></param>
        /// <param name="status"></param>
        /// <param name="dt"></param>
        public void UpdateMachineStatus(string machineName,string status,DateTime dt,bool isPACloseed)
        {
            try
            {
                if (mycon.State == ConnectionState.Open)
                {
                    string dtStr = dt.Year.ToString() + "-" + dt.Month.ToString() + "-" + dt.Day.ToString() + " " +
                                   dt.Hour.ToString() + ":" + dt.Minute.ToString() + ":" + dt.Second.ToString();
                    if (!isPACloseed)//PA没有关闭
                    {
                        if (status == lastMachineStatus.machineStatus)//与上一个状态相同，直接返回
                        {
                            return;
                        }
                        if (lastMachineStatus.machineStatus != null)//不是第一次状态
                        {
                            //写表项机床状态的起始时间
                            string str = "insert into 机床运行状态(机床名,机床状态,起始时间) values('" + machineName + "','" + status + "','" + dtStr + "')";
                            SqlCmd = new SqlCommand(str, mycon);
                            SqlCmd.ExecuteNonQuery();
                            SqlCmd.Dispose();
                            //更新上一次的表项结束时间
                            DateTime startDT = Convert.ToDateTime(lastMachineStatus.lastTime); //开始时间
                            DateTime endDT = Convert.ToDateTime(dtStr); //结束时间
                            TimeSpan dtSpan = endDT - startDT;
                            int timers = (int)dtSpan.TotalSeconds;
                            string updateStr = "update 机床运行状态 set 结束时间='" + dtStr + "',持续时间='" + timers + "',上传网络= 0 where 机床名='" + machineName + "'and 机床状态='" + lastMachineStatus.machineStatus + "' and 起始时间='" + lastMachineStatus.lastTime + "'";
                            SqlCmd = new SqlCommand(updateStr, mycon);
                            SqlCmd.ExecuteNonQuery();
                            SqlCmd.Dispose();

                        }
                        else//是第一次，只写一项起始时间Ok
                        {
                            //写表项机床状态的起始时间
                            string str = "insert into 机床运行状态(机床名,机床状态,起始时间) values('" + machineName + "','" + status + "','" + dtStr + "')";
                            SqlCmd = new SqlCommand(str, mycon);
                            SqlCmd.ExecuteNonQuery();
                            SqlCmd.Dispose();
                        }
                        //记录上次表项的信息
                        lastMachineStatus.machineStatus = status;
                        lastMachineStatus.lastTime = dtStr;
                    }
                    else//PA关闭
                    {
                        //写PA关闭时最后一个状态的结束时间和持续时间
                        DateTime startDT = Convert.ToDateTime(lastMachineStatus.lastTime); //开始时间
                        DateTime endDT = Convert.ToDateTime(dtStr); //结束时间
                        TimeSpan dtSpan = endDT - startDT;
                        int timers = (int)dtSpan.TotalSeconds;
                        string updateStr = "update 机床运行状态 set 结束时间='" + dtStr + "',持续时间='" + timers + "',上传网络= 0 where 机床名='" + machineName + "'and 机床状态='" + lastMachineStatus.machineStatus + "' and 起始时间='" + lastMachineStatus.lastTime + "'";
                        SqlCmd = new SqlCommand(updateStr, mycon);
                        SqlCmd.ExecuteNonQuery();
                        SqlCmd.Dispose();
                        lastMachineStatus.machineStatus = null ;
                        lastMachineStatus.lastTime = null;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
        /// <summary>
        /// 更新报警信息表
        /// </summary>
        /// <param name="machineName"></param>
        /// <param name="alarmNum"></param>
        private void UpdateAlarmInfo(string machineName,string alarmNum,string alarmStyle)
        {
            try
            {
                DateTime dt = DateTime.Now;
                string dtStr = dt.Year.ToString() + "-" + dt.Month.ToString() + "-" + dt.Day.ToString() + " " +
                   dt.Hour.ToString() + ":" + dt.Minute.ToString() + ":" + dt.Second.ToString();
                string str = "insert into 报警信息(机床名,报警号,报警时间,报警类型,上传网络) values('" + machineName + "','" + alarmNum + "','" + dtStr + "','" + alarmStyle + "','0')";
                SqlCmd = new SqlCommand(str, mycon);
                SqlCmd.ExecuteNonQuery();
                SqlCmd.Dispose();
            }
            catch { }

        }
        /// <summary>
        /// 从字符串中提取数字，碰到非数字就返回，包括了小数
        /// </summary>
        /// <param name="tempStr"></param>
        private string GetNumFromStr(string tempStr)
        {
            string number = null;
            foreach (char item in tempStr)
            {
                if ((item >= 48 && item <= 57) || item == 46)
                {
                    number += item;
                }
                else
                {
                    break;
                }
            }
            return number;
        }
        /// <summary>
        ///定时的连接PA
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        void GetMachineStatusTimer_Elapsed(object source, System.Timers.ElapsedEventArgs e)
        {
            if (fmsKeyMonitorThread == null)
            {
                connectDone = new ManualResetEvent(false);
                sendDone = new ManualResetEvent(false);
                receiveDone = new ManualResetEvent(false);
                //启动监控FMS页面Softkey操作线程
                fmsKeyMonitorThread = new Thread(new ThreadStart(fmsKeyMonitorThreadFun));
                fmsKeyMonitorThread.Name = "fmsKeyMonitorThread";
                fmsKeyMonitorThread.Start();
            }
            else
            {
                if (isActiveNCName)
                {
                    isActiveNCName = false;
                    try
                    {
                        //得到当前加工的NC程序
                        string strCmd = "<dir><req>yes</req><sub>exe</sub></dir>\n";
                        byte[] cmdBuffer = Encoding.UTF8.GetBytes(strCmd);
                        SynclientCncSocket.Send(cmdBuffer, cmdBuffer.Length, SocketFlags.None);
                        byte[] bytes = new byte[256];
                        SynclientCncSocket.Receive(bytes, SynclientCncSocket.Available, SocketFlags.None);
                        string strRet = Encoding.UTF8.GetString(bytes);
                        if (strRet.Contains("<name>"))
                        {
                            activeNCName = strRet.Substring(strRet.IndexOf("<name>") + 6, strRet.IndexOf("</name>") - strRet.IndexOf("<name>") - 6);
                            NCPath = strRet.Substring(strRet.IndexOf("<path>") + 6, strRet.IndexOf("</path>") - strRet.IndexOf("<path>") - 6);
                            NCPath = NCPath + "/" + activeNCName;
                            //发送命令到文件服务器
                            FileSend(clientFileSocket, "<fload><fname>" + NCPath + "</fname></fload>\n");

                            sendDoneFile.WaitOne();

                            StateObject state = new StateObject();
                            state.workSocket = clientFileSocket;
                            // Begin receiving the data from the remote device.
                            clientFileSocket.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(FileReceiveCallback), state);
                            receiveDoneFile.WaitOne();
                        }
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("读取加工NC文件出错，请确认机床PA是否运行");
                    }
                }
            }
            if (fmsPlcServerThread == null)
            {
                connectDonePlc = new ManualResetEvent(false);
                sendDonePlc = new ManualResetEvent(false);
                receiveDonePlc = new ManualResetEvent(false);
                // 启动监控FMS页面Softkey操作线程
                fmsPlcServerThread = new Thread(new ThreadStart(fmsPlcServerThreadFun));
                fmsPlcServerThread.Name = "fmsPlcServerThread";
                fmsPlcServerThread.Start();
            }
            if (fmsFileServerThread == null)
            {
                connectDoneFile = new ManualResetEvent(false);
                sendDoneFile = new ManualResetEvent(false);
                receiveDoneFile = new ManualResetEvent(false);
                // 启动监控FMS页面Softkey操作线程
                fmsFileServerThread = new Thread(new ThreadStart(fmsFileServerThreadFun));
                fmsFileServerThread.Name = "fmsFileServerThread";
                fmsFileServerThread.Start();
            }
        }
        /// <summary>
        /// 连接到PA的PLC监控线程函数
        /// </summary>
        private void fmsPlcServerThreadFun()
        {
            try
            {
                // Establish the remote endpoint for the socket.
                machineName = Dns.GetHostEntry(machineIP).HostName.ToString();
                IPAddress ipAddress = IPAddress.Parse(machineIP);
                IPEndPoint remoteEP = new IPEndPoint(ipAddress, 62944);


                // Create a TCP/IP socket.
                clientPlcSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                // Connect to the remote endpoint.
                clientPlcSocket.BeginConnect(remoteEP, new AsyncCallback(PlcConnectCallback), clientPlcSocket);
                connectDonePlc.WaitOne();
                //向PA发送指令
                //PlcSend(clientPlcSocket, "<get><auto>yes</auto><time>100</time><var>FMS_CONTROL_Prg.Parr[583]</var></get>\n");
                PlcSend(clientPlcSocket, "<get><auto>yes</auto><time>100</time><var>.FMS</var></get>\n");
                sendDonePlc.WaitOne();

                StateObject state = new StateObject();
                state.workSocket = clientPlcSocket;
                // Begin receiving the data from the remote device.
                clientPlcSocket.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(PlcReceiveCallback), state);
                receiveDonePlc.WaitOne();

                // Release the socket.
                if (clientPlcSocket.Connected)
                {
                    clientPlcSocket.Shutdown(SocketShutdown.Both);
                    clientPlcSocket.Close();
                }
            }
            catch
            {
                fmsPlcServerThread = null;
            }
        }
        /// <summary>
        /// 连接到PA的文件监控线程函数
        /// </summary>
        private void fmsFileServerThreadFun()
        {
            try
            {
                // Establish the remote endpoint for the socket.
                machineName = Dns.GetHostEntry(machineIP).HostName.ToString();
                IPAddress ipAddress = IPAddress.Parse(machineIP);
                IPEndPoint remoteEP = new IPEndPoint(ipAddress, 62938);


                // Create a TCP/IP socket.
                clientFileSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                // Connect to the remote endpoint.
                clientFileSocket.BeginConnect(remoteEP, new AsyncCallback(FileConnectCallback), clientFileSocket);
                connectDoneFile.WaitOne();
            }
            catch
            {
                fmsFileServerThread = null;
            }
        }
        /// <summary>
        /// 建立Socket通讯
        /// </summary>
        /// <param name="host">主机</param>
        /// <param name="port">端口</param>
        /// <param name="errorMsg">错误信息</param>
        /// <returns></returns>
        public int EstablishSocketConnect(string host, int port, out Socket socket)
        {
            int ret = 0;
            socket = null;
            try
            {
               // IPHostEntry ipHostInfo = Dns.GetHostEntry(host);
               // IPAddress ipAddress = ipHostInfo.AddressList[0];
                IPAddress ipAddress = IPAddress.Parse(host);
                IPEndPoint ipEndPointPLC = new IPEndPoint(ipAddress, port);

                socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                socket.Connect(ipEndPointPLC);
            }
            catch (SocketException ex)
            {
                ret = ex.ErrorCode;
            }

            return ret;
        }
        /// <summary>
        /// 连接到PA的CNC监控线程函数
        /// </summary>
        private void fmsKeyMonitorThreadFun()
        {
            try
            {
                // Establish the remote endpoint for the socket.
                machineName = Dns.GetHostEntry(machineIP).HostName.ToString();
                IPAddress ipAddress = IPAddress.Parse(machineIP);
                IPEndPoint remoteEP = new IPEndPoint(ipAddress, 62937);

                        
                // Create a TCP/IP socket.
                clientCncSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                // Connect to the remote endpoint.
                clientCncSocket.BeginConnect(remoteEP, new AsyncCallback(ConnectCallback), clientCncSocket);
                connectDone.WaitOne();
                //建立同步的SOCKET，用来读取加工NC的路径
                EstablishSocketConnect(machineIP, 62937, out SynclientCncSocket);
                //向PA发送指令
                Send(clientCncSocket, "<ncda><auto>yes</auto><time>50</time><req>yes</req><var>status,mode</var></ncda>\n");
                //sendDone.WaitOne();
                Send(clientCncSocket, "<alarm><auto>yes</auto><time>50</time></alarm>\n");
                //sendDone.WaitOne();
                // Receive the response from the remote device.
                StateObject state = new StateObject();
                state.workSocket = clientCncSocket;
                // Begin receiving the data from the remote device.
                clientCncSocket.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReceiveCallback), state);
                receiveDone.WaitOne();

                // Release the socket.
                if (clientCncSocket.Connected)
                {
                    clientCncSocket.Shutdown(SocketShutdown.Both);
                    clientCncSocket.Close();
                }
            }
            catch(Exception ex)
            {
                fmsKeyMonitorThread = null;
            }
        }
        /// <summary>
        /// 得到加工NC的材料
        /// </summary>
        /// <param name="readStr"></param>
        private void GetNCMaterial(string readStr)
        {
            if (readStr.Contains("Sus"))
            {
                NCInfo.material = "Sus";
            }
            if (readStr.Contains("Steel"))
            {
                NCInfo.material = "Steel";
            }
            if (readStr.Contains("Aluminum"))
            {
                NCInfo.material = "Aluminum";
            }
            if (readStr.Contains("Acid-ms"))
            {
                NCInfo.material = "Acid-ms";
            }
            if (readStr.Contains("Titanium"))
            {
                NCInfo.material = "Titanium";
            }
            if (readStr.Contains("Ms"))
            {
                NCInfo.material = "Ms";
            }
        }
        /// <summary>
        /// 得到NC工件的基本信息
        /// </summary>
        /// <param name="taskName"></param>
        /// <returns></returns>
        private bool GetNCPartInfo(string taskName)
        {
            string fileAccess = taskName;
            try
            {
                FileStream fs = new FileStream(fileAccess, FileMode.Open, FileAccess.Read, FileShare.None);
                StreamReader sr = new StreamReader(fs);
                bool isEnd = false;
                PartInfoList = new List<partInfo>();
                PartInfoList.Clear();
                NCInfo = new information();
                int i = 0;
                while (!isEnd)
                {
                    string lineStr = sr.ReadLine();
                    if (lineStr == null) { isEnd = sr.EndOfStream; continue; }
                    string tempStr = "";
                    if (!isEnd)
                    {
                        if (i < 15)//读取NC文件头，提取信息
                        {
                            //得到板材信息
                            if (lineStr.Contains("*SHEET"))
                            {
                                tempStr = lineStr.Substring(lineStr.IndexOf("*SHEET") + 7);
                                //得到板材长度
                                NCInfo.length = GetNumFromStr(tempStr);
                                tempStr = tempStr.Substring(NCInfo.length.Length + 1);
                                //得到板材的宽度
                                NCInfo.width = GetNumFromStr(tempStr);
                                tempStr = tempStr.Substring(NCInfo.width.Length + 1);
                                //得到板材的厚度
                                NCInfo.thickness = GetNumFromStr(tempStr);
                            }
                            //得到NC程序的材料
                            GetNCMaterial(lineStr);
                        }
                        if (lineStr.IndexOf("PART NAME:") > 0)
                        {
                            GetPartCnt(lineStr);
                        }
                        isEnd = sr.EndOfStream;
                    }
                    i++;
                }
                sr.Close();
                fs.Close();
                return true;
            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                return false;
            }
        }
        /// <summary>
        /// 得到工件零件数
        /// </summary>
        /// <param name="strNC"></param>
        /// <returns></returns>
        public bool GetPartCnt(string strNC)
        {
            int n1 = strNC.IndexOf(':');
            int n2 = strNC.IndexOf(')');
            string strName = strNC.Substring(n1 + 1, n2 - n1 - 1);
            strName.Replace(" ", "");
            for (int i = 0; i < PartInfoList.Count; i++)
            {
                if (PartInfoList[i].partName == strName)
                {
                    PartInfoList[i].partCount++;
                    return false;
                }
            }
            partInfo partItem = new partInfo();
            partItem.partName = strName;
            partItem.partCount = 1;
            PartInfoList.Add(partItem);
            return true;
        }
        /// <summary>
        /// 初始化数据库表格
        /// </summary>
        private void InitiTable()
        {
            if (mycon.State == ConnectionState.Open)
            {
                //判断机床开机时间表是不是存在
                SqlCmd = new SqlCommand("SELECT COUNT(*) FROM sysobjects where name='机床运行状态'", mycon);
                SqlCmd.ExecuteScalar();
                if (0 == Convert.ToInt32(SqlCmd.ExecuteScalar()))
                {
                    SqlCmd.Dispose();
                    SqlCmd = new SqlCommand("create table 机床运行状态(机床名 varchar(30),机床状态 varchar(5),起始时间 datetime,结束时间 datetime,持续时间 int,上传网络 int)", mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
                else
                {
                    SqlCmd.Dispose();
                }
                //判断报警信表
                SqlCmd = new SqlCommand("SELECT COUNT(*) FROM sysobjects where name='报警信息'", mycon);
                SqlCmd.ExecuteScalar();
                if (0 == Convert.ToInt32(SqlCmd.ExecuteScalar()))
                {
                    SqlCmd.Dispose(); ;
                    SqlCmd = new SqlCommand("create table 报警信息(机床名 varchar(30),报警号 varchar(10),报警时间 datetime,报警类型 varchar(10),上传网络 int)", mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
                else
                {
                    SqlCmd.Dispose();
                }
                //判断工件信息表是不是存在
                SqlCmd = new SqlCommand("SELECT COUNT(*) FROM sysobjects where name='工件加工信息'", mycon);
                SqlCmd.ExecuteScalar();
                if (0 == Convert.ToInt32(SqlCmd.ExecuteScalar()))
                {
                    SqlCmd.Dispose();
                    SqlCmd = new SqlCommand("create table 工件加工信息(工件名称 varchar(50),材料 varchar(10),厚度 varchar(10),数量 varchar(10),加工时间 datetime,加工机床 varchar(30),NC名称 varchar(20),上传网络 int)", mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
                else
                {
                    SqlCmd.Dispose();
                }
                //判断板材信息表是不是存在
                SqlCmd = new SqlCommand("SELECT COUNT(*) FROM sysobjects where name='板材信息表'", mycon);
                SqlCmd.ExecuteScalar();
                if (0 == Convert.ToInt32(SqlCmd.ExecuteScalar()))
                {
                    SqlCmd.Dispose();
                    SqlCmd = new SqlCommand("create table 板材信息表(材料 varchar(5),长度 varchar(10),宽度 varchar(10),厚度 varchar(5),加工时间 datetime,加工机床 varchar(30),NC名称 varchar(20),上传网络 int)", mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
                else
                {
                    SqlCmd.Dispose();
                }
                //判断光气消耗表是不是存在
                SqlCmd = new SqlCommand("SELECT COUNT(*) FROM sysobjects where name='光气消耗表'", mycon);
                SqlCmd.ExecuteScalar();
                if (0 == Convert.ToInt32(SqlCmd.ExecuteScalar()))
                {
                    SqlCmd.Dispose();
                    SqlCmd = new SqlCommand("create table 光气消耗表(机床号 varchar(20),开启时间 datetime,关闭时间 datetime,持续时间 int,耗电量 float,氧气量 float,氮气量 float,NC名称 varchar(20),上传网络 int)", mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
                else
                {
                    SqlCmd.Dispose();
                }
                //机床实时状态
                SqlCmd = new SqlCommand("SELECT COUNT(*) FROM sysobjects where name='机床实时状态'", mycon);
                SqlCmd.ExecuteScalar();
                if (0 == Convert.ToInt32(SqlCmd.ExecuteScalar()))
                {
                    SqlCmd.Dispose();
                    SqlCmd = new SqlCommand("create table 机床实时状态(机床名 varchar(30),运行状态 int,机床IP varchar(30))", mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                    string str = "insert into 机床实时状态(机床名,机床IP,运行状态)" +
                                         " values('" + machineName + "'," +
                                                 "'" + machineIP + "'," +
                                                 "'" + 255 + "')";
                    SqlCmd = new SqlCommand(str, mycon);
                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();
                }
                else
                {
                    SqlCmd.Dispose();
                }
            }
        }
    }
}
