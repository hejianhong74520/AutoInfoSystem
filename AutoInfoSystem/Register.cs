using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;


namespace InfoStat
{
    class Register
    {
        /// <summary>
        ///  机启 注册表检查
        /// </summary>
        public void start_with_windows()
        {
            RegistryKey hklm = Registry.LocalMachine;
            RegistryKey run = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
            //未设置开机启动
            if (run.GetValue("信息统计") == null)
            {
                register();
            }

        }

        /// <summary>
        ///  程序开机启动写入注册表
        /// </summary>
        private void register()
        {
            string starupPath = System.Windows.Forms.Application.ExecutablePath;
            //class Micosoft.Win32.RegistryKey. 表示Window注册表 项级节点, 类 注册表装.
            RegistryKey loca = Registry.LocalMachine;
            RegistryKey run = loca.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run");

            try
            {
                run.SetValue("信息统计", starupPath);
                MessageBox.Show("注册表添加功!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                loca.Close();
                run.Close();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
