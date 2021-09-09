using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Windows.Forms;

namespace AutoInfoSystem
{
    class XMLOperator
    {
        XmlDocument xd = null;//XML文件操作

        public Hashtable LoadFile(string filePath)
        {
            Hashtable paraTable = new Hashtable();
            if (xd == null)
            {
                xd = new XmlDocument();
            }

            try
            {
                xd.Load(filePath);

                //获取当前XML文档的根 一级    
                XmlNode rootNode = xd.DocumentElement;
                //获取根节点的所有子节点列表    
                XmlNodeList rootList = rootNode.ChildNodes;
                for (int i = 0; i < rootList.Count; i++)
                {
                    paraTable.Add(rootList[i].Name, rootList[i].InnerText);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return paraTable;
        }

        /// <summary>
        /// 读取XML节点值
        /// </summary>
        /// <param name="fristNode">主节点</param>
        /// <param name="childNode">子节点</param>
        /// <param name="fileName">文件名</param>
        /// <returns></returns>
        public string Reader(string p_strFristNode, string p_strChildNode, string p_strFileName)
        {
            if (xd == null)
            {
                xd = new XmlDocument();
            }
            // string path = Application.StartupPath + "\\" + p_strFileName;
            string path = p_strFileName;
            string reader = "";
            if (File.Exists(path))
            {
                xd.Load(path);
                XmlNode xn = xd.SelectSingleNode(p_strFristNode);
                if (xn != null)
                {
                    reader = xn[p_strChildNode].InnerText.ToString();
                }
            }
            else
            {
                MessageBox.Show("配置文件丢失！请联系管理员！", "警告！");
            }
            return reader;
        }

        /// <summary>
        /// 修改XML节点值
        /// </summary>
        /// <param name="fristNode">主节点</param>
        /// <param name="childNode">子节点</param>
        /// <param name="text">子节点内容</param>
        /// <param name="fileName">文件名</param>
        /// <returns></returns>
        public void AlterNode(string p_strFristNode, string p_strChildNode, string p_strText, string p_strFileName)
        {
            try
            {
                if (xd == null)
                {
                    xd = new XmlDocument();
                }

                xd.Load(p_strFileName);
                XmlNode xn = xd.SelectSingleNode(p_strFristNode);
                xn[p_strChildNode].InnerText = p_strText;
                xd.Save(p_strFileName);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 创建一个新XML文档
        /// </summary>
        /// <param name="p_strFileName">文档路径名</param>
        /// <param name="parameterTable">参数表</param>
        /// <returns></returns>
        public void CreateDocument(string p_strFileName, Hashtable parameterTable)
        {
            try
            {
                // 创建XmlTextWriter类的实例对象
                XmlTextWriter textWriter = new XmlTextWriter(p_strFileName, null);
                textWriter.Formatting = Formatting.Indented;

                // 开始写过程，调用WriteStartDocument方法
                textWriter.WriteStartDocument();

                // 写入说明
                textWriter.WriteComment("Author: liuyc111766");
                textWriter.WriteComment("All Rights Reserved by Han's Laser");

                //创建节点
                textWriter.WriteStartElement("AutoInfoSystemParameters");
                foreach (DictionaryEntry de in parameterTable)
                {
                    textWriter.WriteElementString(de.Key.ToString(), de.Value.ToString());
                }
                textWriter.WriteEndElement();

                // 写文档结束，调用WriteEndDocument方法
                textWriter.WriteEndDocument();

                // 关闭textWriter
                textWriter.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
