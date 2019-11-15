using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Excel;

namespace WindowsFormsApp1
{
    /// <summary>
    /// 数据类型
    /// </summary>
    public enum DataType
    {
        /// <summary>
        /// 服务器用
        /// </summary>
        Server, 
        /// <summary>
        /// 客户端
        /// </summary>
        Client, 
    }
    public partial class Form1 : Form
    {
         
        string[] _excelNameArr;
        string _jsonPath;
        int _startRow = 5;

        public Form1()
        {
            InitializeComponent();
           
        }

 
        /// <summary>
        /// 打开生成Json文件的路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
         
            FolderBrowserDialog dialog = new FolderBrowserDialog
            {
                Description = "请选择Json生成的文件夹"
            }; 
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    dialog.Description = "文件夹路径不能为空!";
                    return;
                }
                else
                {
                    textBox2.Text = _jsonPath = dialog.SelectedPath;
                }
            }
        }

        /// <summary>
        /// 打开excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog
            {
                // 指定打开文本文件（后缀名为xlsx）
                Filter = "文本文件|*.xlsx"
            };
            openDlg.Multiselect = true;
            if (openDlg.ShowDialog() == DialogResult.OK)
            {
                // 显示文件路径名
 
                textBox1.Text = "已选文件数量："+ openDlg.FileNames.Length.ToString();
                _excelNameArr = openDlg.FileNames;
              
            }
 
        }

        /// <summary>
        /// 生成按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (_excelNameArr?.Length > 0)
            {
                foreach (var item in _excelNameArr)
                {                
                    ExcelToJson(item, DataType.Client);
                    ExcelToJson(item, DataType.Server);
                }
            }
            
          
            MessageBox.Show("已完成！","提示");
             
             
        }

        /// <summary>
        /// 客户端 excel转json
        /// </summary>
        private void ExcelToJson(string excelPath, DataType dataType)
        {
            Console.WriteLine(_excelNameArr);
            FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            DataSet dataSet = excelReader.AsDataSet();
            DataTable dt = new DataTable();

            //获得文件名
            string tempStrName = dataSet.Tables[0].Rows[0][0].ToString();
            //生成表头的行号
            int tableHeaderRow = 3;
           
            for ( int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
            {
                //根据数据类型判断表头是否显示
                switch (dataType)
                {
                    case DataType.Server:
                        {             
                            //删除客户端列
                            if (dataSet.Tables[0].Rows[tableHeaderRow][i].ToString().Contains("1"))
                            {
                                dataSet.Tables[0].Columns.RemoveAt(i);
                                i--;
                            }
                            break;
                        }
                    case DataType.Client:
                        {
                            //删除服务器列                              
                            if (dataSet.Tables[0].Rows[tableHeaderRow][i].ToString().Contains("2"))
                            {
                                dataSet.Tables[0].Columns.RemoveAt(i);
                                i--;
                            }
                            break;
                        }
                        default :
                        {  
                                
                            break;
                        }                       
                }
             
            }
            //生成表头起始行号
            int startRow = 1;
            for (int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
            {
                //表头类型判断
                Type tempType = Type.GetType("System.String");
                if (dataSet.Tables[0].Rows[2][i].ToString().Contains("int"))
                {
                    tempType = Type.GetType("System.Int32");
                }
                if (dataSet.Tables[0].Rows[1][i]?.ToString().Length > 0)
                {
                    //共用表头
                    dt.Columns.Add(dataSet.Tables[0].Rows[startRow][i].ToString(), tempType);
                }
            }
             


            //将DataSet的数据重新组织并填充到DataTable
            for (int i = _startRow; i < dataSet.Tables[0].Rows.Count; i++)
            {
               
                if (string.IsNullOrEmpty(dataSet.Tables[0].Rows[i][0].ToString()))
                {
                    break;
                }
                DataRow dr = dt.NewRow();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
           
                    if (string.IsNullOrEmpty(dataSet.Tables[0].Rows[tableHeaderRow][j]?.ToString()))
                    {
                        
                        dr[j] = "";
                    }
                    else
                    {
                        dr[j] = dataSet.Tables[0].Rows[i][j];
                    }       
                }
                dt.Rows.Add(dr);
            }

            //写入Json文件
            string json = JsonConvert.SerializeObject(dt);
                     
            switch (dataType)
            {
                case DataType.Server:
                    {
                        Directory.CreateDirectory(_jsonPath+ "/Server/");
                        File.WriteAllText(_jsonPath + "/Server/" + tempStrName + ".json", json);
                        break;
                    }
                case DataType.Client:
                    {
                        Directory.CreateDirectory(_jsonPath + "/Client/");
                        File.WriteAllText(_jsonPath + "/Client/" + tempStrName + ".json", json);
                        break;
                    }
            }

        }

      
    }
}
