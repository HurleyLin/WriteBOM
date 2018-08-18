using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using Oracle_DLL;
using System.Data.OleDb;
using System.Threading; 

namespace WriteBom
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        static public DataTable ExcelToDS(string Path)
        {
            //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataTable dt = null;
            strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            dt = new DataTable();
            myCommand.Fill(dt);
            return dt;
        }

        static void excelwrite()
        {
            string exepath = System.IO.Directory.GetCurrentDirectory();
            string filepath = exepath + "\\BOM.xls";

            DataTable p = new DataTable();
            DataTable dt = ExcelToDS(filepath);

            //textBox1.Text = (dt.Rows[0][0].ToString() + dt.Rows[0][1].ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string Database;
            //Database = comboBox1.Text;
            if (comboBox1.Text == "二楼MES")
            {
                Database = "2FMES";
            }
            else if (comboBox1.Text == "四楼MES")
            {
                Database = "4FMES";
            }
            else if (comboBox1.Text == "TestMES")
            {
                Database = "TestMES";
            }
            else 
            { 
                MessageBox.Show("请先选择导入数据库！");
                return;
            }
            
            string exepath = System.IO.Directory.GetCurrentDirectory();
            string filepath = exepath + "\\BOM.xls";

            string Date = DateTime.Now.ToString("yyyyMMdd");
            //textBox1.Text = Date;

            DataTable p = new DataTable();
            DataTable dt = ExcelToDS(filepath);
            int number = dt.Rows.Count; //获取数据行数

            //textBox1.Text = dt.Rows[0][0].ToString() + "\t";
            //textBox1.Text = dt.Rows.Count.ToString();
            
            
            for (int i = 2; i < number; i++)
            {

                bool b = ORACLEDLL.Material_Exists(dt.Rows[i][0].ToString(), Database);
                if (b)
                {
                    bool result = ORACLEDLL.InsterMatertal(dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), "PCS", "itemtype_rawmaterial", "0", "0", "SYSTEM", Date, "10000", "1", "0", Database);
                    if (result)
                    {
                        bool temp = ORACLEDLL.BOM_Exists(dt.Rows[0][0].ToString(), dt.Rows[i][0].ToString(), Database);
                        if (temp)
                        {
                            bool result2 = ORACLEDLL.InsterBom(dt.Rows[0][0].ToString(), dt.Rows[i][0].ToString(), dt.Rows[i][0].ToString(), dt.Rows[i][4].ToString(), Date, "0", "0", "29991231", "0", "EA", "SYSTEM", Date, "10000", "1", "1.1", "1", "0", "0", "S", Database);
                            if (result2)
                                textBox1.Text += "物料：" + dt.Rows[i][0].ToString() + " 与产品：" + dt.Rows[0][0].ToString() + " 成功绑定\r\n";
                            else
                                textBox1.Text += "物料：" + dt.Rows[i][0].ToString() + " 录入成功" + " 但是与产品：" + dt.Rows[0][0].ToString() +"绑定失败，请重试！\r\n";

                        }
                        else
                        {
                            textBox1.Text += "物料：" + dt.Rows[i][0].ToString() + " 与产品：" + dt.Rows[0][0].ToString() + " 已有绑定关系\r\n";
                        }
                    }
                    else {
                        textBox1.Text += "物料：" + dt.Rows[i][0].ToString() + " 录入失败，请重试！\r\n";
                    }
                    
                }
                else
                {
                    bool temp = ORACLEDLL.BOM_Exists(dt.Rows[0][0].ToString(), dt.Rows[i][0].ToString(), Database);
                    if (temp)
                    {
                        bool result = ORACLEDLL.InsterBom(dt.Rows[0][0].ToString(), dt.Rows[i][0].ToString(), dt.Rows[i][0].ToString(), dt.Rows[i][4].ToString(), Date, "0", "0", "29991231", "0", "EA", "SYSTEM", Date, "10000", "1", "1.1", "1", "0", "0", "S", Database);
                        if (result)
                            textBox1.Text += "物料：" + dt.Rows[i][0].ToString() + " 与产品：" + dt.Rows[0][0].ToString() + " 成功绑定\r\n";
                        else
                            textBox1.Text += "物料：" + dt.Rows[i][0].ToString() + " 与产品：" + dt.Rows[0][0].ToString() + "绑定失败，请重试！\r\n";

                    }
                    else 
                    {
                        textBox1.Text += "物料：" + dt.Rows[i][0].ToString() + " 与产品：" + dt.Rows[0][0].ToString() + " 已有绑定关系\r\n";
                    }
                }
                textBox1.Text += "\r\n";
                //Thread.Sleep(1000);
            }
            //bool B = ORACLEDLL.InsterMatertal("1202008-0003-8", "PCBA", "69830WVRAIT8M", "PCS", "itemtypel", "0", "0", "SYSTEM", "20180508", "10000", "1", "0");        
            //bool B = ORACLEDLL.InsterBom("1109008-0161-1", "1601002-0537-1", "1601002-0537-1", "1", "20180508", " ", "0", "29991231", "0", "EA", "SYSTEM", "20180508", "10000", "1", "1.1", "1", "0", "0", "S");
        }
    }
}
