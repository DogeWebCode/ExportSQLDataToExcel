using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Reflection;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Data;
using System.Diagnostics;
using HZH_Controls.Forms;
using System.Globalization;

namespace ExportSQLDataToExcel
{
    public partial class Form1 : Form
    {
        public string StartupFilePath { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //設定DatetimePicker為當日日期，並且設定預設的開始時間為00:00:00
            DateTime dStart = this.dtFromDate.Value.Date;
            this.dtFromDate.Value = dStart;
            //設定DatetimePicker為當日日期，並且設定預設的結束時間為23:59:59
            DateTime dEnd = new DateTime(this.dtToDate.Value.Year, this.dtToDate.Value.Month, this.dtToDate.Value.Day, 23, 59, 59);
            this.dtToDate.Value = dEnd;

            using (MySqlConnection conn = new MySqlConnection(@"server=127.0.0.1 ; port=3306 ; pwd = root;user=root ; database = test ; charset = utf8"))
            //測試數據用
            //using (MySqlConnection conn = new MySqlConnection(@"server=127.0.0.1 ; port=3306 ; pwd = root;user=root ; database = bigdata_test ; charset = utf8"))
            {
                using (MySqlDataAdapter sda = new MySqlDataAdapter("SELECT DISTINCT TagName from bwanachglog", conn))
                //測試數據用
                //using (MySqlDataAdapter sda = new MySqlDataAdapter("select * from vote_record", conn))
                {

                    using (System.Data.DataTable dt = new System.Data.DataTable())

                    {

                        //將資料庫中查到的資料填充到DataTable資料表
                        comboBox.Items.Clear();
                        sda.Fill(dt);
                        foreach (DataRow dr in dt.Rows)
                        {
                            comboBox.Items.Add(dr["TagName"].ToString());
                        }
                        conn.Close();
                    }
                }
            }


        }      

        private void btn_Export_Click(object sender, EventArgs e)
        {
            

            Application app = new Application();
            //設定DWName變數為，抓取目前日期為星期幾，變轉化成TW地區的顯示方式，即為"星期X"
            CultureInfo ci = new CultureInfo("zh-tw");            
            string DWName = ci.DateTimeFormat.GetDayName(DateTime.Today.DayOfWeek);

            //儲存視窗各項設定

            string filename = StartupFilePath + "D:\\template.xlsm";
            string newfile = DateTime.Now.ToString(@"D:\yyyy-MM-dd-HH-mm-ss");
            
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel 工作簿.xlsm|*.xlsm";//格式設定
            saveFileDialog.Title = "儲存Excel檔";
            saveFileDialog.RestoreDirectory = true;//是否開啟上次存檔的位置
            saveFileDialog.FileName = DateTime.Now.ToString("yyyyMMdd_HH-mm-ss")+"_"+ DWName + "_"+comboBox.Text;
            saveFileDialog.FilterIndex = 1;

            

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
               
                string path = saveFileDialog.FileName;
                app.Visible = false; //不顯示EXCEL
                app.DisplayAlerts = false; //不提示信息
                app.ScreenUpdating = false; //停止刷新螢幕，加快速度
                Workbooks wbs = app.Workbooks; //取得工作頁
                _Workbook wb = wbs.Open(filename); //開啟excel檔，這邊開啟的為設定好的tamplate
                Sheets shs = wb.Worksheets; //檔案中的Sheets

                using (MySqlConnection conn = new MySqlConnection(@"server=127.0.0.1 ; port=3306 ; pwd = root;user=root ; database = test ; charset = utf8"))
                //測試數據用
                //using (MySqlConnection conn = new MySqlConnection(@"server=127.0.0.1 ; port=3306 ; pwd = root;user=root ; database = bigdata_test ; charset = utf8"))
                {
                    using (MySqlDataAdapter sda = new MySqlDataAdapter($"SELECT * FROM bwanachglog WHERE TagName = '{comboBox.Text}' AND LogDate BETWEEN '{dtFromDate.Value.ToShortDateString()}' AND '{dtToDate.Value.ToShortDateString()}' AND LogTime BETWEEN '{dtFromDate.Value.TimeOfDay}' AND '{dtToDate.Value.TimeOfDay}' ORDER BY bwanachglog.LogTime,bwanachglog.LogDate ASC", conn))
                    //測試數據用
                    //using (MySqlDataAdapter sda = new MySqlDataAdapter("select * from vote_record limit 100000", conn))

                    {
                        using (System.Data.DataTable dt = new System.Data.DataTable())
                        {
                            
                            conn.Open();
                            //將資料庫中查到的資料填充到DataTable資料表
                            sda.Fill(dt);
                            if (dt == null || dt.Rows.Count < 1)
                            {
                                FrmTips.ShowTips(null, "不能保存空數據。");
                                return;
                            }
                            //使用方法，將DataTable里的數據寫入到第2個Sheet中
                            CopyDataToSheet(dt, (_Worksheet)shs[1]);
                            conn.Close();
                        }
                    }
                }

                wb.SaveAs(path); //保存
                wb.Close(); //關閉工作簿
                app.Quit(); //關閉Excel
                           

                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(shs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);                                         
            }
            else
            {
                return;
            }
            GC.Collect();
        }

        private void CopyDataToSheet(System.Data.DataTable dt, _Worksheet Sheet)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();//計時器開始

            

            
            int colCount, rowCount;
            //DataSet ds = new DataSet();
            colCount = dt.Columns.Count;
            rowCount = dt.Rows.Count;
            
            long rowRead = 0;
            Range range;

            

            range = Sheet.get_Range("A2", Missing.Value);
            range = range.get_Resize(rowCount, colCount);
        
            object[,] cellData = new object[rowCount, colCount];


            Form2 fm = new Form2(0, 100);

            fm.Show(this);//把進度條show出來

           
            //迴圈裡的是將SQL的值，先放進二維陣列中。
            for (int iRow = 0; iRow < rowCount; iRow++)
            {
                
                for (int iCol = 0; iCol < colCount; iCol++)
                {
                    rowRead++;
                    float percent = ((float)(100 * rowRead)) / rowCount;
                    fm.setPos((int)percent);
                    cellData[iRow, iCol] = dt.Rows[iRow][iCol].ToString();
                    
                }
                
            }
            range.set_Value(Missing.Value, cellData);//將陣列中的值匯進excel
            
            fm.Close();//關閉進度視窗
            stopWatch.Stop();//計時器停止
            MessageBox.Show("匯出已完成！" + "總用時為：" + stopWatch.ElapsedMilliseconds  + "毫秒");
        }


        //寫進標頭(第一列)
        //range = Sheet.get_Range("A1", Missing.Value);
        //range = range.get_Resize(1, colCount);
        //object[,] headerData = new object[1, colCount];
        //for (int iCol = 0; iCol < colCount; iCol++)
        //{
        //    headerData[0, iCol] = dt.Columns[iCol].ColumnName;
        //}
        //range.set_Value(Missing.Value, headerData);

        //寫進資料(從第二列開始)                  



        private void btnLoad_Click(object sender, EventArgs e)
        {
            using (MySqlConnection conn = new MySqlConnection(@"server=127.0.0.1 ; port=3306 ; pwd = root;user=root ; database = test ; charset = utf8"))
            //測試數據用
            //using (MySqlConnection conn = new MySqlConnection(@"server=127.0.0.1 ; port=3306 ; pwd = root;user=root ; database = bigdata_test ; charset = utf8"))
            {
                using (MySqlDataAdapter sda = new MySqlDataAdapter($"SELECT * FROM bwanachglog WHERE TagName = '{comboBox.Text}' AND LogDate BETWEEN '{dtFromDate.Value}' AND '{dtToDate.Value}' AND LogTime BETWEEN '{dtFromDate.Value.TimeOfDay}' AND '{dtToDate.Value.TimeOfDay}' ORDER BY bwanachglog.LogTime,bwanachglog.LogDate ASC", conn))
                //測試數據用
                //using (MySqlDataAdapter sda = new MySqlDataAdapter("select * from vote_record", conn))
                {
                    conn.Close();
                    using (System.Data.DataTable dt = new System.Data.DataTable())

                    {
                        Stopwatch stopWatch = new Stopwatch();
                        //將資料庫中查到的資料填充到DataTable資料表
                        stopWatch.Start();
                        DataSet ds = new DataSet();
                        sda.Fill(ds);
                        sda.Fill(dt);
                        dataGridView1.DataSource = dt;
                        long totalCount = ds.Tables[0].Rows.Count;
                       
                        

                        lblTip.Text = "共有" + totalCount + "條數據";
                        lblTip.TextAlign = HorizontalAlignment.Center;
                        sda.Dispose();
                        stopWatch.Stop();
                        MessageBox.Show("讀取資料已完成！" + "總用時為：" + stopWatch.ElapsedMilliseconds + "毫秒");
                    }
                }
            }
        } 
    }
}

















