using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportSQLDataToExcel
{
   
    public partial class Form2 : Form

    {

        public Form2(int _Minimum, int _Maximum)//帶入參數，表示數據範圍的最小值及最大值

        {

            InitializeComponent();

            progressBar2.Maximum = _Maximum;//設置範圍最大值

            progressBar2.Value = progressBar2.Minimum = _Minimum;//設置範圍最小值

        }

        public void setPos(int value)//設置當前進度值

        {

            if (value < progressBar2.Maximum)//如果值有效

            {

                progressBar2.Value = value;//設置進度值

                label1.Text = (value * 100 / progressBar2.Maximum).ToString() + "%";//顯示百分比

            }

            Application.DoEvents();//☆一定要加，否則父子視窗都會假死。☆
        }

        private void Form2_Load(object sender, EventArgs e)

        {

            

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)

        {

            this.Owner.Enabled = true;//回復父視窗為可用，上面應該要加一個false在匯出時把父視窗給停用，但不知道為甚麼關掉視窗之後父視窗還是不行用，再研究。

        }

       
    }
}
