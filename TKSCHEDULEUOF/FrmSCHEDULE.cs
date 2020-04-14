using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using FastReport;
using FastReport.Data;

namespace TKSCHEDULEUOF
{
    public partial class FrmSCHEDULE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        DataSet ds = new DataSet();

        public FrmSCHEDULE()
        {
            InitializeComponent();

            timer1.Enabled = true;
            timer1.Interval = 1000 ;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            string RUNTIME = DateTime.Now.ToString("HHmm");
            string HHmm = "0910";

            // DayOfWeek 0 開始 (表示星期日) 到 6 (表示星期六)
            string RUNDATE = DateTime.Now.DayOfWeek.ToString("d");//tmp2 = 4 
            string date = "1";


            if (RUNTIME.Equals(HHmm))
            {
                ADDTOUOFTB_EIP_SCH_MEMO(DateTime.Now.ToString("yyyyMMdd"));
            }
        }

        #region FUNCTION

        public void ADDTOUOFTB_EIP_SCH_MEMO(string STRATDATE)
        {

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            ADDTOUOFTB_EIP_SCH_MEMO(DateTime.Now.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
