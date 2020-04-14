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
using System.Threading;

namespace TKSCHEDULEUOF
{
    public partial class FrmSCHEDULE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();

        public FrmSCHEDULE()
        {
            InitializeComponent();

            timer1.Enabled = true;
            timer1.Interval = 1000*60 ;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            string RUNTIME = DateTime.Now.ToString("HHmm");
            string HHmm = "0900";

            // DayOfWeek 0 開始 (表示星期日) 到 6 (表示星期六)
            string RUNDATE = DateTime.Now.DayOfWeek.ToString("d");//tmp2 = 4 
            string date = "1";


            if (RUNTIME.Equals(HHmm))
            {
                ADDTOUOFTB_EIP_SCH_MEMO_MOC(DateTime.Now.ToString("yyyyMMdd"));
            }
        }

        #region FUNCTION

        public void ADDTOUOFTB_EIP_SCH_MEMO_MOC(string Sday)
        {

            DataSet ds = new DataSet();
            ds = SEARCHMANULINE(Sday);
            Thread.Sleep(1000);
            ds2 = SEARCHMANULINE2(Sday);
            Thread.Sleep(1000);
            ds3 = SEARCHMANULINE3(Sday);
            Thread.Sleep(1000);
            ds4 = SEARCHMANULINE4(Sday);

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                //[CREATE_USER]='7774b96c-6762-45ef-b9d1-fcd718854e9f'，包裝線 MANU90
                //[CREATE_USER]='5ce0f554-8b80-4aed-afea-fcd224cecb81'，新廠製一組 MANU10
                //[CREATE_USER]='0c98530a-b467-4cd4-a411-7279f1e04d0d'，新廠製二組 MANU20
                //[CREATE_USER]='88789ece-41d1-4b48-94f1-6ffab05b05f4'，新廠製三組(手工) MANU30
                //將資料從TKMOC的MOCMANULINE計算出工時，再COPY到UOF的TB_EIP_SCH_MEMO
                //先刪除再新增

                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='7774b96c-6762-45ef-b9d1-fcd718854e9f'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='5ce0f554-8b80-4aed-afea-fcd224cecb81'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='0c98530a-b467-4cd4-a411-7279f1e04d0d'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='88789ece-41d1-4b48-94f1-6ffab05b05f4'", Sday);
                sbSql.AppendFormat(" ");

                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }

                if (ds2.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds2.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }

                if (ds3.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds3.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }

                if (ds4.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds4.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }




                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                    //Console.WriteLine("ADDTOUOFTB_EIP_SCH_MEMO_MOC OK");

                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        //新廠製一組、新廠製二組的桶數
        public DataSet SEARCHMANULINE(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

               

                sbSql.AppendFormat(" SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做24桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做24桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製一組'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE]");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做38桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做38桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製二組'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE]");
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ORDER BY [START_TIME],[SUBJECT]");
                sbSql.AppendFormat(" ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");



                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return ds1;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        return ds1;
                    }

                    return ds1;
                }

            }
            catch
            {
                return ds1;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        //新廠包裝線、新廠製一組、新廠製二組、新廠製三組(手工)的總工時
        public DataSet SEARCHMANULINE2(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠包裝線'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE]");                
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '  AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製一組'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE]");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製二組'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE]");
                sbSql.AppendFormat(" UNION");               
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)   AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製三組(手工)%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製三組(手工)'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE] ");             
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ORDER BY [START_TIME],[SUBJECT]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");



                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    return ds2;
                }
                else
                {
                    if (ds1.Tables["ds2"].Rows.Count >= 1)
                    {
                        return ds2;
                    }

                    return ds2;
                }

            }
            catch
            {
                return ds2;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        //新廠包裝線、新廠製一組、新廠製二組、新廠製三組(手工)的稼動率
        public DataSet SEARCHMANULINE3(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]");
                sbSql.AppendFormat(" FROM (");               
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/16*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/16*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠包裝線'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE]");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製一組'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE] ");
                sbSql.AppendFormat(" UNION");
                 sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製二組'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE]");
                sbSql.AppendFormat(" UNION");                
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/6.5*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/6.5*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製三組(手工)%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製三組(手工)'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],[MANUDATE] ");               
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ORDER BY [START_TIME],[SUBJECT]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");



                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    return ds3;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        return ds3;
                    }

                    return ds3;
                }

            }
            catch
            {
                return ds3;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        //新廠包裝線、新廠製一組、新廠製二組、新廠製三組(手工)的明細
        public DataSet SEARCHMANULINE4(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠包裝線'");
                sbSql.AppendFormat(" UNION");                
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製一組'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製二組'");               
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製三組(手工)%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='新廠製三組(手工)'");
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ORDER BY [START_TIME],[SUBJECT]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");



                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    return ds4;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        return ds4;
                    }

                    return ds4;
                }

            }
            catch
            {
                return ds4;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            ADDTOUOFTB_EIP_SCH_MEMO_MOC(DateTime.Now.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
