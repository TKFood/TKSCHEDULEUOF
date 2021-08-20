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
using System.Xml;
using System.Xml.Linq;
using System.Xml;
using System.Xml.Linq;


namespace TKSCHEDULEUOF
{
    public partial class FrmSCHEDULE : Form
    {
        //測試ID = "";
        //正式ID =""
        //測試DB DBNAME = "UOFTEST";
        //正式DB DBNAME = "UOF";
        string COPID = "b30a5a74-8785-4af3-8f50-e35488db05a3";
        string COPCHANGEID = "";

        string ID = "9cf7d919-c825-4b79-97e3-7f532f4fb8a6";
        string DBNAME = "UOF";

        string OLDTASK_ID = null;
        string NEWTASK_ID = null;
        string ATTACH_ID = null;

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
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();

        /// <summary>
        /// 新廠製一組桶數 BASELIMITHRSBAR1
        /// 新廠製二組桶數 BASELIMITHRSBAR2
        /// 新廠包裝線稼動率時數 BASELIMITHRS9
        /// 新廠製一組稼動率時數 BASELIMITHRS1
        /// 新廠製二組稼動率時數 BASELIMITHRS2
        /// 新廠製三組(手工)稼動率時數 BASELIMITHRS3
        /// </summary>
        decimal BASELIMITHRSBAR1 =0;
        decimal BASELIMITHRSBAR2 = 0;
        decimal BASELIMITHRS1 = 0;
        decimal BASELIMITHRS2 = 0;
        decimal BASELIMITHRS3 = 0;
        decimal BASELIMITHRS9 = 0;

        public FrmSCHEDULE()
        {
            InitializeComponent();

            timer1.Enabled = true;
            timer1.Interval = 1000*60 ;
            timer1.Start();

            timer2.Enabled = true;
            timer2.Interval = 1000 * 60;
            timer2.Start();
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

                BASELIMITHRSBAR1 = SEARCHBASELIMITHRS("新廠製一組桶數");
                BASELIMITHRSBAR1 = Math.Round(BASELIMITHRSBAR1, 0);
                BASELIMITHRSBAR2 = SEARCHBASELIMITHRS("新廠製二組桶數");
                BASELIMITHRSBAR2 = Math.Round(BASELIMITHRSBAR2, 0);

                BASELIMITHRS1 = SEARCHBASELIMITHRS("新廠製一組稼動率時數");
                BASELIMITHRS2 = SEARCHBASELIMITHRS("新廠製二組稼動率時數");
                BASELIMITHRS3 = SEARCHBASELIMITHRS("新廠製三組(手工)稼動率時數");
                BASELIMITHRS9 = SEARCHBASELIMITHRS("新廠包裝線稼動率時數");

                ADDTOUOFTB_EIP_SCH_MEMO_MOC(DateTime.Now.ToString("yyyyMMdd"));
                ADDTOUOFTB_EIP_SCH_MEMO_PUR(DateTime.Now.ToString("yyyyMMdd"));
                ADDTOUOFTB_EIP_SCH_MEMO_COP(DateTime.Now.ToString("yyyyMMdd"));
                UPDATEtb_COMPANYSTATUS1();
                UPDATEtb_COMPANYSTATUS2();
                UPDATEtb_COMPANYOWNER_ID();
                ADDtb_COMPANY();
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            ADDTOUOFOURTAB();
        }

        #region FUNCTION

        //
        public Decimal SEARCHBASELIMITHRS(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT  [ID],[LIMITHRS] FROM [TKMOC].[dbo].[BASELIMITHRS] WHERE [ID]='{0}'
                                    ", ID);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToDecimal(ds1.Tables["ds1"].Rows[0]["LIMITHRS"].ToString());
                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

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


                sbSql.AppendFormat(@" 
                                    SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]
                                    FROM (
                                    SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{1}桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{1}桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                    FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                    LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'
                                    WHERE INVMB.MB001=MOCMANULINE.MB001 
                                    AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}'
                                    AND [MOCMANULINE]. [MANU]='新廠製一組'
                                    AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])      
                                    GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                    UNION
                                    SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{2}桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{2}桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                    FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                    LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'
                                    WHERE INVMB.MB001=MOCMANULINE.MB001   
                                    AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}'
                                    AND [MOCMANULINE]. [MANU]='新廠製二組'
                                    AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                    GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                    ) AS TEMP
                                    ORDER BY [START_TIME],[SUBJECT]
                                    ", DateTime.Now.ToString("yyyyMMdd"), BASELIMITHRSBAR1, BASELIMITHRSBAR2);

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

                sbSql.AppendFormat(@" 
                                     SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]
                                     FROM (
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001   
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='新廠包裝線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]                
                                     UNION
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '  AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001 
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='新廠製一組'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                     UNION
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001   
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='新廠製二組'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                     UNION               
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)   AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製三組(手工)%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001 
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='新廠製三組(手工)'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]              
                                     ) AS TEMP
                                     ORDER BY [START_TIME],[SUBJECT]
                                    ", DateTime.Now.ToString("yyyyMMdd"));

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

               
                sbSql.AppendFormat(@" 
                                 SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]
                                 FROM (               
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{4}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{4}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001   
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='新廠包裝線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                 UNION
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{1}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{1}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001 
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='新廠製一組'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE] 
                                 UNION
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{2}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{2}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001   
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='新廠製二組'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                 UNION                
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{3}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{3}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製三組(手工)%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001 
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='新廠製三組(手工)'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]                
                                 ) AS TEMP
                                 ORDER BY [START_TIME],[SUBJECT]
                                 ", DateTime.Now.ToString("yyyyMMdd"), BASELIMITHRS1, BASELIMITHRS2, BASELIMITHRS3, BASELIMITHRS9);

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
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],'新廠製三組(手工)'+TA001+'-'+TA002+TA034+CONVERT(NVARCHAR,CONVERT(INT,TA015))+TA007 AS [DESCRIPTION],CONVERT(NVARCHAR,TA003,112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,TA003,112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],'新廠製三組(手工)'+TA001+'-'+TA002+TA034+CONVERT(NVARCHAR,CONVERT(INT,TA015))+TA007 AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(" WHERE  TA021='04'");
                sbSql.AppendFormat(" AND TA003>='{0}'", DateTime.Now.ToString("yyyyMMdd"));
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

        public void ADDTOUOFTB_EIP_SCH_MEMO_PUR(string Sday)
        {

            DataSet ds5 = new DataSet();
            ds5 = SEARCHMANULINE5(Sday);
          

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
        
                //先刪除再新增

                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='701e642b-c4d5-43ce-8289-c7dffb7ba016'", Sday);
                sbSql.AppendFormat(" ");

                if (ds5.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds5.Tables[0].Rows)
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

        //採購明細
        public DataSet SEARCHMANULINE5(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'701e642b-c4d5-43ce-8289-c7dffb7ba016' AS [CREATE_USER],TD005+'-'+CONVERT(NVARCHAR,(TD008-TD015))+' '+TD009 AS [DESCRIPTION],CONVERT(NVARCHAR,[TD012],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[TD012],112) AS [START_TIME],TD005+'-'+CONVERT(NVARCHAR,(TD008-TD015))+' '+TD009  AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'701e642b-c4d5-43ce-8289-c7dffb7ba016' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[PURTD],[TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE TD004=MB001");
                sbSql.AppendFormat(" AND TD016 IN ('N')");
                sbSql.AppendFormat(" AND (TD004 LIKE '1%' OR TD004 LIKE '2%')");
                sbSql.AppendFormat(" AND TD012>='{0}'",Sday);
                sbSql.AppendFormat(" ORDER BY TD012");
                sbSql.AppendFormat(" ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5= new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");



                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    return ds5;
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {
                        return ds5;
                    }

                    return ds5;
                }

            }
            catch
            {
                return ds5;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTOUOFTB_EIP_SCH_MEMO_COP(string Sday)
        {

            DataSet ds6 = new DataSet();
            DataSet ds7 = new DataSet();
            DataSet ds8 = new DataSet();
            ds6 = SEARCHMANULINE6(Sday);
            ds7 = SEARCHMANULINE7(Sday);
            ds8 = SEARCHMANULINE8(Sday);



            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                //先刪除再新增

                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='4eda2bfc-bf4b-4df2-a39c-1cc46e68598a'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='8e841f56-0a77-4b5c-9c7e-1fd05b089900'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='e6a83ac9-5ab4-4c5b-af50-1936a694ffe8'", Sday);
                sbSql.AppendFormat(" ");

                if (ds6.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds6.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }
                if (ds7.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds7.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }
                if (ds8.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds8.Tables[0].Rows)
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

        //訂單-國內
        public DataSet SEARCHMANULINE6(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'4eda2bfc-bf4b-4df2-a39c-1cc46e68598a' AS [CREATE_USER],TC053+'-'+TD005+'-'+CONVERT(NVARCHAR,CONVERT(INT,(TD008-TD009)))+' ('+TD010+') 贈品'+CONVERT(NVARCHAR,CONVERT(INT,(TD024-TD025)))+' ('+TD010+')' AS [DESCRIPTION],CONVERT(NVARCHAR,[TD013],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[TD013],112) AS [START_TIME],TC053+'-'+TD005+'-'+CONVERT(NVARCHAR,CONVERT(INT,(TD008-TD009)))+' ('+TD010+') 贈品'+CONVERT(NVARCHAR,CONVERT(INT,(TD024-TD025)))+' ('+TD010+')' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'4eda2bfc-bf4b-4df2-a39c-1cc46e68598a' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[COPTC],[TK].[dbo].[COPTD],[TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD004=MB001");
                sbSql.AppendFormat(" AND TD016 IN ('N')");
                sbSql.AppendFormat(" AND TD001 IN ('A221','A223','A227','A229')");
                sbSql.AppendFormat(" AND TD013>='{0}'",Sday);
                sbSql.AppendFormat(" ORDER BY TD013,TC001,TC002");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter6 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder6 = new SqlCommandBuilder(adapter6);
                sqlConn.Open();
                ds6.Clear();
                adapter6.Fill(ds6, "ds6");



                if (ds6.Tables["ds6"].Rows.Count == 0)
                {
                    return ds6;
                }
                else
                {
                    if (ds6.Tables["ds6"].Rows.Count >= 1)
                    {
                        return ds6;
                    }

                    return ds6;
                }

            }
            catch
            {
                return ds6;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataSet SEARCHMANULINE7(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'8e841f56-0a77-4b5c-9c7e-1fd05b089900' AS [CREATE_USER],TC053+'-'+TD005+'-'+CONVERT(NVARCHAR,CONVERT(INT,(TD008-TD009)))+' ('+TD010+') 贈品'+CONVERT(NVARCHAR,CONVERT(INT,(TD024-TD025)))+' ('+TD010+')' AS [DESCRIPTION],CONVERT(NVARCHAR,[TD013],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[TD013],112) AS [START_TIME],TC053+'-'+TD005+'-'+CONVERT(NVARCHAR,CONVERT(INT,(TD008-TD009)))+' ('+TD010+') 贈品'+CONVERT(NVARCHAR,CONVERT(INT,(TD024-TD025)))+' ('+TD010+')' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'8e841f56-0a77-4b5c-9c7e-1fd05b089900' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[COPTC],[TK].[dbo].[COPTD],[TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD004=MB001");
                sbSql.AppendFormat(" AND TD016 IN ('N')");
                sbSql.AppendFormat(" AND TD001 IN ('A222','A227')");
                sbSql.AppendFormat(" AND TD013>='{0}'", Sday);
                sbSql.AppendFormat(" ORDER BY TD013,TC001,TC002");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "ds7");



                if (ds7.Tables["ds7"].Rows.Count == 0)
                {
                    return ds7;
                }
                else
                {
                    if (ds7.Tables["ds7"].Rows.Count >= 1)
                    {
                        return ds7;
                    }

                    return ds7;
                }

            }
            catch
            {
                return ds7;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataSet SEARCHMANULINE8(string Sday)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'e6a83ac9-5ab4-4c5b-af50-1936a694ffe8' AS [CREATE_USER],TC053+'-'+TD005+'-'+CONVERT(NVARCHAR,CONVERT(INT,(TD008-TD009)))+' ('+TD010+') 贈品'+CONVERT(NVARCHAR,CONVERT(INT,(TD024-TD025)))+' ('+TD010+')' AS [DESCRIPTION],CONVERT(NVARCHAR,[TD013],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[TD013],112) AS [START_TIME],TC053+'-'+TD005+'-'+CONVERT(NVARCHAR,CONVERT(INT,(TD008-TD009)))+' ('+TD010+') 贈品'+CONVERT(NVARCHAR,CONVERT(INT,(TD024-TD025)))+' ('+TD010+')' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'e6a83ac9-5ab4-4c5b-af50-1936a694ffe8' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[COPTC],[TK].[dbo].[COPTD],[TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD004=MB001");
                sbSql.AppendFormat(" AND TD016 IN ('N')");
                sbSql.AppendFormat(" AND TD001 IN ('A224','A225','A226','A228')");
                sbSql.AppendFormat(" AND TD013>='{0}'", Sday);
                sbSql.AppendFormat(" ORDER BY TD013,TC001,TC002");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                sqlConn.Open();
                ds8.Clear();
                adapter8.Fill(ds8, "ds8");



                if (ds8.Tables["ds8"].Rows.Count == 0)
                {
                    return ds8;
                }
                else
                {
                    if (ds8.Tables["ds8"].Rows.Count >= 1)
                    {
                        return ds8;
                    }

                    return ds8;
                }

            }
            catch
            {
                return ds8;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEtb_COMPANYSTATUS1()
        {
            DataSet dsCOMPA = new DataSet();
            dsCOMPA = SERACHCOMPA();

            if (dsCOMPA.Tables[0].Rows.Count > 0)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    //更新          
                    foreach (DataRow dr in dsCOMPA.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(@" UPDATE [HJ_BM_DB].[dbo].[tb_COMPANY]");
                        sbSql.AppendFormat(@" SET [STATUS]='1'");
                        sbSql.AppendFormat(@" WHERE [ERPNO]='{0}' ", dr["MA001"].ToString());
                        sbSql.AppendFormat(@" ");
                    }

                    sbSql.AppendFormat(@" ");

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

        }

        public DataSet SERACHCOMPA()
        {
            DataSet ds = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //找出在ERP是被停用的客戶，但是在UOF沒有被停用
                sbSql.AppendFormat(" SELECT MA001");
                sbSql.AppendFormat(" FROM [TK].dbo.COPMA");
                sbSql.AppendFormat(" WHERE ISNULL(UDF01,'')<>'Y'");
                sbSql.AppendFormat(" AND MA001 IN (SELECT [ERPNO] FROM [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE ISNULL([ERPNO],'')<>''  )");
                sbSql.AppendFormat(" AND MA001 IN (SELECT [ERPNO] FROM [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE ISNULL([ERPNO],'')<>'' AND [STATUS]='2' )");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");



                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    return ds;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        return ds;
                    }

                    return ds;
                }

            }
            catch
            {
                return ds;
            }
            finally
            {
                sqlConn.Close();
            }

        }

        public void UPDATEtb_COMPANYSTATUS2()
        {
            DataSet dsCOMPASTOP = new DataSet();
            dsCOMPASTOP=SERACHCOMPASTOP();

            if (dsCOMPASTOP.Tables[0].Rows.Count>0)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    //更新          
                    foreach (DataRow dr in dsCOMPASTOP.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(@" UPDATE [HJ_BM_DB].[dbo].[tb_COMPANY]");
                        sbSql.AppendFormat(@" SET [STATUS]='2'");
                        sbSql.AppendFormat(@" WHERE [ERPNO]='{0}' ", dr["MA001"].ToString());
                        sbSql.AppendFormat(@" ");
                    }

                    sbSql.AppendFormat(@" ");

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

        }

        public DataSet SERACHCOMPASTOP()
        {
            DataSet ds = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //找出在ERP是被停用的客戶，但是在UOF沒有被停用
                sbSql.AppendFormat(" SELECT MA001");
                sbSql.AppendFormat(" FROM [TK].dbo.COPMA");
                sbSql.AppendFormat(" WHERE UDF01='Y'");
                sbSql.AppendFormat(" AND MA001 IN (SELECT [ERPNO] FROM [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE ISNULL([ERPNO],'')<>''  )");
                sbSql.AppendFormat(" AND MA001 IN (SELECT [ERPNO] FROM [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE ISNULL([ERPNO],'')<>'' AND [STATUS]='1' )");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");



                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    return ds;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        return ds;
                    }

                    return ds;
                }

            }
            catch
            {
                return ds;
            }
            finally
            {
                sqlConn.Close();
            }

        }

        public void ADDtb_COMPANY()
        {
            DataSet dsCOPMA = new DataSet();
            dsCOPMA = SERACHdsCOPMA();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                //更新          
                foreach (DataRow dr in dsCOPMA.Tables[0].Rows)
                {                   
                    sbSql.AppendFormat(@" 
                                        INSERT INTO [HJ_BM_DB].[dbo].[tb_COMPANY]
                                        ([COMPANY_NAME],[ERPNO],[TAX_NUMBER],[PHONE],[FAX],[COUNTRY],[CITY],[TOWN],[ADDRESS],[OVERSEAS_ADDR]
                                        ,[EMAIL],[WEBSITE],[FACEBOOK],[INDUSTRY],[TURNOVER],[WORKER_NUMBER],[EST_DATE],[PARENT_ID],[UPDATE_DATETIME],[CREATE_DATETIME]
                                        ,[CREATE_USER_ID],[UPDATE_USER_ID],[NOTE],[OWNER_ID],[LAST_CONTACT_DATE],[STATUS])
                                        VALUES
                                        ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'
                                        ,'{10}','{11}','{12}','{13}','{14}','{15}','{16}',{17},'{18}','{19}'
                                        ,'{20}','{21}','{22}','{23}','{24}','{25}')
                                        ", dr["COMPANY_NAME"].ToString(), dr["ERPNO"].ToString(), dr["TAX_NUMBER"].ToString(), dr["PHONE"].ToString(), dr["FAX"].ToString(), dr["COUNTRY"].ToString(), dr["CITY"].ToString(), dr["TOWN"].ToString(), dr["ADDRESS"].ToString(), dr["OVERSEAS_ADDR"].ToString()
                                        , dr["EMAIL"].ToString(), dr["WEBSITE"].ToString(), dr["FACEBOOK"].ToString(), dr["INDUSTRY"].ToString(), dr["TURNOVER"].ToString(), dr["WORKER_NUMBER"].ToString(), dr["EST_DATE"].ToString(),"NULL", dr["UPDATE_DATETIME"].ToString(), dr["CREATE_DATETIME"].ToString()
                                        , dr["CREATE_USER_ID"].ToString(), dr["UPDATE_USER_ID"].ToString(), dr["NOTE"].ToString(), dr["OWNER_ID"].ToString(), dr["LAST_CONTACT_DATE"].ToString(), dr["STATUS"].ToString());
                }
                sbSql.AppendFormat(@" ");

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

        public void UPDATEtb_COMPANYOWNER_ID()
        {
            //DataSet dsCOMPAMA016 = new DataSet();
            //dsCOMPAMA016 = SERACHCOMPAMA016();

            //try
            //{
            //    connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
            //    sqlConn = new SqlConnection(connectionString);

            //    sqlConn.Close();
            //    sqlConn.Open();
            //    tran = sqlConn.BeginTransaction();

            //    sbSql.Clear();

            //    //更新          
            //    foreach (DataRow dr in dsCOMPAMA016.Tables[0].Rows)
            //    {
            //        sbSql.AppendFormat(@" UPDATE [HJ_BM_DB].[dbo].[tb_COMPANY]");
            //        sbSql.AppendFormat(@" SET [OWNER_ID]='{0}'", dr["USER_ID"].ToString());
            //        sbSql.AppendFormat(@" WHERE [ERPNO]='{0}' ", dr["ERPNO"].ToString());
            //        sbSql.AppendFormat(@" ");
            //    }
            //    sbSql.AppendFormat(@" ");

            //    cmd.Connection = sqlConn;
            //    cmd.CommandTimeout = 60;
            //    cmd.CommandText = sbSql.ToString();
            //    cmd.Transaction = tran;
            //    result = cmd.ExecuteNonQuery();

            //    if (result == 0)
            //    {
            //        tran.Rollback();    //交易取消
            //    }
            //    else
            //    {
            //        tran.Commit();      //執行交易  

            //    }

            //}
            //catch
            //{

            //}

            //finally
            //{
            //    sqlConn.Close();
            //}
        }

        public DataSet SERACHCOMPAMA016()
        {
            DataSet ds = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //找出在ERP是被停用的客戶，但是在UOF沒有被停用
                sbSql.AppendFormat(" SELECT [ERPNO],[OWNER_ID],MA001,MA016,[USER_ACCOUNT],[USER_ID]");
                sbSql.AppendFormat(" FROM [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY],[TK].dbo.COPMA");
                sbSql.AppendFormat(" LEFT JOIN [192.168.1.223].[HJ_BM_DB].[dbo].[tb_USER] ON [tb_USER].[USER_ACCOUNT]=COPMA.MA016");
                sbSql.AppendFormat(" WHERE MA001=[ERPNO] ");
                sbSql.AppendFormat(" AND ISNULL(MA016,'')<>''");
                sbSql.AppendFormat(" AND [OWNER_ID]<>[USER_ID]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");



                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    return ds;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        return ds;
                    }

                    return ds;
                }

            }
            catch
            {
                return ds;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataSet SERACHdsCOPMA()
        {
            DataSet ds = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //               
                sbSql.AppendFormat(@" 
                                    SELECT 
                                    MA002 [COMPANY_NAME]
                                    ,MA001 [ERPNO]
                                    ,MA010 [TAX_NUMBER]
                                    ,MA006 [PHONE]
                                    ,MA008 [FAX]
                                    ,'Taiwan (台灣)' [COUNTRY]
                                    ,'' [CITY]
                                    ,'' [TOWN]
                                    ,'' [ADDRESS]
                                    ,'' [OVERSEAS_ADDR]
                                    ,MA009 [EMAIL]
                                    ,'' [WEBSITE]
                                    ,'' [FACEBOOK]
                                    ,'' [INDUSTRY]
                                    ,'0' [TURNOVER]
                                    ,'0' [WORKER_NUMBER]
                                    ,'' [EST_DATE]
                                    ,'' [PARENT_ID]
                                    ,CONVERT(nvarchar,GETDATE(),111)  [UPDATE_DATETIME]
                                    ,CONVERT(nvarchar,GETDATE(),111)  [CREATE_DATETIME]
                                    ,[USER_ID]  [CREATE_USER_ID]
                                    ,[USER_ID]  [UPDATE_USER_ID]
                                    ,''[NOTE]
                                    ,[USER_ID] [OWNER_ID]
                                    ,'' [LAST_CONTACT_DATE]
                                    ,'1' [STATUS]
                                    FROM [TK].dbo.COPMA
                                    LEFT JOIN [192.168.1.223].[HJ_BM_DB].[dbo].[tb_USER] ON [tb_USER].[USER_ACCOUNT]=MA016
                                    WHERE MA001 NOT IN (SELECT [ERPNO] FROM [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE ISNULL([ERPNO] ,'')<>'')
                                    AND MA001 NOT LIKE '1%'
                                    AND MA001 NOT LIKE '299%'
                                    AND MA001 NOT LIKE '399%'
                                    AND MA001 NOT LIKE '4%'
                                    AND MA001 NOT LIKE '5%'
                                    AND MA001 NOT LIKE '6%'
                                    AND MA001 NOT LIKE '7%'
                                    AND MA001 NOT LIKE '910%'
                                    AND MA001 NOT LIKE '990%'
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");



                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    return ds;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        return ds;
                    }

                    return ds;
                }

            }
            catch
            {
                return ds;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTB_WKF_EXTERNAL_TASK(string TA001, string TA002)
        {
            DataTable DT = SEARCHPURTAPURTB(TA001, TA002);
            DataTable DTUPFDEP = SEARCHUOFDEP(DT.Rows[0]["TA012"].ToString());

            string account = DT.Rows[0]["TA012"].ToString();
            string groupId = DT.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
            string fillerName= DT.Rows[0]["MV002"].ToString();
            string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = DT.Rows[0]["TA001"].ToString().Trim() + DT.Rows[0]["TA002"].ToString().Trim() ;

            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");
           
            //正式的id
            Form.SetAttribute("formVersionId", ID);

            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點 FormFieldValue
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //ID 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "ID");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //QC	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QC");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["QC"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //DEPNO 變更版本	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "DEPNO");
            FieldItem.SetAttribute("fieldValue", DEPNAME);
            FieldItem.SetAttribute("realValue", DEPNO);
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA001 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA001");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA001"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA002 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA003 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA003"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA012 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA012");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA012"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //MV002 姓名	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "MV002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["MV002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA006 單頭備註	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA006");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA006"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TB 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TB");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點 DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            //DataGrid 加入至 TB 節點底下
            XmlNode TB = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='TB']");
            TB.AppendChild(DataGrid);

           
            foreach (DataRow od in DT.Rows)
            {
                // 新增 Row
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", (rowscounts).ToString());

                //Row	TB004
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB004");
                Cell.SetAttribute("fieldValue", od["TB004"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB005
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB005");
                Cell.SetAttribute("fieldValue", od["TB005"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB006
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB006");
                Cell.SetAttribute("fieldValue", od["TB006"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB007
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB007");
                Cell.SetAttribute("fieldValue", od["TB007"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB009
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB009");
                Cell.SetAttribute("fieldValue", od["TB009"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	SUMLA011 可用庫存數量
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "SUMLA011");
                Cell.SetAttribute("fieldValue", od["SUMLA011"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB011
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB011");
                Cell.SetAttribute("fieldValue", od["TB011"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB010
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB010");
                Cell.SetAttribute("fieldValue", od["TB010"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	MA002
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MA002");
                Cell.SetAttribute("fieldValue", od["MA002"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB012
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB012");
                Cell.SetAttribute("fieldValue", od["TB012"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='TB']/DataGrid");
                DataGridS.AppendChild(Row);

            }

            ////用ADDTACK，直接啟動起單
            //ADDTACK(Form);

            //ADD TO DB
            string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [{0}].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{1}')
                                        ", DBNAME, EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }





        }

        public void ADDTACK(XmlElement Form)
        {
            //Ede.Uof.WKF.Utility.TaskUtilityUCO taskUCO = new Ede.Uof.WKF.Utility.TaskUtilityUCO();

            //string result = taskUCO.WebService_CreateTask(Form.OuterXml);

            //XElement resultXE = XElement.Parse(result);

            //string status = "";
            //string formNBR = "";
            //string error = "";

            //if (resultXE.Element("Status").Value == "1")
            //{
            //    status = "起單成功!";
            //    formNBR = resultXE.Element("FormNumber").Value;
            //    NEWTASK_ID = formNBR;

            //    //Logger.Write("TEST", status + formNBR);

            //}
            //else
            //{
            //    status = "起單失敗!";
            //    error = resultXE.Element("Exception").Element("Message").Value;

            //    //Logger.Write("TEST", status + error + "\r\n" + Form.OuterXml);

            //    throw new Exception(status + error + "\r\n" + Form.OuterXml);

            //}
        }


        public DataTable SEARCHPURTAPURTB(string TA001,string TA002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //庫存數量看LA009 IN ('20004','20006','20008','20019','20020'

                sbSql.AppendFormat(@"  
                                   SELECT CREATOR,TA001,TA002,TA003,TA012,TB004,TB005,TB006,TB007,TB009,TB011,TA006,TB012,MV002,UDF03 QC
                                    ,USER_GUID,NAME
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    ,TB010,MA002,SUMLA011
                                    FROM 
                                    (
                                    SELECT PURTA.CREATOR,TA001,TA002,TA003,TA012,TB004,TB005,TB006,TB007,TB009,TB011,TA006,TB012,TB010,PURTA.UDF03
                                    ,[TB_EB_USER].USER_GUID,NAME
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=TA012) AS 'MV002'
                                    ,(SELECT TOP 1 MA002 FROM [TK].dbo.PURMA WHERE MA001=TB010) AS 'MA002'
                                    ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA001=TB004 AND LA009 IN ('20004','20006','20008','20019','20020')) AS SUMLA011
                                    FROM [TK].dbo.PURTB,[TK].dbo.PURTA
                                    LEFT JOIN [192.168.1.223].[{0}].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= TA012 COLLATE Chinese_Taiwan_Stroke_BIN
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA001='{1}' AND TA002='{2}'
                                    ) AS TEMP
                              
                                    ", DBNAME, TA001, TA002);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }


        public void ADDTOUOFOURTAB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                DataSet ds1 = new DataSet();
                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT TA001,TA002
                                    FROM [TK].dbo.PURTA
                                    WHERE TA007='N' AND UDF01='Y'
                                    ORDER BY TA001,TA002
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    foreach (DataRow dr in ds1.Tables["ds1"].Rows)
                    {
                        ADDTB_WKF_EXTERNAL_TASK(dr["TA001"].ToString().Trim(), dr["TA002"].ToString().Trim());
                    }
                        

                    //ADDTB_WKF_EXTERNAL_TASK("A311", "20210415007");
                }
                else
                {

                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            UPDATEPURTAUDF01();
        }

        public void UPDATEPURTAUDF01()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                              
                sbSql.AppendFormat(@"
                                    UPDATE  [TK].dbo.PURTA  
                                    SET UDF01 = 'UOF',TA016='1'
                                    WHERE TA007 = 'N' AND UDF01 = 'Y'       
                                    ");

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

        public DataTable SEARCHUOFDEP(string ACCOUNT)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [GROUP_NAME] AS 'DEPNAME'
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]+','+[GROUP_NAME]+',False' AS 'DEPNO'
                                    ,[TB_EB_USER].[USER_GUID]
                                    ,[ACCOUNT]
                                    ,[NAME]
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]
                                    ,[TITLE_ID]     
                                    ,[GROUP_NAME]
                                    ,[GROUP_CODE]
                                    FROM [192.168.1.223].[{0}].[dbo].[TB_EB_USER],[192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP],[192.168.1.223].[{0}].[dbo].[TB_EB_GROUP]
                                    WHERE [TB_EB_USER].[USER_GUID]=[TB_EB_EMPL_DEP].[USER_GUID]
                                    AND [TB_EB_EMPL_DEP].[GROUP_ID]=[TB_EB_GROUP].[GROUP_ID]
                                    AND ISNULL([TB_EB_GROUP].[GROUP_CODE],'')<>''
                                    AND [ACCOUNT]='{1}'
                              
                                    ", DBNAME, ACCOUNT);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDCOPTCCOPTD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                DataSet ds1 = new DataSet();
                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT TC001,TC002
                                    FROM [TK].dbo.COPTC
                                    WHERE TC027='N' AND UDF02='Y'
                                    ORDER BY TC001,TC002
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    foreach (DataRow dr in ds1.Tables["ds1"].Rows)
                    {
                        ADDTB_WKF_EXTERNAL_TASK_COPTCCOPTD(dr["TC001"].ToString().Trim(), dr["TC002"].ToString().Trim());
                    }


                    //ADDTB_WKF_EXTERNAL_TASK("A311", "20210415007");
                }
                else
                {

                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            //UPDATECOPTCUDF02();
        }

        public void ADDTB_WKF_EXTERNAL_TASK_COPTCCOPTD(string TC001,string TC002)
        {

            DataTable DT = SEARCHCOPTCCOPTD(TC001, TC002);
            DataTable DTUPFDEP = SEARCHUOFDEP(DT.Rows[0]["TC006"].ToString());

            string account = DT.Rows[0]["TC006"].ToString();
            string groupId = DT.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
            string fillerName = DT.Rows[0]["NAME"].ToString();
            string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = DT.Rows[0]["TC001"].ToString().Trim() + DT.Rows[0]["TC002"].ToString().Trim();

            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //正式的id
            Form.SetAttribute("formVersionId", COPID);

            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點 FormFieldValue
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //ID 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "ID");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);



            //建立節點FieldItem
            //TC001 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC001");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC001"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC002 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC003 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC003"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC004 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC004");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC004"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC053 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC053");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC053"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC006 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC006");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC006"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //MV002 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "MV002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC015 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC015");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC015"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC008 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC008");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC008"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC009 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC009");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC009"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC045 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC045");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC045"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC029 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC029");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC029"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC030 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC030");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC030"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC041 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC041");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC041"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC016 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC016");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NEWTC016"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC124 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC124");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC124"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC031 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC031");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC031"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC043 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC043");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC043"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC044 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC044");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC044"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC046 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC046");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC046"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC018 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC018");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC018"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC010 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC010");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC010"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC012 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC012");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC012"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC035 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC035");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC035"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC054 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC054");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC054"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC055 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC055");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC055"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC065 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC065");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC065"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC042 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC042");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC042"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC014 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC014");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC014"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC019 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC019");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC019"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC032 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC032");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC032"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC033 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC033");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC033"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC039 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC039");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC039"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC121 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC121");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NEWTC016"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC094 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC094");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC094"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC063 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC063");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC063"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC115 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC115");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC115"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC116 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC116");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC116"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //MOC 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "MOC");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //PUR 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "PUR");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);



            //建立節點FieldItem
            //DETAILS 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "DETAILS");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點 DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            //DataGrid 加入至 TB 節點底下
            XmlNode DETAILS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']");
            DETAILS.AppendChild(DataGrid);

           

            foreach (DataRow od in DT.Rows)
            {
                // 新增 Row
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", (rowscounts).ToString());

                //Row	UDF01
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "UDF01");
                Cell.SetAttribute("fieldValue", od["UDF01"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD003
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD003");
                Cell.SetAttribute("fieldValue", od["TD003"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD004
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD004");
                Cell.SetAttribute("fieldValue", od["TD004"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD005
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD005");
                Cell.SetAttribute("fieldValue", od["TD005"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD006
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD006");
                Cell.SetAttribute("fieldValue", od["TD006"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD008
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD008");
                Cell.SetAttribute("fieldValue", od["TD008"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD024
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD024");
                Cell.SetAttribute("fieldValue", od["TD024"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD009
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD009");
                Cell.SetAttribute("fieldValue", od["TD009"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD025
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD025");
                Cell.SetAttribute("fieldValue", od["TD025"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD003
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD003");
                Cell.SetAttribute("fieldValue", od["TD003"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD010
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD010");
                Cell.SetAttribute("fieldValue", od["TD010"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD011
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD011");
                Cell.SetAttribute("fieldValue", od["TD011"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD026
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD026");
                Cell.SetAttribute("fieldValue", od["TD026"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD012
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD012");
                Cell.SetAttribute("fieldValue", od["TD012"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD003
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD003");
                Cell.SetAttribute("fieldValue", od["TD003"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD013
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD013");
                Cell.SetAttribute("fieldValue", od["TD013"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD017
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD017");
                Cell.SetAttribute("fieldValue", od["TD017"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD018
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD018");
                Cell.SetAttribute("fieldValue", od["TD018"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD019
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD019");
                Cell.SetAttribute("fieldValue", od["TD019"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD020
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD020");
                Cell.SetAttribute("fieldValue", od["TD020"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);


                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']/DataGrid");
                DataGridS.AppendChild(Row);

            }

            ////用ADDTACK，直接啟動起單
            //ADDTACK(Form);

            //ADD TO DB
            string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [{0}].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{1}')
                                        ", DBNAME, EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }
        }


        public void UPDATECOPTCUDF02()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"
                                    UPDATE  [TK].dbo.COPTC 
                                    SET UDF02 = 'UOF'
                                    WHERE TC027='N' AND UDF02='Y'     
                                    ");

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
        public DataTable SEARCHCOPTCCOPTD(string TC001, string TC002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //庫存數量看LA009 IN ('20004','20006','20008','20019','20020'

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME
                                    ,sync_date,sync_time,sync_mark,sync_count,DataUser,DataGroup
                                    ,TC001,TC002,TC003,TC004,TC005,TC006,TC007,TC008,TC009,TC010
                                    ,TC011,TC012,TC013,TC014,TC015,TC016,TC017,TC018,TC019,TC020
                                    ,TC021,TC022,TC023,TC024,TC025,TC026,TC027,TC028,TC029,TC030
                                    ,TC031,TC032,TC033,TC034,TC035,TC036,TC037,TC038,TC039,TC040
                                    ,TC041,TC042,TC043,TC044,TC045,TC046,TC047,TC048,TC049,TC050
                                    ,TC051,TC052,TC053,TC054,TC055,TC056,TC057,TC058,TC059,TC060
                                    ,TC061,TC062,TC063,TC064,TC065,TC066,TC067,TC068,TC069,TC070
                                    ,TC071,TC072,TC073,TC074,TC075,TC076,TC077,TC078,TC079,TC080
                                    ,TC081,TC082,TC083,TC084,TC085,TC086,TC087,TC088,TC089,TC090
                                    ,TC091,TC092,TC093,TC094,TC095,TC096,TC097,TC098,TC099,TC100
                                    ,TC101,TC102,TC103,TC104,TC105,TC106,TC107,TC108,TC109,TC110
                                    ,TC111,TC112,TC113,TC114,TC115,TC116,TC117,TC118,TC119,TC120
                                    ,TC121,TC122,TC123,TC124,TC125,TC126,TC127,TC128,TC129,TC130
                                    ,TC131,TC132,TC133,TC134,TC135,TC136,TC137,TC138,TC139,TC140
                                    ,TC141,TC142,TC143,TC144,TC145,TC146
                                    ,UDF01,UDF02,UDF03,UDF04,UDF05,UDF06,UDF07,UDF08,UDF09,UDF10
                                    ,TD001,TD002,TD003,TD004,TD005,TD006,TD007,TD008,TD009,TD010
                                    ,TD011,TD012,TD013,TD014,TD015,TD016,TD017,TD018,TD019,TD020
                                    ,TD021,TD022,TD023,TD024,TD025,TD026,TD027,TD028,TD029,TD030
                                    ,TD031,TD032,TD033,TD034,TD035,TD036,TD037,TD038,TD039,TD040
                                    ,TD041,TD042,TD043,TD044,TD045,TD046,TD047,TD048,TD049,TD050
                                    ,TD051,TD052,TD053,TD054,TD055,TD056,TD057,TD058,TD059,TD060
                                    ,TD061,TD062,TD063,TD064,TD065,TD066,TD067,TD068,TD069,TD070
                                    ,TD071,TD072,TD073,TD074,TD075,TD076,TD077,TD078,TD079,TD080
                                    ,TD081,TD082,TD083,TD084,TD085,TD086,TD087,TD088,TD089,TD090
                                    ,TD091,TD092,TD093,TD094,TD095,TD096,TD097,TD098,TD099,TD100
                                    ,TD101,TD102,TD103,TD104,TD105,TD106,TD107,TD108,TD109,TD110
                                    ,TD111,TD112,TD113
                                    ,COPTDUDF01,COPTDUDF02,COPTDUDF03,COPTDUDF04,COPTDUDF05,COPTDUDF06,COPTDUDF07,COPTDUDF08,COPTDUDF09,COPTDUDF10,TD200
                                    ,USER_GUID,NAME
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    ,MA002
                                    ,CASE WHEN TC016='1' THEN '1.應稅內含'  ELSE (CASE WHEN TC016='2' THEN '2.應稅外加'  ELSE (CASE WHEN TC016='3' THEN '3.零稅率'  ELSE (CASE WHEN TC016='4' THEN '4.免稅'  ELSE (CASE WHEN TC016='9' THEN '9.不計稅'  ELSE '' END) END) END) END ) END AS 'NEWTC016'
                                    ,CASE WHEN TC121='1' THEN '1.二聯式' ELSE (CASE WHEN TC121='2' THEN '2.三聯式' ELSE (CASE WHEN TC121='3' THEN '3.二聯式收銀機發票' ELSE (CASE WHEN TC121='4' THEN '4.三聯式收銀機發票' ELSE (CASE WHEN TC121='5' THEN '5.電子計算機發票' ELSE (CASE WHEN TC121='6' THEN '6.免用統一發票' ELSE (CASE WHEN TC121='7' THEN '7.電子發票' ELSE '' END) END) END) END) END) END) END AS 'NEWTC121'
                                    FROM 
                                    (
                                    SELECT [COPTC].[COMPANY],[COPTC].[CREATOR],[COPTC].[USR_GROUP],[COPTC].[CREATE_DATE],[COPTC].[MODIFIER],[COPTC].[MODI_DATE],[COPTC].[FLAG],[COPTC].[CREATE_TIME],[COPTC].[MODI_TIME],[COPTC].[TRANS_TYPE],[COPTC].[TRANS_NAME]
                                    ,[COPTC].[sync_date],[COPTC].[sync_time],[COPTC].[sync_mark],[COPTC].[sync_count],[COPTC].[DataUser],[COPTC].[DataGroup]
                                    ,[COPTC].[TC001],[COPTC].[TC002],[COPTC].[TC003],[COPTC].[TC004],[COPTC].[TC005],[COPTC].[TC006],[COPTC].[TC007],[COPTC].[TC008],[COPTC].[TC009],[COPTC].[TC010]
                                    ,[COPTC].[TC011],[COPTC].[TC012],[COPTC].[TC013],[COPTC].[TC014],[COPTC].[TC015],[COPTC].[TC016],[COPTC].[TC017],[COPTC].[TC018],[COPTC].[TC019],[COPTC].[TC020]
                                    ,[COPTC].[TC021],[COPTC].[TC022],[COPTC].[TC023],[COPTC].[TC024],[COPTC].[TC025],[COPTC].[TC026],[COPTC].[TC027],[COPTC].[TC028],[COPTC].[TC029],[COPTC].[TC030]
                                    ,[COPTC].[TC031],[COPTC].[TC032],[COPTC].[TC033],[COPTC].[TC034],[COPTC].[TC035],[COPTC].[TC036],[COPTC].[TC037],[COPTC].[TC038],[COPTC].[TC039],[COPTC].[TC040]
                                    ,[COPTC].[TC041],[COPTC].[TC042],[COPTC].[TC043],[COPTC].[TC044],[COPTC].[TC045],[COPTC].[TC046],[COPTC].[TC047],[COPTC].[TC048],[COPTC].[TC049],[COPTC].[TC050]
                                    ,[COPTC].[TC051],[COPTC].[TC052],[COPTC].[TC053],[COPTC].[TC054],[COPTC].[TC055],[COPTC].[TC056],[COPTC].[TC057],[COPTC].[TC058],[COPTC].[TC059],[COPTC].[TC060]
                                    ,[COPTC].[TC061],[COPTC].[TC062],[COPTC].[TC063],[COPTC].[TC064],[COPTC].[TC065],[COPTC].[TC066],[COPTC].[TC067],[COPTC].[TC068],[COPTC].[TC069],[COPTC].[TC070]
                                    ,[COPTC].[TC071],[COPTC].[TC072],[COPTC].[TC073],[COPTC].[TC074],[COPTC].[TC075],[COPTC].[TC076],[COPTC].[TC077],[COPTC].[TC078],[COPTC].[TC079],[COPTC].[TC080]
                                    ,[COPTC].[TC081],[COPTC].[TC082],[COPTC].[TC083],[COPTC].[TC084],[COPTC].[TC085],[COPTC].[TC086],[COPTC].[TC087],[COPTC].[TC088],[COPTC].[TC089],[COPTC].[TC090]
                                    ,[COPTC].[TC091],[COPTC].[TC092],[COPTC].[TC093],[COPTC].[TC094],[COPTC].[TC095],[COPTC].[TC096],[COPTC].[TC097],[COPTC].[TC098],[COPTC].[TC099],[COPTC].[TC100]
                                    ,[COPTC].[TC101],[COPTC].[TC102],[COPTC].[TC103],[COPTC].[TC104],[COPTC].[TC105],[COPTC].[TC106],[COPTC].[TC107],[COPTC].[TC108],[COPTC].[TC109],[COPTC].[TC110]
                                    ,[COPTC].[TC111],[COPTC].[TC112],[COPTC].[TC113],[COPTC].[TC114],[COPTC].[TC115],[COPTC].[TC116],[COPTC].[TC117],[COPTC].[TC118],[COPTC].[TC119],[COPTC].[TC120]
                                    ,[COPTC].[TC121],[COPTC].[TC122],[COPTC].[TC123],[COPTC].[TC124],[COPTC].[TC125],[COPTC].[TC126],[COPTC].[TC127],[COPTC].[TC128],[COPTC].[TC129],[COPTC].[TC130]
                                    ,[COPTC].[TC131],[COPTC].[TC132],[COPTC].[TC133],[COPTC].[TC134],[COPTC].[TC135],[COPTC].[TC136],[COPTC].[TC137],[COPTC].[TC138],[COPTC].[TC139],[COPTC].[TC140]
                                    ,[COPTC].[TC141],[COPTC].[TC142],[COPTC].[TC143],[COPTC].[TC144],[COPTC].[TC145],[COPTC].[TC146]
                                    ,[COPTC].[UDF01],[COPTC].[UDF02],[COPTC].[UDF03],[COPTC].[UDF04],[COPTC].[UDF05],[COPTC].[UDF06],[COPTC].[UDF07],[COPTC].[UDF08],[COPTC].[UDF09],[COPTC].[UDF10]
                                    ,[COPTD].[TD001],[COPTD].[TD002],[COPTD].[TD003],[COPTD].[TD004],[COPTD].[TD005],[COPTD].[TD006],[COPTD].[TD007],[COPTD].[TD008],[COPTD].[TD009],[COPTD].[TD010]
                                    ,[COPTD].[TD011],[COPTD].[TD012],[COPTD].[TD013],[COPTD].[TD014],[COPTD].[TD015],[COPTD].[TD016],[COPTD].[TD017],[COPTD].[TD018],[COPTD].[TD019],[COPTD].[TD020]
                                    ,[COPTD].[TD021],[COPTD].[TD022],[COPTD].[TD023],[COPTD].[TD024],[COPTD].[TD025],[COPTD].[TD026],[COPTD].[TD027],[COPTD].[TD028],[COPTD].[TD029],[COPTD].[TD030]
                                    ,[COPTD].[TD031],[COPTD].[TD032],[COPTD].[TD033],[COPTD].[TD034],[COPTD].[TD035],[COPTD].[TD036],[COPTD].[TD037],[COPTD].[TD038],[COPTD].[TD039],[COPTD].[TD040]
                                    ,[COPTD].[TD041],[COPTD].[TD042],[COPTD].[TD043],[COPTD].[TD044],[COPTD].[TD045],[COPTD].[TD046],[COPTD].[TD047],[COPTD].[TD048],[COPTD].[TD049],[COPTD].[TD050]
                                    ,[COPTD].[TD051],[COPTD].[TD052],[COPTD].[TD053],[COPTD].[TD054],[COPTD].[TD055],[COPTD].[TD056],[COPTD].[TD057],[COPTD].[TD058],[COPTD].[TD059],[COPTD].[TD060]
                                    ,[COPTD].[TD061],[COPTD].[TD062],[COPTD].[TD063],[COPTD].[TD064],[COPTD].[TD065],[COPTD].[TD066],[COPTD].[TD067],[COPTD].[TD068],[COPTD].[TD069],[COPTD].[TD070]
                                    ,[COPTD].[TD071],[COPTD].[TD072],[COPTD].[TD073],[COPTD].[TD074],[COPTD].[TD075],[COPTD].[TD076],[COPTD].[TD077],[COPTD].[TD078],[COPTD].[TD079],[COPTD].[TD080]
                                    ,[COPTD].[TD081],[COPTD].[TD082],[COPTD].[TD083],[COPTD].[TD084],[COPTD].[TD085],[COPTD].[TD086],[COPTD].[TD087],[COPTD].[TD088],[COPTD].[TD089],[COPTD].[TD090]
                                    ,[COPTD].[TD091],[COPTD].[TD092],[COPTD].[TD093],[COPTD].[TD094],[COPTD].[TD095],[COPTD].[TD096],[COPTD].[TD097],[COPTD].[TD098],[COPTD].[TD099],[COPTD].[TD100]
                                    ,[COPTD].[TD101],[COPTD].[TD102],[COPTD].[TD103],[COPTD].[TD104],[COPTD].[TD105],[COPTD].[TD106],[COPTD].[TD107],[COPTD].[TD108],[COPTD].[TD109],[COPTD].[TD110]
                                    ,[COPTD].[TD111],[COPTD].[TD112],[COPTD].[TD113]
                                    ,[COPTD].[UDF01] AS 'COPTDUDF01',[COPTD].[UDF02] AS 'COPTDUDF02',[COPTD].[UDF03] AS 'COPTDUDF03',[COPTD].[UDF04] AS 'COPTDUDF04',[COPTD].[UDF05] AS 'COPTDUDF05',[COPTD].[UDF06] AS 'COPTDUDF06',[COPTD].[UDF07] AS 'COPTDUDF07',[COPTD].[UDF08] AS 'COPTDUDF08',[COPTD].[UDF09] AS 'COPTDUDF09',[COPTD].[UDF10] AS 'COPTDUDF10',[COPTD].[TD200] AS 'TD200'
                                    ,[TB_EB_USER].USER_GUID,NAME
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=TC006) AS 'MV002'
                                    ,(SELECT TOP 1 MA002 FROM [TK].dbo.COPMA WHERE MA001=TC004) AS 'MA002'
                                    FROM [TK].dbo.COPTD,[TK].dbo.COPTC
                                    LEFT JOIN [192.168.1.223].[{0}].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= TC006 COLLATE Chinese_Taiwan_Stroke_BIN
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC001='{1}' AND TC002='{2}'
                                    ) AS TEMP   
                              
                                    ", DBNAME, TC001, TC002);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
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
            BASELIMITHRSBAR1 = SEARCHBASELIMITHRS("新廠製一組桶數");
            BASELIMITHRSBAR1 = Math.Round(BASELIMITHRSBAR1,0);
            BASELIMITHRSBAR2 = SEARCHBASELIMITHRS("新廠製二組桶數");
            BASELIMITHRSBAR2 = Math.Round(BASELIMITHRSBAR2, 0);

            BASELIMITHRS1 = SEARCHBASELIMITHRS("新廠製一組稼動率時數");
            BASELIMITHRS2 = SEARCHBASELIMITHRS("新廠製二組稼動率時數");
            BASELIMITHRS3 = SEARCHBASELIMITHRS("新廠製三組(手工)稼動率時數");
            BASELIMITHRS9 = SEARCHBASELIMITHRS("新廠包裝線稼動率時數");

            ADDTOUOFTB_EIP_SCH_MEMO_MOC(DateTime.Now.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDTOUOFTB_EIP_SCH_MEMO_PUR(DateTime.Now.ToString("yyyyMMdd"));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDTOUOFTB_EIP_SCH_MEMO_COP(DateTime.Now.ToString("yyyyMMdd"));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            UPDATEtb_COMPANYSTATUS1();
            UPDATEtb_COMPANYSTATUS2();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            ADDtb_COMPANY();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //UPDATEtb_COMPANYOWNER_ID();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            ADDTOUOFOURTAB();
            //ADDTB_WKF_EXTERNAL_TASK("A311", "20210415007");
        }
        private void button8_Click(object sender, EventArgs e)
        {
            ADDCOPTCCOPTD();
        }

        #endregion


    }
}
