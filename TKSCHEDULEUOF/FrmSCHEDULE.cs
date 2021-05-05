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

            string account = DT.Rows[0]["CREATOR"].ToString();
            string groupId = DT.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
            string fillerName= DT.Rows[0]["NAME"].ToString();
            string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();
            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");
            //測試的id
            Form.SetAttribute("formVersionId", "dd7e61e3-d77f-40e4-8408-007d2b5fb92e");
            
            //正式的id
            //Form.SetAttribute("formVersionId", "1cc71c35-0a2c-490c-b733-f887b7975b17");

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

                //Row	TB011
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB011");
                Cell.SetAttribute("fieldValue", od["TB011"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='TB']/DataGrid");
                DataGridS.AppendChild(Row);

            }


            //ADD TO DB
            string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            StringBuilder queryString = new StringBuilder();

            ////UOFTEST
            ///
            queryString.AppendFormat(@" INSERT INTO [UOFTEST].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
                                        VALUES (NEWID(),@XML,2)
                                        ");

            //UOF
            //
            //queryString.AppendFormat(@" INSERT INTO [UOF].dbo.TB_WKF_EXTERNAL_TASK
            //                             (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
            //                            VALUES (NEWID(),@XML,2)
            //                            ");

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

                sbSql.AppendFormat(@"  
                                    SELECT PURTA.CREATOR,TA001,TA002,TA003,TA012,TB004,TB005,TB006,TB007,TB009,TB011
                                    ,[TB_EB_USER].USER_GUID,GROUP_ID,TITLE_ID,NAME
                                    FROM [TK].dbo.PURTA,[TK].dbo.PURTB
                                    LEFT JOIN [192.168.1.223].[UOFTEST].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= CREATOR COLLATE Chinese_Taiwan_Stroke_BIN
                                    LEFT JOIN [192.168.1.223].[UOFTEST].[dbo].[TB_EB_EMPL_DEP] ON [TB_EB_EMPL_DEP].USER_GUID=[TB_EB_USER].USER_GUID
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA001='{0}' AND TA002='{1}'
                                    ", TA001, TA002);


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
                    ADDTB_WKF_EXTERNAL_TASK(ds1.Tables["ds1"].Rows[0]["TA001"].ToString().Trim(), ds1.Tables["ds1"].Rows[0]["TA002"].ToString().Trim());

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
                                    SET UDF01 = 'UOF'
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

        #endregion

       
    }
}
