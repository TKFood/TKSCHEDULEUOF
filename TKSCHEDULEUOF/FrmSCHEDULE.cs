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
using TKITDLL;


namespace TKSCHEDULEUOF
{
    public partial class FrmSCHEDULE : Form
    {
        //測試ID = "";
        //正式ID =""
        //測試DB DBNAME = "UOFTEST";
        //正式DB DBNAME = "UOF";
        string COPID = "3b2ca47e-3b9c-4633-855b-e5d3e58278ee";
        string COPCHANGEID = "8c637ad2-adcf-48ef-b665-1860eba83566";

        string PURID = "cbf3035c-2b72-4416-b4b3-534ea8763460";
        string DBNAME = "UOF";

        string OLDTASK_ID = null;
        string NEWTASK_ID = null;
        string ATTACH_ID = null;
        string COPTCUDF01 = "N";

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
        /// 製一線桶數 BASELIMITHRSBAR1
        /// 製二線桶數 BASELIMITHRSBAR2
        /// 包裝線稼動率時數 BASELIMITHRS9
        /// 製一線稼動率時數 BASELIMITHRS1
        /// 製二線稼動率時數 BASELIMITHRS2
        /// 手工線稼動率時數 BASELIMITHRS3
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

                BASELIMITHRSBAR1 = SEARCHBASELIMITHRS("製一線桶數");
                BASELIMITHRSBAR1 = Math.Round(BASELIMITHRSBAR1, 0);
                BASELIMITHRSBAR2 = SEARCHBASELIMITHRS("製二線桶數");
                BASELIMITHRSBAR2 = Math.Round(BASELIMITHRSBAR2, 0);

                BASELIMITHRS1 = SEARCHBASELIMITHRS("製一線稼動率時數");
                BASELIMITHRS2 = SEARCHBASELIMITHRS("製二線稼動率時數");
                BASELIMITHRS3 = SEARCHBASELIMITHRS("手工線稼動率時數");
                BASELIMITHRS9 = SEARCHBASELIMITHRS("包裝線稼動率時數");

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                //[CREATE_USER]='7774b96c-6762-45ef-b9d1-fcd718854e9f'，包裝線 MANU90
                //[CREATE_USER]='5ce0f554-8b80-4aed-afea-fcd224cecb81'，製一線 MANU10
                //[CREATE_USER]='0c98530a-b467-4cd4-a411-7279f1e04d0d'，製二線 MANU20
                //[CREATE_USER]='88789ece-41d1-4b48-94f1-6ffab05b05f4'，手工線 MANU30
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

        //製一線、製二線的桶數
        public DataSet SEARCHMANULINE(string Sday)
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]
                                    FROM (
                                    SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{1}桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{1}桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                    FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                    LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'
                                    WHERE INVMB.MB001=MOCMANULINE.MB001 
                                    AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}'
                                    AND [MOCMANULINE]. [MANU]='製一線'
                                    AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])      
                                    GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                    UNION
                                    SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{2}桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{2}桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                    FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                    LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'
                                    WHERE INVMB.MB001=MOCMANULINE.MB001   
                                    AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}'
                                    AND [MOCMANULINE]. [MANU]='製二線'
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
        //包裝線、製一線、製二線、手工線的總工時
        public DataSet SEARCHMANULINE2(string Sday)
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                                     AND [MOCMANULINE]. [MANU]='包裝線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]                
                                     UNION
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '  AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001 
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='製一線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                     UNION
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001   
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='製二線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                     UNION               
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)   AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001 
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='手工線'
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

        //包裝線、製一線、製二線、手工線的稼動率
        public DataSet SEARCHMANULINE3(string Sday)
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                                 AND [MOCMANULINE]. [MANU]='包裝線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                 UNION
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{1}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{1}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001 
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='製一線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE] 
                                 UNION
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{2}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{2}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001   
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='製二線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                 UNION                
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{3}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{3}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001 
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='手工線'
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

        //包裝線、製一線、製二線、手工線的明細
        public DataSet SEARCHMANULINE4(string Sday)
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(" SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='包裝線'");
                sbSql.AppendFormat(" UNION");                
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製一線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製二線'");               
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='手工線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],'手工線'+TA001+'-'+TA002+TA034+CONVERT(NVARCHAR,CONVERT(INT,TA015))+TA007 AS [DESCRIPTION],CONVERT(NVARCHAR,TA003,112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,TA003,112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],'手工線'+TA001+'-'+TA002+TA034+CONVERT(NVARCHAR,CONVERT(INT,TA015))+TA007 AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
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
                //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                    //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                    //sqlConn = new SqlConnection(connectionString);

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                    //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                    //sqlConn = new SqlConnection(connectionString);

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
            Form.SetAttribute("formVersionId", PURID);

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
            ////string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            //sqlConn = new SqlConnection(connectionString);

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp22"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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

            UPDATECOPTCUDF02();
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

            string BA = DT.Rows[0]["BA"].ToString();
            string BANAME = DT.Rows[0]["BANAME"].ToString();
            string BA_USER_GUID = DT.Rows[0]["BA_USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = DT.Rows[0]["TC001"].ToString().Trim() + DT.Rows[0]["TC002"].ToString().Trim();

            int rowscounts = 0;

            COPTCUDF01 = "N";

            foreach (DataRow od in DT.Rows)
            {
                if(od["COPTDUDF01"].ToString().Equals("Y"))
                {
                    COPTCUDF01 = "Y";
                    break;
                }
                else
                {
                    COPTCUDF01 = "N";
                }
            }

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
            //COPTCUDF01 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "COPTCUDF01");
            FieldItem.SetAttribute("fieldValue", COPTCUDF01);
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

            //建立userset
            var xElement = new XElement(
                  new XElement("UserSet",
                      new XElement("Element", new XAttribute("type", "user"),
                          new XElement("userId", fillerUserGuid)
                          )
                          )
                        );



            //XmlDocument doc = new XmlDocument();
            //XmlElement UserSet = doc.CreateElement("UserSet");

            //XmlElement Element = doc.CreateElement("Element");
            //Element.SetAttribute("type", "user");//設定屬性
            //UserSet.AppendChild(Element);

            //XmlElement userId = doc.CreateElement("userId", fillerUserGuid);
            //Element.AppendChild(userId);

            //建立節點FieldItem
            //TC006 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC006");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NAME"].ToString()+"("+DT.Rows[0]["TC006"].ToString()+")");
            FieldItem.SetAttribute("realValue", xElement.ToString());
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

            //建立userset
            var xElement_BA = new XElement(
                  new XElement("UserSet",
                      new XElement("Element", new XAttribute("type", "user"),
                          new XElement("userId", BA_USER_GUID)
                          )
                          )
                        );

            //建立節點FieldItem
            //BA 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "BA");
            FieldItem.SetAttribute("fieldValue",BANAME + "(" + BA + ")");
            FieldItem.SetAttribute("realValue", xElement_BA.ToString());
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //BANAME 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "BANAME");
            FieldItem.SetAttribute("fieldValue", BANAME);
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
                Cell.SetAttribute("fieldValue", od["COPTDUDF01"].ToString());
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
            //string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            //sqlConn = new SqlConnection(connectionString);

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

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
                //connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


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
                                    ,BA
                                    ,BANAME
                                    ,(SELECT TOP 1 [USER_GUID] FROM [192.168.1.223].[UOF].[dbo].[TB_EB_USER] WHERE [ACCOUNT]=BA COLLATE Chinese_Taiwan_Stroke_BIN) AS 'BA_USER_GUID'
    
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
                                    ,(SELECT TOP 1 COPMA.UDF04 FROM [TK].dbo.COPMA,[TK].dbo.CMSMV WHERE COPMA.UDF04=CMSMV.MV001 AND COPMA.MA001=TC004) AS 'BA'
                                    ,(SELECT TOP 1 CMSMV.MV002 FROM [TK].dbo.COPMA,[TK].dbo.CMSMV WHERE COPMA.UDF04=CMSMV.MV001 AND COPMA.MA001=TC004) AS 'BANAME'

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

        public void ADDCOPTECOPTF()
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                DataSet ds1 = new DataSet();
                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT TE001,TE002,TE003
                                    FROM [TK].dbo.COPTE
                                    WHERE TE029='N' AND UDF01='Y'
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
                        ADDTB_WKF_EXTERNAL_TASK_COPTECOPTF(dr["TE001"].ToString().Trim(), dr["TE002"].ToString().Trim(), dr["TE003"].ToString().Trim());
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

            UPDATECOPTEUDF01();
        }

        public void ADDTB_WKF_EXTERNAL_TASK_COPTECOPTF(string TE001, string TE002,string TE003)
        {

            DataTable DT = SEARCHCOPTECOPTF(TE001, TE002, TE003);
            DataTable DTUPFDEP = SEARCHUOFDEP(DT.Rows[0]["TE009"].ToString());

            string account = DT.Rows[0]["TE009"].ToString();
            string groupId = DT.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
            string fillerName = DT.Rows[0]["NAME"].ToString();
            string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();

            string BA = DT.Rows[0]["BA"].ToString();
            string BANAME = DT.Rows[0]["BANAME"].ToString();
            string BA_USER_GUID = DT.Rows[0]["BA_USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = DT.Rows[0]["TE001"].ToString().Trim() + DT.Rows[0]["TE002"].ToString().Trim() + DT.Rows[0]["TE003"].ToString().Trim();

            int rowscounts = 0;

            COPTCUDF01 = "N";

            foreach (DataRow od in DT.Rows)
            {
                if (od["COPTDUDF01"].ToString().Equals("Y"))
                {
                    COPTCUDF01 = "Y";
                    break;
                }
                else
                {
                    COPTCUDF01 = "N";
                }
            }

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //正式的id
            Form.SetAttribute("formVersionId", COPCHANGEID);

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
            //TE006 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE006");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE006"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE001 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE001");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE001"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE002 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE003 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE003"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //COPTCUDF01 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "COPTCUDF01");
            FieldItem.SetAttribute("fieldValue", COPTCUDF01);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //10
            //建立節點FieldItem
            //TE038 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE038");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE038"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE007 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE007");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE007"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //MA002 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "MA002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["MA002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE011 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE011");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE011"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE111 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE111");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE111"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE012 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE012");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE012"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE112 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE112");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE112"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE041 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE041");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE041"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE041C 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE041C");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE017"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem
            //TE137 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE137");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE137"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //建立節點FieldItem

            //20
            //TE137C 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE137C");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE117"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE015 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE015");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE015"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE115 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE115");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE115"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE018 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE018");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NEWTE018"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE118 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE118");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NEWTE118"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE008 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE008");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE008"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //CMSME002A 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "CMSME002A");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["CMSME002A"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE108 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE108");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE108"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //CMSME002B 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "CMSME002B");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["CMSME002B"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立userset
            var xElement = new XElement(
                  new XElement("UserSet",
                      new XElement("Element", new XAttribute("type", "user"),
                          new XElement("userId", BA_USER_GUID)
                          )
                          )
                        );

            //TE009 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE009");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["CMSMV002A"].ToString() +"("+ DT.Rows[0]["TE009"].ToString()+")");
            FieldItem.SetAttribute("realValue", xElement.ToString());
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //30
            //CMSMV002A 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "CMSMV002A");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["CMSMV002A"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE109 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE109");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE109"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //CMSMV002B 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "CMSMV002B");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["CMSMV002B"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立userset
            var xElement_BA = new XElement(
                  new XElement("UserSet",
                      new XElement("Element", new XAttribute("type", "user"),
                          new XElement("userId", BA_USER_GUID)
                          )
                          )
                        );

            //建立節點FieldItem
            //BA 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "BA");
            FieldItem.SetAttribute("fieldValue", BANAME + "(" + BA + ")");
            FieldItem.SetAttribute("realValue", xElement_BA.ToString());
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //BANAME 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "BANAME");
            FieldItem.SetAttribute("fieldValue", BANAME);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //TE040 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE040");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE040"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE136 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE136");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE136"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE013 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE013");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE013"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE113 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE113");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE113"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);
            //TE050 	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE050");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE050"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //UDF05
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "UDF05");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["COPTCUDF05"].ToString());
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

         
                //Row	TF004
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF004");
                Cell.SetAttribute("fieldValue", od["TF004"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //10
                //Row	TF005
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF005");
                Cell.SetAttribute("fieldValue", od["TF005"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF006
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF006");
                Cell.SetAttribute("fieldValue", od["TF006"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF007
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF007");
                Cell.SetAttribute("fieldValue", od["TF007"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF009
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF009");
                Cell.SetAttribute("fieldValue", od["TF009"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF020
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF020");
                Cell.SetAttribute("fieldValue", od["TF020"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF010
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF010");
                Cell.SetAttribute("fieldValue", od["TF010"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF013
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF013");
                Cell.SetAttribute("fieldValue", od["TF013"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF021
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF021");
                Cell.SetAttribute("fieldValue", od["TF021"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF014
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF014");
                Cell.SetAttribute("fieldValue", od["TF014"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF015
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF015");
                Cell.SetAttribute("fieldValue", od["TF015"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //20
                //Row	TF024
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF024");
                Cell.SetAttribute("fieldValue", od["TF024"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF025
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF025");
                Cell.SetAttribute("fieldValue", od["TF025"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF018
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF018");
                Cell.SetAttribute("fieldValue", od["TF018"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF105
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF105");
                Cell.SetAttribute("fieldValue", od["TF105"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF106
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF106");
                Cell.SetAttribute("fieldValue", od["TF106"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF107
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF107");
                Cell.SetAttribute("fieldValue", od["TF107"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF109
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF109");
                Cell.SetAttribute("fieldValue", od["TF109"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF120
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF120");
                Cell.SetAttribute("fieldValue", od["TF120"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF110
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF110");
                Cell.SetAttribute("fieldValue", od["TF110"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF015
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF113");
                Cell.SetAttribute("fieldValue", od["TF113"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //30
                //Row	TF121
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF121");
                Cell.SetAttribute("fieldValue", od["TF121"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF114
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF114");
                Cell.SetAttribute("fieldValue", od["TF114"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF115
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF115");
                Cell.SetAttribute("fieldValue", od["TF115"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF126
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF126");
                Cell.SetAttribute("fieldValue", od["TF126"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);
                //Row	TF127
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF127");
                Cell.SetAttribute("fieldValue", od["TF127"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                Row.AppendChild(Cell);





                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']/DataGrid");
                DataGridS.AppendChild(Row);

            }

            ////用ADDTACK，直接啟動起單
            //ADDTACK(Form);

            //ADD TO DB
            //string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            //sqlConn = new SqlConnection(connectionString);

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

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


        public void UPDATECOPTEUDF01()
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"
                                    UPDATE  [TK].dbo.COPTE 
                                    SET UDF01 = 'UOF'
                                    WHERE TE029='N' AND UDF01='Y'                                             

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
        public DataTable SEARCHCOPTECOPTF(string TE001, string TE002,string TE003)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //庫存數量看LA009 IN ('20004','20006','20008','20019','20020'

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME]
                                    ,[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]
                                    ,[TE001],[TE002],[TE003],[TE004],[TE005],[TE006],[TE007],[TE008],[TE009],[TE010]
                                    ,[TE011],[TE012],[TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020]
                                    ,[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030]
                                    ,[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040]
                                    ,[TE041],[TE042],[TE043],[TE044],[TE045],[TE046],[TE047],[TE048],[TE049],[TE050]
                                    ,[TE051],[TE052],[TE053],[TE054],[TE055],[TE056],[TE057],[TE058],[TE059],[TE060]
                                    ,[TE061],[TE062],[TE063],[TE064],[TE065],[TE066],[TE067],[TE068],[TE069],[TE070]
                                    ,[TE071],[TE072],[TE073],[TE074],[TE075],[TE076],[TE077],[TE078],[TE079],[TE080]
                                    ,[TE081],[TE082],[TE083],[TE084],[TE085],[TE086],[TE087],[TE088]
                                    ,[TE103],[TE107],[TE108],[TE109],[TE110]
                                    ,[TE111],[TE112],[TE113],[TE114],[TE115],[TE116],[TE117],[TE118],[TE119],[TE120]
                                    ,[TE121],[TE122],[TE123],[TE124],[TE125],[TE126],[TE127],[TE128],[TE129],[TE130]
                                    ,[TE131],[TE132],[TE133],[TE134],[TE135],[TE136],[TE137],[TE138],[TE139],[TE140]
                                    ,[TE141],[TE142],[TE143],[TE144],[TE145],[TE146],[TE147],[TE148],[TE149],[TE150]
                                    ,[TE151],[TE152],[TE163],[TE164],[TE165],[TE166],[TE167],[TE168],[TE169],[TE170]
                                    ,[TE171],[TE172],[TE173],[TE174],[TE175],[TE176],[TE177],[TE178],[TE179],[TE180]
                                    ,[TE181],[TE182],[TE183],[TE184],[TE185],[TE186],[TE187],[TE188],[TE189],[TE190]
                                    ,[TE191],[TE192],[TE193],[TE194],[TE195],[TE196],[TE197],[TE198],[TE199]
                                    ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]

                                    ,[TF001],[TF002],[TF003],[TF004],[TF005],[TF006],[TF007],[TF008],[TF009],[TF010]
                                    ,[TF011],[TF012],[TF013],[TF014],[TF015],[TF016],[TF017],[TF018],[TF019],[TF020]
                                    ,[TF021],[TF022],[TF023],[TF024],[TF025],[TF026],[TF027],[TF028],[TF029],[TF030]
                                    ,[TF031],[TF032],[TF034],[TF035],[TF036],[TF037],[TF038],[TF039],[TF040],[TF041]
                                    ,[TF042],[TF043],[TF044],[TF045],[TF046],[TF048],[TF049],[TF050]
                                    ,[TF051],[TF052],[TF053],[TF054],[TF055],[TF056],[TF057],[TF058],[TF059],[TF060]
                                    ,[TF061],[TF062],[TF063],[TF064],[TF065],[TF066],[TF067],[TF068],[TF069],[TF070]
                                    ,[TF071],[TF072],[TF073],[TF074],[TF075],[TF076],[TF077],[TF078],[TF079],[TF080]
                                    ,[TF104],[TF105],[TF106],[TF107],[TF108],[TF109],[TF110]
                                    ,[TF111],[TF112],[TF113],[TF114],[TF115],[TF116],[TF117],[TF120]
                                    ,[TF121],[TF122],[TF123],[TF124],[TF125],[TF126],[TF127],[TF128],[TF129],[TF130]
                                    ,[TF131],[TF132],[TF133],[TF134],[TF135],[TF136],[TF137],[TF138],[TF139],[TF140]
                                    ,[TF141],[TF142],[TF143],[TF144],[TF145],[TF146],[TF147],[TF148],[TF149],[TF150]
                                    ,[TF151],[TF152],[TF153],[TF154],[TF155],[TF156],[TF157],[TF158],[TF159],[TF160]
                                    ,[TF161],[TF162],[TF163],[TF164],[TF165],[TF166],[TF167],[TF168],[TF169],[TF170]
                                    ,[TF171],[TF172],[TF173],[TF174],[TF175],[TF176],[TF177],[TF178],[TF179],[TF180]
                                    ,[TF181],[TF182],[TF183],[TF184],[TF185],[TF186],[TF187],[TF188],[TF189],[TF190]
                                    ,[TF191],[TF192],[TF193],[TF194],[TF195],[TF196],[TF197],[TF198],[TF199]
                                    ,[TF200],[TF300]

                                    ,COPTFUDF01,COPTFUDF02,COPTFUDF03,COPTFUDF04,COPTFUDF05,COPTFUDF06,COPTFUDF07,COPTFUDF08,COPTFUDF09,COPTFUDF10

                                    ,USER_GUID,NAME
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    ,MA002
                                    ,CASE WHEN TE018='1' THEN '1.應稅內含'  ELSE (CASE WHEN TE018='2' THEN '2.應稅外加'  ELSE (CASE WHEN TE018='3' THEN '3.零稅率'  ELSE (CASE WHEN TE018='4' THEN '4.免稅'  ELSE (CASE WHEN TE018='9' THEN '9.不計稅'  ELSE '' END) END) END) END ) END AS 'NEWTE018'
                                    ,CASE WHEN TE118='1' THEN '1.應稅內含'  ELSE (CASE WHEN TE118='2' THEN '2.應稅外加'  ELSE (CASE WHEN TE118='3' THEN '3.零稅率'  ELSE (CASE WHEN TE118='4' THEN '4.免稅'  ELSE (CASE WHEN TE118='9' THEN '9.不計稅'  ELSE '' END) END) END) END ) END AS 'NEWTE118'
                                    ,(SELECT TOP 1 ME002 FROM [TK].dbo.CMSME WHERE ME001=TE008) AS 'CMSME002A'
                                    ,(SELECT TOP 1 ME002 FROM [TK].dbo.CMSME WHERE ME001=TE108) AS 'CMSME002B'
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=TE009) AS 'CMSMV002A'
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=TE109) AS 'CMSMV002B'
                                    ,(SELECT TOP 1 COPTC.UDF05 FROM [TK].dbo.COPTC WHERE TC001=TE001 AND TC002=TE002) AS 'COPTCUDF05'
                                    ,ISNULL((SELECT TOP 1 COPTD.UDF01 FROM [TK].dbo.COPTD WHERE TD001=TE001 AND TD002=TE002 AND COPTD.UDF01='Y'),'N') AS 'COPTDUDF01'
                                    ,BA
                                    ,BANAME
                                    ,(SELECT TOP 1 [USER_GUID] FROM [192.168.1.223].[UOF].[dbo].[TB_EB_USER] WHERE [ACCOUNT]=BA COLLATE Chinese_Taiwan_Stroke_BIN) AS 'BA_USER_GUID'

                                    FROM 
                                    (

                                    SELECT 
                                    [COPTE].[COMPANY],[COPTE].[CREATOR],[COPTE].[USR_GROUP],[COPTE].[CREATE_DATE],[COPTE].[MODIFIER],[COPTE].[MODI_DATE],[COPTE].[FLAG],[COPTE].[CREATE_TIME],[COPTE].[MODI_TIME],[COPTE].[TRANS_TYPE],[COPTE].[TRANS_NAME]
                                    ,[COPTE].[sync_date],[COPTE].[sync_time],[COPTE].[sync_mark],[COPTE].[sync_count],[COPTE].[DataUser],[COPTE].[DataGroup]
                                    ,[COPTE].[TE001],[COPTE].[TE002],[COPTE].[TE003],[COPTE].[TE004],[COPTE].[TE005],[COPTE].[TE006],[COPTE].[TE007],[COPTE].[TE008],[COPTE].[TE009],[COPTE].[TE010]
                                    ,[COPTE].[TE011],[COPTE].[TE012],[COPTE].[TE013],[COPTE].[TE014],[COPTE].[TE015],[COPTE].[TE016],[COPTE].[TE017],[COPTE].[TE018],[COPTE].[TE019],[COPTE].[TE020]
                                    ,[COPTE].[TE021],[COPTE].[TE022],[COPTE].[TE023],[COPTE].[TE024],[COPTE].[TE025],[COPTE].[TE026],[COPTE].[TE027],[COPTE].[TE028],[COPTE].[TE029],[COPTE].[TE030]
                                    ,[COPTE].[TE031],[COPTE].[TE032],[COPTE].[TE033],[COPTE].[TE034],[COPTE].[TE035],[COPTE].[TE036],[COPTE].[TE037],[COPTE].[TE038],[COPTE].[TE039],[COPTE].[TE040]
                                    ,[COPTE].[TE041],[COPTE].[TE042],[COPTE].[TE043],[COPTE].[TE044],[COPTE].[TE045],[COPTE].[TE046],[COPTE].[TE047],[COPTE].[TE048],[COPTE].[TE049],[COPTE].[TE050]
                                    ,[COPTE].[TE051],[COPTE].[TE052],[COPTE].[TE053],[COPTE].[TE054],[COPTE].[TE055],[COPTE].[TE056],[COPTE].[TE057],[COPTE].[TE058],[COPTE].[TE059],[COPTE].[TE060]
                                    ,[COPTE].[TE061],[COPTE].[TE062],[COPTE].[TE063],[COPTE].[TE064],[COPTE].[TE065],[COPTE].[TE066],[COPTE].[TE067],[COPTE].[TE068],[COPTE].[TE069],[COPTE].[TE070]
                                    ,[COPTE].[TE071],[COPTE].[TE072],[COPTE].[TE073],[COPTE].[TE074],[COPTE].[TE075],[COPTE].[TE076],[COPTE].[TE077],[COPTE].[TE078],[COPTE].[TE079],[COPTE].[TE080]
                                    ,[COPTE].[TE081],[COPTE].[TE082],[COPTE].[TE083],[COPTE].[TE084],[COPTE].[TE085],[COPTE].[TE086],[COPTE].[TE087],[COPTE].[TE088]
                                    ,[COPTE].[TE103],[COPTE].[TE107],[COPTE].[TE108],[COPTE].[TE109],[COPTE].[TE110]
                                    ,[COPTE].[TE111],[COPTE].[TE112],[COPTE].[TE113],[COPTE].[TE114],[COPTE].[TE115],[COPTE].[TE116],[COPTE].[TE117],[COPTE].[TE118],[COPTE].[TE119],[COPTE].[TE120]
                                    ,[COPTE].[TE121],[COPTE].[TE122],[COPTE].[TE123],[COPTE].[TE124],[COPTE].[TE125],[COPTE].[TE126],[COPTE].[TE127],[COPTE].[TE128],[COPTE].[TE129],[COPTE].[TE130]
                                    ,[COPTE].[TE131],[COPTE].[TE132],[COPTE].[TE133],[COPTE].[TE134],[COPTE].[TE135],[COPTE].[TE136],[COPTE].[TE137],[COPTE].[TE138],[COPTE].[TE139],[COPTE].[TE140]
                                    ,[COPTE].[TE141],[COPTE].[TE142],[COPTE].[TE143],[COPTE].[TE144],[COPTE].[TE145],[COPTE].[TE146],[COPTE].[TE147],[COPTE].[TE148],[COPTE].[TE149],[COPTE].[TE150]
                                    ,[COPTE].[TE151],[COPTE].[TE152],[COPTE].[TE163],[COPTE].[TE164],[COPTE].[TE165],[COPTE].[TE166],[COPTE].[TE167],[COPTE].[TE168],[COPTE].[TE169],[COPTE].[TE170]
                                    ,[COPTE].[TE171],[COPTE].[TE172],[COPTE].[TE173],[COPTE].[TE174],[COPTE].[TE175],[COPTE].[TE176],[COPTE].[TE177],[COPTE].[TE178],[COPTE].[TE179],[COPTE].[TE180]
                                    ,[COPTE].[TE181],[COPTE].[TE182],[COPTE].[TE183],[COPTE].[TE184],[COPTE].[TE185],[COPTE].[TE186],[COPTE].[TE187],[COPTE].[TE188],[COPTE].[TE189],[COPTE].[TE190]
                                    ,[COPTE].[TE191],[COPTE].[TE192],[COPTE].[TE193],[COPTE].[TE194],[COPTE].[TE195],[COPTE].[TE196],[COPTE].[TE197],[COPTE].[TE198],[COPTE].[TE199]
                                    ,[COPTE].[UDF01],[COPTE].[UDF02],[COPTE].[UDF03],[COPTE].[UDF04],[COPTE].[UDF05],[COPTE].[UDF06],[COPTE].[UDF07],[COPTE].[UDF08],[COPTE].[UDF09],[COPTE].[UDF10]

                                    ,[COPTF].[TF001],[COPTF].[TF002],[COPTF].[TF003],[COPTF].[TF004],[COPTF].[TF005],[COPTF].[TF006],[COPTF].[TF007],[COPTF].[TF008],[COPTF].[TF009],[COPTF].[TF010]
                                    ,[COPTF].[TF011],[COPTF].[TF012],[COPTF].[TF013],[COPTF].[TF014],[COPTF].[TF015],[COPTF].[TF016],[COPTF].[TF017],[COPTF].[TF018],[COPTF].[TF019],[COPTF].[TF020]
                                    ,[COPTF].[TF021],[COPTF].[TF022],[COPTF].[TF023],[COPTF].[TF024],[COPTF].[TF025],[COPTF].[TF026],[COPTF].[TF027],[COPTF].[TF028],[COPTF].[TF029],[COPTF].[TF030]
                                    ,[COPTF].[TF031],[COPTF].[TF032],[COPTF].[TF034],[COPTF].[TF035],[COPTF].[TF036],[COPTF].[TF037],[COPTF].[TF038],[COPTF].[TF039],[COPTF].[TF040],[COPTF].[TF041]
                                    ,[COPTF].[TF042],[COPTF].[TF043],[COPTF].[TF044],[COPTF].[TF045],[COPTF].[TF046],[COPTF].[TF048],[COPTF].[TF049],[COPTF].[TF050]
                                    ,[COPTF].[TF051],[COPTF].[TF052],[COPTF].[TF053],[COPTF].[TF054],[COPTF].[TF055],[COPTF].[TF056],[COPTF].[TF057],[COPTF].[TF058],[COPTF].[TF059],[COPTF].[TF060]
                                    ,[COPTF].[TF061],[COPTF].[TF062],[COPTF].[TF063],[COPTF].[TF064],[COPTF].[TF065],[COPTF].[TF066],[COPTF].[TF067],[COPTF].[TF068],[COPTF].[TF069],[COPTF].[TF070]
                                    ,[COPTF].[TF071],[COPTF].[TF072],[COPTF].[TF073],[COPTF].[TF074],[COPTF].[TF075],[COPTF].[TF076],[COPTF].[TF077],[COPTF].[TF078],[COPTF].[TF079],[COPTF].[TF080]
                                    ,[COPTF].[TF104],[COPTF].[TF105],[COPTF].[TF106],[COPTF].[TF107],[COPTF].[TF108],[COPTF].[TF109],[COPTF].[TF110]
                                    ,[COPTF].[TF111],[COPTF].[TF112],[COPTF].[TF113],[COPTF].[TF114],[COPTF].[TF115],[COPTF].[TF116],[COPTF].[TF117],[COPTF].[TF120]
                                    ,[COPTF].[TF121],[COPTF].[TF122],[COPTF].[TF123],[COPTF].[TF124],[COPTF].[TF125],[COPTF].[TF126],[COPTF].[TF127],[COPTF].[TF128],[COPTF].[TF129],[COPTF].[TF130]
                                    ,[COPTF].[TF131],[COPTF].[TF132],[COPTF].[TF133],[COPTF].[TF134],[COPTF].[TF135],[COPTF].[TF136],[COPTF].[TF137],[COPTF].[TF138],[COPTF].[TF139],[COPTF].[TF140]
                                    ,[COPTF].[TF141],[COPTF].[TF142],[COPTF].[TF143],[COPTF].[TF144],[COPTF].[TF145],[COPTF].[TF146],[COPTF].[TF147],[COPTF].[TF148],[COPTF].[TF149],[COPTF].[TF150]
                                    ,[COPTF].[TF151],[COPTF].[TF152],[COPTF].[TF153],[COPTF].[TF154],[COPTF].[TF155],[COPTF].[TF156],[COPTF].[TF157],[COPTF].[TF158],[COPTF].[TF159],[COPTF].[TF160]
                                    ,[COPTF].[TF161],[COPTF].[TF162],[COPTF].[TF163],[COPTF].[TF164],[COPTF].[TF165],[COPTF].[TF166],[COPTF].[TF167],[COPTF].[TF168],[COPTF].[TF169],[COPTF].[TF170]
                                    ,[COPTF].[TF171],[COPTF].[TF172],[COPTF].[TF173],[COPTF].[TF174],[COPTF].[TF175],[COPTF].[TF176],[COPTF].[TF177],[COPTF].[TF178],[COPTF].[TF179],[COPTF].[TF180]
                                    ,[COPTF].[TF181],[COPTF].[TF182],[COPTF].[TF183],[COPTF].[TF184],[COPTF].[TF185],[COPTF].[TF186],[COPTF].[TF187],[COPTF].[TF188],[COPTF].[TF189],[COPTF].[TF190]
                                    ,[COPTF].[TF191],[COPTF].[TF192],[COPTF].[TF193],[COPTF].[TF194],[COPTF].[TF195],[COPTF].[TF196],[COPTF].[TF197],[COPTF].[TF198],[COPTF].[TF199]
                                    ,[COPTF].[TF200],[COPTF].[TF300]

                                    ,[COPTF].[UDF01] AS 'COPTFUDF01',[COPTF].[UDF02] AS 'COPTFUDF02',[COPTF].[UDF03] AS 'COPTFUDF03',[COPTF].[UDF04] AS 'COPTFUDF04',[COPTF].[UDF05] AS 'COPTFUDF05',[COPTF].[UDF06] AS 'COPTFUDF06',[COPTF].[UDF07] AS 'COPTFUDF07',[COPTF].[UDF08] AS 'COPTFUDF08',[COPTF].[UDF09] AS 'COPTFUDF09',[COPTF].[UDF10] AS 'COPTFUDF10'
                                    ,[TB_EB_USER].USER_GUID,NAME
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=TE009) AS 'MV002'
                                    ,(SELECT TOP 1 MA002 FROM [TK].dbo.COPMA WHERE MA001=TE007) AS 'MA002'
                                    ,(SELECT TOP 1 COPMA.UDF04 FROM [TK].dbo.COPMA,[TK].dbo.CMSMV WHERE COPMA.UDF04=CMSMV.MV001 AND COPMA.MA001=TE007) AS 'BA'
                                    ,(SELECT TOP 1 CMSMV.MV002 FROM [TK].dbo.COPMA,[TK].dbo.CMSMV WHERE COPMA.UDF04=CMSMV.MV001 AND COPMA.MA001=TE007) AS 'BANAME'

                                    FROM [TK].dbo.COPTF,[TK].dbo.COPTE
                                    LEFT JOIN [192.168.1.223].[{0}].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= TE009 COLLATE Chinese_Taiwan_Stroke_BIN
                                    WHERE TE001=TF001 AND TE002=TF002 AND TE003=TF003
                                    AND TE001='{1}' AND TE002='{2}' AND TE003='{3}'
                                    ) AS TEMP   
                              
                                    ", DBNAME, TE001, TE002,TE003);


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
            BASELIMITHRSBAR1 = SEARCHBASELIMITHRS("製一線桶數");
            BASELIMITHRSBAR1 = Math.Round(BASELIMITHRSBAR1,0);
            BASELIMITHRSBAR2 = SEARCHBASELIMITHRS("製二線桶數");
            BASELIMITHRSBAR2 = Math.Round(BASELIMITHRSBAR2, 0);

            BASELIMITHRS1 = SEARCHBASELIMITHRS("製一線稼動率時數");
            BASELIMITHRS2 = SEARCHBASELIMITHRS("製二線稼動率時數");
            BASELIMITHRS3 = SEARCHBASELIMITHRS("手工線稼動率時數");
            BASELIMITHRS9 = SEARCHBASELIMITHRS("包裝線稼動率時數");

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
        private void button9_Click(object sender, EventArgs e)
        {
            ADDCOPTECOPTF();
        }

        #endregion


    }
}
