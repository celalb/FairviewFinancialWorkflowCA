using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared
{
    public class DbUtility
    {
        int ErrCode;
        string ErrMsg;
        public SqlConnection cn;

        SqlTransaction transaction;
        SqlConnection DBConnection
        {
            get
            {

                if (cn != null && cn.State == ConnectionState.Open)
                    return cn;
                else
                {
                    try
                    {
                        cn = new SqlConnection(ProgData.sqlConnectionString);
                        cn.Open();
                        transaction = cn.BeginTransaction();

                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }
                }

                return cn;

            }

        }
        public void Close(bool h)
        {
            if (h)
                transaction.Commit();
            else
                transaction.Rollback();
            cn.Close();

        }
        public void CreateTables()
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append(" CREATE TABLE[dbo].[redi_ConsBP_Header]([DocEntry][int] NOT NULL IDENTITY PRIMARY KEY,[CardCode] [nvarchar](15) NULL,");
                qry.Append("[DocDate] [datetime] NULL, [DocDueDate] [datetime] NULL,[DocRef] [nvarchar](32) NULL,[UserSign] [smallint] NULL,[CreateDate] [datetime] NULL) ON[PRIMARY]");
                SqlCommand command = new SqlCommand(qry.ToString(), DBConnection);
                command.ExecuteNonQuery();
                command.Dispose();
            }
            catch (Exception er)
            {

            }
            qry.Clear();
            try
            {
                qry.Append(" CREATE TABLE[dbo].[redi_ConsBP_Trans]([DocEntry][int] NULL,[DocNum] [int] NULL,[CardCode] [nvarchar](15) NULL,[CardName] [nvarchar](100) NULL,");
                qry.Append("[DocType] [varchar](2) NOT NULL,[DocDate] [datetime] NULL,[DocDueDate] [datetime] NULL,[DaysPastDue] [varchar](30) NULL,");
                qry.Append("[DocTotal] [Double](19, 6) NULL,[BalDue] [Double](19, 6) NULL,[TotalDisc] [Double](19, 6) NULL,[TotalPay] [Double](19, 6) NULL,");
                qry.Append("[PayOnAccount] [Double](19, 6) NULL,[Selected] [char](1) NULL) ON[PRIMARY]");

                SqlCommand command = new SqlCommand(qry.ToString(), DBConnection);
                command.ExecuteNonQuery();
                command.Dispose();
            }
            catch (Exception er)
            {

            }
            cn.Close();

        }
        public void UpdatePayAcct(int docEntry, int docNum, int Linenum, string cardCode, string account, string docType, DateTime postDate,
                           string LineMemo, Double Debit, Double Credit, int user, string posted, int SumLinenum)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append(" UPDATE  [dbo].[redi_ConsBP_PayAcct]  SET PostDate=@PostDate,CardCode=@CardCode ");
                qry.Append(",Account=@Account,DocType=@DocType,DocNum=@DocNum,LineMemo=@LineMemo,Debit=@Debit,Credit=@Credit,UserSign=@UserSign,UpdateDate=@UpdateDate ");
                qry.Append(",Posted=@Posted WHERE DocEntry=@DocEntry AND LineNum=@LineNum AND SumLinenum = @SumLinenum");
                using (SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {
                    SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                    SqlCmd.Parameters.Add("@DocNum", SqlDbType.Int).Value = docNum;
                    SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                    SqlCmd.Parameters.Add("@LineNum", SqlDbType.Int).Value = Linenum;
                    SqlCmd.Parameters.Add("@SumLinenum", SqlDbType.Int).Value = SumLinenum;
                    SqlCmd.Parameters.Add("@Doctype", SqlDbType.VarChar, 2).Value = docType;
                    SqlCmd.Parameters.Add("@Account", SqlDbType.VarChar, 15).Value = account;
                    SqlCmd.Parameters.Add("@PostDate", SqlDbType.DateTime).Value = postDate;
                    SqlCmd.Parameters.Add("@Posted", SqlDbType.VarChar, 1).Value = posted;
                    SqlCmd.Parameters.Add("@LineMemo", SqlDbType.VarChar, 15).Value = LineMemo;
                    SqlCmd.Parameters.Add("@Debit", SqlDbType.Decimal, 19);
                    SqlCmd.Parameters["@Debit"].Precision = 19;
                    SqlCmd.Parameters["@Debit"].Scale = 6;
                    SqlCmd.Parameters["@Debit"].Value = Convert.ToDecimal(Debit);
                    SqlCmd.Parameters.Add("@Credit", SqlDbType.Decimal, 19);
                    SqlCmd.Parameters["@Credit"].Precision = 19;
                    SqlCmd.Parameters["@Credit"].Scale = 6;
                    SqlCmd.Parameters["@Credit"].Value = Convert.ToDecimal(Credit);


                    SqlCmd.Parameters.Add("@userSign", SqlDbType.Int).Value = user;

                    SqlCmd.Parameters.Add("@UpdateDate", SqlDbType.DateTime).Value = DateTime.Now;


                    SqlCmd.ExecuteNonQuery();
                    SqlCmd.Dispose();

                }
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("UpdatePayAcct = " + ex.Message + qry.ToString());

            }

        }

        public void InsertPayAcct(int docEntry, int docNum, int Linenum, string cardCode, string account, string docType, DateTime postDate,
                           string LineMemo, Double Debit, Double Credit, int user, string posted, int SumLinenum)
        {
            string qry = "";
            if (posted == null)
                posted = "";
            try
            {
                bool exist = false;
                qry = $"select DocEntry FROM [redi_ConsBP_PayAcct] where DocEntry = {docEntry.ToString()} AND LineNum = {Linenum.ToString()} AND SumLinenum = {SumLinenum.ToString()}";
                SqlCommand SqlCmd = null;
                using (SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {
                    /* SqlCmd.Parameters.Add("@LineNum", SqlDbType.Int).Value = Linenum;
                     SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                     SqlCmd.Parameters.Add("@SumLinenum", SqlDbType.Int).Value = SumLinenum;
                     SqlCmd.Prepare();*/

                    using (SqlDataReader reader = SqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (reader.HasRows)
                            {
                                exist = true;


                            }
                        }
                    }
                }
                if (exist)
                {
                    this.UpdatePayAcct(docEntry, docNum, Linenum, cardCode, account, docType, postDate,
                               LineMemo, Debit, Credit, user, posted, SumLinenum);
                    return;
                }

                qry = " Insert INTO [dbo].[redi_ConsBP_PayAcct] ([DocEntry] ,[CardCode],[PostDate],[LineNum],[Account],[DocType]";
                qry += ",[DocNum],[LineMemo],[Debit],[Credit],[Posted] ,[UserSign],[CreateDate] ,[UpdateDate],[SumLinenum])";
                qry += " Values";
                qry += "(@DocEntry,@CardCode,@PostDate,@LineNum,@Account,@DocType";
                qry += ",@DocNum,@LineMemo,@Debit ,@Credit,@Posted,@UserSign,@CreateDate,@UpdateDate,@SumLinenum)";
                SqlCmd = new SqlCommand(qry, DBConnection, transaction);
                SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                SqlCmd.Parameters.Add("@PostDate", SqlDbType.DateTime).Value = postDate;
                SqlCmd.Parameters.Add("@LineNum", SqlDbType.Int).Value = Linenum;
                SqlCmd.Parameters.Add("@Account", SqlDbType.VarChar, 15).Value = account;
                SqlCmd.Parameters.Add("@SumLinenum", SqlDbType.Int).Value = SumLinenum;
                SqlCmd.Parameters.Add("@Doctype", SqlDbType.VarChar, 2).Value = docType;
                SqlCmd.Parameters.Add("@DocNum", SqlDbType.Int).Value = docNum;
                SqlCmd.Parameters.Add("@LineMemo", SqlDbType.VarChar, 15).Value = LineMemo;
                SqlParameter parameter = new SqlParameter("@Debit", SqlDbType.Decimal);
                parameter.Precision = 19;
                parameter.Scale = 6;
                parameter.Value = Convert.ToDecimal(Debit);
                SqlCmd.Parameters.Add(parameter);
                parameter = new SqlParameter("@Credit", SqlDbType.Decimal);
                parameter.Precision = 19;
                parameter.Scale = 6;
                parameter.Value = Convert.ToDecimal(Credit);
                SqlCmd.Parameters.Add(parameter);
                SqlCmd.Parameters.Add("@Posted", SqlDbType.VarChar, 1).Value = posted;

                SqlCmd.Parameters.Add("@UserSign", SqlDbType.Int).Value = user;
                SqlCmd.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = DateTime.Now;
                SqlCmd.Parameters.Add("@UpdateDate", SqlDbType.DateTime).Value = DateTime.Now;


                SqlCmd.Prepare();
                SqlCmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("InsertPayAcct = " + ex.Message + qry.ToString());

            }

        }
        public void UpdateSumm(int docEntry, int linenum, string cardCode, int postedDocEntry, string pmess, string postOption, double payOnAcc
                   , string selp, int pdDocEntry, string PdMess)
        {
            StringBuilder qry = new StringBuilder();
            try
            {

                qry = new StringBuilder();
                qry.Append(" UPDATE  [dbo].[redi_ConsBP_Summ] SET CardCode=@CardCode,PostedDocEntry=@PostedDocEntry,PostedMess=@PostedMess,PostOption=@PostOption,PayOnAcc=@PayOnAcc ");
                qry.Append(" ,PostDPay=@PostDPay,PostDPayMessage=@PostDPayMessage,SelDraft=@SelDraft ");

                qry.Append(" WHERE DocEntry=@DocEntry AND LineNum=@LineNum");

                using (SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {
                    SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                    SqlCmd.Parameters.Add("@LineNum", SqlDbType.Int).Value = linenum;

                    SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;

                    SqlCmd.Parameters.Add("@PostedDocEntry", SqlDbType.Int).Value = postedDocEntry;
                    SqlCmd.Parameters.Add("@PostedMess", SqlDbType.VarChar, 50).Value = pmess;
                    SqlCmd.Parameters.Add("@PostOption", SqlDbType.VarChar, 5).Value = postOption;
                    SqlCmd.Parameters.Add("@PayOnAcc", SqlDbType.Decimal, 19);
                    SqlCmd.Parameters["@PayOnAcc"].Precision = 19;
                    SqlCmd.Parameters["@PayOnAcc"].Scale = 6;
                    SqlCmd.Parameters["@PayOnAcc"].Value = payOnAcc;
                    SqlCmd.Parameters.Add("@PostDPay", SqlDbType.Int).Value = pdDocEntry;
                    SqlCmd.Parameters.Add("@PostDPayMessage", SqlDbType.VarChar, 50).Value = PdMess;
                    SqlCmd.Parameters.Add("@SelDraft", SqlDbType.VarChar, 1).Value = selp;
                    SqlCmd.Prepare();
                    SqlCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("UPDATESumm = " + ex.Message + qry.ToString());

            }

        }
        public void deleteSumm(int docEntry, int linenum)
        {
            StringBuilder qry = new StringBuilder();
            try
            {

                qry = new StringBuilder();
                qry.Append(" DELETE FROM   [dbo].[redi_ConsBP_Summ] ");
                qry.Append(" WHERE DocEntry=@DocEntry AND LineNum=@LineNum");

                using (SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {
                    SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                    SqlCmd.Parameters.Add("@LineNum", SqlDbType.Int).Value = linenum;

                    SqlCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("UPDATESumm = " + ex.Message + qry.ToString());

            }

        }
        public void InsertSumm(int docEntry, int linenum, string cardCode, int postedDocEntry, string pmess, string postOption, double payOnAcc, string selp, int pdDocEntry, string PdMess)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append("select Count(*) FROM [redi_ConsBP_Summ] where DocEntry = @DocEntry AND LineNum = @LineNum");
                SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                SqlCmd.Parameters.Add("@LineNum", SqlDbType.Int).Value = linenum;
                SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                SqlDataReader reader = SqlCmd.ExecuteReader();
                if (reader.Read() && Convert.ToInt32(reader.GetValue(0)) > 0)
                {
                    reader.Close();
                    SqlCmd.Dispose();
                    this.UpdateSumm(docEntry, linenum, cardCode, postedDocEntry, pmess, postOption, payOnAcc, selp, pdDocEntry, PdMess);
                    return;
                }
                reader.Close();
                SqlCmd.Dispose();
                qry = new StringBuilder();
                qry.Append(" Insert INTO [dbo].[redi_ConsBP_Summ] ([DocEntry],[LineNum],[CardCode],[PostedDocEntry] ,[PostedMess],[PostOption],[PayOnAcc],[PostDPay],[PostDPayMessage],[SelDraft])");
                qry.Append(" Values");
                qry.Append("(@DocEntry,@LineNum,@CardCode,@PostedDocEntry,@PostedMess,@PostOption,@PayOnAcc,@PostDPay,@PostDPayMessage,@SelDraft)");

                SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                SqlCmd.Parameters.Add("@LineNum", SqlDbType.Int).Value = linenum;

                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;

                SqlCmd.Parameters.Add("@PostedDocEntry", SqlDbType.Int).Value = postedDocEntry;
                SqlCmd.Parameters.Add("@PostedMess", SqlDbType.VarChar, 50).Value = pmess;
                SqlCmd.Parameters.Add("@PostOption", SqlDbType.VarChar, 1).Value = postOption;
                SqlCmd.Parameters.Add("@PayOnAcc", SqlDbType.Decimal, 19);
                SqlCmd.Parameters["@PayOnAcc"].Precision = 19;
                SqlCmd.Parameters["@PayOnAcc"].Scale = 6;
                SqlCmd.Parameters["@PayOnAcc"].Value = payOnAcc;
                SqlCmd.Parameters.Add("@PostDPay", SqlDbType.Int).Value = pdDocEntry;
                SqlCmd.Parameters.Add("@PostDPayMessage", SqlDbType.VarChar, 50).Value = PdMess;
                SqlCmd.Parameters.Add("@SelDraft", SqlDbType.VarChar, 1).Value = selp;
                SqlCmd.Prepare();
                SqlCmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("InsertSumm = " + ex.Message + qry.ToString());

            }

        }
        private void deleteTrans(int hdocEntry, int docEntry, int linenum)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append("Delete FROM [redi_ConsBP_Trans] where HDocEntry = @HDocEntry AND DocEntry = @DocEntry");
                using (SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {
                    SqlCmd.Parameters.Add("@HDocEntry", SqlDbType.Int).Value = hdocEntry;
                    SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                    SqlCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("Delete trans = " + ex.Message + qry.ToString());

            }
        }
        public void InsertTrans(int hdocEntry, int docEntry, int linenum, int docNum, string cardCode, string cardName, string docType, DateTime docDate,
                              DateTime dueDate, string daysPastDue, Double docTotal, Double balDue, Double TotalDisc, Double TotalPay, string selected, string posted)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append("select Count(*) FROM [redi_ConsBP_Trans] where HDocEntry = @HDocEntry AND DocEntry = @DocEntry");
                SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                SqlCmd.Parameters.Add("@HDocEntry", SqlDbType.Int).Value = hdocEntry;
                SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                SqlDataReader reader = SqlCmd.ExecuteReader();
                if (reader.Read() && Convert.ToInt32(reader.GetValue(0)) > 0)
                {
                    reader.Close();
                    SqlCmd.Dispose();
                    if (selected == "Y")
                        this.UpdateTrans(hdocEntry, docEntry, linenum, docNum, cardCode, cardName, docType, docDate,
                                       dueDate, daysPastDue, docTotal, balDue, TotalDisc, TotalPay, selected, posted);
                    else
                        deleteTrans(hdocEntry, docEntry, linenum);
                    return;
                }
                reader.Close();
                SqlCmd.Dispose();
                if (selected != "Y")
                    return;
                qry = new StringBuilder();

                qry.Append(" Insert INTO [dbo].[redi_ConsBP_Trans] ([DocEntry],[DocNum] ,[CardCode] ,[CardName] ,");
                qry.Append("[DocType],[DocDate] ,[DocDueDate] ,[DaysPastDue]  ,");
                qry.Append("[DocTotal] ,[TotalDisc] ,[TotalPay] ,[BalDue] ");
                qry.Append(",[Posted],[HDocEntry],[Linenum] ");
                qry.Append(",[Selected] ");

                qry.Append(") Values");
                qry.Append("(@DocEntry,@DocNum,@CardCode,@CardName,@Doctype,@DocDate,@DocDueDate,@DaysPastDue,@DocTotal,@TotalDisc,@TotalPay,@BalDue,@Posted,@HDocEntry,@Linenum,@Selected)");
                SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                SqlCmd.Parameters.Add("@DocNum", SqlDbType.Int).Value = docNum;
                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                SqlCmd.Parameters.Add("@CardName", SqlDbType.VarChar, 100).Value = cardName;
                SqlCmd.Parameters.Add("@Doctype", SqlDbType.VarChar, 2).Value = docType;
                SqlCmd.Parameters.Add("@DocDate", SqlDbType.DateTime).Value = docDate;
                SqlCmd.Parameters.Add("@DocDueDate", SqlDbType.DateTime).Value = dueDate;
                SqlCmd.Parameters.Add("@DaysPastDue", SqlDbType.NVarChar, 3).Value = daysPastDue;

                SqlCmd.Parameters.Add("@DocTotal", SqlDbType.Decimal, 19);
                SqlCmd.Parameters["@DocTotal"].Precision = 19;
                SqlCmd.Parameters["@DocTotal"].Scale = 6;
                SqlCmd.Parameters["@DocTotal"].Value = Convert.ToDecimal(docTotal);
                SqlCmd.Parameters.Add("@TotalDisc", SqlDbType.Decimal, 19);
                SqlCmd.Parameters["@TotalDisc"].Precision = 19;
                SqlCmd.Parameters["@TotalDisc"].Scale = 6;
                SqlCmd.Parameters["@TotalDisc"].Value = Convert.ToDecimal(TotalDisc);
                SqlCmd.Parameters.Add("@TotalPay", SqlDbType.Decimal, 19);
                SqlCmd.Parameters["@TotalPay"].Precision = 19;
                SqlCmd.Parameters["@TotalPay"].Scale = 6;
                SqlCmd.Parameters["@TotalPay"].Value = Convert.ToDecimal(TotalPay);
                SqlCmd.Parameters.Add("@BalDue", SqlDbType.Decimal, 19);
                SqlCmd.Parameters["@BalDue"].Precision = 19;
                SqlCmd.Parameters["@BalDue"].Scale = 6;
                SqlCmd.Parameters["@BalDue"].Value = Convert.ToDecimal(balDue);
                SqlCmd.Parameters.Add("@Posted", SqlDbType.Char, 1).Value = posted;
                SqlCmd.Parameters.Add("@Selected", SqlDbType.Char, 1).Value = selected;
                SqlCmd.Parameters.Add("@HDocEntry", SqlDbType.Int).Value = hdocEntry;
                SqlCmd.Parameters.Add("@Linenum", SqlDbType.Int).Value = linenum;
                SqlCmd.Prepare();
                SqlCmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception(ex.Message + qry.ToString());

            }

        }
        public void UpdateTransPost(int hdocEntry, string cardCode, string selected, string posted)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append(" UPDATE  [dbo].[redi_ConsBP_Trans] SET ");

                qry.Append("Posted=@Posted");

                qry.Append(" WHERE HDocEntry = @HDocEntry AND CardCode=@CardCode AND Selected=@Selected");
                SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                SqlCmd.Parameters.Add("@HDocEntry", SqlDbType.Int).Value = hdocEntry;

                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                SqlCmd.Parameters.Add("@Posted", SqlDbType.Char, 1).Value = posted;
                SqlCmd.Parameters.Add("@Selected", SqlDbType.Char, 1).Value = selected;
                SqlCmd.Prepare();
                SqlCmd.ExecuteNonQuery();
                SqlCmd.Dispose();
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception(ex.Message + qry.ToString());

            }

        }


        public void UpdateTrans(int hdocEntry, int docEntry, int linenum, int docNum, string cardCode, string cardName, string docType, DateTime docDate,
                             DateTime dueDate, string daysPastDue, Double docTotal, Double balDue, Double TotalDisc, Double TotalPay, string selected, string posted)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append(" UPDATE  [dbo].[redi_ConsBP_Trans] SET ");

                qry.Append("DocNum=@DocNum,CardCode=@CardCode,CardName=@CardName,Doctype=@Doctype,DocDate=@DocDate,DocDueDate=@DocDueDate,DaysPastDue=@DaysPastDue");
                qry.Append(",DocTotal=@DocTotal,TotalDisc=@TotalDisc,TotalPay=@TotalPay,BalDue=@BalDue,Selected=@Selected ");
                qry.Append(" WHERE HDocEntry = @HDocEntry AND Linenum =@Linenum");
                using (SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {
                    SqlCmd.Parameters.Add("@HDocEntry", SqlDbType.Int).Value = hdocEntry;
                    SqlCmd.Parameters.Add("@Linenum", SqlDbType.Int).Value = linenum;
                    SqlCmd.Parameters.Add("@DocNum", SqlDbType.Int).Value = docNum;
                    SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                    SqlCmd.Parameters.Add("@CardName", SqlDbType.VarChar, 100).Value = cardName;
                    SqlCmd.Parameters.Add("@Doctype", SqlDbType.VarChar, 2).Value = docType;
                    SqlCmd.Parameters.Add("@DocDate", SqlDbType.DateTime).Value = docDate;
                    SqlCmd.Parameters.Add("@DocDueDate", SqlDbType.DateTime).Value = dueDate;
                    SqlCmd.Parameters.Add("@DaysPastDue", SqlDbType.NVarChar, 3).Value = daysPastDue;

                    SqlCmd.Parameters.Add("@DocTotal", SqlDbType.Decimal, 19);
                    SqlCmd.Parameters["@DocTotal"].Precision = 19;
                    SqlCmd.Parameters["@DocTotal"].Scale = 6;
                    SqlCmd.Parameters["@DocTotal"].Value = Convert.ToDecimal(docTotal);
                    SqlCmd.Parameters.Add("@TotalDisc", SqlDbType.Decimal, 19);
                    SqlCmd.Parameters["@TotalDisc"].Precision = 19;
                    SqlCmd.Parameters["@TotalDisc"].Scale = 6;
                    SqlCmd.Parameters["@TotalDisc"].Value = Convert.ToDecimal(TotalDisc);
                    SqlCmd.Parameters.Add("@TotalPay", SqlDbType.Decimal, 19);
                    SqlCmd.Parameters["@TotalPay"].Precision = 19;
                    SqlCmd.Parameters["@TotalPay"].Scale = 6;
                    SqlCmd.Parameters["@TotalPay"].Value = Convert.ToDecimal(TotalPay);
                    SqlCmd.Parameters.Add("@BalDue", SqlDbType.Decimal, 19);
                    SqlCmd.Parameters["@BalDue"].Precision = 19;
                    SqlCmd.Parameters["@BalDue"].Scale = 6;
                    SqlCmd.Parameters["@BalDue"].Value = Convert.ToDecimal(balDue);
                    // SqlCmd.Parameters.Add("@Posted", SqlDbType.Char, 1).Value = posted;
                    SqlCmd.Parameters.Add("@Selected", SqlDbType.Char, 1).Value = selected;

                    SqlCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception(ex.Message + qry.ToString());

            }

        }

        public int InsertHeader(string cardCode, DateTime docDate, DateTime dueDate, string docRef, int user, int docNum)
        {
            StringBuilder qry = new StringBuilder();
            int retEnrty = 0;
            try
            {
                qry.Append(" Insert INTO [dbo].[redi_ConsBP_Header] ([CardCode], ");
                qry.Append("[DocDate] , [DocDueDate] ,[DocRef] ,[UserSign] ,[CreateDate],[DocNum]) Values");
                qry.Append("(@CardCode,@DocDate,@DocDueDate,@DocRef,@userSign,@CreateDate,@DocNum)  SELECT SCOPE_IDENTITY()");
                SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);

                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                SqlCmd.Parameters.Add("@DocDate", SqlDbType.DateTime).Value = docDate;
                SqlCmd.Parameters.Add("@DocDueDate", SqlDbType.DateTime).Value = dueDate;
                SqlCmd.Parameters.Add("@DocRef", SqlDbType.VarChar, 32).Value = docRef;
                SqlCmd.Parameters.Add("@userSign", SqlDbType.Int).Value = user;
                SqlCmd.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = DateTime.Now;
                SqlCmd.Parameters.Add("@DocNum", SqlDbType.Int).Value = docNum;
                SqlCmd.Prepare();
                var x = SqlCmd.ExecuteScalar();
                retEnrty = Convert.ToInt32(x);

            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception(ex.Message + qry.ToString());

            }

            return retEnrty;

        }
        public int UpdateHeader(string cardCode, DateTime docDate, DateTime dueDate, string docRef, int docentry)
        {
            StringBuilder qry = new StringBuilder();
            int retEnrty = 0;
            try
            {
                qry.Append(" UPDATE [dbo].[redi_ConsBP_Header] SET ");
                qry.Append(" CardCode = @CardCode,DocDate=@DocDate,DocDueDate=@DocDueDate,DocRef=@DocRef");
                qry.Append(" WHERE DocEntry=@DocEntry");
                using (SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {

                    SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                    SqlCmd.Parameters.Add("@DocDate", SqlDbType.DateTime).Value = docDate;
                    SqlCmd.Parameters.Add("@DocDueDate", SqlDbType.DateTime).Value = dueDate;
                    SqlCmd.Parameters.Add("@DocRef", SqlDbType.VarChar, 32).Value = docRef;
                    SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docentry;

                    SqlCmd.Prepare();
                    retEnrty = Convert.ToInt32(SqlCmd.ExecuteScalar());
                }
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception(ex.Message + qry.ToString());

            }

            return retEnrty;

        }
        public string[] CreateJEntry(List<JE> ljes,string cardcode)
        {

            string[] ret = { "-1", "" };
            int jenumber = 0;
            JournalEntries journalEntries = (JournalEntries)ProgData.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
            journalEntries.AutoVAT = BoYesNoEnum.tYES;
            journalEntries.DueDate = DateTime.Now;
            journalEntries.TaxDate = DateTime.Now;
            journalEntries.ReferenceDate = DateTime.Now;

            int ix = -1;

            foreach (var item in ljes)
            {
                ix++;
                journalEntries.Lines.SetCurrentLine(ix);
                if (item.CardCode == item.Account)
                    journalEntries.Lines.ShortName = item.CardCode;
                else
                {
                    string acct = item.Account.Replace("-", "");
                    string qry = $"select AcctCode from OACT where FormatCode = '{acct}' ";
                    SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRec.DoQuery(qry);
                    try
                    {
                        acct = oRec.Fields.Item(0).Value.ToString();
                    } catch (Exception e)
                    {
                        acct = item.Account;
                    }
                    journalEntries.Lines.AccountCode = acct;
                    
                }
               
                journalEntries.Lines.Credit = item.Credit;
                journalEntries.Lines.Debit = item.Debit;
                journalEntries.Lines.LineMemo = item.LineMemo;
                journalEntries.Lines.BPLID = 1;
                Logger.Log(journalEntries.Lines.Line_ID.ToString() + "=" + journalEntries.Lines.AccountCode + journalEntries.Lines.LineMemo);

                // journalEntries.Lines.Docn = item.Docnum
                journalEntries.Lines.Add();
            }

            int RetVal = journalEntries.Add();
            int batchnum = 0;
            if (RetVal != 0)
            {
                ProgData.B1Company.GetLastError(out ErrCode, out ErrMsg);
                ret[0] = RetVal.ToString();
                ret[1] = ErrCode.ToString() + "-" + ErrMsg;
            } else
            {
                jenumber = journalEntries.JdtNum;
                ret[0] = jenumber.ToString();

                if (jenumber == 0)
                {
                    string JENumber = "";
                    ProgData.B1Company.GetNewObjectCode(out JENumber);
                    ret[0] = JENumber;
                }
                ret[1] = ret[0];
            }
            return ret;

        }
        public int CreateJVoucher(string CardCode, double amount)
        {
            int ret = 0;
            JournalVouchers jVoucher = (JournalVouchers)ProgData.B1Company.GetBusinessObject(BoObjectTypes.oJournalVouchers);
            jVoucher.JournalEntries.AutoVAT = BoYesNoEnum.tYES;
            jVoucher.JournalEntries.DueDate = DateTime.Now;
            jVoucher.JournalEntries.TaxDate = DateTime.Now;
            jVoucher.JournalEntries.ReferenceDate = DateTime.Now;
            jVoucher.JournalEntries.Lines.SetCurrentLine(0);
            jVoucher.JournalEntries.Lines.ShortName = CardCode;
            jVoucher.JournalEntries.Lines.Credit = amount;

            jVoucher.JournalEntries.Lines.Add();
            jVoucher.JournalEntries.Lines.SetCurrentLine(1);
            ret = jVoucher.Add();
            if (!ret.Equals(0))
            {

            }
            return ret;
        }
        public void DeletePayAcct(int DocEntry, int Linenum)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append("Delete FROM [redi_ConsBP_PayAcct] where DocEntry = @DocEntry AND SumLinenum = @SumLinenum");
                using (SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction))
                {
                    SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = DocEntry;
                    SqlCmd.Parameters.Add("@SumLinenum", SqlDbType.Int).Value = Linenum;
                    SqlCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("Delete PAyacct = " + ex.Message + qry.ToString());

            }

        }
        public void DeleteDraftPayment(int DPEntry)
        {
            Payments payments = (Payments)ProgData.B1Company.GetBusinessObject(BoObjectTypes.oPaymentsDrafts);
            payments.GetByKey(DPEntry);
            int RetVal =  payments.Remove();
            if (RetVal != 0)
            {
                ProgData.B1Company.GetLastError(out ErrCode, out ErrMsg);

                ProgData.B1Application.MessageBox(" can not delete DraftPayment  DocEntry ="+DPEntry.ToString()+ " Error = " +ErrCode.ToString() + "-" + ErrMsg);
              //  throw new Exception("Delete PAyacct = " + " can not delete DraftPayment  DocEntry =" + DPEntry.ToString() + " Error = " + ErrCode.ToString() + "-" + ErrMsg);
            }
        }
        public string[] UpdateDraftPayment(int DPEntry, string cardCode, double amount, string acctnum, DateTime docdate, string reference, List<grpPaymentdetail> grpdetail)
        {
            string[] ret = new string[2] { "-1", "" };
            Payments payments = (Payments)ProgData.B1Company.GetBusinessObject(BoObjectTypes.oPaymentsDrafts);
            payments.GetByKey(DPEntry);
            payments.TransferSum = amount;
            payments.LocalCurrency = BoYesNoEnum.tYES;
            payments.TransferAccount = acctnum;
            payments.TransferDate = docdate;
            payments.TransferReference = reference;
            int ix = 0;
            foreach (var item in grpdetail)
            {
                payments.Invoices.SetCurrentLine(ix);
                ix++;
                if (payments.Invoices.DocEntry == item.DocDocEntry)
                {
                    payments.Invoices.TotalDiscount = item.TotalDisc;
                   
                } else
                {
                    payments.Invoices.DocEntry = item.DocDocEntry;
                    if (item.DocType == "IN")
                    {
                        payments.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                    }
                    else
                    if (item.DocType == "CM")
                    {
                        payments.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                    }
                    else
                      if (item.DocType == "JE")
                        payments.Invoices.InvoiceType = BoRcptInvTypes.it_JournalEntry;
                    payments.Invoices.TotalDiscount = item.TotalDisc;

                    payments.Invoices.Add();

                }
            }
            int RetVal = payments.Update();

            int docentry = -1;
            if (RetVal != 0)
            {
                ProgData.B1Company.GetLastError(out ErrCode, out ErrMsg);
                ret[0] = RetVal.ToString();
                ret[1] = ErrCode.ToString() + "-" + ErrMsg;
            }
            else
            {
                docentry = payments.DocEntry;
                ret[0] = docentry.ToString();
                ret[1] = payments.DocNum.ToString();
                if (docentry == 0)
                {
                    string sDocEntry = "";
                    ProgData.B1Company.GetNewObjectCode(out sDocEntry);
                    ret[0] = sDocEntry;
                }
                ret[1] = ret[0];
            }
            return ret;
        }
    
        public string[] CreateDraftPayment(string cardCode,double amount,string acctnum,DateTime docdate,string reference, List<grpPaymentdetail> grpdetail)
        {
            string[] ret = new string[2] { "-1", "" };
            string qry = $"SELECT \"Series\" FROM \"NNM1\" " +
                    $"WHERE \"ObjectCode\"='24' AND \"Locked\"='N' AND \"GroupCode\" = 1 AND (\"BPLId\" Is null or \"BPLId\" = 1) ORDER BY \"BPLId\" desc ";
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(qry);

            Payments payments = (Payments)ProgData.B1Company.GetBusinessObject(BoObjectTypes.oPaymentsDrafts);
             
              payments.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
              payments.DocType = BoRcptTypes.rCustomer;
               payments.Series = Convert.ToInt32(oRec.Fields.Item(0).Value);
              payments.CardCode = cardCode;
              payments.DocDate = DateTime.Now;
              payments.TransferSum = amount;
              payments.LocalCurrency = BoYesNoEnum.tYES;
            payments.TransferAccount = acctnum;
            payments.TransferDate = docdate;
            payments.TransferReference = reference;

            // payments.DocCurrency = currency;
            int ix = 0;
            foreach (var item in grpdetail)
            {
                payments.Invoices.SetCurrentLine(ix);
                ix++;
                payments.Invoices.DocEntry = item.DocDocEntry;
                if (item.DocType == "IN")
                {
                    payments.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                }
                else
                if (item.DocType == "CM")
                {
                    payments.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                }
                else
                  if (item.DocType == "JE")
                    payments.Invoices.InvoiceType = BoRcptInvTypes.it_JournalEntry;
                payments.Invoices.TotalDiscount = item.TotalDisc;
            
                payments.Invoices.Add();
            }
           
            int RetVal = payments.Add();
            
            int docentry = -1;
            if (RetVal != 0)
            {
                ProgData.B1Company.GetLastError(out ErrCode, out ErrMsg);
                ret[0] = RetVal.ToString();
                ret[1] = ErrCode.ToString() + "-" + ErrMsg;
            }
            else
            {
                docentry = payments.DocEntry;
                ret[0] = docentry.ToString();
                ret[1] = payments.DocNum.ToString();
                if (docentry == 0)
                {
                    string sDocEntry = "";
                    ProgData.B1Company.GetNewObjectCode(out sDocEntry);
                    ret[0] = sDocEntry;
                }
                ret[1] = ret[0];
            }
            return ret;
        }
    }

}
