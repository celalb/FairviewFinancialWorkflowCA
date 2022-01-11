using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FairviewFinancialWorkflowCA
{
    class DbUtility
    {
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
                    } catch (Exception e)
                    {
                        Logger.Log(e);
                    }
                }

                return cn;

            }

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
                qry.Append("[DocTotal] [Double](19, 6) NULL,[BalDue] [Double](19, 6) NULL,[DiscTotal] [Double](19, 6) NULL,[PayTotal] [Double](19, 6) NULL,");
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

        public void InsertPayAcct(int docEntry, int docNum, string cardCode, string docType, DateTime refDate,
                           string LineMemo, Decimal Debit, Decimal Credit,string user)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append(" Insert INTO [dbo].[redi_ConsBP_PayAcct] ([DocEntry] ,[CardCode],[RefDate],[LineNum],[Account],[DocType]");
                 qry.Append(",[DocNum],[LineMemo],[Debit],[Credit],[UserSign],[CreateDate] ,[UpdateDate])");
                qry.Append(" Values");
                qry.Append("(@DocEntry,@CardCode,@RefDate,@LineNum,@Account,@DocType");
                qry.Append(",@DocNum,@LineMemo,@Debit,@Credit,@UserSign,@CreateDate,@UpdateDate)");
                SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                SqlCmd.Parameters.Add("@DocNum", SqlDbType.Int).Value = docNum;
                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
               
                SqlCmd.Parameters.Add("@Doctype", SqlDbType.VarChar, 2).Value = docType;
                SqlCmd.Parameters.Add("@RefDate", SqlDbType.DateTime).Value = refDate;
                SqlCmd.Parameters.Add("@LineMemo", SqlDbType.VarChar, 15).Value = LineMemo;
                SqlCmd.Parameters.Add("@Debit", SqlDbType.Decimal,19);
                SqlCmd.Parameters["@Debit"].Precision = 19;
                SqlCmd.Parameters["@Debit"].Scale = 6;
                SqlCmd.Parameters["@Debit"].Value = Debit;
                SqlCmd.Parameters.Add("@Credit", SqlDbType.Decimal,19);
                SqlCmd.Parameters["@Credit"].Precision = 19;
                SqlCmd.Parameters["@Credit"].Scale = 6;
                SqlCmd.Parameters["@Credit"].Value = Credit;
               
              
                SqlCmd.Parameters.Add("@userSign", SqlDbType.Int).Value = user;
                SqlCmd.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = DateTime.Now;
                SqlCmd.Parameters.Add("@@UpdateDate", SqlDbType.DateTime).Value = DateTime.Now;
              
                SqlCmd.Prepare();
                SqlCmd.ExecuteScalar();

            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception("InsertPayAcct = " + ex.Message+ qry.ToString());

            }

        }


        public void InsertTrans(int docEntry,int docNum,string cardCode,string cardName,string docType,DateTime docDate,
                              DateTime dueDate,string daysPastDue, Double docTotal, Double balDue, Double discTotal,Double payTotal,string selected)
        {
            StringBuilder qry = new StringBuilder();
            try
            {
                qry.Append(" Insert INTO [dbo].[redi_ConsBP_Trans] ([DocEntry],[DocNum] ,[CardCode] ,[CardName] ,");
                qry.Append("[DocType],[DocDate] ,[DocDueDate] ,[DaysPastDue]  ,");
                qry.Append("[DocTotal] ,[BalDue] ,[DiscTotal] ,[PayTotal] ,");
                qry.Append("[PayOnAccount] ,[Selected] ");
                qry.Append(") Values");
                qry.Append("(@DocEntry,@DocNum,@CardCode,@CardName,@Doctype,@DocDate,@DocDueDate,@DocPastDue,@DocTotal,@BalDue,@DiscTotal,@PayTotal,@PayOnAccount,@Selected)");
                SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                SqlCmd.Parameters.Add("@DocEntry", SqlDbType.Int).Value = docEntry;
                SqlCmd.Parameters.Add("@DocNum", SqlDbType.Int).Value = docNum;
                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                SqlCmd.Parameters.Add("@CardName", SqlDbType.VarChar, 100).Value = cardName;
                SqlCmd.Parameters.Add("@Doctype", SqlDbType.VarChar, 2).Value = docType;
                SqlCmd.Parameters.Add("@DocDate", SqlDbType.DateTime).Value = docDate;
                SqlCmd.Parameters.Add("@DocDueDate", SqlDbType.DateTime).Value = dueDate;
                SqlCmd.Parameters.Add("@DaysPastDue", SqlDbType.NVarChar,3).Value = daysPastDue;

                SqlCmd.Parameters.Add("@DocTotal", SqlDbType.Decimal,19);
                SqlCmd.Parameters["@DocTotal"].Precision = 19;
                SqlCmd.Parameters["@DocTotal"].Scale = 6;
                SqlCmd.Parameters["@DocTotal"].Value = Convert.ToDecimal(docTotal);
                SqlCmd.Parameters.Add("@BalDue", SqlDbType.Decimal,19);
                SqlCmd.Parameters["@BalDue"].Precision = 19;
                SqlCmd.Parameters["@BalDue"].Scale = 6;
                SqlCmd.Parameters["@BalDue"].Value = Convert.ToDecimal(balDue);
                SqlCmd.Parameters.Add("@DiscTotal", SqlDbType.Decimal,19);
                SqlCmd.Parameters["@DiscTotal"].Precision = 19;
                SqlCmd.Parameters["@DiscTotal"].Scale = 6;
                SqlCmd.Parameters["@DiscTotal"].Value = Convert.ToDecimal(discTotal);
                SqlCmd.Parameters.Add("@PayTotal", SqlDbType.Decimal,19);
                SqlCmd.Parameters["@PayTotal"].Precision = 19;
                SqlCmd.Parameters["@PayTotal"].Scale = 6;
                SqlCmd.Parameters["@PayTotal"].Value = Convert.ToDecimal(payTotal);
                SqlCmd.Parameters.Add("@Selected", SqlDbType.Char,1).Value = selected;
                SqlCmd.Prepare();
               SqlCmd.ExecuteScalar();

            }
            catch (Exception ex)
            {
                if (transaction != null)
                    transaction.Rollback();
                cn.Close();
                throw new Exception( ex.Message + qry.ToString());

            }

        }
        public int InsertHeader(string cardCode,DateTime docDate,DateTime dueDate,string docRef,int user)
        {
            StringBuilder qry = new StringBuilder();
            int retEnrty = 0;
            try
            {
                qry.Append(" Insert INTO [dbo].[redi_ConsBP_Header] ([CardCode] ");
                qry.Append("[DocDate] , [DocDueDate] ,[DocRef] ,[UserSign] ,[CreateDate]) Values");
                qry.Append("(@CardCode,@DocDate.@DocDueDate,@DocRef,@userSign,@CreateDate)");
                SqlCommand SqlCmd = new SqlCommand(qry.ToString(), DBConnection, transaction);
                
                SqlCmd.Parameters.Add("@CardCode", SqlDbType.VarChar, 15).Value = cardCode;
                SqlCmd.Parameters.Add("@DocDate", SqlDbType.DateTime).Value = docDate;
                SqlCmd.Parameters.Add("@DocDueDate", SqlDbType.DateTime).Value = dueDate;
                SqlCmd.Parameters.Add("@DocRef", SqlDbType.VarChar, 32).Value = docRef;
                SqlCmd.Parameters.Add("@userSign", SqlDbType.Int).Value = user;
                SqlCmd.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = DateTime.Now;
                SqlCmd.Prepare();
                retEnrty = Convert.ToInt32(SqlCmd.ExecuteScalar());

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
        public void CreateJVoucher(string CardCode, double amount)
        {
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
            if (!jVoucher.Add().Equals(0))
            {
            }
        }
        public void CrateDraftPayment()
        {
          /*  Payments payments = (Payments)ProgData.B1Company.GetBusinessObject(BoObjectTypes.oPaymentsDrafts);
            payments.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
            payments.DocType = BoRcptTypes.rCustomer;
            payments.CardCode = cardCode;
            payments.DocDate = DateTime.Now;
            payments.CashSum = amount;
            payments.LocalCurrency = BoYesNoEnum.tYES;
            payments.DocCurrency = currency;
            payments.Series = seriesNumber;
            payments.Invoices.DocEntry = docEntry;
            payments.Invoices.DocLine = 0;
            payments.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
            payments.Invoices.SumApplied = amount;
            if (!payments.Add() == 0)
            {

            }
          */
        }
    }
}
