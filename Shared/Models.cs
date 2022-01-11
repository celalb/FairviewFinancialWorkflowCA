using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared
{
    public struct grpPaymentdetail
    {
       
        public int DocDocEntry { get; set; }
      
        public string DocType { get; set; }
        public string DocNum { get; set; }
        public Double BalDue { get; set; }
        public Double TotalDisc { get; set; }
        public Double TotalPay { get; set; }
    }
    public class Payment
    {
        public string Sel { get; set; }
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string DocType { get; set; }
        public string DocNum { get; set; }
        public int DocEntry { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DueDate { get; set; }
        public string DPastDue { get; set; }
        public Double DocTotal { get; set; }
        public Double BalDue { get; set; }
        public Double TotalDisc { get; set; }
        public Double TotalPay { get; set; }
        public int Row { get; set; }
       
        public int Linenum { get; set; }

        public int PostedDocEntry { get; set; }

    }
    public struct Grouppay
    {
        public int DocEntry { get; set; }
        public string CardCode { get; set; }
        public string DocNum { get; set; }
        public Double DocTotal { get; set; }
        public Double TotalDisc { get; set; }
        public Double TotalPay { get; set; }
        public Double PayAcc { get; set; }
        public string SelP { get; set; }
        public string SelJ { get; set; }
        public int Row { get; set; }
        public int Linenum { get; set; }
        public int PostedJDocEntry { get; set; }
        public int PostedDDocEntry { get; set; }
        public string JMess { get; set; }
        public string DMess { get; set; }
        public List<grpPaymentdetail> paymentdetail { get; set; }
    }
    public struct JE
    {

        public int LineNum { get; set; }
        public string Account { get; set; }
        public string CardCode { get; set; }
        public string DocType { get; set; }
        public int Docnum { get; set; }
        public string LineMemo { get; set; }
        public double Debit { get; set; }
        public Double Credit { get; set; }
        public DateTime PostDate { get; set; }
        public string  Posted { get; set; }
        public int SumLinenum { get; set; }
        public int DocEntry { get; set; }
    }
    public struct JesKey
    {
        public int DocEntry { get; set; }
        public string CardCode { get; set; }
    }

}
