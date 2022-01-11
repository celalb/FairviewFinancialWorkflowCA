using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FairviewFinancialWorkflowCA
{
    class Logger
    {
        public static string exeFolder;
        public static string firm;
        public static void Log(string message)
        {
            try
            {


                string path = exeFolder;
                path += @"\" + firm + @"\";

                string filename = path + "process" + DateTime.Now.ToString("yyyyMMdd") + ".Log";


                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);



                var sr = File.AppendText(filename);
                TextWriter w = (TextWriter)sr;
                w.WriteLine(message + " Tarih = " + DateTime.Now.ToShortTimeString());
                w.Close();




            }
            catch (Exception errt)
            {
                Console.WriteLine(errt.Message);
                Debug.WriteLine(errt.Message);

            }

            Console.WriteLine(message);
            Debug.WriteLine(message);
        }

        public static void Log(Exception ex)
        {
            try
            {


                string path = exeFolder;
                path += @"\" + firm + @"\";

                string filename = path + "Process" + DateTime.Now.ToString("yyyyMMdd") + ".Log";


                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);



                var sr = File.AppendText(filename);
                TextWriter w = (TextWriter)sr;
                w.WriteLine(ex.Message + " Source = " + ex.StackTrace + " Tarih = " + DateTime.Now.ToShortTimeString());
               
                w.Close();




            }
            catch (Exception errt)
            {
                Console.WriteLine(errt.Message);
                Debug.WriteLine(errt.Message);

            }

            Console.WriteLine(ex.Message);
            Debug.WriteLine(ex.Message);
        }
    }

}
