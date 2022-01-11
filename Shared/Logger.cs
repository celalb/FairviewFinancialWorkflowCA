﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared
{
    public class Logger
    {
        public static string exeFolder;
        public static string firm;
        public static void Log(string message)
        {
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory;


                path += @"\" + firm + @"\";
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                string flname = path + "info.log";

                var sr = File.AppendText(flname);
                TextWriter w = (TextWriter)sr;
                w.WriteLine(message);
                w.Flush();
                w.Close();
                //////////

            }
            catch (Exception errt)
            {
                Console.WriteLine(errt.Message);

            }
            Console.WriteLine(message);

        }

        public static void Log(Exception ex)
        {
            try
            {


                string path = AppDomain.CurrentDomain.BaseDirectory;


                path += @"\" + firm + @"\";
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                string flname = path + "Error.log";

                var sr = File.AppendText(flname);
                TextWriter w = (TextWriter)sr;
                w.WriteLine("Error="+ ex.Message + "\n\t" + DateTime.Now.ToString("yyyyMMdd HH-mm") + "\n\tSource = " + ex.StackTrace);
                w.Flush();
                w.Close();




            }
            catch (Exception errt)
            {
                Console.WriteLine(errt.Message);

            }

        }

    }

}
