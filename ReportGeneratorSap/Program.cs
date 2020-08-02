using System;
using ReportGeneratorSap.Services;

namespace ReportGeneratorSap
{
    class Program
    {
        public static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Uploading File please waiting....");
                var serviceDropBox = new ServiceDropbox();
           
                Console.WriteLine("Processing...!!");
                var result = serviceDropBox.convertoAllToXlsx();
                if (result) serviceDropBox.processUpload().Wait();
                else
                {
                    string date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    Console.WriteLine($"{date} Error => files are not converted!!!");
                    logger.Trace($"{date} Error => files are not converted!!!");
                }
                //Console.ReadKey();
            }
            catch (Exception ex )
            {
                string date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                Console.WriteLine(ex.Message);
                if(ex.InnerException != null)
                {
                    Console.WriteLine(ex.InnerException.Message);
                    Console.WriteLine($"{date} Error => {ex.InnerException.Message}");
                    logger.Trace($"{date} Error => {ex.InnerException.Message}");
                }
                Console.WriteLine($"{date} Error => {ex.Message}");
                logger.Trace($"{date} Error => {ex.Message}");
                //Console.ReadKey();
            }
        }


    }



}
