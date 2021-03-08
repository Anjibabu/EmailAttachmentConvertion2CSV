using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Quartz;
using Quartz.Impl;
using excel = Microsoft.Office.Interop.Excel;

namespace EmailAttachments
{

    class Program
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            //EmailReadFromOutLook();
            try
            {
                ISchedulerFactory schedulerFactory = new StdSchedulerFactory();
                IScheduler scheduler = schedulerFactory.GetScheduler();

                IJobDetail jobDetail = JobBuilder.Create<EmailReadJob>()
                    .WithIdentity("TestJob")
                    .Build();
                ITrigger trigger = TriggerBuilder.Create()
                    .ForJob(jobDetail)
                    .WithSimpleSchedule(x => x.WithIntervalInSeconds(int.Parse(ConfigurationManager.AppSettings["FirstInterval"])).RepeatForever())
                    // .WithCronSchedule("10 * * * * ?")
                    .WithIdentity("TestTrigger")
                    .StartNow()
                    .Build();
                scheduler.ScheduleJob(jobDetail, trigger);
                scheduler.Start();

            }
            catch (ArgumentException e)
            {
                log.ErrorFormat("Error=", e.Message);
            }

            log.Info("Done");
            //Console.ReadLine();
        }

    }
    public class EmailReadJob : IJob
    {


        private static Workbook mWorkBook;
        private static Sheets mWorkSheets;
        private static Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;
        private static string ErrorMessage = string.Empty;
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static string BasePath
        {
            get
            {
                return ConfigurationManager.AppSettings["folderPath"];
            }
        }
        public void Execute(IJobExecutionContext context)
        {
            log.InfoFormat("{0} - Started", context.Trigger.Description);
            EmailReadFromOutLook();
            log.InfoFormat("{0} - End ", context.Trigger.Description);
        }

        private static void EmailReadFromOutLook()
        {
            // Console.WriteLine("Running..");
            Microsoft.Office.Interop.Outlook.Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;
            log.Info("Emails Reading from OutLook started.");
            try
            {
                outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                string Filter = "[ReceivedTime] >= Today";

                Items mis = inboxFolder.Items.Restrict(Filter);
                int cnt = mis.Count; ;

                log.InfoFormat("email Count :{0} ", mis.Count);
                DeleteAllFiles();
                foreach (MailItem item in mis)
                {
                    if (item.Attachments.Count > 0)
                    {
                        foreach (Attachment attach in item.Attachments)
                        {
                            SaveFile(attach);
                        }
                    }
                    Marshal.ReleaseComObject(item);
                }
            }
            //Error handler.
            catch (System.Exception e)
            {
                log.ErrorFormat("{0} Exception caught: {0} ", e);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        public static void DeleteAllFiles()
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(BasePath);

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        public static void SaveAsCSV(string sourcePath)
        {
            excel.Application xlApp = new excel.Application();
            excel.Workbook xlWorkBook = xlApp.Workbooks.Open(sourcePath);
            xlApp.Visible = true;
            foreach (excel.Worksheet sht in xlWorkBook.Worksheets)
            {
                sht.Select();
                xlWorkBook.SaveAs(string.Format("{0}.csv",Path.Combine( BasePath, sht.Name)), excel.XlFileFormat.xlCSV, excel.XlSaveAsAccessMode.xlNoChange);

            }
            xlWorkBook.Close(false);
        }
        public static void SaveFile(Attachment attach)
        {
            var fileName = BasePath + attach.FileName;
            string fileExtension = Path.GetExtension(fileName);
            string filenameWithNameWithExtension = Path.GetFileName(fileName);
            string csvFilePath = Path.Combine(BasePath, filenameWithNameWithExtension);
            if (fileExtension == ".xls" || fileExtension == ".xlsx")
            {
                if (File.Exists(csvFilePath))
                {
                    File.Delete(csvFilePath);
                }
                attach.SaveAsFile(csvFilePath);
                if (File.Exists(csvFilePath))
                {
                    File.Delete(csvFilePath);
                }
            }
            else
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
                attach.SaveAsFile(fileName);
            }
        }
        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }



}
