using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace CreateExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {


                string emailFrom = ConfigurationManager.AppSettings["emailFrom"].ToString();

                string emailTo = ConfigurationManager.AppSettings["emailTo"].ToString();
                string emailSubject = ConfigurationManager.AppSettings["emailSubject"] + "- " + DateTime.UtcNow;

                string smtpPort = ConfigurationManager.AppSettings["smtpPort"].ToString();
                string smtpServer = ConfigurationManager.AppSettings["smtpServer"].ToString();
                string emailPwd = ConfigurationManager.AppSettings["emailPwd"].ToString();
                string emailBody = ConfigurationManager.AppSettings["Body"];

                String sDate = DateTime.UtcNow.ToString();
                DateTime datevalue = (Convert.ToDateTime(sDate.ToString()));

                String dy = datevalue.Day.ToString();
                String mn = datevalue.Month.ToString();
                String yy = datevalue.Year.ToString();
                string min = datevalue.Minute.ToString();
                string sec = datevalue.Second.ToString();
                string hours = datevalue.Hour.ToString();
                string url = @"\" + yy + @"\" + mn + @"\" + dy;
                string year = @"\" + yy;
                var timeFormat = dy + mn + yy + "-" + hours + min + sec;

                var fileName = ConfigurationManager.AppSettings["FilePath"].ToString() + ConfigurationManager.AppSettings["emailSubject"] + timeFormat + ".xlsx";

                DataTable dataTable = new DataTable("OneVisionAgency_GE_Properties");
                dataTable.Columns.Add("Sl.No", typeof(Int32));
                dataTable.Columns.Add("Message Name", typeof(string));
                dataTable.Columns.Add("Application Name", typeof(string));
                dataTable.Columns.Add("Success Archive FolderLocation", typeof(string));
                dataTable.Columns.Add("Direction", typeof(string));
                dataTable.Columns.Add("Files Count", typeof(Int32));
                dataTable.Columns.Add("Failed Count", typeof(Int32));

                MessageSectionGroup group = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).SectionGroups["MessageServerConfiguration"] as MessageSectionGroup;

                //List<Package> packages =
                //    new List<Package>
                //    {

                //        //new Package { SerialNumber = "1", MessageName = "APERAK COPARAN 01B", ApplicationName = "MSC.EDI.OVA.DE.IN.APKCOP01B",  ServerLocation = @"\\Ch001-ov-a\de-edi01\BZTFiles\IN\APERAK\APK01B_COPARN\Archive" ,DirectionName = "IN" ,filesCount = 1 ,failedFileCount = 0 },

                //        //new Package { SerialNumber = "2", MessageName = "APERAK IFTMCS 00B", ApplicationName = "MSC.EDI.OVA.DE.IN.APKMCS00B", ServerLocation = @"\\Ch001-ov-a\de-edi01\BZTFiles\IN\APERAK\APK00B_IFTMCS\Archive",DirectionName = "IN" ,filesCount = 1 ,failedFileCount = 0 }


                //    };

                //ExcelFacade excelFacade1 = new ExcelFacade();

                //List<string> headerName = new List<string> { "OneVision EDi Application", "Date/Time" };

                //excelFacade1.setColorforHeader<Package>(fileName,  "BIZTALK", headerName);

                //List<Package> packages = new List<Package>();
                foreach (ConfigurationSection section in group.Sections)

                {
                    if (section.GetType() == typeof(MessageSection))
                    {
                        MessageSection c = (MessageSection)section;
                        MessageCollections coll = c.Messages;
                        foreach (MessageElement item in coll)
                        {
                            int filesCount = 0;
                            int failedFileCount = 0;

                            if (Directory.Exists(item.ServerLocation + url))
                            {
                                var file = new DirectoryInfo(item.ServerLocation + url).GetFiles().Where(x => (x.Attributes & FileAttributes.Hidden) == 0).ToList();
                                filesCount = file.Count;
                            }

                            if (((item.DirectionName == "OUT") || (item.DirectionName == "TRANS")) && (Directory.Exists(item.FailedFileCountPath + url)))
                            {
                                var file = new DirectoryInfo(item.FailedFileCountPath).GetFiles().Where(x => (x.Attributes & FileAttributes.Hidden) == 0).ToList();
                                failedFileCount = file.Count;
                            }
                            dataTable.Rows.Add(item.SerialNumber, item.MessageName, item.ApplicationName, item.ServerLocation, item.DirectionName, filesCount, failedFileCount);

                        //    List<Package> packages =
                        //                       new List<Package>()
                        //                       {
                        //                           new Package() { SerialNumber = item.SerialNumber, MessageName = item.MessageName, ApplicationName = item.ApplicationName, ServerLocation = item.ServerLocation, DirectionName = item.DirectionName, filesCount = filesCount, failedFileCount = failedFileCount },
                        //};
                        //    //packages.Add(SerialNumber = item.SerialNumber);


                            //List<string> headerNames = new List<string> { "OneVision EDi Application", "Date/Time","S.NO", "Message Name", "Application Name", "Success Archive FolderLocation", "Direction", "Files Count", "Failed Count" };

                            //ExcelFacade excelFacade = new ExcelFacade();
                            //excelFacade.Create<Package>(fileName, packages, "BIZTALK", headerNames);



                        }
                    }
                }

                List<Package> packages = new List<Package>();
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    Package student = new Package();
                    student.SerialNumber = Convert.ToInt32(dataTable.Rows[i]["Sl.No"]);
                    student.MessageName = dataTable.Rows[i]["Message Name"].ToString();
                    student.ApplicationName = dataTable.Rows[i]["Application Name"].ToString();
                    student.ServerLocation = dataTable.Rows[i]["Success Archive FolderLocation"].ToString();

                    student.DirectionName = dataTable.Rows[i]["Direction"].ToString();

                    student.filesCount = Convert.ToInt32(dataTable.Rows[i]["Files Count"]);
                    student.failedFileCount = Convert.ToInt32(dataTable.Rows[i]["Failed Count"]);
                    packages.Add(student);
                }

                List<string> headerNames = new List<string> { "OneVision EDi Application", "Date/Time", "S.NO", "Message Name", "Application Name", "Success Archive FolderLocation", "Direction", "Files Count", "Failed Count" };

                ExcelFacade excelFacade = new ExcelFacade();
                excelFacade.Create<Package>(fileName, packages, "BIZTALK", headerNames);

                //DataSet ds = new DataSet() { EnforceConstraints = false };
                //ds.Tables.Add(dataTable);
                //var table = ds.Tables[0].Copy();
                ////packages = table;


                //SendTo(emailFrom, emailTo, emailSubject, smtpPort, smtpServer, emailPwd, fileName, emailBody);
                Console.WriteLine("Completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.Read();


        }

        private static void SendTo(string emailFrom, string emailTo, string emailSubject, string smtpPort, string smtpServer, string emailPwd, string sFile, string emailBody)
        {
            SmtpClient SmtpServer = new SmtpClient(smtpServer);
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(emailFrom);
            foreach (var recipient in emailTo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
            {
                mail.To.Add(recipient);
            }

            mail.Subject = emailSubject;
            mail.Body = emailBody;

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(sFile);
            mail.Attachments.Add(attachment);

            SmtpServer.Port = System.Convert.ToInt16(smtpPort);
            SmtpServer.Credentials = new System.Net.NetworkCredential(emailFrom, emailPwd);
            SmtpServer.EnableSsl = true;
            SmtpServer.Timeout = 7200000;
            SmtpServer.Send(mail);
        }
    }
    public class Package
    {
        public string Title { get; set; }
        public string DateTime { get; set; }
        public int SerialNumber { get; set; }

        public string MessageName { get; set; }

        public string ApplicationName { get; set; }
        public string ServerLocation { get; set; }

        public string DirectionName { get; set; }
        public int filesCount { get; set; }
        public int failedFileCount { get; set; }


    }

}
