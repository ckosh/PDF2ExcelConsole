﻿using GrabNadlanLicense;
using OpenPop.Mime;
using OpenPop.Pop3;
using PDF2excelConsole.Properties;
using PDF2ExcelVsto;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using static GrabNadlanLicense.ClassMongoDBPDF;

namespace PDF2excelConsole
{
    internal class ClassMailManager
    {
        public Pop3Client objPop3Client;
        public string host;
        public string user;
        public string userdebug;
        public string password;
        public string passworddebug;
        public int port;
        bool useSsl;
        bool debugMode;
        string[] PdfFileNames;
        string TempFolder;
        public string customerMail;
        int customertype;
        public int NumberOfPDFFiles;
        public int numberOfOwners;
        public List<int> listofOwners = new List<int>();

        static Dictionary<string, DateTime> messageTime = new Dictionary<string, DateTime>();

        public ClassMailManager(bool debug, int PopType)
        {
            objPop3Client = new Pop3Client();
            host = "mail.grabnadlan.co.il";
            user = "tabu2excel@grabnadlan.co.il";
            userdebug = "chaim.koshizky@grabnadlan.co.il";
            passworddebug = "3O5!RvHQ6Y5q";
            password = "L#(Bcw^Wi7{7";
            port = 995;
            useSsl = true;
            debugMode = debug;
            if ( PopType == 0)
            {
                port = 110;
                useSsl = false;
            }
            else if ( PopType == 1)
            {
                port = 995;
                useSsl = true;
            }

        }


        public void connectToPop()
        {
            try
            {
                objPop3Client.Connect(host, port, useSsl);
                objPop3Client.Authenticate(user, password);
            }
            catch (Exception e)
            {
            }
        }
        public int getNumberOfmessages()
        {

            string localUser = "";
            string localPassword = "";
            if (debugMode)
            {
                localUser = userdebug;
                localPassword = passworddebug;
            }
            else
            {
                localUser = user;
                localPassword = password;
            }

            //            Log.Info("get number of messages ");
            

            objPop3Client.Connect(host, port, useSsl);
            objPop3Client.Authenticate(localUser, localPassword);


            int j = objPop3Client.GetMessageCount();
            objPop3Client.Disconnect();
            objPop3Client.Dispose();
//            Log.Info("after  getNumberOfmessages " + j.ToString());
            return j;
        }

        public int checkMailContent(int firstMail, string tempFolder, bool debugMode, int obsticalMinutes)
        {
            int iret = 0;
            string resultExcelFile;
            TempFolder = tempFolder;
            bool DebugMode = debugMode;
            MessageModel message = new MessageModel();
            message = GetEmailContent(firstMail, DebugMode);

            if (!removeObstical(message, obsticalMinutes)) //  over 30 minutes in que - remove
            {
                deleteMail(firstMail, DebugMode);
                string sub = "ארעה שגיאה בניתוח הנסחים - שלח את הקבצים שנית ל - grabnadlan@gmail.com";
                sendMail(message.FromID.ToLower(), sub, null, sub);
                return 0;
            }

            if (message.Subject == "Registration")
            {
                var text = message.Body;
                bool bret;
                string[] stringSeparators = new string[] { "\n" };
                string[] lines = text.Split(stringSeparators, StringSplitOptions.None);
                string mail = getRequestedEmail("E_Mail:", lines).ToLower();
                mail = mail.Trim();
                bret = CheckUserRegistered(mail);
                if (bret)
                {
                    string sub = "שגיאת רישום - משתמש כבר רשום ";
                    sendMail(mail, sub, null, sub);
                }
                else
                {
                    bool bret0 = ClassUtils.verifyMailAddress(mail);
                    if (bret0)
                    {
                        PDFCustomers cust = new PDFCustomers();
                        cust.Mail = mail;
                        cust.Office_Name = getRequestedEmail("Office:", lines).Trim();
                        cust.Phone = getRequestedEmail("Phone:", lines).Trim();
                        cust.customerStatus = 0;
                        cust.User_Name = getRequestedEmail("Name:", lines).Trim();

                        bool ret = false;
                        Vba2VSTO vba2VSTO = new Vba2VSTO();
                        ret = vba2VSTO.SavePDFCustomerToDB(cust);
                        if (ret)
                        {
                            customerMail = message.FromID;
                            string sub = "אישור רישום לשרות הסבת נסחי טאבו לאקסל";
                            ConfirmationMail(mail, sub);
                            deleteAllFilesFromDirectory(TempFolder);
                            string body0 = cust.Mail + Environment.NewLine;
                            body0 = body0 + cust.Office_Name + Environment.NewLine;
                            body0 = body0 + cust.Phone;

                            ConfirmationMailSimple("grabnadlan@gmail.com", sub, body0);

                        }
                    }
                    else
                    {
                        ConfirmationMailSimple("grabnadlan@gmail.com", "שגיאת כתובת מייל", mail);
                    }
                }
                deleteMail(firstMail, DebugMode);
                return 0;
            }
            customerMail = message.FromID.ToLower();
            customertype = GetCustomerType(customerMail);
            if (!CheckUserRegistered(customerMail))
            {
                ///  return mail non registered
                ///  delete mail 
                string sub = "משתמש אינו רשום במערכת";
                sendMail(customerMail, sub, null, sub);
                deleteMail(firstMail, DebugMode);
                deleteAllFilesFromDirectory(TempFolder);
                return 0;
            }
            if (!CheckUserPermission(customerMail))
            {
                ///  return mail non registered
                ///  delete mail 
                string sub = "משתמש אינו מאושר לשימוש";
                sendMail(customerMail, sub, null, sub);
                deleteMail(firstMail, DebugMode);
                deleteAllFilesFromDirectory(TempFolder);
                return 0;
            }
            NumberOfPDFFiles = 0;
            numberOfOwners = 0;
            if (message.Attachment != null)
            {
                NumberOfPDFFiles = message.Attachment.Count;
            }
            if (NumberOfPDFFiles == 0)
            {
                string sub = "!מייל לא מכיל נסחים";
                sendMail(customerMail, sub, null, sub);
                deleteMail(firstMail, DebugMode);
                return 0;
            }

            ClassProcessPDF processPDF = new ClassProcessPDF(PdfFileNames, DebugMode, TempFolder);
//            Log.Info("after process files  ");
            resultExcelFile = processPDF.convert();
//            Log.Info("after convert files  ");
            listofOwners = processPDF.getTotalNumberOfOwners();
            numberOfOwners = listofOwners.Sum();
            double totalcost = getTotalCost(NumberOfPDFFiles, listofOwners);

            return iret;
        }
        public class MessageModel
        {
            public string MessageID;
            public string FromID;
            public string FromName;
            public string Subject;
            public string Body;
            public string Html;
            public string FileName;
            public List<MessagePart> Attachment;
        }

        public MessageModel GetEmailContent(int intMessageNumber, bool debug)
        {
            objPop3Client.Connect(host, port, useSsl);
            string localUser = "";
            string localPassword = "";
            if (debug)
            {
                localUser = userdebug;
                localPassword = passworddebug;
            }
            else
            {
                localUser = user;
                localPassword = password;
            }
            objPop3Client.Authenticate(localUser, localPassword);

            MessageModel message = new MessageModel();
            OpenPop.Mime.Message objMessage;
            MessagePart plainTextPart = null, HTMLTextPart = null;
            objMessage = objPop3Client.GetMessage(intMessageNumber);
            message.MessageID = objMessage.Headers.MessageId == null ? "" : objMessage.Headers.MessageId.Trim();
            message.FromID = objMessage.Headers.From.Address.Trim();
            message.FromName = objMessage.Headers.From.DisplayName.Trim();
            message.Subject = objMessage.Headers.Subject.Trim();

            plainTextPart = objMessage.FindFirstPlainTextVersion();
            message.Body = (plainTextPart == null ? "" : plainTextPart.GetBodyAsText().Trim());

            HTMLTextPart = objMessage.FindFirstHtmlVersion();
            message.Html = (HTMLTextPart == null ? "" : HTMLTextPart.GetBodyAsText().Trim());


            List<MessagePart> attachment = objMessage.FindAllAttachments();
            if (attachment.Count > 0)
            {
                PdfFileNames = new string[attachment.Count];
                for (int j = 0; j < attachment.Count; j++)
                {
                    byte[] content = attachment[j].Body;
                    string[] stringParts = attachment[j].FileName.Split(new char[] { '.' });
                    string stringType = stringParts[1];
                    BinaryWriter Writer = null;
                    string Name = TempFolder + "\\" + attachment[j].FileName;
                    PdfFileNames[j] = attachment[j].FileName;
                    try
                    {
                        // Create a new stream to write to the file
                        Writer = new BinaryWriter(File.OpenWrite(Name));

                        // Writer raw data                
                        Writer.Write(content);
                        Writer.Flush();
                        Writer.Close();
                    }
                    catch
                    {
                    }

                }
                message.FileName = attachment[0].FileName.Trim();
                message.Attachment = attachment;

            }
            objPop3Client.Disconnect();
            return message;
        }

        private bool removeObstical(MessageModel msg, int obsticalMinutes)
        {
            bool bret = true;
            string mesID = msg.MessageID;
            DateTime dateTime = DateTime.Now;

            string key;
            DateTime tim;


            if (messageTime.Count == 0)
            {
//                Log.Info("Message time stamp " + msg.MessageID + " " + dateTime.ToString());
                messageTime.Add(msg.MessageID, dateTime);
                return bret; ;
            }
            else
            {
                var first = messageTime.First();
                key = first.Key;
                tim = first.Value;
            }

            if (key == mesID)
            {
                var diff = DateTime.Now - tim;
                if (diff.Minutes > obsticalMinutes)
                {
//                    Log.Info("Message check time difference " + msg.MessageID + " " + diff.Minutes.ToString());
                    messageTime.Clear();
                    bret = false;
                    return bret;
                }
            }
            return bret;
        }
        public void deleteMail(int j, bool debug)
        {
            string localUser = "";
            string localPassword = "";
            if (debug)
            {
                localUser = userdebug;
                localPassword = passworddebug;
            }
            else
            {
                localUser = user;
                localPassword = password;
            }
            objPop3Client.Connect(host, port, useSsl);
            objPop3Client.Authenticate(localUser, localPassword);
            objPop3Client.DeleteMessage(j);
            objPop3Client.Disconnect();
        }
        public void sendMail(string to, string subject, string filePath, string Body)
        {
            MailMessage message = new MailMessage(user, to);
            message.Subject = subject;
            message.Body = Body;

            System.Net.Mail.Attachment data;
            if (filePath != null)
            {
                data = new System.Net.Mail.Attachment(filePath); //, MediaTypeNames.Application.Octet
                message.Attachments.Add(data);
            }

            string localUser = "";
            string localPassword = "";
            if (debugMode)
            {
                localUser = userdebug;
                localPassword = passworddebug;
            }
            else
            {
                localUser = user;
                localPassword = password;
            }
            SmtpClient client = new SmtpClient(host);
            client.Credentials = new System.Net.NetworkCredential(localUser, localPassword);
            try
            {
                client.Send(message);
            }
            catch (Exception e)
            {
            }
            message.Dispose();
            client.Dispose();

        }
        public string getRequestedEmail(string param, string[] lines)
        {
            string reqString = "";
            int paramLenth = param.Length;
            foreach (string s in lines)
            {
                int pos = s.IndexOf(param, 0);
                if (pos > -1)
                {
                    reqString = s.Substring(paramLenth);
                    reqString = reqString.Replace("\r", string.Empty);
                    break;
                }
            }
            return reqString;
        }
        public bool CheckUserRegistered(string userMail)
        {
            bool ret = false;
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            ret = vba2VSTO.GetRegisteredCustomerFromMongoDBPDF(userMail);

            return ret;
        }
        public void ConfirmationMail(string to, string subject)
        {
            MailMessage message = new MailMessage(user, to);
            message.Subject = subject;
            message.IsBodyHtml = true;
            ResourceManager rm = Resources.ResourceManager;
            Bitmap myImage = (Bitmap)rm.GetObject("registrationReply");
            var filePath = TempFolder + "\\registrationReply.png";
            myImage.Save(filePath);
            var inlineLogo = new LinkedResource(filePath, "image/png");
            inlineLogo.ContentId = Guid.NewGuid().ToString();
            string body = string.Format(@"<img src= ""cid:{0}"" />", inlineLogo.ContentId);
            var view = AlternateView.CreateAlternateViewFromString(body, null, "text/html");
            view.LinkedResources.Add(inlineLogo);
            message.AlternateViews.Add(view);


            string localUser = "";
            string localPassword = "";
            if (debugMode)
            {
                localUser = userdebug;
                localPassword = passworddebug;
            }
            else
            {
                localUser = user;
                localPassword = password;
            }
            SmtpClient client = new SmtpClient(host);
            client.Credentials = new System.Net.NetworkCredential(localUser, localPassword);
            try
            {
                client.Send(message);
            }
            catch (Exception e)
            {

            }
            message.Dispose();
            client.Dispose();
        }
        public void deleteAllFilesFromDirectory(string directory)
        {
            DirectoryInfo di = new DirectoryInfo(directory);
            FileInfo[] files = di.GetFiles();
            foreach (FileInfo file in files)
            {
                file.Delete();
            }
        }
        public void ConfirmationMailSimple(string to, string subject, string body)
        {
            MailMessage message = new MailMessage(user, to);
            message.Subject = subject;
            message.IsBodyHtml = false;
            message.Body = body;
            SmtpClient client = new SmtpClient(host);
            client.Credentials = new System.Net.NetworkCredential("tabu2excel@grabnadlan.co.il", "L#(Bcw^Wi7{7");
            try
            {
                client.Send(message);
            }
            catch (Exception e)
            {

            }
            message.Dispose();
            client.Dispose();
        }
        public int GetCustomerType(string customerMail)
        {
            int iret = 0;
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            iret = vba2VSTO.GetDBPDFCustomerType(customerMail);
            return iret;
        }
        public bool CheckUserPermission(string userMail)
        {
            bool ret = false;
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            ret = vba2VSTO.GetRegisteredPermissionFromMongoDBPDF(userMail);
            return ret;
        }
        public double getTotalCost(int NumberOfPDFFiles, List<int> numberOfOwners)
        {
            double totalCost = 0;
            for (int i = 0; i < NumberOfPDFFiles; i++)
            {
                totalCost = totalCost + 5.0;
                if (numberOfOwners[i] > 25)
                {
                    totalCost = totalCost + (numberOfOwners[i] - 25) * 0.2;
                }
            }
            return totalCost;

        }
    }
}