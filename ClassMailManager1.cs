using GrabNadlanLicense;
using OpenPop.Mime;
using OpenPop.Pop3;
using PDF2excelConsole.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using static GrabNadlanLicense.ClassMongoDBPDF;

namespace PDF2excelConsole
{
    internal class ClassMailManager1
    {
        public Pop3Client objPop3Client;
        public string host;
        public string user;
        public string userdebug;
        public string password;
        public string passworddebug;
        public int port;
        private bool portOpen;
        private int numberOfMails;
        bool useSsl;
        string[] PdfFileNames;
        public string localUser;
        public string localPassword;
        public List<int> listofOwners = new List<int>();
        List<string> oldMailID;
        MessageModel1 message;
        OpenPop.Mime.Message objMessage;
        string tempFolder;
        bool debugMode;
        int messageItem;

        int NumberOfPDFFiles = 0;
        int numberOfOwners = 0;
        string userMail = "";

        static Dictionary<string, DateTime> messageTime = new Dictionary<string, DateTime>();

        public ClassMailManager1(bool debug, int Portnumber, bool ssl, string temp)
        {
            localUser = "";
            localPassword = "";
            user = "tabu2excel@grabnadlan.co.il";
            userdebug = "chaim.koshizky@grabnadlan.co.il";
            passworddebug = "T3mp22dJCg21dC";
            password = "L#(Bcw^Wi7{7";
            portOpen = false;
            numberOfMails = 0;
            tempFolder = temp;
            message = new MessageModel1();
            debugMode = debug;
            oldMailID = new List<string>();

            host = "mail.grabnadlan.co.il";

            port = Portnumber;
            useSsl = ssl;

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
        }

        public string ConnectToPop3()
        {
            userMail = "";
            objPop3Client = new Pop3Client();
            
            try
            {
                objPop3Client.Connect(host, port, useSsl);
                objPop3Client.Authenticate(localUser, localPassword);
            }
            catch(Exception ex)
            {
                return userMail;
            }



            if (oldMailID.Count > 0)
            {
                for (int i = 0; i < oldMailID.Count; i++)
                {
                    if ( DeleteMessageByMessageId(oldMailID[i]) )
                    {
                        Console.WriteLine("delete last error deletion " + i.ToString() );
                    }
                }
            }


            numberOfMails = objPop3Client.GetMessageCount();
            portOpen = true;
                

            if (numberOfMails > 0)
            {
                /// read mail content
                for ( messageItem = 1; messageItem <= numberOfMails;  messageItem++)
                {
                    if (oldMailID.Count > 0  && isOldMessage(messageItem))
                    {
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
                objMessage = objPop3Client.GetMessage(messageItem);
                message.MessageID = objMessage.Headers.MessageId == null ? "" : objMessage.Headers.MessageId.Trim();
                message.FromID = objMessage.Headers.From.Address.Trim().ToLower();
                message.FromName = objMessage.Headers.From.DisplayName.Trim();
                message.Subject = objMessage.Headers.Subject.Trim();
                MessagePart plainTextPart = null, HTMLTextPart = null;
                plainTextPart = objMessage.FindFirstPlainTextVersion();
                message.Body = (plainTextPart == null ? "" : plainTextPart.GetBodyAsText().Trim());
                HTMLTextPart = objMessage.FindFirstHtmlVersion();
                message.Html = (HTMLTextPart == null ? "" : HTMLTextPart.GetBodyAsText().Trim());
                List<MessagePart> att = objMessage.FindAllAttachments();
                List<MessagePart> attachment = new List<MessagePart>();
                for ( int j = 0; j < att.Count; j++)
                {
                    if (  Path.GetExtension(att[j].FileName).ToLower() == ".pdf" )
                    {
                        attachment.Add(att[j]);
                    }
                }

                userMail = message.FromID;

                if (attachment.Count > 0)
                {
                    PdfFileNames = new string[attachment.Count];
                    for (int j = 0; j < attachment.Count; j++)
                    {
                        byte[] content = attachment[j].Body;
                        string[] stringParts = attachment[j].FileName.Split(new char[] { '.' });
                        string stringType = stringParts[1];
                        BinaryWriter Writer = null;
                        string Name = tempFolder + "\\" + attachment[j].FileName;
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
                /// parse content
                if (message.Subject == "Registration")
                {
                    if (!RegisterANewUser())
                    {
                        userMail = "0";
                    }
                }
                else
                {
                    if (!CheckUserRegistered())
                    {
                        quitWithnote(userMail, "משתמש אינו רשום לשימוש");
                    }
                    if (!CheckUserPermission())
                    {
                        quitWithnote(userMail, "משתמש אינו מאושר לשימוש");
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
                        sendMail(userMail, sub, null, sub);
                        userMail = "0";
                    }

                }
            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ClosePop3(userMail);
           
            return userMail;
        }

        public void ClosePop3(string userMail)
        {
            if ( userMail != null && userMail.Length > 0 )
            {
                try
                {
                    if (objPop3Client.GetMessageCount() > 0)
                    {
                        deleteMail(messageItem);
                    }
                    
                }
                catch(Exception e)
                {
                    oldMailID.Add(objPop3Client.GetMessageHeaders(messageItem).MessageId);
                } 
                // try delete old mail
                //if ( oldMailID.Count > 0)
                //{
                //    for ( int i = 0; i < oldMailID.Count; i++)
                //    {
                //        DeleteMessageByMessageId(oldMailID[i]);
                //    }
                    
                //}

                objPop3Client.Disconnect();
                objPop3Client.Dispose();
                portOpen = false;

            }
        }

        public int getNumberOfMails()
        {
            int iret = 0;
            if (portOpen)
            {
                iret = numberOfMails;
            }
            return iret;
        }
        public string getSubject()
        {
            string ret = "";

            if (!message.Equals(null))
            {
                ret = message.Subject;
            }
            return ret;
        }


        public bool RegisterANewUser()
        {
            bool bret = false;
            var text = message.Body;
            string[] stringSeparators = new string[] { "\n" };
            string[] lines = text.Split(stringSeparators, StringSplitOptions.None);
            string mail = getRequestedEmail("E_Mail:", lines).ToLower();
            mail = mail.Trim();
            bret = CheckUserRegistered(mail);
            if (bret)
            {
                string sub = "שגיאת רישום - משתמש כבר רשום ";
                bret = false;
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
                        string sub = "אישור רישום לשרות הסבת נסחי טאבו לאקסל";
                        ConfirmationMail(mail, sub);
                        deleteAllFilesFromDirectory(tempFolder);
                        string body0 = cust.Mail + Environment.NewLine;
                        body0 = body0 + cust.Office_Name + Environment.NewLine;
                        body0 = body0 + cust.Phone;
                        ConfirmationMailSimple("grabnadlan@gmail.com", sub, body0);
                        Console.WriteLine("new registration " + cust.Mail + " " + cust.Office_Name);

                    }
                }
                else
                {
                    ConfirmationMailSimple("grabnadlan@gmail.com", "שגיאת כתובת מייל", mail);
                }
            }
            return bret;
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
        public void deleteAllFilesFromDirectory(string directory)
        {
            DirectoryInfo di = new DirectoryInfo(directory);
            FileInfo[] files = di.GetFiles();
            foreach (FileInfo file in files)
            {
                file.Delete();
            }
        }

        public bool CheckUserPermission()
        {
            bool ret = false;
            string ss = message.FromID;
            ss = ss.ToLower();
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            ret = vba2VSTO.GetRegisteredPermissionFromMongoDBPDF(ss);
            return ret;
        }
        public bool CheckUserRegistered()
        {
            bool ret = false;
            string ss = message.FromID;
            ss = ss.ToLower();

            Vba2VSTO vba2VSTO = new Vba2VSTO();
            ret = vba2VSTO.GetRegisteredCustomerFromMongoDBPDF(ss);

            return ret;
        }
        public void savecBillingData(string e_mail, int numberOfFiles, string excelFileName, int numOfOwners, double totalCost, string project)
        {
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            vba2VSTO.SaveBillingData(e_mail, numberOfFiles, excelFileName, numOfOwners, totalCost, project);
        }

        public void quitWithnote(string mailaddress, string note)
        {

            sendMail(mailaddress, note, null, note);
            deleteMail(messageItem);
            deleteAllFilesFromDirectory(tempFolder);
        }

        public void deleteMail(int num)
        {
            try
            {
                bool bb = false;
                if (bb) throw new Exception("ggg");
                objPop3Client.DeleteMessage(num);
            }
            catch(Exception ex)
            {
                oldMailID.Add(objPop3Client.GetMessageHeaders(num).MessageId);
                Console.WriteLine("Error message delete " + num.ToString() + " " + objPop3Client.GetMessageHeaders(num).MessageId );
            }
        }

        public class MessageModel1
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
        public void ConfirmationMail(string to, string subject)
        {
            MailMessage message = new MailMessage(user, to);
            message.Subject = subject;
            message.IsBodyHtml = true;
            ResourceManager rm = Resources.ResourceManager;
            Bitmap myImage = (Bitmap)rm.GetObject("registrationReply");
            var filePath = tempFolder + "\\registrationReply.png";
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
                Console.WriteLine("Error sending mail to " + to );
            }
            message.Dispose();
            client.Dispose();
        }
        public string[] GetPdfFiles()
        {
            return PdfFileNames;
        }

        public bool DeleteMessageByMessageId( string messageId)
        {
            int messageCount = objPop3Client.GetMessageCount();

            for (int messageItem = 1; messageItem <= messageCount; messageItem++)
            {
                if (objPop3Client.GetMessageHeaders(messageItem).MessageId == messageId)
                {
                    try
                    {
                        objPop3Client.DeleteMessage(messageItem);
                    }
                    catch(Exception e)
                    {
                        return false;
                    }
                    
                    return true;
                }
            }
            return false;
        }
        public bool isOldMessage( int messageItem)
        {
            bool bret = false;
            objMessage = objPop3Client.GetMessage(messageItem);
            for (int i = 0; i < oldMailID.Count; i++)
            {
                if (oldMailID[i].Equals(objPop3Client.GetMessageHeaders(messageItem).MessageId))
                {
                    bret = true;
                    return bret;
                }
            }
            
            return bret;
        }
        public string getProject()
        {
            string project = "";
            char[] charsToTrim = {' ' };
            string[] subs = message.Subject.Split('-');
            if (subs.Length > 1)
            {
                project = subs[1].Trim(charsToTrim);
            }
            return project;
        }
    }
}

