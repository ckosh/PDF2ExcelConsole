using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        bool useSsl;
        bool debugMode;
        string[] PdfFileNames;
        string TempFolder;
        public string customerMail;
        int customertype;
        public int NumberOfPDFFiles;
        public int numberOfOwners;
        public string localUser ;
        public string localPassword;
        public List<int> listofOwners = new List<int>();

        static Dictionary<string, DateTime> messageTime = new Dictionary<string, DateTime>();

        public ClassMailManager1(bool debug, int PopType)
        {
            localUser = "";
            localPassword = "";
            objPop3Client = new Pop3Client();
            host = "mail.grabnadlan.co.il";
            user = "tabu2excel@grabnadlan.co.il";
            passworddebug = "3O5!RvHQ6Y5q";
            password = "L#(Bcw^Wi7{7";
            port = 995;
            useSsl = true;
            debugMode = debug;
        }

        public bool fetchMailContent()
        {
            bool fret = false;

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

            objPop3Client.Connect(host, port, useSsl);
            objPop3Client.Authenticate(localUser, localPassword);

            int numOfMessages = objPop3Client.GetMessageCount();

            if ( numOfMessages > 0)
            {
                MessageModel1 message = new MessageModel1();
                OpenPop.Mime.Message objMessage;
                MessagePart plainTextPart = null, HTMLTextPart = null;
                objMessage = objPop3Client.GetMessage(1);
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


            }

            return fret;
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

    }
}

