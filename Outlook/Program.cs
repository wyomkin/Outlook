using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Reflection;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook1
{
    class Mail1
    {
        public string Topic { get; set; }
        public string sender { get; set; }
        public string senderEmailAdr { get; set; }
        public DateTime sendDate { get; set; }


        public List<Mail1> answeredByList = new List<Mail1>();


    }
    class Program
    {
        static string getSenderEmailAddress(Outlook.MailItem mail)
        {
            Outlook.AddressEntry sender = mail.Sender;
            string SenderEmailAddress = "";

            if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry || sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                Outlook.ExchangeUser exchUser = sender.GetExchangeUser();
                if (exchUser != null)
                {
                    SenderEmailAddress = exchUser.PrimarySmtpAddress;
                }
            }
            else
            {
                SenderEmailAddress = mail.SenderEmailAddress;
            }

            return SenderEmailAddress;
        }
        static Outlook.Application GetApplicationObject()
        {

            Outlook.Application application = null;

            // Check whether there is an Outlook process running.
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {

                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            else
            {

                // If not, create a new instance of Outlook and sign in to the default profile.
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "", Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;
        }
        static List<Mail1> checkMailsInFolderForAnswer(string folder)
        {
            Outlook.Folder defFold = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            Outlook.Folder tempFold = null;
            foreach (Outlook.Folder subfolder in defFold.Folders)
            {
                if (subfolder.Name == folder)
                {
                    tempFold = subfolder;
                    break;
                }

            }

            string filter = DateTime.Now.AddHours(-DateTime.Now.Hour).AddMinutes(-DateTime.Now.Minute).AddHours(-3).ToString("dd/MM/yyyy HH:mm").Replace(".", "/");
            Outlook.Items items = tempFold.Items.Restrict($"[CreationTime]>'{filter}'");


            List<Mail1> mails1TempList = new List<Mail1>();
            foreach (Outlook.MailItem mail in items)
            {
                Mail1 mailTemp = new Mail1();
                mailTemp.sender = mail.SenderName;
                mailTemp.Topic = mail.ConversationTopic;
                mailTemp.senderEmailAdr = getSenderEmailAddress(mail);
                mailTemp.sendDate = mail.CreationTime;

                Outlook.Conversation tc = mail.GetConversation();
                Outlook.Table table1 = tc.GetTable();
                table1.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x5D0A001F"); //email от кого письмо
                table1.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x3FF8001F"); //От кого письмо
                //Console.WriteLine(table1.GetRowCount());
                if (table1.GetRowCount() > 1)
                {
                    List<Mail1> listOfMembers = new List<Mail1>();
                    while (!table1.EndOfTable)
                    {
                        Mail1 t = new Mail1();
                        Outlook.Row row = table1.GetNextRow();
                        t.sendDate = row["CreationTime"];
                        t.sender = row["http://schemas.microsoft.com/mapi/proptag/0x3FF8001F"];
                        t.senderEmailAdr = row["http://schemas.microsoft.com/mapi/proptag/0x5D0A001F"];
                        t.sender = row["http://schemas.microsoft.com/mapi/proptag/0x3FF8001F"];

                        if (needToBesnaweredByThisList.Contains(t.senderEmailAdr))
                        {
                            if (t.sender != mailTemp.sender)
                            {
                                mailTemp.answeredByList.Add(t);
                            }


                        }
                    }

                    mails1TempList.Add(mailTemp);

                }
                else { mails1TempList.Add(mailTemp); }
            }

            return mails1TempList;
            #region commented
            // For this example, you will work only with 
            //MailItem. Other item types such as
            //MeetingItem and PostItem can participate 
            //in Conversation.
            //if (selectedItem is Outlook.MailItem)
            //{
            //    // Cast selectedItem to MailItem.
            //    Outlook.MailItem mailItem =
            //        selectedItem as Outlook.MailItem; ;
            //    // Determine store of mailItem.
            //    Outlook.Folder folder = mailItem.Parent
            //        as Outlook.Folder;
            //    Outlook.Store store = folder.Store;
            //    if (store.IsConversationEnabled == true)
            //    {
            //        // Obtain a Conversation object.
            //        Outlook.Conversation conv =
            //            mailItem.GetConversation();
            //        // Check for null Conversation.
            //        if (conv != null)
            //        {
            //            // Obtain Table that contains rows 
            //            // for each item in Conversation.
            //            Outlook.Table table = conv.GetTable();
            //            Debug.WriteLine("Conversation Items Count: " +
            //                table.GetRowCount().ToString());
            //            Debug.WriteLine("Conversation Items from Table:");
            //            while (!table.EndOfTable)
            //            {
            //                Outlook.Row nextRow = table.GetNextRow();
            //                Console.WriteLine(nextRow["Subject"]
            //                    + " Modified: "
            //                    + nextRow["LastModificationTime"]);
            //            }
            //            Debug.WriteLine("Conversation Items from Root:");
            //            // Obtain root items and enumerate Conversation.
            //            Outlook.SimpleItems simpleItems
            //                = conv.GetRootItems();
            //            foreach (object item in simpleItems)
            //            {
            //                // In this example, enumerate only MailItem type.
            //                // Other types such as PostItem or MeetingItem
            //                // can appear in Conversation.
            //                if (item is Outlook.MailItem)
            //                {
            //                    Outlook.MailItem mail = item
            //                        as Outlook.MailItem;
            //                    Outlook.Folder inFolder =
            //                        mail.Parent as Outlook.Folder;
            //                    string msg = mail.Subject
            //                        + " in folder " + inFolder.Name;
            //                    Debug.WriteLine(msg);
            //                }
            //                // Call EnumerateConversation 
            //                // to access child nodes of root items.
            //                EnumerateConversation(item, conv);
            //            }
            //        }
            //    }
            //}
            #endregion
        }

        static void EnumerateConversation(object item, Outlook.Conversation conversation)
        {
            Outlook.SimpleItems items = conversation.GetChildren(item);
            if (items.Count > 0)
            {
                foreach (object myItem in items)
                {
                    // In this example, enumerate only MailItem type.
                    // Other types such as PostItem or MeetingItem
                    // can appear in Conversation.
                    if (myItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                            myItem as Outlook.MailItem;
                        Outlook.Folder inFolder =
                            mailItem.Parent as Outlook.Folder;
                        string msg = mailItem.Subject
                            + " in folder " + inFolder.Name;
                        Debug.WriteLine(msg);
                    }
                    // Continue recursion.
                    EnumerateConversation(myItem, conversation);
                }
            }
        }


        static Outlook.Application app;
        static Outlook.NameSpace ns;

        static List<string> foldersToCheckList;
        static List<string> needToBesnaweredByThisList;

        static void checkTask(string connectionString)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string qstring = String.Format("Delete from dbo.outlook_check ");
                    Console.WriteLine($"{DateTime.Now}: {qstring}");
                    SqlCommand command = new SqlCommand(qstring, connection);
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception e) { Console.WriteLine($"{DateTime.Now }: БДшка в отвале :< "); }




            foreach (string folder in foldersToCheckList)
            {
                List<Mail1> templist = checkMailsInFolderForAnswer(folder);
                foreach (Mail1 mail in templist)
                {
                    if (mail.answeredByList.Count == 0)
                    {
                        if (DateTime.Now - mail.sendDate > TimeSpan.FromMinutes(5))
                        {
                            try
                            {
                                using (SqlConnection connection = new SqlConnection(connectionString))
                                {
                                    connection.Open();
                                    string qstring = String.Format("INSERT into dbo.outlook_check VALUES('{0}','{1}','{2}','{3}')", DateTime.Now, mail.Topic, mail.sendDate, mail.sender);
                                    Console.WriteLine($"{DateTime.Now}: {qstring}");
                                    SqlCommand command = new SqlCommand(qstring, connection);
                                    command.ExecuteNonQuery();
                                }
                            }
                            catch (Exception e) { Console.WriteLine($"{DateTime.Now }: БДшка в отвале :< "); }
                        }
                    }
                   
                }
            }


        }
        static void Main(string[] args)
        {
            Console.Title = "Outlook Checker";
            Console.SetBufferSize(Console.BufferWidth, 1000);

            string connectionString = "Data Source=(local);Initial Catalog=Infinity;" + "Integrated Security=true";
            //заполняем лист папок для поиска
            needToBesnaweredByThisList = new List<string>{ "a.gzogyan@leadtop.org", "g.dorin@leadtop.org", "d.dumchikov@leadtop.org", "kadorozhkin@leadtop.org", "eoduyzarov@leadtop.org",
            "y.evdokimova@leadtop.org","g.klimov@leadtop.org","r.konovalov@leadtop.org","dakrasnogrudskiy@leadtop.org","aamkrtchyan@leadtop.org","dnpavlyuchenkov@leadtop.org","m.sivtsov@leadtop.org",
            "aasolovyanov@leadtop.org","a.tserkunov@leadtop.org","aachernogurskikh@leadtop.org"};
            foldersToCheckList = new List<string> { "Халакоева Марьяна", "Урюпин" };

            //Находим процесс и неймспей
            app = Program.GetApplicationObject();
            ns = app.GetNamespace("MAPI");

            var timer = new System.Threading.Timer((e) => { checkTask(connectionString); }, null, TimeSpan.Zero, TimeSpan.FromMinutes(5));


            Console.ReadLine();

        }
    }
}
