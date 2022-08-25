using System;
using System.Reflection;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace ReadingEmail
{
    public class Program
    {
        public static int Main(string[] args)
        {
            try
            {

                Outlook.Application oApp = new Outlook.Application();

                // Get the MAPI namespace.
                Outlook.NameSpace oNS = oApp.GetNamespace("MAPI");
                oNS.Logon("email address placeholder", "password placeholder", false, true);

                //Get the Inbox folder.
                Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                String user = oNS.CurrentUser.EntryID;

                //Get the Items collection in the Inbox folder.
                Outlook.Items oItems = oInbox.Items;

                Console.WriteLine("Total Emails are: " + oItems.Count);
                Outlook.MailItem oMsg = (Outlook.MailItem)oItems.GetFirst();

                Outlook.Items unreadItems = oInbox.Items.Restrict("[Unread]=true");

                int readItems = oItems.Count - unreadItems.Count;

                //Console.WriteLine(oMsg.Subject);
                //Console.WriteLine(oMsg.SenderName);
                //Console.WriteLine(oMsg.ReceivedTime);
                //Console.WriteLine(oMsg.Body);

                //int AttachCnt = oMsg.Attachments.Count;
                //Console.WriteLine("Attachments: " + AttachCnt.ToString());

                //if (AttachCnt > 0)
                //{
                //    for (int i = 1; i <= AttachCnt; i++)
                //        Console.WriteLine(i.ToString() + "-" + oMsg.Attachments[i].DisplayName);
                //}
                //
                //for (int i = 0; i < oItems.Count; i++)
                //{
                //    if (oItems.GetNext() is Outlook.MailItem)
                //    {
                //        oMsg = (Outlook.MailItem)oItems.GetNext();
                //        Console.WriteLine(oMsg.Subject);
                //        Console.WriteLine(oMsg.SenderName);
                //        Console.WriteLine(oMsg.ReceivedTime);
                //        Console.WriteLine(oMsg.Body);
                //        AttachCnt = oMsg.Attachments.Count;
                //        if (AttachCnt > 0)
                //        {
                //            for (int j = 1; j <= AttachCnt; j++)
                //                Console.WriteLine(j.ToString() + "-" + oMsg.Attachments[j].DisplayName);
                //        }
                //
                //        Console.WriteLine("Attachments: " + AttachCnt.ToString());
                //        Console.WriteLine("CURRENT EMAIL # IS: " + i);
                //    }
                //    else
                //    {
                //        oItems.GetNext();
                //        Console.WriteLine("NOT AN EMAIL");
                //        Console.WriteLine("CURRENT EMAIL # IS: " + i);
                //    }
                //
                //
                //}

                Console.WriteLine(string.Format("Read Emails in Inbox = {0}", readItems));
                Console.WriteLine(string.Format("Unread Emails in Inbox = {0}", unreadItems.Count));

                oNS.Logoff();

                oMsg = null;
                oItems = null;
                oInbox = null;
                oNS = null;
                oApp = null;
            }

            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught: ", e);


            }

            return 0;

        }
    }
}