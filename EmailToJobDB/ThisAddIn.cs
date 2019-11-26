using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using EmailToJobDB.EmailDatabase;
using EmailHandler.DataTypes;

namespace EmailToJobDB
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        EmailContext ctx;
        string user;
        IEnumerable<Job> jobs;


        void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Initialise();
            GetFirstEmails();
            SetupHandler();
        }

        private void Initialise()
        {
            ctx = new EmailContext();
            var user = this.Application.ActiveExplorer().Session.CurrentUser.Name;
            jobs = ctx.Jobs;
        }

        private void GetFirstEmails()
        {
            Outlook.MAPIFolder inbox = this.Application.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.Items unreadItems = inbox.Items.Restrict("[Unread]=true");
            if (unreadItems.Count == 0)
            {
                return;
            }
            MessageBox.Show(
                string.Format("Unread items in Inbox = {0}", unreadItems.Count));

            foreach (object Item in unreadItems)
            {
                Items_ItemAdd(Item);
            }
            ctx.SaveChanges();
            // There's a weird bug which causes some emails to initially not save, so this
            // function has to be recalled until all the emails are loaded.
            GetFirstEmails();
        }

        private void SetupHandler()
        {
            // New message received handler.
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAddSave);
        }

        void Items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                Email email = new Email
                {
                    User = user,
                    Body = mail.Body,
                    Sender = mail.Sender.Address,
                    Subject = mail.Subject,
                    DateRetrieved = mail.ReceivedTime,
                };

                foreach (Job job in jobs)
                {
                    if (mail.MessageClass == "IPM.Note" && mail.Subject.ToUpper().Contains(job.JobName.ToUpper()))
                    {
                        email.Job = job;
                    }
                    mail.UnRead = false;
                }
                ctx.Emails.Add(email);
            }
        }

        void Items_ItemAddSave(object Item)
        {
            Items_ItemAdd(Item);
            ctx.SaveChanges();
        }

        void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            ctx.Dispose();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
