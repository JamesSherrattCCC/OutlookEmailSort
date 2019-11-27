using EmailHandler.DataTypes;
using EmailToJobDB.EmailDatabase;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace EmailToJobDB
{
    class EmailRetriever
    {
        private Outlook.Application _application;
        private Outlook.NameSpace _outlookNameSpace;
        private Outlook.MAPIFolder _inbox;
        private Outlook.Items _items;
        private EmailContext _ctx;
        private string _user;
        private IEnumerable<Job> jobs;

        public EmailRetriever(Outlook.Application application)
        {
            _application = application;
            SetupApplication();

        } 

        public void CloseDBConnection()
        {
            _ctx.Dispose();
        }

        private void SetupApplication()
        {
            Initialise();
            GetFirstEmails();
            SetupHandler();
        }

        private void Initialise()
        {
            _ctx = new EmailContext();
            _user = _application.ActiveExplorer().Session.CurrentUser.Name;
            jobs = _ctx.Jobs;
        }

        private void GetFirstEmails()
        {
            Outlook.MAPIFolder inbox = this._application.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.Items unreadItems = inbox.Items.Restrict("[Unread]=true");
            MessageBox.Show(
                string.Format("Unread items in Inbox = {0}", unreadItems.Count));
            if (unreadItems.Count == 0)
            {
                return;
            }

            foreach (object Item in unreadItems)
            {
                Items_ItemAdd(Item);
            }
            _ctx.SaveChanges();
            // There's a weird bug which causes some emails to initially not save or reload as null, so this
            // function has to be recalled until all the emails are loaded.
            GetFirstEmails();
        }

        private void SetupHandler()
        {
            // New message received handler.
            _outlookNameSpace = _application.GetNamespace("MAPI");
            _inbox = _outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            _items = _inbox.Items;
            _items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAddSave);
        }

        void Items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                Email email = mail.ToEmail();
                email.User = _user;

                foreach (Job job in jobs)
                {
                    if (mail.MessageClass == "IPM.Note" && mail.Subject.ToUpper().Contains(job.JobName.ToUpper()))
                    {
                        email.Job = job;
                    }
                    mail.UnRead = false;
                }
                _ctx.Emails.Add(email);
            }
        }

        void Items_ItemAddSave(object Item)
        {
            Items_ItemAdd(Item);
            _ctx.SaveChanges();
        }
    }
}
