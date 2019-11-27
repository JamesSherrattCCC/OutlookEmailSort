using EmailHandler.DataTypes;
using EmailToJobDB.EmailDatabase;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailToJobDB
{
    public partial class Email2JobDB
    {
        private EmailRetriever _retriever;

        void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _retriever = new EmailRetriever(Application);
        }



        void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            _retriever.CloseDBConnection();
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
