using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using System;

namespace BCC_myself_OutlookAddIn {
    public partial class ThisAddIn {
        Outlook.MailItem myMail;

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            Application.ItemLoad += new Outlook.ApplicationEvents_11_ItemLoadEventHandler(Application_ItemLoad);
        }

        public void Application_ItemLoad(object Item) {
            if (Item is Outlook.MailItem) {
                myMail = Item as Outlook.MailItem;
                (new Thread(() => {
                    Thread.Sleep(100);
                    try {
                        Type dummy = Application.ActiveInspector().GetType();
                        if (myMail.Sender == null) {
                            myMail.BCC = myMail.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                        }
                    } catch (Exception) {
                        ;
                    }
                })).Start();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}