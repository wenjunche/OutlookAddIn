using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using OpenFin.Notifications.Constants;
using OpenFin.Notifications;

namespace FirstOutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        Task task;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

            NotificationClient.InitializeAsync().ContinueWith(x =>
            {
                Console.WriteLine("NotificationClient ready");
            });

        }

        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by my first add-in";
                    mailItem.Body = "This text was added by my first add-in";
                }

            }
        }

        void items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                Console.WriteLine("new email: subject {0} from: {1}", mail.Subject, mail.SenderEmailAddress);
                string Title = "OpenFin OutLook AddIn";

                task = NotificationClient.CreateNotificationAsync(Title, new NotificationOptions
                {
                    Title = mail.SenderEmailAddress,
                    Body = mail.Subject,
                    Category = "Category",
                    Icon = "https://openfin.co/favicon-32x32.png",
                    OnNotificationSelect = new Dictionary<string, object>
                    {
                        {"task", "selected" }
                    },
                    Buttons = new ButtonOptions[] {
                        new ButtonOptions {
                            Title = "Button1",
                            IconUrl = "https://openfin.co/favicon-32x32.png",
                            OnNotificationButtonClick = new Dictionary<string, object>
                                {
                                    { "btn", "button1" }
                                },
                            IsCallToActionButton = false
                        }
                    },
                    IsStickyNotification = true,
                    NotificationIndicator = new NotificationIndicator
                    {
                        IndicatorType = IndicatorType.Success
                    },
                    Expires = DateTime.Now.AddSeconds(60),
                    CustomData = new Dictionary<string, object>()
                });

                try
                {
                    task.Wait();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("exception {0}", ex);
                }
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
