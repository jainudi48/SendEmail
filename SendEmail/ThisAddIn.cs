using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using Newtonsoft.Json;

namespace SendEmail
{
    public partial class ThisAddIn
    {
   

        private void CreateMailItem()
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            
            mailItem.Subject = "This is the subject";
            mailItem.To = "jain.udit48@outlook.com";
            mailItem.Body = "This is the message.";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
            mailItem.Send();
        }

        private void WriteEmployeeProfilesToFile(EmployeeProfile empProfile)
        {
            
        }

        //private void ReadEmployeeProfileFromConsole()
        //{
        //    string name = Console.ReadLine();
        //    DateTime dob = Console.ReadLine();
        //    DateTime doj = Console.ReadLine();
        //    EmployeeProfile empProfile = new EmployeeProfile(name, dob, doj);

        //}

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CreateMailItem();
            //SendEmailtoContacts();
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
