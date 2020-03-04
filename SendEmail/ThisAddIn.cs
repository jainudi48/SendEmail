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
        string serializedFileName;
        private void SendWish(EmployeeProfiles shortlistedEmpProfiles)
        {
            
            
            foreach(var item in shortlistedEmpProfiles.listOfEmployeeProfiles)
            {
                // conditions for birthday
                if(item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.Date == DateTime.Now.Date)
                {
                    SendBirthDayWishForToday(item);
                }
                if(item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.Date < DateTime.Now.Date)
                {
                    SendBirthdayWishBelated(item);
                }
                if(item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.Date > DateTime.Now.Date)
                {
                    SendBirthdayWishInAdvance(item);
                }

                // conditions for service delivery anniversary
                if(item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining == DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishForToday(item);
                }
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining == DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishBelated(item);
                }
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining == DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishInAdvance(item);
                }
            }
        }

        private void SendServiceAnniversaryWishInAdvance(EmployeeProfile emp)
        {
            string name = emp.EmpName;
            string email = DecorateEmailFromAlias(emp.Alias);
            string yearsWorking = (emp.DateOfJoining.Year - DateTime.Now.Year).ToString();

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "Happy Service Anniversary " + name; ;
            mailItem.To = email;
            mailItem.Body = "Happy " + yearsWorking + " years of Service Anniversary in advance!";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        private void SendServiceAnniversaryWishBelated(EmployeeProfile emp)
        {
            string name = emp.EmpName;
            string email = DecorateEmailFromAlias(emp.Alias);
            string yearsWorking = (emp.DateOfJoining.Year - DateTime.Now.Year).ToString();

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "Happy Service Anniversary " + name; ;
            mailItem.To = email;
            mailItem.Body = "Happy Belated " + yearsWorking + " years of Service Anniversary!";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        private void SendServiceAnniversaryWishForToday(EmployeeProfile emp)
        {
            string name = emp.EmpName;
            string email = DecorateEmailFromAlias(emp.Alias);
            string yearsWorking = (emp.DateOfJoining.Year - DateTime.Now.Year).ToString();

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "Happy Service Anniversary " + name; ;
            mailItem.To = email;
            mailItem.Body = "Happy " + yearsWorking + " years of Service Anniversary!";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        private void SendBirthdayWishInAdvance(EmployeeProfile emp)
        {
            string name = emp.EmpName;
            string email = DecorateEmailFromAlias(emp.Alias);

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "Happy Birthday " + name; ;
            mailItem.To = email;
            mailItem.Body = "Happy Birthday in advance!";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        private void SendBirthdayWishBelated(EmployeeProfile emp)
        {
            string name = emp.EmpName;
            string email = DecorateEmailFromAlias(emp.Alias);

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "Happy Birthday " + name; ;
            mailItem.To = email;
            mailItem.Body = "Happy Belated Birthday!";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        private void SendBirthDayWishForToday(EmployeeProfile emp)
        {
            string name = emp.EmpName;
            string email = DecorateEmailFromAlias(emp.Alias);

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "Happy Birthday " + name; ;
            mailItem.To = email;
            mailItem.Body = "Happy Birthday!";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        private string DecorateEmailFromAlias(string alias)
        {
            return alias + "@microsoft.com";
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            serializedFileName = @"C:\Repos\WishTheEmployee\WishTheEmployee\bin\Debug\EmployeeProfilesDatabase.txt";

            var empProfiles = GetEmployeeProfiles();
            if (empProfiles.listOfEmployeeProfiles.Count == 0)
                return;

            var listOfEmpProfilesNeededToBeSentEmail = (EmployeeProfiles)GetEmpProfilesNeededToBeSentEmail(empProfiles);
            if (listOfEmpProfilesNeededToBeSentEmail.listOfEmployeeProfiles.Count == 0)
                return;

            SendWish(listOfEmpProfilesNeededToBeSentEmail);
        }

        private object GetEmpProfilesNeededToBeSentEmail(EmployeeProfiles empProfiles)
        {
            EmployeeProfiles shortlistedEmpProfiles = new EmployeeProfiles();
            shortlistedEmpProfiles.listOfEmployeeProfiles = new List<EmployeeProfile>();

            foreach(var item in empProfiles.listOfEmployeeProfiles)
            {
                // conditions for wish to be sent on same day
                bool c1 = item.DateOfBirthday.Date.Day == DateTime.Now.Date.Day && item.DateOfBirthday.Date.Month == DateTime.Now.Date.Month;
                bool c2 = item.DateOfJoining.Date.Day == DateTime.Now.Date.Day && item.DateOfJoining.Date.Month == DateTime.Now.Date.Month;

                // conditions for advance wish
                bool c3 = (item.DateOfBirthday.Date.Day == DateTime.Now.Date.AddDays(1).Day && item.DateOfBirthday.Date.Month == DateTime.Now.Date.AddDays(1).Month) && (DateTime.Now.Date.AddDays(1).DayOfWeek == DayOfWeek.Saturday);
                bool c4 = (item.DateOfJoining.Date.Day == DateTime.Now.Date.AddDays(1).Day && item.DateOfJoining.Date.Month == DateTime.Now.Date.AddDays(1).Month) && (DateTime.Now.Date.AddDays(1).DayOfWeek == DayOfWeek.Saturday);
                bool c5 = (item.DateOfBirthday.Date.Day == DateTime.Now.Date.AddDays(2).Day && item.DateOfBirthday.Date.Month == DateTime.Now.Date.AddDays(2).Month) && (DateTime.Now.Date.AddDays(2).DayOfWeek == DayOfWeek.Sunday);
                bool c6 = (item.DateOfJoining.Date.Day == DateTime.Now.Date.AddDays(2).Day && item.DateOfJoining.Date.Month == DateTime.Now.Date.AddDays(2).Month) && (DateTime.Now.Date.AddDays(2).DayOfWeek == DayOfWeek.Sunday);

                // conditions for missed wish
                // below two conditions need to be checked
                bool c7 = (item.DateOfBirthday.Date >= DateTime.Now.Date.AddDays(-7)) && (item.DateOfBirthday.Date < DateTime.Now.Date);
                bool c8 = (item.DateOfJoining.Date >= DateTime.Now.Date.AddDays(-7)) && (item.DateOfJoining.Date < DateTime.Now.Date);

                // conditions to check status flag 
                bool c9 = item.birthdayWishSentForCurrentYear == false;
                bool c10 = item.serviceAnniversaryWishSentForCurrentYear == false;

                
                /*
                 * Below is the pseudo code for the below condition
                 * 
                if
                (
                    (
                        1. dob matches current date 
                        OR
                        2. doj matches current date
                        OR
                        3. current day +1 is dob AND dob day is saturday 
                        OR
                        4. current day +1 is doj AND doj day is saturday 
                        OR
                        5. current day +2 is dob AND dob day is sunday 
                        OR
                        6. current day +1 is doj AND doj day is sunday 
                        OR
                        7. dob is lesser or equals to current day - 7
                        OR 
                        8. doj is lesser or equals to current day - 7
                    )
                    AND
                    (
                        9. status of dob is false
                        OR
                        10. status of doj is false
                    )
                )*/

                

                if ( (c1 || c2 || c3 || c4 || c5 || c6 || c7 || c8) && (c9 || c10) )
                {
                    shortlistedEmpProfiles.listOfEmployeeProfiles.Add(item);
                }
            }

            return shortlistedEmpProfiles;
        }

        private EmployeeProfiles GetEmployeeProfiles()
        {
            string existingJsonString;
            EmployeeProfiles empProfiles = new EmployeeProfiles();
            empProfiles.listOfEmployeeProfiles = new List<EmployeeProfile>();

            if (!File.Exists(serializedFileName))
            {
                FileStream fs = File.Create(serializedFileName);
                fs.Close();
            }
            else
            {
                existingJsonString = File.ReadAllText(serializedFileName);
                if (!existingJsonString.Equals(""))
                    empProfiles = JsonConvert.DeserializeObject<EmployeeProfiles>(existingJsonString);
            }
            return empProfiles;
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
