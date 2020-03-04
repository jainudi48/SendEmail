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
            
        }

        private void SendServiceAnniversaryWishInAdvance(EmployeeProfile emp)
        {
            string name = emp.EmpName;
            string email = DecorateEmailFromAlias(emp.Alias);
            string yearsWorking = (DateTime.Now.Year - emp.DateOfJoining.Year).ToString();

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
            string yearsWorking = (DateTime.Now.Year - emp.DateOfJoining.Year).ToString();

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
            string yearsWorking = (DateTime.Now.Year - emp.DateOfJoining.Year).ToString();

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

            var listOfEmpProfilesNeededToBeSentEmailForBirthdays = (EmployeeProfiles)GetEmpProfilesNeededToBeSentEmailForBirthdays(empProfiles);
            if (listOfEmpProfilesNeededToBeSentEmailForBirthdays.listOfEmployeeProfiles.Count != 0)
                SendWishBirthdays(listOfEmpProfilesNeededToBeSentEmailForBirthdays);

            var listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries = (EmployeeProfiles)GetEmpProfilesNeededToBeSentEmailForServiceDeliveries(empProfiles);
            if (listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries.listOfEmployeeProfiles.Count != 0)
                SendWishServiceDeliveries(listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries);

        }

        private void SendWishServiceDeliveries(EmployeeProfiles listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries)
        {
            int yearsDiffJoining = 0;

            foreach (var item in listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries.listOfEmployeeProfiles)
            {
                yearsDiffJoining = DateTime.Now.Year - item.DateOfJoining.Year;

                // conditions for service delivery anniversary
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining.AddYears(yearsDiffJoining).Date == DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishForToday(item);
                }
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining.AddYears(yearsDiffJoining).Date < DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishBelated(item);
                }
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining.AddYears(yearsDiffJoining).Date > DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishInAdvance(item);
                }
            }
        }

        private void SendWishBirthdays(EmployeeProfiles listOfEmpProfilesNeededToBeSentEmailForBirthdays)
        {
            int yearsDiffBirthday = 0;

            foreach (var item in listOfEmpProfilesNeededToBeSentEmailForBirthdays.listOfEmployeeProfiles)
            {
                yearsDiffBirthday = DateTime.Now.Year - item.DateOfBirthday.Year;

                // conditions for birthday
                if (item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.AddYears(yearsDiffBirthday).Date == DateTime.Now.Date)
                {
                    SendBirthDayWishForToday(item);
                }
                if (item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.AddYears(yearsDiffBirthday).Date < DateTime.Now.Date)
                {
                    SendBirthdayWishBelated(item);
                }
                if (item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.AddYears(yearsDiffBirthday).Date > DateTime.Now.Date)
                {
                    SendBirthdayWishInAdvance(item);
                }
            }
        }

        private object GetEmpProfilesNeededToBeSentEmailForBirthdays(EmployeeProfiles empProfiles)
        {
            EmployeeProfiles shortlistedEmpProfilesForBirthdays = new EmployeeProfiles();
            shortlistedEmpProfilesForBirthdays.listOfEmployeeProfiles = new List<EmployeeProfile>();


            int yearDiffBirthday = 0;
           

            foreach (var item in empProfiles.listOfEmployeeProfiles)
            {
                yearDiffBirthday = DateTime.Now.Year - item.DateOfBirthday.Year;
           

                // conditions for wish to be sent on same day
                bool c1 = item.DateOfBirthday.Date.Day == DateTime.Now.Date.Day && item.DateOfBirthday.Date.Month == DateTime.Now.Date.Month;
           

                // conditions for advance wish
                bool c3 = (item.DateOfBirthday.Date.Day == DateTime.Now.Date.AddDays(1).Day && item.DateOfBirthday.Date.Month == DateTime.Now.Date.AddDays(1).Month) && (DateTime.Now.Date.AddDays(1).DayOfWeek == DayOfWeek.Saturday);
           
                bool c5 = (item.DateOfBirthday.Date.Day == DateTime.Now.Date.AddDays(2).Day && item.DateOfBirthday.Date.Month == DateTime.Now.Date.AddDays(2).Month) && (DateTime.Now.Date.AddDays(2).DayOfWeek == DayOfWeek.Sunday);
           

                // conditions for missed wish
                // Added the year difference b/w actual date of occasion and current date
                bool c7 = (item.DateOfBirthday.Date.AddYears(yearDiffBirthday) >= DateTime.Now.Date.AddDays(-7)) && (item.DateOfBirthday.Date.AddYears(yearDiffBirthday) < DateTime.Now.Date);
           

                // conditions to check status flag 
                bool c9 = item.birthdayWishSentForCurrentYear == false;
           

                
                /*
                 * Below is the pseudo code for the below condition
                 * 
                if
                (
                    (
                        1. dob matches current date 
                        OR
                        3. current day +1 is dob AND dob day is saturday 
                        OR
                        5. current day +2 is dob AND dob day is sunday 
                        OR
                        7. dob is lesser or equals to current day - 7
                    )
                    AND
                    (
                        9. status of dob is false
                    )
                )*/

                

                if ((c1 || c3 || c5 || c7) && (c9))
                {
                    shortlistedEmpProfilesForBirthdays.listOfEmployeeProfiles.Add(item);
                }
            }

            return shortlistedEmpProfilesForBirthdays;
        }

        private object GetEmpProfilesNeededToBeSentEmailForServiceDeliveries(EmployeeProfiles empProfiles)
        {
            EmployeeProfiles shortlistedEmpProfilesForServiceDeliveries = new EmployeeProfiles();
            shortlistedEmpProfilesForServiceDeliveries.listOfEmployeeProfiles = new List<EmployeeProfile>();


            
            int yearDiffJoining = 0;

            foreach (var item in empProfiles.listOfEmployeeProfiles)
            {
            
                yearDiffJoining = DateTime.Now.Year - item.DateOfJoining.Year;

                // conditions for wish to be sent on same day
            
                bool c2 = item.DateOfJoining.Date.Day == DateTime.Now.Date.Day && item.DateOfJoining.Date.Month == DateTime.Now.Date.Month;

                // conditions for advance wish
            
                bool c4 = (item.DateOfJoining.Date.Day == DateTime.Now.Date.AddDays(1).Day && item.DateOfJoining.Date.Month == DateTime.Now.Date.AddDays(1).Month) && (DateTime.Now.Date.AddDays(1).DayOfWeek == DayOfWeek.Saturday);
            
                bool c6 = (item.DateOfJoining.Date.Day == DateTime.Now.Date.AddDays(2).Day && item.DateOfJoining.Date.Month == DateTime.Now.Date.AddDays(2).Month) && (DateTime.Now.Date.AddDays(2).DayOfWeek == DayOfWeek.Sunday);

                // conditions for missed wish
                // Added the year difference b/w actual date of occasion and current date
            
                bool c8 = (item.DateOfJoining.Date.AddYears(yearDiffJoining) >= DateTime.Now.Date.AddDays(-7)) && (item.DateOfJoining.Date.AddYears(yearDiffJoining) < DateTime.Now.Date);

                // conditions to check status flag 
            
                bool c10 = item.serviceAnniversaryWishSentForCurrentYear == false;


                /*
                 * Below is the pseudo code for the below condition
                 * 
                if
                (
                    (
                        2. doj matches current date
                        OR
                        4. current day +1 is doj AND doj day is saturday 
                        OR
                        6. current day +2 is doj AND doj day is sunday 
                        OR
                        8. doj is lesser or equals to current day - 7
                    )
                    AND
                    (
                        10. status of doj is false
                    )
                )*/



                if ((c2 || c4 || c6 || c8) && (c10))
                {
                    shortlistedEmpProfilesForServiceDeliveries.listOfEmployeeProfiles.Add(item);
                }
            }

            return shortlistedEmpProfilesForServiceDeliveries;
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
