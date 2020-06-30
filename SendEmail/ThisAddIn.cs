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
        string CUR_USER_NAME;
        const string LOCAL_USER = @"C:\Users\";
        const string MICROSOFT_OUTLOOK = @"\AppData\Local\Microsoft\Outlook\";
        EmployeeProfiles empProfiles;



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
            CUR_USER_NAME = Environment.UserName;
            serializedFileName = LOCAL_USER + CUR_USER_NAME + MICROSOFT_OUTLOOK + @"EmployeeProfilesDatabase.txt";

            empProfiles = GetEmployeeProfiles();
            if (empProfiles.listOfEmployeeProfiles.Count == 0)
                return;

            var listOfEmpProfilesNeededToBeSentEmailForBirthdays = (EmployeeProfiles)GetEmpProfilesNeededToBeSentEmailForBirthdays(empProfiles, out EmployeeProfiles shortlistedEmpsHavingBdaysThisMonth);
            if (listOfEmpProfilesNeededToBeSentEmailForBirthdays.listOfEmployeeProfiles.Count != 0)
            {
                SendWishBirthdays(listOfEmpProfilesNeededToBeSentEmailForBirthdays);
            }
            if(shortlistedEmpsHavingBdaysThisMonth.listOfEmployeeProfiles.Count != 0)
            {
                SendMonthlyReminderForBdays(shortlistedEmpsHavingBdaysThisMonth);
            }
                
            var listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries = (EmployeeProfiles)GetEmpProfilesNeededToBeSentEmailForServiceDeliveries(empProfiles, out EmployeeProfiles shortlistedEmpsHavingServiceDeliveriesThisMonth);
            if (listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries.listOfEmployeeProfiles.Count != 0)
            {
                SendWishServiceDeliveries(listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries);
            }
            if (shortlistedEmpsHavingServiceDeliveriesThisMonth.listOfEmployeeProfiles.Count != 0)
            {
                SendMonthlyReminderForServiceDeliveries(shortlistedEmpsHavingServiceDeliveriesThisMonth);
            }

        }

        private void SendMonthlyReminderForBdays(EmployeeProfiles shortlistedEmpsHavingBdaysThisMonth)
        {
            string str = string.Empty;
            str = "<HTML><head><style>table, th, td {border: 1px solid black;}</style></head>";
            str += "<table><tr><th>ALIAS</th><th>NAME</th><th>BIRTHDAY</th></tr>";
            foreach(var item in shortlistedEmpsHavingBdaysThisMonth.listOfEmployeeProfiles)
            {
                str += "<tr><td>" +
                    item.Alias +
                    "</td><td>" +
                    item.EmpName +
                    "</td><td>" +
                    item.DateOfBirthday +
                    "</td></tr>";
            }

            str += "</body></html>";

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "MONTHLY BIRTHDAY REMINDER!!";
            mailItem.To = Application.Session.CurrentUser.
            AddressEntry.GetExchangeUser().PrimarySmtpAddress;
            mailItem.HTMLBody = str;

            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            mailItem.Send();
        }

        private void SendMonthlyReminderForServiceDeliveries(EmployeeProfiles shortlistedEmpsHavingServiceDeliveriesThisMonth)
        {
            string str = string.Empty;
            str = "<HTML><head><style>table, th, td {border: 1px solid black;}</style></head>";
            str += "<table><tr><th>ALIAS</th><th>NAME</th><th>SERVICE ANNIVERSARY</th></tr>";
            foreach (var item in shortlistedEmpsHavingServiceDeliveriesThisMonth.listOfEmployeeProfiles)
            {
                str += "<tr><td>" +
                    item.Alias +
                    "</td><td>" +
                    item.EmpName +
                    "</td><td>" +
                    item.DateOfJoining +
                    "</td></tr>";
            }

            str += "</body></html>";

            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "MONTHLY SERVICE DELIVERY REMINDER!!";
            mailItem.To = Application.Session.CurrentUser.
            AddressEntry.GetExchangeUser().PrimarySmtpAddress;
            mailItem.HTMLBody = str;

            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            mailItem.Send();
        }


        private void SendWishServiceDeliveries(EmployeeProfiles listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries)
        {
            int yearsDiffJoining = 0;

            foreach (var item in listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries.listOfEmployeeProfiles)
            {
                yearsDiffJoining = DateTime.Now.Year - item.DateOfJoining.Year;

                // conditions for service delivery anniversary

                if (item.birthdayWishSentForCurrentYear == false && item.DateOfJoining.AddYears(yearsDiffJoining).Date == DateTime.Now.Date.AddDays(1).Date)
                {
                    SendTomorrowServiceDeliveryReminderToManager(item);
                }
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining.AddYears(yearsDiffJoining).Date == DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishForToday(item);
                    item.serviceAnniversaryWishSentForCurrentYear = true;
                }
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining.AddYears(yearsDiffJoining).Date < DateTime.Now.Date)
                {
                    SendServiceAnniversaryWishBelated(item);
                    item.serviceAnniversaryWishSentForCurrentYear = true;
                }
                if (item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining.AddYears(yearsDiffJoining).Date > DateTime.Now.Date && (item.DateOfJoining.AddYears(yearsDiffJoining).DayOfWeek == DayOfWeek.Saturday || item.DateOfJoining.AddYears(yearsDiffJoining).DayOfWeek == DayOfWeek.Sunday))
                {
                    SendServiceAnniversaryWishInAdvance(item);
                    item.serviceAnniversaryWishSentForCurrentYear = true;
                }

            }

            UpdateDbAfterWishSent(listOfEmpProfilesNeededToBeSentEmailForServiceDeliveries);
        }

        private void SendTomorrowServiceDeliveryReminderToManager(EmployeeProfile emp)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "REMINDER!! Service Delivery -> " + emp.EmpName;
            mailItem.To = Application.Session.CurrentUser.
            AddressEntry.GetExchangeUser().PrimarySmtpAddress;
            mailItem.HTMLBody = "<HTML>Hey " +
                "<br><br><h2>" +
                emp.EmpName +
                "'s SERVICE ANNIVERSARY!</h2><br><br>" + "<h4>Name: " +
                emp.EmpName +
                "<br>Joining Date: " +
                emp.DateOfJoining +
                "</h4><br><br>" + "It's been " +
                (DateTime.Now.Year - emp.DateOfJoining.Year).ToString() +
                " successful years!";

            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            //mailItem.Display(false);
            mailItem.Send();
        }

        private void SendWishBirthdays(EmployeeProfiles listOfEmpProfilesNeededToBeSentEmailForBirthdays)
        {
            int yearsDiffBirthday = 0;

            foreach (var item in listOfEmpProfilesNeededToBeSentEmailForBirthdays.listOfEmployeeProfiles)
            {
                yearsDiffBirthday = DateTime.Now.Year - item.DateOfBirthday.Year;

                // conditions for birthday
                if (item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.AddYears(yearsDiffBirthday).Date == DateTime.Now.Date.AddDays(1).Date)
                {
                    SendTomorrowBirthdaysReminderToManager(item);
                }
                if (item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.AddYears(yearsDiffBirthday).Date == DateTime.Now.Date)
                {
                    SendBirthDayWishForToday(item);
                    item.birthdayWishSentForCurrentYear = true;
                }
                if (item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.AddYears(yearsDiffBirthday).Date < DateTime.Now.Date)
                {
                    SendBirthdayWishBelated(item);
                    item.birthdayWishSentForCurrentYear = true;
                }
                if (item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.AddYears(yearsDiffBirthday).Date > DateTime.Now.Date && (item.DateOfBirthday.AddYears(yearsDiffBirthday).DayOfWeek == DayOfWeek.Saturday || item.DateOfBirthday.AddYears(yearsDiffBirthday).DayOfWeek == DayOfWeek.Sunday))
                {
                    SendBirthdayWishInAdvance(item);
                    item.birthdayWishSentForCurrentYear = true;
                }
                
            }

            UpdateDbAfterWishSent(listOfEmpProfilesNeededToBeSentEmailForBirthdays);
        }

        private void UpdateDbAfterWishSent(EmployeeProfiles empPros)
        {
            int index = -1;
            foreach(var item in empPros.listOfEmployeeProfiles)
            {
                index = empProfiles.listOfEmployeeProfiles.FindIndex(m => m.Alias == item.Alias);
                if(index>=0 && index<empProfiles.listOfEmployeeProfiles.Count)
                {
                    empProfiles.listOfEmployeeProfiles[index].birthdayWishSentForCurrentYear = item.birthdayWishSentForCurrentYear;
                    empProfiles.listOfEmployeeProfiles[index].serviceAnniversaryWishSentForCurrentYear = item.serviceAnniversaryWishSentForCurrentYear;
                }
            }
            
            if(File.Exists(serializedFileName))
            {
                string jsonString = JsonConvert.SerializeObject(empProfiles, Formatting.Indented);
                File.WriteAllText(serializedFileName, jsonString);
            }
        }

        private void SendTomorrowBirthdaysReminderToManager(EmployeeProfile emp)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = "REMINDER!! Birthday -> " + emp.EmpName;
            mailItem.To = Application.Session.CurrentUser.
            AddressEntry.GetExchangeUser().PrimarySmtpAddress;
            mailItem.HTMLBody = "<HTML><h4>Hey</h4> " +
                "<br><br><h2>" +
                emp.EmpName +
                "'s BIRTHDAY!</h2><br><br>" + "<h4>Name: " +
                emp.EmpName +
                "<br>Birth Date: " +
                emp.DateOfBirthday +
                "</h4><br><br>" +
                emp.EmpName +
                " has become " + 
                (DateTime.Now.Year - emp.DateOfBirthday.Year).ToString() +
                " years old!";

            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            //mailItem.Display(false);
            mailItem.Send();
        }

        private object GetEmpProfilesNeededToBeSentEmailForBirthdays(EmployeeProfiles empProfiles, out EmployeeProfiles shortlistedEmpsHavingBdaysThisMonth)
        {
            EmployeeProfiles shortlistedEmpProfilesForBirthdays = new EmployeeProfiles();
            shortlistedEmpProfilesForBirthdays.listOfEmployeeProfiles = new List<EmployeeProfile>();

            shortlistedEmpsHavingBdaysThisMonth = new EmployeeProfiles();
            shortlistedEmpsHavingBdaysThisMonth.listOfEmployeeProfiles = new List<EmployeeProfile>();

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


                // conditions for birthdays tomorrow
                bool c11 = (item.DateOfBirthday.Date.Day == DateTime.Now.Date.AddDays(1).Day && item.DateOfBirthday.Date.Month == DateTime.Now.Date.AddDays(1).Month);

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
                        OR
                        11. curent day +1 is dob 
                    )
                    AND
                    (
                        9. status of dob is false
                    )
                )*/

                

                if ((c1 || c3 || c5 || c7 || c11) && (c9))
                {
                    shortlistedEmpProfilesForBirthdays.listOfEmployeeProfiles.Add(item);
                }

                if(DateTime.Now.Day == 1 && item.birthdayWishSentForCurrentYear == false && item.DateOfBirthday.Month == DateTime.Now.Month)
                {
                    shortlistedEmpsHavingBdaysThisMonth.listOfEmployeeProfiles.Add(item);
                }
            }

            return shortlistedEmpProfilesForBirthdays;
        }

        private object GetEmpProfilesNeededToBeSentEmailForServiceDeliveries(EmployeeProfiles empProfiles, out EmployeeProfiles shortlistedEmpsHavingServiceDeliveriesThisMonth)
        {
            EmployeeProfiles shortlistedEmpProfilesForServiceDeliveries = new EmployeeProfiles();
            shortlistedEmpProfilesForServiceDeliveries.listOfEmployeeProfiles = new List<EmployeeProfile>();

            shortlistedEmpsHavingServiceDeliveriesThisMonth = new EmployeeProfiles();
            shortlistedEmpsHavingServiceDeliveriesThisMonth.listOfEmployeeProfiles = new List<EmployeeProfile>();

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

                
                // conditions for service deliveries tomorrow
                bool c12 = (item.DateOfJoining.Date.Day == DateTime.Now.Date.AddDays(1).Day && item.DateOfJoining.Date.Month == DateTime.Now.Date.AddDays(1).Month);

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
                        OR
                        12. current day +1 is doj
                    )
                    AND
                    (
                        10. status of doj is false
                    )
                )*/



                if ((c2 || c4 || c6 || c8 || c12) && (c10))
                {
                    shortlistedEmpProfilesForServiceDeliveries.listOfEmployeeProfiles.Add(item);
                }

                if (DateTime.Now.Day == 1 && item.serviceAnniversaryWishSentForCurrentYear == false && item.DateOfJoining.Month == DateTime.Now.Month)
                {
                    shortlistedEmpsHavingServiceDeliveriesThisMonth.listOfEmployeeProfiles.Add(item);
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
