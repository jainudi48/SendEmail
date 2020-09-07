# SendEmail
Outlook AddIn for celebrating birthdays and joinings

# Features
1. Beautifully creates an email draft for the managers to send wishes to their direct reports on their birthdays and service delivery anniversaries.
2. Sends a monthly reminder list to the manager on 1st of each month listing all the employees whose birthdays and service deliveries are falling in the corresponding month.
3. 1st of every month, one should receive email with a list of all the birthdays and service deliveries in this month.
4. send reminder 1 day ago for birthdays and service deliveries.

# Future Changes
1. Remove birth year from all the emails 
2. In drafted email, add the date of joining/birth so manager has more visibility around the dates (without year for birthdays)

# Bug
1. Flags are not resetting to false at the year end. 
Logic:
  a. Reset flags for service delivery and birthday at year end.
  b. If anyone's wishing date lies in 1st week of the year then advance wishing logic should be rechecked.
