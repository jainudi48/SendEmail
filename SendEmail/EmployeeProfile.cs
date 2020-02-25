using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendEmail
{
    class EmployeeProfile
    {
        private string Alias;
        private string EmpName;
        private DateTime DateOfBirthday;
        private DateTime DateOfJoining;

        public EmployeeProfile(string alias, string empName, DateTime dateOfBirthday, DateTime dateOfJoining)
        {
            Alias = alias;
            EmpName = empName;
            DateOfBirthday = dateOfBirthday;
            DateOfJoining = dateOfJoining;
        }
    }
}
