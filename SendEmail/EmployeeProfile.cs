using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WishTheEmployee
{
    public class EmployeeProfile
    {
        public string Alias { get; set; }
        public string EmpName { get; set; }
        public DateTime DateOfBirthday { get; set; }
        public DateTime DateOfJoining { get; set; }
        public bool birthdayWishSentForCurrentYear { get; set; }
        public bool serviceAnniversaryWishSentForCurrentYear { get; set; }
    }

    public class EmployeeProfiles
    {
        public List<EmployeeProfile> employeeProfiles { get; set; }
    }
}
