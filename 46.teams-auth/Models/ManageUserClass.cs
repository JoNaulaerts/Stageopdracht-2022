using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TeamsAuth.Models
{
    public class ManageUserClass
    {
        public string currentPassword { get; set; }
        public string newPassword { get; set; }
        public string contolNewPassword { get; set; }

        public bool accountEnabled { get; set; }
        public string displayName { get; set; }
        public string mailNickname { get; set; }
        public string userPrincipalName { get; set; }

    public void resetValues()
        {
            currentPassword = null;
            newPassword = null;
            contolNewPassword = null;

            accountEnabled = false;
            displayName = null;
            mailNickname = null;
            userPrincipalName = null;
        }
    }
}
