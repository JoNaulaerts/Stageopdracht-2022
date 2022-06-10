using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TeamsAuth.Models
{
    public class ManageGroupClass
    {
        public string groupId { get; set; }
        public string ownerId { get; set; }
        public string ownerPrincipalName { get; set; }
        public string groupDisplayName { get; set; }
        public string groupDescription { get; set; }
        public string groupMailNickname { get; set; }
        public bool confirmDeleteGroup { get; set; }
        public bool? addTeam { get; set; }
        public bool teamExists { get; set; }


        public void resetValues()
        {
            groupId = null;
            ownerId = null;
            ownerPrincipalName = null;
            groupDisplayName = null;
            groupDescription = null;
            groupMailNickname = null;
            confirmDeleteGroup = false;
            addTeam = false;
            teamExists = true;
        }
    }
}
