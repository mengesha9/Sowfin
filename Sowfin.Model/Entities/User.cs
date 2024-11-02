using System.Collections.Generic;
using Sowfin.Model;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.Model
{
    public class User
    {

        public long Id { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdateAt { get; set; }
        public int Deleted { get; set; }
        public string Email { get; set; }
        public int RoleId { get; set; }
        public int Active { get; set; }
    }
}