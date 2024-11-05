using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.User
{
    public class UserLoginViewModel
    {
        public string Email { get; set; }
        public string Password { get; set; }
    }

    public class UserCheck
    {
        public string username  { get; set; }
        public string email { get; set; }
    }
}
