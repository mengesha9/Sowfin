using System.Collections.Generic;
using Sowfin.Model;

namespace Sowfin.Mocker
{
    public class Pack
    {
        public List<User> Users { get; set; } = new List<User>();
        public List<Story> Stories { get; set; } = new List<Story>();
    }
}