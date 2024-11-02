using System.Collections.Generic;

namespace Sowfin.API.ViewModels
{
    public class UserDrafts
    {
      public string Username { get; set; }
      public List<DraftViewModel> Drafts { get; set; }
    }
    public class SharedDraftsViewModel
    {
        public List<UserDrafts> UsersDrafts { get; set; }
    }
}