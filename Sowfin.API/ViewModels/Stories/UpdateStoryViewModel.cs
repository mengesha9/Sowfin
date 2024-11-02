using System.Collections.Generic;

namespace Sowfin.API.ViewModels
{
    public class UpdateStoryViewModel
    {
        public string Title { get; set; }
        public string Content { get; set; }
        public List<string> Tags { get; set; }
    }
}