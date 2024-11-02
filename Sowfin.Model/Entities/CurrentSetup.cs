using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
  public class CurrentSetup
    {
        public long Id { get; set; }  
        public string CurrentYear { get; set; }
        public string EndYear { get; set; }
        public long? InitialSetupId { get; set; }
        public long? UserId { get; set; }
        public string CIKNumber { get; set; }
        public string Company { get; set; }
    }
}
