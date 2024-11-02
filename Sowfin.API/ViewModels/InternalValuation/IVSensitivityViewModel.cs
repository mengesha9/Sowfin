using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Sowfin.API.ViewModels.InternalValuation;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class IVSensitivityViewModel
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public long? SensitivityId { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }

    }
    public class KeyValueViewModel
    {
        public string Name { get; set; }
        public List<string> Key { get; set; }
        public List<string> Value { get; set; }

    }
    public class KeyValueListViewModel
    {
        public List<KeyValueViewModel> KeyValueVM { get; set; }

    }
}
