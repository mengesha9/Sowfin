using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class IVSensitivity
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public long? SensitivityId { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }
    }
}