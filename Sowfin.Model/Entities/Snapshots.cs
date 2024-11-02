using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class Snapshots
    {
        public long Id { get; set; }
        public string SnapShot { get; set; }
        public string Description { get; set; }
        public long UserId { get; set; }
        public long ProjectId { get; set; }
        public string SnapShotType { get; set; }
        public string NPV { get; set; }
        public string CNPV { get; set; }
    }
}
