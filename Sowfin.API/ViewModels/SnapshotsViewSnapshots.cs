using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels
{
    public class SnapshotsViewSnapshots
    {
        public SnapshotsViewSnapshots()
        {

        }
        public long Id { get; set; }
        public string SnapShot { get; set; }
        public string Description { get; set; }
        public long UserId { get; set; }
        public long ProjectId { get; set; }
        public string SnapShotType { get; set; }
        public string NVP { get; set; }
        public string CNVP { get; set; }
    }
}
