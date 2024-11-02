using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class CurrentSetupSnapshotRepository : EntityBaseRepository2<CurrentSetupSnapshot>, ICurrentSetupSnapshot
    {
        public CurrentSetupSnapshotRepository(FindataContext context) : base(context) { }

    }
}
