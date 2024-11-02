using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class ValuesRepository : EntityBaseRepository2<RawHistoricalValues>, IValues
    {
        public ValuesRepository(FindataContext context) : base(context) { }
    }
}
