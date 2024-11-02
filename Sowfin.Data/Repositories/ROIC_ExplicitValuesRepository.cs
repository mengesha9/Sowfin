using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class ROIC_ExplicitValuesRepository : EntityBaseRepository2<ROIC_ExplicitValues>, IROIC_ExplicitValues
    {
        public ROIC_ExplicitValuesRepository(FindataContext context) : base(context) { }
    }
}
