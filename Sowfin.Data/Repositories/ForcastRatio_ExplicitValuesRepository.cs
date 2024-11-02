using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class ForcastRatio_ExplicitValuesRepository : EntityBaseRepository2<ForcastRatio_ExplicitValues>, IForcastRatio_ExplicitValues
    {
        public ForcastRatio_ExplicitValuesRepository(FindataContext context) : base(context) { }
    }
}
