using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public  class Integrated_ExplicitValuesRepository : EntityBaseRepository2<Integrated_ExplicitValues>, IIntegrated_ExplicitValues
    {
        public Integrated_ExplicitValuesRepository(FindataContext context) : base(context) { }
    }
}
