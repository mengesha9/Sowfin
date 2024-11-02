using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class Reorganized_ExplicitValuesRepository : EntityBaseRepository2<Reorganized_ExplicitValues>, IReorganized_ExplicitValues
    {
        public Reorganized_ExplicitValuesRepository(FindataContext context) : base(context) { }
    }
}
