using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class ReorganizedValuesRepository : EntityBaseRepository2<ReorganizedValues>, IReorganizedValues
    {
        public ReorganizedValuesRepository(FindataContext context) : base(context) { }
    }
}
