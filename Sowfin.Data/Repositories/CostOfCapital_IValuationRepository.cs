using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class CostOfCapital_IValuationRepository : EntityBaseRepository2<CostOfCapital_IValuation>, ICostOfCapital_IValuation
    {
        public CostOfCapital_IValuationRepository(FindataContext context) : base(context) { }
    }
}
