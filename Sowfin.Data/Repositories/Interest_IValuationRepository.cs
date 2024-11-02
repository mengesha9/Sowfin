using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public  class Interest_IValuationRepository : EntityBaseRepository2<Interest_IValuation>, IInterest_IValuation
    {
        public Interest_IValuationRepository(FindataContext context) : base(context) { }
    }
}
