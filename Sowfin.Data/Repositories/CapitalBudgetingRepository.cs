using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class CapitalBudgetingRepository : EntityBaseRepository2<CapitalBudgeting>, ICapitalBudgeting
    {
        public CapitalBudgetingRepository(FindataContext context) : base(context) { }
    }
}
