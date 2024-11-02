using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
      public class PayoutPolicy_ScenarioDatasRepository : EntityBaseRepository2<PayoutPolicy_ScenarioDatas>, IPayoutPolicy_ScenarioDatas
    {
        public PayoutPolicy_ScenarioDatasRepository(FindataContext context) : base(context) { }
    }
}
