using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
        public class PayoutPolicy_ScenarioValuesRepository : EntityBaseRepository2<PayoutPolicy_ScenarioValues>, IPayoutPolicy_ScenarioValues
    {
        public PayoutPolicy_ScenarioValuesRepository(FindataContext context) : base(context) { }
    }
}
