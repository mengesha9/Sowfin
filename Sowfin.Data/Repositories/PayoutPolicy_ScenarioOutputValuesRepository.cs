using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class PayoutPolicy_ScenarioOutputValuesRepository : EntityBaseRepository2<PayoutPolicy_ScenarioOutputValues>, IPayoutPolicy_ScenarioOutputValues
    {
        public PayoutPolicy_ScenarioOutputValuesRepository(FindataContext context) : base(context)
        {
        }
    }
}
