﻿using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class PayoutPolicy_ScenarioOutputDatasRepository : EntityBaseRepository2<PayoutPolicy_ScenarioOutputDatas> , IPayoutPolicy_ScenarioOutputDatas
    {
        public PayoutPolicy_ScenarioOutputDatasRepository(FindataContext context) : base(context)
        {
        }
    }
}
