using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class InitialSetup_IValuationRepository : EntityBaseRepository2<InitialSetup_IValuation>, IInitialSetup_IValuation
    {
        public InitialSetup_IValuationRepository(FindataContext context) : base(context) { }
    }
}
