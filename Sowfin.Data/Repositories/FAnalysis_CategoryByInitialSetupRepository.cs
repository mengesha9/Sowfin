using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class FAnalysis_CategoryByInitialSetupRepository : EntityBaseRepository2<FAnalysis_CategoryByInitialSetup>, IFAnalysis_CategoryByInitialSetup
    {
        public FAnalysis_CategoryByInitialSetupRepository(FindataContext context) : base(context)
        {
        }
    }
}
