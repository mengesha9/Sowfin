using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class InitialSetup_FAnalysisRepository : EntityBaseRepository2<InitialSetup_FAnalysis>, IInitialSetup_FAnalysis
    {
        public InitialSetup_FAnalysisRepository(FindataContext context) : base(context) { }

    }
}
