using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class MixedSubValues_FAnalysisRepository : EntityBaseRepository2<MixedSubValues_FAnalysis>, IMixedSubValues_FAnalysis
    {
        public MixedSubValues_FAnalysisRepository(FindataContext context) : base(context)
        {
        }
    }
}
