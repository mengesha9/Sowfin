using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
   public  class MixedSubDatas_FAnalysisRepository : EntityBaseRepository2<MixedSubDatas_FAnalysis>, IMixedSubDatas_FAnalysis
    {
        public MixedSubDatas_FAnalysisRepository(FindataContext context) : base(context)
        {
        }
    }


}
