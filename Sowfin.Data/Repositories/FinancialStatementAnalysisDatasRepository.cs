using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class FinancialStatementAnalysisDatasRepository : EntityBaseRepository2<FinancialStatementAnalysisDatas>, IFinancialStatementAnalysisDatas
    {
        public FinancialStatementAnalysisDatasRepository(FindataContext context) : base(context) { }
    }
}
