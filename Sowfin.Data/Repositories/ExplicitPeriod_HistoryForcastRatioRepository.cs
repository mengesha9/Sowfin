using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
        public class ExplicitPeriod_HistoryForcastRatioRepository : EntityBaseRepository2<ExplicitPeriod_HistoryForcastRatio>, IExplicitPeriod_HistoryForcastRatio
    {
        public ExplicitPeriod_HistoryForcastRatioRepository(FindataContext context) : base(context) { }
    }
}
