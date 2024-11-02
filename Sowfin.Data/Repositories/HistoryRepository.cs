using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class HistoryRepository : EntityBaseRepository2<History>, IHistory
    {
        public HistoryRepository(FindataContext context) : base(context) { }
    }
}
