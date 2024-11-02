using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class AssetsEquityDatasRepository : EntityBaseRepository2<AssetsEquityDatas>, IAssetsEquityDatas
    {
        public AssetsEquityDatasRepository(FindataContext context) : base(context) { }
    }
}
