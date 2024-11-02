using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class ROICDatasRepository : EntityBaseRepository2<ROICDatas>, IROICDatas
    {
        public ROICDatasRepository(FindataContext context) : base(context) { }
    }
}
