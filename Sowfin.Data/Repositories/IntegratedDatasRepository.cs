using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class IntegratedDatasRepository : EntityBaseRepository2<IntegratedDatas>, IIntegratedDatas
    {
        public IntegratedDatasRepository(FindataContext context) : base(context)
        {
        }
    }
}
