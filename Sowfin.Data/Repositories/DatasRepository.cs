using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public  class DatasRepository : EntityBaseRepository2<LineItemInfo>, IDatas
    {
        public DatasRepository(FindataContext context) : base(context) { }
    }
}
