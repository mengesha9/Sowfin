using Sowfin.Data;
using Sowfin.Data.Repositories;
using Microsoft.EntityFrameworkCore;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class LineItemInfoRepository : EntityBaseRepository2<LineItemInfo>, ILineItemInfoRepository
    {
        
        public LineItemInfoRepository(FindataContext context) : base(context)
        {
            
        }

       
    }
}
