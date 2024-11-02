using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class CalculateBetaRepository : EntityBaseRepository2<CalculateBeta>, ICalculateBeta
    {
        public CalculateBetaRepository(FindataContext context) : base(context) { }
    }
}
