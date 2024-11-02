using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
     public class TaxRates_IValuationRepository : EntityBaseRepository2<TaxRates_IValuation>, ITaxRates_IValuation
    {
        public TaxRates_IValuationRepository(FindataContext context) : base(context) { }
    }
}
