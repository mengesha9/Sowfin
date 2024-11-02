using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;


namespace Sowfin.Data.Repositories
{
    public class BusinessUnitRepository : EntityBaseRepository2<BusinessUnit>, IBusinessUnit
    {
        public BusinessUnitRepository(FindataContext context) : base(context) { }
    }
}
