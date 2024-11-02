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
    public class EdgarDataRepository : EntityBaseRepository2<EdgarData>, IEdgarDataRepository
    {
        private readonly FindataContext _context;
        public EdgarDataRepository(FindataContext context) : base(context)
        {
            _context = context;
        }

        public IEnumerable<EdgarData> GetEdgar(String cik, int? startYear = null, int? endYear = null)
        {
            return _context.EdgarData.FromSqlRaw("Select * from \"EdgarView\"({0}, {1}, {2})", cik, startYear, endYear); // changed from .query<EdgarData>() to .EdgarData
        }

        public IEnumerable<EdgarDataByCategory> GetEdgarByCategory(String cik, int? startYear = null, int? endYear = null)
        {
            return _context.EdgarDataByCategory.FromSqlRaw("Select * from \"EdgarViewByCategory\"({0}, {1}, {2})", cik, startYear, endYear); // changed from .query<EdgarDataByCategory>() to .EdgarDataByCategory
        }

    }
}
