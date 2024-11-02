using Sowfin.Data;
using Sowfin.Data.Repositories;
using Microsoft.EntityFrameworkCore;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System.Collections.Generic;


namespace Sowfin.Data.Repositories
{
    public class ProcessedDataRepository : EntityBaseRepository2<ProcessedData>, IProcessedDataRepository
    {
        private readonly FindataContext _context;
        public ProcessedDataRepository(FindataContext context) : base(context)
        {
            _context = context;
        }

        public IEnumerable<ProcessedData> GetProcessedData()
        {
            return _context.ProcessedData.FromSqlRaw("Select * from \"ProcessedData\"() "); // changed from .query<ProcessedData>() to .ProcessedData
        }
    }
}
