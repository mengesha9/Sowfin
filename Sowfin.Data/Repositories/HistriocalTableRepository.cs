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
    public class HistriocalTableRepository : EntityBaseRepository2<HistoricalTable>, IHistriocalTable
    {
        private readonly FindataContext _context;
        public HistriocalTableRepository(FindataContext context) : base(context)
        {
            _context = context;
        }

        public async void DelateAll()
        {
            await _context.HistoricalTable.FromSql($"Truncate histriocal_table").LoadAsync();
            await _context.SaveChangesAsync();
           
        }
    }
}
