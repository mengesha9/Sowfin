
using Sowfin.Data.Abstract;
using Sowfin.Data.Repositories;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Abstract
{
    public interface IHistriocalTable : IEntityBaseRepository<HistoricalTable>
    {
        void DelateAll();
    }
}
