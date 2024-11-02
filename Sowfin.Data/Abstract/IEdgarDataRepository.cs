using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System.Collections.Generic;
using System;

namespace Sowfin.Data.Abstract
{
    public interface IEdgarDataRepository : IEntityBaseRepository<EdgarData>
    {
        IEnumerable<EdgarData> GetEdgar(String cik, int? startYear = null, int? endYear = null);       
        IEnumerable<EdgarDataByCategory> GetEdgarByCategory(String cik, int? startYear = null, int? endYear = null);
    }
}
