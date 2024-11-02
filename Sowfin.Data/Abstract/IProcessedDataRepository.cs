using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System.Collections.Generic;


namespace Sowfin.Data.Abstract
{
    public interface IProcessedDataRepository : IEntityBaseRepository<ProcessedData>
    {
        IEnumerable<ProcessedData> GetProcessedData();
    }
}
