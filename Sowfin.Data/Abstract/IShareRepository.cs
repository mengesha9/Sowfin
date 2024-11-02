using System.Collections.Generic;
using Sowfin.Model;

namespace Sowfin.Data.Abstract
{
    public interface IShareRepository : IEntityBaseRepository<Share>
    {
        List<Story> StoriesSharedToUser(string userId);
    }
}