using System.Collections.Generic;
using System.Linq;
using Sowfin.Data.Abstract;
using Sowfin.Model;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;

namespace Sowfin.Data.Repositories
{
    public class ShareRepository : EntityBaseRepository<Share>, IShareRepository
    {
      public ShareRepository(BlogContext context) : base(context) {}

      public List<Story> StoriesSharedToUser(string userId)
      {
        return _context.Set<Share>()
          .Where(s => s.UserId == userId)
          .Select(s => s.Story)
          .Include(s => s.Owner)
          .ToList();
      }
  }
}