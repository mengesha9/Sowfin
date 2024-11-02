using Sowfin.Data.Abstract;
using Sowfin.Model;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;

namespace Sowfin.Data.Repositories
{
    public class LikeRepository : EntityBaseRepository<Like>, ILikeRepository
    {
        public LikeRepository(BlogContext context) : base(context) {}
    }
}