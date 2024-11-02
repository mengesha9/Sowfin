using Sowfin.Data.Abstract;
using Sowfin.Model;

namespace Sowfin.Data.Repositories
{
    public class StatementRepository : EntityBaseRepository2<Statement>, IStatementRepository 
    {
        public StatementRepository (FindataContext context) : base (context) { }

        /*
        public bool IsInvited(string storyId, string userId)
        {
            var story = this.GetSingle(s => s.Id == storyId, s => s.Shares);
            return story.Shares.Exists(s => s.UserId == userId);
        }

        public bool IsOwner(string storyId, string userId)
        {
            var story = this.GetSingle(s => s.Id == storyId);
            return story.OwnerId == userId;
        }
        */
  }
}