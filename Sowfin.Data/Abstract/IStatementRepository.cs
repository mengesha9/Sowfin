using Sowfin.Model;

namespace Sowfin.Data.Abstract
{
    public interface IStatementRepository: IEntityBaseRepository<Statement>
    {
        //bool IsOwner(string storyId, string userId);
        //bool IsInvited(string storyId, string userId);
    }
}