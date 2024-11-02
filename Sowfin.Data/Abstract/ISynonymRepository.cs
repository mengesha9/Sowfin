using Sowfin.Model;

namespace Sowfin.Data.Abstract
{
    public interface ISynonymRepository: IEntityBaseRepository<Synonyms>
    {
        //bool IsOwner(string storyId, string userId);
        //bool IsInvited(string storyId, string userId);
    }
}