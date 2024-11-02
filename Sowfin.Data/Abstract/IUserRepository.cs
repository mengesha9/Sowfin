using Sowfin.Data.Abstract;
using Sowfin.Model;

namespace Sowfin.Data.Abstract
{
    public interface IUserRepository : IEntityBaseRepository<User>
    {
        //bool IsUsernameUniq(string username);
        bool isEmailUniq(string email);
    }
}