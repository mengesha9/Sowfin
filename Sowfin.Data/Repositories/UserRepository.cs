using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Model;
using Sowfin.Data.Abstract;

namespace Sowfin.Data.Repositories
{
    public class UserRepository : EntityBaseRepository2<User>, IUserRepository
    {
        public UserRepository(FindataContext context) : base(context) { }

        public bool isEmailUniq(string email)
        {
            var user = this.GetSingle(u => u.Email == email);
            return user == null;
        }

        public bool IsUsernameUniq(string username)
        {
           var user = this.GetSingle(u => u.UserName == username);
           return user == null;
        }
    }
}