using System.Collections.Generic;
using Sowfin.Data.Abstract;
using Sowfin.Model;

namespace Sowfin.Data.Repositories
{
    public class FindataRepository : EntityBaseRepository2<Findata>, IFindataRepository
    {
        public FindataRepository(FindataContext context) : base(context) { }




        /*public bool isEmailUniq (string email) {
            var user = this.GetSingle(u => u.Email == email);
            return user == null;
        }

        public bool IsUsernameUniq (string username) {
            var user = this.GetSingle(u => u.Username == username);
            return user == null;
        }*/
    }
}
