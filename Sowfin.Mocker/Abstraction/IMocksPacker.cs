using System.Collections.Generic;
using System.Threading.Tasks;
using Sowfin.Mocker;

namespace Sowfin.Mocker.Abstraction
{
    public interface IMocksPacker
    {
        Task<Pack> GetPack(List<string> usernames);
    }
}