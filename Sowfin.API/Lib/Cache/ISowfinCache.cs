using System;
using System.Collections.Generic;

namespace Sowfin.API.Lib
{
    public interface ISowfinCache
    {
        void Set<T>(string key, T value);

        void Set<T>(string key, T value, TimeSpan timeout);

        //IList GetAll<IList>(string key);

        T Get<T>(string key);

        bool Remove(string key);

        bool IsInCache(string key);
    }
}