using System;
using System.Configuration;
using ServiceStack.Redis;
using Sowfin.API.Lib;

namespace Sowfin.API.Lib
{
    public class SowfinCache : ISowfinCache
    {
        private RedisEndpoint _endPoint {get; set;}

        public SowfinCache(string redisServer, string redisPort)
        {
            _endPoint = new RedisEndpoint(redisServer, Convert.ToInt32(redisPort));
        }

        public void Set<T>(string key, T value)
        {
            this.Set(key, value, TimeSpan.Zero);
        }

        public void Set<T>(string key, T value, TimeSpan timeout)
        {
            using (RedisClient client = new RedisClient(_endPoint))
            {
                client.As<T>().SetValue(key, value, timeout);
            }
        }

        public T Get<T>(string key)
        {
            T result = default(T);

            using (RedisClient client = new RedisClient(_endPoint))
            {
                var wrapper = client.As<T>();

                result = wrapper.GetValue(key);
            }

            return result;
        }

        //public IList GetAll<IList>(string keyPattern)
        //{
        //    IList result = default(IList);
        //    var allKeyPattern = keyPattern + "*";
        //    using (RedisClient client = new RedisClient(_endPoint))
        //    {
        //        var wrapper = client.As<IList>();

        //        result = wrapper.SearchKeys(allKeyPattern).GetValue;
        //    }

        //    return result;

        //}

        public bool Remove(string key)
        {
            bool removed = false;

            using (RedisClient client = new RedisClient(_endPoint))
            {
                removed = client.Remove(key);
            }

            return removed;
        }

        public bool IsInCache(string key)
        {
            bool isInCache = false;

            using (RedisClient client = new RedisClient(_endPoint))
            {
                isInCache = client.ContainsKey(key);
            }

            return isInCache;
        }


    }
}
