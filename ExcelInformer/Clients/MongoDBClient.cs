using MongoDB.Driver;
using System;
using System.Threading.Tasks;

namespace ExcelInformer.Clients
{
    internal class MongoDBClient : IChannelClient
    {


        private MongoClient mongoClient;
        private IMongoDatabase ConnectedDatabase;

        public string Url => throw new NotImplementedException();

        public string Database => throw new NotImplementedException();

        #region not implemented
        public Task<bool> Connect(string url)
        {
            throw new NotImplementedException();
        }
        #endregion not implemented

        public Task<bool> Connect(string url, string database)
        {
            throw new NotImplementedException();
        }

        public bool Connected()
        {
            throw new NotImplementedException();
        }

        public Task<string> Read(string address)
        {
            throw new NotImplementedException();
        }
    }
}
