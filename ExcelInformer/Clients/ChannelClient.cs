
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelInformer.Clients
{
    public class ChannelClient
    {
        private IChannelClient _channelClient;
        public string Url;

        public ChannelClient(IChannelClient channelClient)
        {
            _channelClient = channelClient;
        }

        public async Task<bool> Connect(string url)
        {
            Url = url;
            return await _channelClient.Connect(url);
        }

        public async Task<bool> Connect(string url, string database)
        {
            Url = url;
            return await _channelClient.Connect(url, database);
        }

        public async Task<string> Read(string address)
        {
            return await _channelClient.Read(address);
        }

        public bool Connected()
        {
            return _channelClient.Connected();
        }
    }
}
