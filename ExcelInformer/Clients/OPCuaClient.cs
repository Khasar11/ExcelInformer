using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Workstation.ServiceModel.Ua;
using Workstation.ServiceModel.Ua.Channels;

namespace ExcelInformer.Clients
{
    internal class OPCuaClient : IChannelClient
    {

        private ApplicationDescription clientDescription = new ApplicationDescription
        {
            ApplicationName = "Excel",
            ApplicationUri = $"urn:{System.Net.Dns.GetHostName()}:Excel",
            ApplicationType = ApplicationType.Client
        };

        private DirectoryStore certificateStore = new DirectoryStore(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Workstation.Excel", "pki"));

        private ClientSessionChannel client;

        public async Task<bool> Connect(string url)
        {
            try
            {
                client = new ClientSessionChannel(
                                    clientDescription,
                                    certificateStore,
                                    new AnonymousIdentity(), // no user identity, handle later
                                    url,
                                    SecurityPolicyUris.Basic256);

                await client.OpenAsync();
                return true;
            } 
            catch (Exception e) {
                return false;
            }
        }

        #region not implemented
        public Task<bool> Connect(string url, string database)
        {
            throw new NotImplementedException();
        }
        #endregion not implemented

        public bool Connected()
        {
            return (client.State == CommunicationState.Opened);
        }

        public async Task<string> Read(string address)
        {
            ReadRequest request = new ReadRequest()
            {
                NodesToRead = new[] 
                {
                    new ReadValueId()
                    {
                        NodeId = NodeId.Parse(address),
                        AttributeId = AttributeIds.Value
                    }
                }
            };
            try
            {
                ReadResponse response = await client.ReadAsync(request);
                return response.Results[0].Value+"";
            } catch (Exception e)
            {
                return e+"";
            }
        }
    }
}
