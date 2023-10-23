
using System.Threading.Tasks;

namespace ExcelInformer.Clients
{
    public interface IChannelClient // makes implementation of different db types easier
    {
        Task<string> Read(string address); // returns null if read failure
        Task<bool> Connect(string url); // false if connection failed, url only based 
        Task<bool> Connect(string url, string database); // false if connection failed, database based
        bool Connected(); // is channel connected?
    }
}
