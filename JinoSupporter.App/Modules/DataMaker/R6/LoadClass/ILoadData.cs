using System.Data;

namespace DataMaker.R6.LoadClass
{
    public interface ILoadData
    {
        public Task<DataTable> Load(string fPath, string[] Header, char delimiter = '\t');
    }
}
