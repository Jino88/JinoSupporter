using System.Data;

namespace DataMaker.R6.SaveClass
{
    public interface ISaveData
    {
        public Task Save(DataTable Table, string path);
        public Task Save(List<DataTable> Tables, string path);

    }

}
