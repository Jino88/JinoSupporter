using DataMaker.R6.SQLProcess;
using System.Data;

namespace DataMaker.R6.PreProcessor
{
    public interface IMakeTableFromTxt
    {

        public void Make(string TableName, string txtPath, List<string> Columns);
        public DataTable MakeDataTable(string txtPath, List<string> Columns);
    }
}
