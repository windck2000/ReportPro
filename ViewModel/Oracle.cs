using Oracle.ManagedDataAccess.Client;
using System.Data;

namespace ReportPro
{
    static class Oraclecnn
    {
        public static DataSet OracleCnn(string sql)
        {
            OracleConnection oracleConnection = new OracleConnection(@"Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.3.3)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME = MESDG)));User Id=SAJET;Password=tech;");
            oracleConnection.Open();
            OracleCommand oracleCommand = new OracleCommand(sql, oracleConnection);
            oracleCommand.CommandText = sql;
            OracleDataAdapter oracleDataAdapter = new OracleDataAdapter(oracleCommand);
            DataSet dataSet = new DataSet();
            oracleDataAdapter.Fill(dataSet);
            oracleConnection.Close();
            return dataSet;

        }
    }
}
