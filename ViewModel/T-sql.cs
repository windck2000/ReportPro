using System.Data;
using System.Data.SqlClient;

namespace ReportPro.ViewModel
{
    class T_sql
    {
        public static DataSet Sqldata(string S)
        {
            DataSet dataSet = new DataSet();
            SqlConnection sqlConnection = new SqlConnection("Server=192.168.2.37;Database=factory;uid=sa;pwd=DSC@dsc");
            {

                SqlCommand sqlCommand = new SqlCommand(S, sqlConnection);
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                sqlDataAdapter.Fill(dataSet);
            }
            return dataSet;
        }
    }
}
