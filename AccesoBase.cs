using System.Data;
using System.Data.Odbc;

namespace CargaRemitos
{
    class AccesoBase
    {

        OdbcDataReader dr2;
        OdbcConnection conexion;
        OdbcCommand comando;
        OdbcDataReader dr;
        DataTable dt;
        string cadenaConexion;

        public OdbcDataReader pdr
        {
            get { return dr; }
            set { dr = value; }
        }
        public OdbcDataReader pdr2
        {
            get { return dr2; }
            set { dr2 = value; }
        }

        public string pcadenaconexion
        {
            get { return cadenaConexion; }
            set { cadenaConexion = value; }
        }

        public OdbcCommand pcomando
        {
            get { return comando; }
            set { comando = value; }

        }
        public AccesoBase()
        {
            conexion = new OdbcConnection();
            comando = new OdbcCommand();
            dt = new DataTable();
            dr = null;
            cadenaConexion = "";
        }

        public AccesoBase(string strconexion)
        {
            conexion = new OdbcConnection(strconexion);
            comando = new OdbcCommand();
            dt = new DataTable();
            dr = null;
            cadenaConexion = strconexion;

        }

        public void conectar()
        {
            conexion.ConnectionString = cadenaConexion;
            conexion.Open();
            comando.Connection = conexion;
            comando.CommandType = CommandType.Text;

        }
                

        public void comand()
        {
            comando = new OdbcCommand();


        }
        public void desconectar()
        {
            conexion.Close();
            conexion.Dispose();

        }

      
        public void actualizar(string sql)
        {
            conectar();
            comando.CommandText = sql;
            comando.ExecuteNonQuery();
            desconectar();

        }
        public void leerTabla4()
        {

            conectar();
            comando.CommandText = "select * from PARAMETROS where TIPODOCUMENTO = 'CS' ";
            dr2 = comando.ExecuteReader();

        }
        public void leerTabla3()
        {

            conectar();
            comando.CommandText = "select * from TEMP_REMITOS";
            dr = comando.ExecuteReader();

        }

        public void leerTabla(string art)
        {

            conectar();
            comando.CommandText = "select CODIGOARTICULO , DESCRIPCION from ARTICULOS where CODIGOPARTICULAR = '"+art+"'";
            dr = comando.ExecuteReader();

        }


        public void leerTabla2()
        {

            conectar();
            comando.CommandText = "select VALOR from CONTADORES where DESCRIPCION = 'REMITO INTERNO' ";
            dr2 = comando.ExecuteReader();

        }

      

    }
}
