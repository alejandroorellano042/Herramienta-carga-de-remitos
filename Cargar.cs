using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace CargaRemitos
{
    class Cargar
    {

        public void cargar(string cadena1, string cadena2)
        {

            const int tam2 = 500000;
            Comprobante[] va = new Comprobante[tam2];

            const int tam = 500000;
            RemitoTemp[] vp = new RemitoTemp[tam];


            AccesoBase Base1 = new AccesoBase(cadena1);//origen
            AccesoBase Base2 = new AccesoBase(cadena2);//destino



            int c = 0;
            string sql;
            Base2.leerTabla4();


            DateTime fecha = DateTime.Now;

            if (Base2.pdr2.Read())
            {

                Comprobante l = new Comprobante();
                if (!Base2.pdr2.IsDBNull(0))
                    l.pDocumento = Base2.pdr2.GetString(0);
                if (!Base2.pdr2.IsDBNull(1))
                    l.pValor = Base2.pdr2.GetDouble(1);

                Base2.desconectar();


                Base1.leerTabla3();
                while (Base1.pdr.Read())
                {

                    RemitoTemp p = new RemitoTemp();
                    if (!Base1.pdr.IsDBNull(0))
                        p.ptipoComprobante = Base1.pdr.GetString(0);
                    else
                        break;
                    p.pnumeroComprobante = Base1.pdr.GetDouble(1);
                    p.plinea = Base1.pdr.GetDouble(2);
                    p.pcodigoArticulo = Base1.pdr.GetString(3);
                    p.pdescripcion = Base1.pdr.GetString(4);
                    p.pcantidad = Base1.pdr.GetDouble(5);
                    p.pdescuento = Base1.pdr.GetDouble(6);
                    p.pprecioUnitario = Base1.pdr.GetDouble(7);
                    p.pprecioTotal = Base1.pdr.GetDouble(8);
                    p.pgarantia = Base1.pdr.GetDouble(9);
                    p.pinteres = Base1.pdr.GetDouble(10);
                    p.pcantidadRemitida = Base1.pdr.GetDouble(11);
                    p.plote = Base1.pdr.GetString(12);
                    p.pesConjunto = Base1.pdr.GetInt32(13);
                    p.pfechaModificacion = Base1.pdr.GetDate(14);
                    p.pcodigousuario = Base1.pdr.GetString(15);



                    double pasalo = Math.Truncate(p.pcantidad * 1000) / 1000;
                    NumberFormatInfo nf = new CultureInfo("es-Es", false).NumberFormat;
                    nf.NumberDecimalSeparator = ".";
                    string can = p.pcantidad.ToString("", nf);
                    string cant = string.Format(can, nf, pasalo);



                    string fec = p.pfechaModificacion.ToString("dd.MMM.yyyy", CultureInfo.InvariantCulture);




                    sql = "insert into CORRECCIONESSTOCKMANUALES  values (" + l.pValor + "," + "'" + fec + "'" + ",'"
                                                                            + p.pcodigousuario + "'," + cant + ", 0 ,'"
                                                                            + p.pcodigoArticulo
                                                                            + "','000','STOCK ACTUAL','Transferencia Empresa1-Empresa2 remito:" + p.pnumeroComprobante + "', " + "'"
                                                                            + fec + "'" + ",'001', 0 ," + "'" + fec + "'" + ")";



                    Base2.actualizar(sql);
                    l.pValor = l.pValor + 1;
                    sql = "UPDATE PARAMETROS p SET p.VALOR =" + l.pValor + " where tipodocumento ='CS'";
                    Base2.actualizar(sql);




                    sql = "UPDATE STOCK  SET STOCKACTUAL  = STOCKACTUAL + " + can + " where CODIGOARTICULO=" + p.pcodigoArticulo;

                    Base2.actualizar(sql);


                    sql = "UPDATE CASILLEROS c  SET c.STOCKACTUAL = c.STOCKACTUAL + " + can + " where c.CODIGOARTICULO=" + p.pcodigoArticulo + "and CODIGODEPOSITO = '001'";
                    Base2.actualizar(sql);


                    vp[c] = p;
                    c++;



                }

            }



            Base1.pdr.Close();
            Base1.desconectar();
            Base2.desconectar();

            sql = "delete from TEMP_REMITOS";
            Base1.actualizar(sql);

            Base1.pdr.Close();
            Base1.desconectar();

        }
    }
}
