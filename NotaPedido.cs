using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CargaRemitos
{
    class NotaPedido
    {
        string codigoarticulo;

        public string pcodigoarticulo
        {
            get { return codigoarticulo; }
            set { codigoarticulo = value; }

        }

        double cantidad;

        public double pcantidad
        {
            get { return cantidad; }
            set { cantidad = value; } 

        }

        string codigounico;

        public string pcodigounico
        {
            get { return codigounico; }
            set { codigounico = value; }
        }

        double valorRI;

        public double pvaloRI
        {
            get { return valorRI; }
            set { valorRI = value; }
        }

        string descripcion;
        public string pdescripcion
        {
            get { return descripcion; }
            set { descripcion = value; }
        }


    }
}
