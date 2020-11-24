using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CargaRemitos
{
    class Comprobante
    {

        string Documento;

        public string pDocumento
        {
            get { return Documento; }
            set { Documento = value; }
        }

        double Valor;

        public double pValor
        {
            get { return Valor; }
            set { Valor = value; }
        }

    }
}
