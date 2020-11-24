using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using System.Globalization;
using _excel = Microsoft.Office.Interop.Excel;

namespace CargaRemitos
{
    public partial class Form1 : Form
    {
        
        _Application excel = new _excel.Application();
        Workbook libro;
        Worksheet hoja;
        Range rango;
       

        const int tam = 300000;
        NotaPedido[] vp = new NotaPedido[tam];
        


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            habilitar();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        public string Excelruta;
        //BOTON ARCHIVO
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult resultado = openFileDialog1.ShowDialog();
            if (resultado == DialogResult.OK)
            {
                Excelruta = openFileDialog1.FileName;
                
                textBox1.Text = Excelruta;

            }

            if(textBox1.Text== Excelruta)
            {
                button2.Enabled = true;
                button2.Focus();
            }
        }






       

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

            string conexion = comboBox2.ToString();
            switch (conexion)
            {
                case "Empresa1":
                    

                    break;
                case "Empresa2":
                    
                    break;




            }
        }
        
        public string cadenaconexion;
        public string empresa;
        public string codigocliente;
        public string resultado;
        public string codigoarti;



        /*BOTON VERIFICAR
            se configura la cadena de conexion dependiendo el puesto que tendra la herramienta.*/
        private void button2_Click(object sender, EventArgs e)
        {
            comboBox2.Enabled = true;
            btnCancelar.Enabled = true;           
                       
            libro = excel.Workbooks.Open(Excelruta);
            
            hoja = (Worksheet)libro.Worksheets.Item[1];
            rango = hoja.UsedRange;
                     
            
            AccesoBase cargar = new AccesoBase("dsn = Firebird; Uid = SYSDBA; Pwd = 3122414422");                       

            int rows = 2;



            try
            {
                for (int row = 2; row <= rows; row++)
                {
                    NotaPedido n = new NotaPedido();


                    if (hoja.Cells[row, 1].value2 != null)
                    {

                        n.pcodigoarticulo = (rango.Cells[row, 1] as Range).Value2.ToString();
                        n.pcantidad = Convert.ToDouble((rango.Cells[row, 3] as Range).Value2.ToString());
                        DataGridViewRow fila = new DataGridViewRow();
                        fila.CreateCells(listBox1);
                        fila.Cells[0].Value = n.pcodigoarticulo;
                        fila.Cells[1].Value = n.pcantidad.ToString();
                        listBox1.Rows.Add(fila);





                    }
                    else
                    {

                        MessageBox.Show("Verificacion Completa items:"+(rows-2).ToString(), "EMPRESA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;

                    }



                    rows++;
                    button3.Enabled = true;
                    button2.Enabled = true;
                    comboBox2.Focus();
                }

                libro.Close();
                excel.Quit();

               

            }
            catch (Exception ex)
            {

                MessageBox.Show("Error en archivo Excel, fila " + rows.ToString() + ", código articulo y/o cantidad vacío o formato incorrecto. Verificar e intentar nuevamente", "Flexxus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                libro.Close();
                excel.Quit();
                listBox1.Rows.Clear();
                textBox1.Clear();
                habilitar();

            }
            //VERIFICA SI EXISTE CODIGO ARTICULO DE ARCHIVO EXCEL
            for (int rowe = 0; rowe < listBox1.RowCount; rowe++)
            {
                NotaPedido na = new NotaPedido();

                AccesoBase cargar2 = new AccesoBase("dsn=Firebird;Uid=SYSDBA;Pwd=3122414422");
           
                na.pcodigoarticulo = listBox1.Rows[rowe].Cells[0].Value.ToString();
                try { 
                    if (na.pcodigoarticulo.ToString() != null)
                    {

                        cargar2.leerTabla(na.pcodigoarticulo.ToUpper());
                        cargar2.pdr.Read();

                        codigoarti = cargar2.pdr.GetString(0);
                        resultado = cargar2.pdr.GetString(1);
                        
                        
                    }
                    else
                    {
                        MessageBox.Show("codigo no existe");
                        
                    }
                }
                
                catch
                    {
                        if (codigoarti != null)
                        { break; }
                        else
                    {
                        MessageBox.Show("Error en archivo Excel, fila " + (rowe+2).ToString() + ", código articulo no existe. Verificar e intentar nuevamente", "Flexxus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        listBox1.Rows.Clear();
                        textBox1.Clear();
                        habilitar();
                    }
                    }
                codigoarti = null;
            }

          




        }

        string codigodeposito;
        private string cadenaconexion2;

        //BOTON EJECUTAR
        private void button3_Click(object sender, EventArgs e)
        {
            if (validar())
            {
                if (comboBox2.SelectedIndex == 0)
                {
                    cadenaconexion = "dsn=Firebird;Uid=SYSDBA;Pwd=3122414422";
                    cadenaconexion2 = "dsn=Firebird1;Uid=SYSDBA;Pwd=3122414422";
                    empresa = "Empresa2";
                    codigocliente = "15874";
                    codigodeposito = "002";

                }
                else if (comboBox2.SelectedIndex == 1)
                {
                    cadenaconexion = "dsn=Firebird1;Uid=SYSDBA;Pwd=3122414422";
                    cadenaconexion2 = "dsn=Firebird;Uid=SYSDBA;Pwd=3122414422";
                    empresa = "Empresa1";
                    codigocliente = "20343";
                    codigodeposito = "005";
                }



                AccesoBase cargar = new AccesoBase(cadenaconexion);


                for (int row = 0; row < listBox1.RowCount; row++)
                {
                    NotaPedido n = new NotaPedido();
                    string sql;


                    n.pcodigoarticulo = listBox1.Rows[row].Cells[0].Value.ToString();
                    n.pcantidad = Convert.ToDouble(listBox1.Rows[row].Cells[1].Value.ToString());
                    try
                    {
                        cargar.leerTabla(n.pcodigoarticulo.ToUpper());
                        cargar.pdr.Read();
                        n.pcodigounico = cargar.pdr.GetString(0);
                        n.pdescripcion = cargar.pdr.GetString(1);
                    }
                    catch
                    {
                        MessageBox.Show("Error en archivo Excel, fila " + (row + 2).ToString() + ", código articulo no existe en empresa de destino. Verificar e intentar nuevamente", "Flexxus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    cargar.pdr.Close();
                    cargar.desconectar();
                    cargar.leerTabla2();
                    cargar.pdr2.Read();
                    n.pvaloRI = cargar.pdr2.GetDouble(0);
                    cargar.pdr2.Close();
                    cargar.desconectar();


                    double pasalo = Math.Truncate(n.pcantidad * 1000) / 1000;
                    NumberFormatInfo nf = new CultureInfo("es-Es", false).NumberFormat;
                    nf.NumberDecimalSeparator = ".";
                    string can = n.pcantidad.ToString("", nf);
                    string cant = string.Format(can, nf, pasalo);

                    DateTime fecha = DateTime.Now;

                    string fec = fecha.ToString("dd.MMM.yyyy", CultureInfo.InvariantCulture);
                    sql = "INSERT INTO CABEZACOMPROBANTES VALUES ('RI', " + n.pvaloRI + ", '" + codigocliente + "', '" + fec + "', '" + empresa + "', 'Calle 1', 21, 0, 0, 0, 0, 0, 1, '1-JAN-1900 14:58:51', '-', 'CF', 0, '', '0', '" + fec + "', 0, 0, '  -        -', 276274.86656, 1, 0, 3, 0, 0, '- -', 1, '" + fec + "', 'RI', 'PESOS', 1, 0, 0, 0, '" + codigocliente + "', '" + empresa + "', '-1', '-', 0, 1, 1, -1, -1, 0, " + n.pvaloRI + ", 2, 'CASA CENTRAL', '001', 1, '.CONTADO', 0, 0, 0, 0, 0, 0, '-1', 1, 0, 0, 0)";
                    cargar.actualizar(sql);
                    sql = "INSERT INTO CUERPOCOMPROBANTES  VALUES ('RI'," + n.pvaloRI + ", 1, '" + n.pcodigounico + "', '" + n.pdescripcion + "', " + cant + ", 0, 0, 0, 0, 0, 0, '000', 0, '" + fec + "', '" + codigodeposito + "', 00, 0, '" + n.pcodigoarticulo + "', 0, 0, 0.00, 0, 00, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)";
                    cargar.actualizar(sql);
                    sql = "Update CONTADORES set VALOR = VALOR + " + 1 + " where DESCRIPCION ='REMITO INTERNO'";
                    cargar.actualizar(sql);
                    sql = "UPDATE STOCK  SET STOCKACTUAL  = STOCKACTUAL - " + cant + " where CODIGOARTICULO='" + n.pcodigounico + "'";
                    cargar.actualizar(sql);
                    sql = "UPDATE CASILLEROS  SET STOCKACTUAL = STOCKACTUAL - " + cant + " where CODIGOARTICULO=" + n.pcodigounico + " and CODIGODEPOSITO = '" + codigodeposito + "'";
                    cargar.actualizar(sql);
                    cargar.desconectar();
                    cargar.pdr.Close();
                    cargar.pdr2.Close();

                    


                }
                Cargar ca = new Cargar();
                ca.cargar(cadenaconexion, cadenaconexion2);



                limpiar();
                habilitar();
                
                MessageBox.Show("Carga en destino OK", "Flexxus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                
                
               

            }
            }
         
            private bool validar()
        {
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccionar Empresa de destino");
                comboBox2.Focus();
                return false;

            }

            if (openFileDialog1 == null)
            {
                MessageBox.Show("Seleccionar ruta de destino");
                button1.Focus();
                return false;
            }
            return true;
        }


        private void limpiar()
        {
            comboBox2.SelectedIndex = -1;
            listBox1.Rows.Clear();
            textBox1.Text = "";
             

        }
        private void habilitar()
        {
            comboBox2.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            btnCancelar.Enabled = false;
            
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Desea cancelar?", "EMPRESA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                limpiar();
                habilitar();
        }
            else {
                return;
        }
        }
   
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Desea salir de la aplicacion?", "EMPRESA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
            
        }
    }





    }


    

  

    
    


