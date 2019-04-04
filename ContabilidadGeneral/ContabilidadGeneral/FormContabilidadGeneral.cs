/*
    Guillermo Arriaga Garcia: Sistemas de Informacion Empresarial: Sistema Contable Basico.  Julio 2018
    Tiempo de programacion y diseno: 14 horas.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ContabilidadGeneral
{
    public partial class FormContabilidadGeneral : Form
    {
        string[] CuentasActivoCirculante;
        string[] CuentasActivoFijo;
        string[] CuentasActivoDiferido;
        string[] CuentasPasivoCortoPlazo;
        string[] CuentasPasivoLargoPlazo;
        string[] CuentasPasivoDiferido;
        string[] CuentasCapital;
        string[] CuentasEstadoResultados;

        /*
    "BalanceGral"
    "EdoResultados"

    "Caja",
    "Bancos",
    "Inversiones Temporales",
    "Inventario, Mercancias, Almacen",
    "Clientes",
    "Documentos por Cobrar",
    "Deudores Diversos",
    "Anticipo a Proveedores"
    "Terrenos",
    "Edificios",
    "Mobiliario y Equipo",
    "Equipo de Computo",
    "Equipo de Entrega",
    "Deposito en Garantia",
    "Inversiones Permanentes"
    "Gastos Investigacion y Desarrollo",
    "Gastos Preoperativos de Org-Admn",
    "Gastos de Mercadotecnia",
    "Gastos de Organizacion",
    "Gastos de Instalacion",
    "Papeleria y Utiles",
    "Propaganda y Publicidad",
    "Primas de Seguros",
    "Rentas Pagadas por Anticipado",
    "Intereses Pagados por Anticipado"

    "Proveedores",
    "Documentos por Pagar",
    "Acreedores Diversos",
    "Anticipo de Clientes",
    "Gastos Pendientes de Pago",
    "Impuestos Pendientes de Pago"
    "Acreedores Hipotecarios",
    "Documentos por Pagar a Largo Plazo",
    "Cuentas por Pagar a Largo Plazo"
    "Rentas Cobradas por Anticipado",
    "Intereses Cobrados por Anticipado"

    "Capital Social",
    "Utilidad Antes de Impuestos"

    "Ventas Netas",
    "Inventario Inicial",
    "Compras",
    "Inventario Final",
    "Gastos de Venta",
    "Gastos de Administracion",
    "Gastos Financieros",
    "Otros Productos"
    */

        public FormContabilidadGeneral()
        {
            InitializeComponent();

            CuentasActivoCirculante = new string[] 
            {
                "Caja",
                "Bancos",
                "Inversiones Temporales",
                "Clientes",
                "Docs. por Cobrar",
                "Deudores Diversos",
                "Anticipo a Proveedores",
                "InventarioMercancAlmacen"  // Si se agregan cuentas, que inventario quede al final, pues se pondra directamente su valor
            };
            CuentasActivoFijo = new string[]
            {
                "Terrenos",
                "Edificios",
                "Mobiliario y Equipo",
                "Equipo de Computo",
                "Equipo de Entrega",
                "Deposito en Garantia",
                "Inversiones Permanentes"
            };
            CuentasActivoDiferido = new string[]
            {
                "Gastos Invest. y Desar.",
                "Gastos Preoper.Org-Admn",
                "Gastos de Mercadotecnia",
                "Gastos de Organizacion",
                "Gastos de Instalacion",
                "Papeleria y Utiles",
                "Propaganda y Publicidad",
                "Primas de Seguros",
                "Rentas Pagadas Anticip.",
                "Intereses Pagados Antic."
            };
            CuentasPasivoCortoPlazo = new string[]
            {
                "Proveedores",
                "Docs. por Pagar",
                "Acreedores Diversos",
                "Anticipo de Clientes",
                "Gastos PagoPendiente",
                "Impuestos PagoPendiente"
            };
            CuentasPasivoLargoPlazo = new string[]
            {
                "Acreedores Hipotecarios",
                "Docs. porPagar L Plazo",
                "Cuentas porPagar L Plazo"
            };
            CuentasPasivoDiferido = new string[]
            {
                "Rentas CobradasAnticip.",
                "Intereses CobradosAntic."
            };
            CuentasCapital = new string[]
            {
                "Capital Social",
                "Utilidad AntesImpuestos"
            };
            CuentasEstadoResultados = new string[]
            {
                "Ventas",
                // "Inventario Inicial",  // Se toma la cuenta de activo correspondiente
                "Compras",
                //"Inventario Final",     // Se recibe del usuario
                "Gastos de Venta",
                "Gastos de Admin.",
                "Gastos Financieros",
                "Otros Productos"
            };

            // Ajuste de fechas por defecto a un mes hasta hoy
            dtpFechaFin.Value = DateTime.Today;
            dtpFechaInicio.Value = DateTime.Today.AddMonths(-1);

            // Deshabilitar ordenamiento de columnas
            foreach (DataGridViewColumn column in dgvBalanceGeneral.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
            foreach (DataGridViewColumn column in dgvCuentasT.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
            foreach (DataGridViewColumn column in dgvEdoResultados.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
            foreach (DataGridViewColumn column in dgvRazonesFinancieras.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void VaciadoDGV()
        {
            dgvRazonesFinancieras.Rows.Clear();
            dgvEdoResultados.Rows.Clear();
            dgvBalanceGeneral.Rows.Clear();
            dgvCuentasT.Rows.Clear();
        }

        private void cbTipoCuenta_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbCuenta.Items.Clear();
            if (cbTipoCuenta.SelectedIndex == 1)
            {
                cbCuenta.Items.AddRange(CuentasEstadoResultados);
                cbCuenta.SelectedIndex = 0;
            }
            else
            {
                cbCuenta.Items.AddRange(CuentasActivoCirculante);
                cbCuenta.Items.AddRange(CuentasActivoFijo);
                cbCuenta.Items.AddRange(CuentasActivoDiferido);
                cbCuenta.Items.AddRange(CuentasPasivoCortoPlazo);
                cbCuenta.Items.AddRange(CuentasPasivoLargoPlazo);
                cbCuenta.Items.AddRange(CuentasPasivoDiferido);
                cbCuenta.Items.AddRange(CuentasCapital);
                cbCuenta.SelectedIndex = 0;
            }
            //cbCuenta.DroppedDown = true;
        }

        private void btInfo_Click(object sender, EventArgs e)
        {
            MessageBox.Show("SISTEMA DE CONTABILIDAD NO DETALLADA con fines pedagogicos.\n\n\n" +
                "El usuario debe ingresar la informacion de las cuentas de diario de modo que en cada asiento contable correspondan las sumas de su debe y de su haber.\n" +
                "Se recomienda suministrar el nombre de la empresa y las fechas solicitadas.\n\n" +
                "En el asiento del balance inicial contiene el valor de inventario inicial, sin embargo, se pide que se ponga tambien este valor en la seccion de inventario (inicial , final) debajo de la tabla a llenar.\n\n" +
                "El sistema calcula las cuentas T, el estado de resultados y el balance general no detallados, ademas de brindar razones financieras estandar.\n\n" +
                "Incluye un programa para realizar la contabilidad de costos de una empresa con la posibilidad de calcular varias proyecciones de venta con varios porcentajes de ganancia cada una.\n\n" +
                "Programa realizado por Guillermo Arriaga Garcia en julio de 2018. Contacto en guillermoarriagag@gmail.com"
            );
        }

        private void btEjemplo_Click(object sender, EventArgs e)
        {
            tbEmpresa.Text = "Comercializadora SA de CV";
            dtpFechaFin.Value = DateTime.Today;
            dtpFechaInicio.Value = DateTime.Today.AddMonths(-1);


            dgvAsientosContables.Rows.Clear();

            dgvAsientosContables.Rows.Add("a", "x", "BalanceGral", "Caja", "8,000.00", "");
            dgvAsientosContables.Rows.Add("a", "x", "BalanceGral", "Bancos", "185,000.00", "");
            dgvAsientosContables.Rows.Add("a", "x", "BalanceGral", "InventarioMercancAlmacen", "220,000.00", "");
            dgvAsientosContables.Rows.Add("a", "x", "BalanceGral", "Docs. por Cobrar", "58,000.00", "");
            dgvAsientosContables.Rows.Add("a", "x", "BalanceGral", "Equipo de Entrega", "300,000.00", "");
            dgvAsientosContables.Rows.Add("a", "x", "BalanceGral", "Proveedores", "",  "115,000.00");
            dgvAsientosContables.Rows.Add("a", "x", "BalanceGral", "Capital Social", "", "656,000.00");

            dgvAsientosContables.Rows.Add("b", "x", "EdoResultados", "Compras", "75,000.00", "");
            dgvAsientosContables.Rows.Add("b", "x", "BalanceGral", "Bancos", "", "40,000.00");
            dgvAsientosContables.Rows.Add("b", "x", "BalanceGral", "Proveedores", "", "35,000.00");

            dgvAsientosContables.Rows.Add("c", "x", "EdoResultados", "Gastos de Venta", "8,400.00", "");
            dgvAsientosContables.Rows.Add("c", "x", "EdoResultados", "Gastos de Admin.", "3,600.00", "");
            dgvAsientosContables.Rows.Add("c", "x", "BalanceGral", "Bancos", "", "12,000.00");

            dgvAsientosContables.Rows.Add("d", "x", "EdoResultados", "Gastos de Venta", "400.00", "");
            dgvAsientosContables.Rows.Add("d", "x", "BalanceGral", "Caja", "", "400.00");

            dgvAsientosContables.Rows.Add("e", "x", "BalanceGral", "Bancos", "80,000.00", "");
            dgvAsientosContables.Rows.Add("e", "x", "BalanceGral", "Docs. por Cobrar", "35,000.00", "");
            dgvAsientosContables.Rows.Add("e", "x", "BalanceGral", "Equipo de Entrega", "", "105,000.00");
            dgvAsientosContables.Rows.Add("e", "x", "EdoResultados", "Otros Productos", "", "10,000.00");

            dgvAsientosContables.Rows.Add("f", "x", "BalanceGral", "Bancos", "50,000.00", "");
            dgvAsientosContables.Rows.Add("f", "x", "BalanceGral", "Clientes", "30,000.00", "");
            dgvAsientosContables.Rows.Add("f", "x", "EdoResultados", "Ventas", "", "80,000.00");

            dgvAsientosContables.Rows.Add("g", "x", "BalanceGral", "Bancos", "17,200.00", "");
            dgvAsientosContables.Rows.Add("g", "x", "EdoResultados", "Ventas", "", "16,300.00");
            dgvAsientosContables.Rows.Add("g", "x", "EdoResultados", "Otros Productos", "", "900.00");

            dgvAsientosContables.Rows.Add("h", "x", "BalanceGral", "Bancos", "31,500.00", "");
            dgvAsientosContables.Rows.Add("h", "x", "BalanceGral", "Docs. por Cobrar", "", "30,000.00");
            dgvAsientosContables.Rows.Add("h", "x", "EdoResultados", "Otros Productos", "", "1,500.00");

            dgvAsientosContables.Rows.Add("i", "x", "BalanceGral", "Proveedores", "40,000.00", "");
            dgvAsientosContables.Rows.Add("i", "x", "EdoResultados", "Gastos Financieros", "1,200.00", "");
            dgvAsientosContables.Rows.Add("i", "x", "BalanceGral", "Bancos", "", "41,200.00");

            dgvAsientosContables.Rows.Add("j", "x", "BalanceGral", "Caja", "1,500.00", "");
            dgvAsientosContables.Rows.Add("j", "x", "EdoResultados", "Otros Productos", "", "1,500.00");

            tbInventarioInicial.Text = "220,000.00";
            tbInventarioFinal.Text = "240,000.00";

            btLimpiarAsiento_Click(null,null);
            /*
                "Caja",
                "Bancos",
                "Inversiones Temporales",
                "Clientes",
                "Docs. por Cobrar",
                "Deudores Diversos",
                "Anticipo a Proveedores",
                "InventarioMercancAlmacen"
                "Terrenos",
                "Edificios",
                "Mobiliario y Equipo",
                "Equipo de Computo",
                "Equipo de Entrega",
                "Deposito en Garantia",
                "Inversiones Permanentes"
                "Gastos Invest. y Desar.",
                "Gastos Preoper.Org-Admn",
                "Gastos de Mercadotecnia",
                "Gastos de Organizacion",
                "Gastos de Instalacion",
                "Papeleria y Utiles",
                "Propaganda y Publicidad",
                "Primas de Seguros",
                "Rentas Pagadas Anticip.",
                "Intereses Pagados Antic."

                "Proveedores",
                "Docs. por Pagar",
                "Acreedores Diversos",
                "Anticipo de Clientes",
                "Gastos PagoPendiente",
                "Impuestos PagoPendiente"
                "Acreedores Hipotecarios",
                "Docs. porPagar L Plazo",
                "Cuentas porPagar L Plazo"
                "Rentas CobradasAnticip.",
                "Intereses CobradosAntic."

                "Capital Social",
                "Utilidad AntesImpuestos"

                "Ventas",
                "Compras",
                "Gastos de Venta",
                "Gastos de Admin.",
                "Gastos Financieros",
                "Otros Productos"
             
             */
        }

        private void btAbrirContaCostos_Click(object sender, EventArgs e)
        {
            string ruta = Application.StartupPath + "\\ContabilidadDeCostos.exe";

            if (System.IO.File.Exists(ruta))
            {
                System.Diagnostics.Process.Start(ruta);
            }
            else   // Si no existe lo saca (copia) de dentro de este mismo ejecutable pues esta embebido el solicitado
            {
                System.IO.File.WriteAllBytes(ruta, Properties.Resources.ContabilidadCostos);
                System.Diagnostics.Process.Start(ruta);
            }
        }

        private void btImprimirTexto_Click(object sender, EventArgs e)
        {
            // Creacion de Texto
            Font titulo1 = new Font("Consolas", 14.0f, FontStyle.Bold);
            Font titulo2 = new Font("Consolas", 10.0f, FontStyle.Bold);
            Font texto1 = new Font("Consolas", 8.0f, FontStyle.Regular);
            Font texto2 = new Font("Consolas", 7.0f, FontStyle.Regular);
            RichTextBox informe = new RichTextBox();

            informe.Font = texto2;
            informe.DeselectAll();

            informe.SelectionAlignment = HorizontalAlignment.Center;
            informe.SelectionFont = titulo1;
            informe.AppendText("________________| |   CONTABILIDAD   | |________________\n\n");
            informe.SelectionFont = titulo2;
            informe.AppendText("Asientos Contables, Cuentas T, \n" +
                "Estado de Resultados, Balance General y Razones Financieras.\n" +
                "______________________________________________________________________________\n\n\n\n\n");

            informe.DeselectAll();
            informe.SelectionFont = titulo2;
            informe.AppendText("Empresa: " + tbEmpresa.Text +"\n " +
                "Balance Inicial: " + dtpFechaInicio.Value.ToString("dd/MM/yyyy") + "  " +
                "Balance Final: " + dtpFechaFin.Value.ToString("dd/MM/yyyy")+ "\n\n\n\n\n");


            informe.DeselectAll();
            informe.SelectionFont = titulo1;
            informe.AppendText("* * *   ASIENTOS CONTABLES   * * *\n\n");
            informe.DeselectAll();
            informe.SelectionFont = texto2;
            informe.AppendText(DGVaSTRING(dgvAsientosContables)); // Metodo que convierte tabla a texto rellenando con espacios
            informe.AppendText("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");

            informe.DeselectAll();
            informe.SelectionFont = titulo1;
            informe.AppendText("* * *   CUENTAS T   * * *\n\n");
            informe.DeselectAll();
            informe.SelectionFont = texto2;
            informe.AppendText(DGVaSTRING(dgvCuentasT)); // Metodo que convierte tabla a texto rellenando con espacios
            informe.AppendText("\n\n\n");

            informe.DeselectAll();
            informe.SelectionFont = titulo1;
            informe.AppendText("* * *   ESTADO DE RESULTADOS   * * *\n\n");
            informe.DeselectAll();
            informe.SelectionFont = titulo2;
            informe.AppendText("Empresa: " + tbEmpresa.Text + "\nPeriodo: del " + dtpFechaInicio.Value.ToString("dd/MM/yyyy") + " al " + dtpFechaFin.Value.ToString("dd/MM/yyyy") + "\n\n");
            informe.DeselectAll();
            informe.SelectionFont = texto1;
            informe.AppendText(DGVaSTRING(dgvEdoResultados)); // Metodo que convierte tabla a texto rellenando con espacios
            informe.AppendText("\n\n\n");
            
            informe.DeselectAll();
            informe.SelectionFont = titulo1;
            informe.AppendText("* * *   BALANCE GENERAL   * * *\n\n");
            informe.DeselectAll();
            informe.SelectionFont = titulo2;
            informe.AppendText("Empresa: " + tbEmpresa.Text + "\nAl " + dtpFechaFin.Value.ToString("dd/MM/yyyy") + "\n\n");
            informe.DeselectAll();
            informe.SelectionFont = texto1;
            informe.AppendText(DGVaSTRING(dgvBalanceGeneral)); // Metodo que convierte tabla a texto rellenando con espacios
            informe.AppendText("\n\n");
            
            informe.DeselectAll();
            informe.SelectionFont = titulo1;
            informe.AppendText("* * *   RAZONES FINANCIERAS   * * *\n\n");
            informe.DeselectAll();
            informe.SelectionFont = texto2;
            informe.AppendText(DGVaSTRING(dgvRazonesFinancieras)); // Metodo que convierte tabla a texto rellenando con espacios


            // Tratamiento de archivo, nombre unico por fecha y hora
            try
            {
                string ruta = Application.StartupPath + "\\InformeContaGeneral" + DateTime.Now.ToString("-yyyyMMdd-HH'h'mm'm'ss's'") + ".rtf";
                if (System.IO.File.Exists(ruta)) System.IO.File.Delete(ruta);
                informe.SaveFile(ruta);
                System.Diagnostics.Process.Start(ruta);
            }
            catch
            {
                MessageBox.Show("Debe cerrar el archivo en Word para desplegar el actual.\n\nVuelva a intentarlo.");
            }
        }


        private void btnAgregarCuenta_Click(object sender, EventArgs e)
        {
            VaciadoDGV();
            
            // Revision de cantidad solo numerica
            try
            {
                Convert.ToDouble(tbCantidad.Text);
            }
            catch
            {
                MessageBox.Show("La cantidad no es solo numerica. \n\nVuelva a intentarlo.");
                return;
            }

            // Revision de datos necesarios suministrados
            if ( tbNumAsiento.Text == "" ||
                cbTipoCuenta.Text == "" ||
                cbCuenta.Text == "" ||
                tbCantidad.Text == ""
              )
            {
                MessageBox.Show("Falta algun dato necesario. \n\nVuelva a intentarlo.");
                return;
            }

            string cantidad1 = "", cantidad2 = "";
            if (rbDebe.Checked) cantidad1 = tbCantidad.Text;
            else cantidad2 = tbCantidad.Text;

            dgvAsientosContables.Rows.Add(
                tbNumAsiento.Text,
                tbComentario.Text,
                cbTipoCuenta.Text,
                cbCuenta.Text,
                cantidad1,
                cantidad2
            );

            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, dgvAsientosContables.Rows.Count - 1];
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[4], ListSortDirection.Descending);
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[0], ListSortDirection.Ascending);
            int fila = dgvAsientosContables.SelectedRows[0].Index;

            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, 0];    // Para que muestre la deseada habiendo un cambio de seleccion si es necesario
            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, dgvAsientosContables.Rows.Count - 1];

            dgvAsientosContables.Rows[fila].Selected = true;
            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, fila];
        }

        private void dgvAsientosContables_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int fila = dgvAsientosContables.CurrentRow.Index;

            tbNumAsiento.Text = dgvAsientosContables[0, fila].Value.ToString();
            tbComentario.Text = dgvAsientosContables[1, fila].Value.ToString();
            cbTipoCuenta.Text = dgvAsientosContables[2, fila].Value.ToString();
            cbCuenta.Text = dgvAsientosContables[3, fila].Value.ToString();

            if (dgvAsientosContables[4, fila].Value.ToString() == "")
            {
                tbCantidad.Text = dgvAsientosContables[5, fila].Value.ToString();
                rbHaber.Checked = true;
            }
            else
            {
                tbCantidad.Text = dgvAsientosContables[4, fila].Value.ToString();
                rbDebe.Checked = true;
            }

            btEliminar.Enabled = true;
            btActualizarCuenta.Enabled = true;
        }

        private void btActualizarCuenta_Click(object sender, EventArgs e)
        {
            VaciadoDGV();

            if (dgvAsientosContables.Rows.Count == 0)
            {
                MessageBox.Show("No hay renglon alguno para modificar.");
                return;
            }

            if (DialogResult.Cancel == MessageBox.Show("El renglon seleccionado en la tabla sera modificado con la informacion de este recuadro. \n\n¿Desea continuar?", "Modificacion", MessageBoxButtons.OKCancel))
            {
                return;
            }

            // Revision de cantidad solo numerica
            try
            {
                Convert.ToDouble(tbCantidad.Text);
            }
            catch
            {
                MessageBox.Show("La cantidad no es solo numerica. \n\nVuelva a intentarlo.");
                return;
            }

            // Revision de datos necesarios suministrados
            if (tbNumAsiento.Text == "" ||
                cbTipoCuenta.Text == "" ||
                cbCuenta.Text == "" ||
                tbCantidad.Text == ""
              )
            {
                MessageBox.Show("Falta algun dato necesario. \n\nVuelva a intentarlo.");
                return;
            }

            string cantidad1 = "", cantidad2 = "";
            if (rbDebe.Checked) cantidad1 = tbCantidad.Text;
            else cantidad2 = tbCantidad.Text;

            dgvAsientosContables.Rows.RemoveAt(dgvAsientosContables.CurrentRow.Index);

            dgvAsientosContables.Rows.Add(
                tbNumAsiento.Text,
                tbComentario.Text,
                cbTipoCuenta.Text,
                cbCuenta.Text,
                cantidad1,
                cantidad2
            );

            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, dgvAsientosContables.Rows.Count - 1];
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[4], ListSortDirection.Descending);
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[0], ListSortDirection.Ascending);
            int fila = dgvAsientosContables.SelectedRows[0].Index;

            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, 0];    // Para que muestre la deseada habiendo un cambio de seleccion si es necesario
            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, dgvAsientosContables.Rows.Count - 1];

            dgvAsientosContables.Rows[fila].Selected = true;
            dgvAsientosContables.CurrentCell = dgvAsientosContables[0, fila];
        }

        private void btEliminar_Click(object sender, EventArgs e)
        {
            VaciadoDGV();

            if (dgvAsientosContables.Rows.Count == 0)
            {
                MessageBox.Show("No hay renglon alguno para eliminar.");
                return;
            }

            if (DialogResult.Cancel == MessageBox.Show("El renglon seleccionado en la tabla sera eliminado. \n\n¿Desea continuar?", "Eliminacion", MessageBoxButtons.OKCancel))
            {
                return;
            }

            dgvAsientosContables.Rows.RemoveAt(dgvAsientosContables.CurrentRow.Index);
            btLimpiarAsiento_Click(null, null);
        }


        private void btnCalcular_Click(object sender, EventArgs e)
        {
            // Revision de cantidad solo numerica en VALORES DE INVENTARIO
            try
            {
                Convert.ToDouble(tbInventarioInicial.Text);
                Convert.ToDouble(tbInventarioFinal.Text);
            }
            catch
            {
                MessageBox.Show("Una cantidad de inventario inicial o final no es solo numerica. \n\nVuelva a intentarlo.");
                return;
            }

            // Revision de datos necesarios suministrados
            if (tbInventarioInicial.Text == "" ||
                tbInventarioFinal.Text == ""
              )
            {
                MessageBox.Show("Falta algun dato sobre el valor del inventario (final o inicial). \n\nVuelva a intentarlo.");
                return;
            }

            // Revision de que haya informacion en asientos contables
            if (dgvAsientosContables.Rows.Count < 1)
            {
                MessageBox.Show("No se han agregado algun asiento contable. \n\nVuelva a intentarlo.");
                return;
            }


            VaciadoDGV();
            
            // Revision de asientos con sumas correspondientes en el debe y haber
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[0], ListSortDirection.Ascending);

            double debe = 0.0, haber= 0.0;
            string id = dgvAsientosContables[0, 0].Value.ToString();

            for (int fila=0; fila < dgvAsientosContables.Rows.Count; fila++)
            {
                if (id != dgvAsientosContables[0, fila].Value.ToString())
                {
                    if (debe != haber)
                    {
                        MessageBox.Show("El asiento " + id + " no suma lo mismo en el debe (" + debe.ToString() + ") que en el haber (" + haber.ToString() +"). \n\nVuelva a intentarlo.");
                        return;
                    }
                    else
                    {
                        debe = 0;
                        haber = 0;
                        id = dgvAsientosContables[0, fila].Value.ToString();
                    }
                }
                if(dgvAsientosContables[4, fila].Value.ToString() != "") debe += Convert.ToDouble(dgvAsientosContables[4, fila].Value.ToString());
                if (dgvAsientosContables[5, fila].Value.ToString() != "") haber += Convert.ToDouble(dgvAsientosContables[5, fila].Value.ToString());
            }
            // Revision del ultimo asiento
            if (debe != haber)
            {
                MessageBox.Show("El asiento " + id + " no suma lo mismo en el debe (" + debe.ToString() + ") que en el haber (" + haber.ToString() + "). \n\nVuelva a intentarlo.");
                return;
            }

            MessageBox.Show("Los asientos cuadran adecuadamente.");

            //==============================================================================================================
            // Creacion total de tablas de Edo Resultados y Balance General para ir llenando con saldos de cuentas T
            // Despues del llenado, se eliminaran los conceptos vacios, se haran sumas y demas calculos.

            // EDO DE RESULTADOS
            dgvEdoResultados.Rows.Add("Ventas Netas", "", "", "100");
            dgvEdoResultados.Rows.Add("    Inventario Inicial", tbInventarioInicial.Text, "", "");
            dgvEdoResultados.Rows.Add("    Compras", "", "", "");
            dgvEdoResultados.Rows.Add("", "_______________", "", "");
            dgvEdoResultados.Rows.Add("  Total Mercancias Disponibles", "", "", "");
            dgvEdoResultados.Rows.Add("  Inventario Final", tbInventarioFinal.Text, "", "");
            dgvEdoResultados.Rows.Add("", "_______________", "", "");
            dgvEdoResultados.Rows.Add("Costo de lo vendido", "", "", "");
            dgvEdoResultados.Rows.Add("", "", "_______________", "");
            dgvEdoResultados.Rows.Add("UTILIDAD BRUTA", "", "", "");
            dgvEdoResultados.Rows.Add("    Gastos de Venta", "", "", "");
            dgvEdoResultados.Rows.Add("    Gastos Administrativos", "", "", "");
            dgvEdoResultados.Rows.Add("", "_______________", "", "");
            dgvEdoResultados.Rows.Add("  Total Gastos Operativos", "", "", "");
            dgvEdoResultados.Rows.Add("", "", "_______________", "");
            dgvEdoResultados.Rows.Add("UTILIDAD OPERATIVA", "", "", "");
            dgvEdoResultados.Rows.Add("  Gastos Financieros", "", "", "");
            dgvEdoResultados.Rows.Add("", "", "_______________", "");
            dgvEdoResultados.Rows.Add("UTILIDAD FINANCIERA", "", "", "");
            dgvEdoResultados.Rows.Add("  Otros Productos", "", "", "");
            dgvEdoResultados.Rows.Add("", "", "_______________", "");
            dgvEdoResultados.Rows.Add("UTILIDAD ANTES DE IMPUESTOS", "", "", "");
            dgvEdoResultados.Rows.Add("  ISR(30%)", "", "", "");
            dgvEdoResultados.Rows.Add("  PTU(10%)", "", "", "");
            dgvEdoResultados.Rows.Add("", "", "_______________", "");
            dgvEdoResultados.Rows.Add("UTILIDAD NETA", "", "", "");

            // BALANCE GENERAL
            dgvBalanceGeneral.Rows.Add("============================", "ACTIVO       =====", "==================", "");
            dgvBalanceGeneral.Rows.Add("CIRCULANTE", "", "", "");
            for(int i=0; i < CuentasActivoCirculante.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasActivoCirculante[i], "", "", "");

            dgvBalanceGeneral[1, dgvBalanceGeneral.Rows.Count - 1].Value = tbInventarioFinal.Text;   // Este valor se pone directo, por lo que Inventario debe mantenerse como la ultima cuenta de activo circulante
            
            dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            dgvBalanceGeneral.Rows.Add("  TOTAL de Activo Circulante", "", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("FIJO", "", "", "");
            for (int i = 0; i < CuentasActivoFijo.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasActivoFijo[i], "", "", "");
            dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            dgvBalanceGeneral.Rows.Add("  TOTAL de Activo Fijo", "", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("DIFERIDO", "", "", "");
            for (int i = 0; i < CuentasActivoDiferido.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasActivoDiferido[i], "", "", "");
            dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            dgvBalanceGeneral.Rows.Add("  TOTAL de Activo Diferido", "", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("", "TOTAL ACTIVO", "0.00", "100");
            dgvBalanceGeneral.Rows.Add("", "", "", "");

            dgvBalanceGeneral.Rows.Add("============================", "PASIVO       =====", "==================", "");
            dgvBalanceGeneral.Rows.Add("A CORTO PLAZO", "", "", "");
            for (int i = 0; i < CuentasPasivoCortoPlazo.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasPasivoCortoPlazo[i], "", "", "");
            dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            dgvBalanceGeneral.Rows.Add("  TOTAL de Pasivo a Corto Plazo", "", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("A LARGO PLAZO", "", "", "");
            for (int i = 0; i < CuentasPasivoLargoPlazo.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasPasivoLargoPlazo[i], "", "", "");
            dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            dgvBalanceGeneral.Rows.Add("  TOTAL de Pasivo a Largo Plazo", "", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("OTROS PASIVOS", "", "", "");
            for (int i = 0; i < CuentasPasivoDiferido.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasPasivoDiferido[i], "", "", "");
            dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            dgvBalanceGeneral.Rows.Add("  TOTAL de Otros Pasivos", "", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("", "TOTAL PASIVO", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");

            dgvBalanceGeneral.Rows.Add("============================", "CAPITAL CONTABLE  ", "==================", "");
            for (int i = 0; i < CuentasCapital.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasCapital[i], "0.00", "", "");
            dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("", "TOTAL CAPITAL CONTABLE", "0.00", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            dgvBalanceGeneral.Rows.Add("", "", "", "");
            dgvBalanceGeneral.Rows.Add("", "TOTAL PASIVO Y CAPITAL", "0.00", "");


            //==============================================================================================================
            // Creacion de cuentas T
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[4], ListSortDirection.Descending);
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[3], ListSortDirection.Ascending);

            // dgvCuentasT.Rows.Clear();   // Ya se hizo al inicio de este metodo
            debe = 0;
            haber = 0;
            
            id = dgvAsientosContables[3, 0].Value.ToString();
            string izq = "", der = "";
            dgvCuentasT.Rows.Add("======================", "=========", "===============", "===============", "=========");

            for (int fila = 0; fila < dgvAsientosContables.Rows.Count; fila++)
            {
                if (id != dgvAsientosContables[3, fila].Value.ToString())
                {
                    // Cierre de cuenta T
                    if (debe >= haber)
                    {
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "Suma", debe.ToString("N2"), haber.ToString("N2"), "");
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "S Deudor", (debe-haber).ToString("N2"), "", "");
                        dgvCuentasT.Rows.Add("======================", "=========", "===============", "===============", "=========");
                        AgregarCuenta(id, debe - haber);
                    }
                    else
                    {
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "", debe.ToString("N2"), haber.ToString("N2"), "Suma");
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "", "", (haber - debe).ToString("N2"), "S Acreedor");
                        dgvCuentasT.Rows.Add("======================", "=========", "===============", "===============", "=========");
                        AgregarCuenta(id, haber - debe);
                    }

                    // Inicio de la siguiente
                    debe = 0;
                    haber = 0;
                    id = dgvAsientosContables[3, fila].Value.ToString();
                }

                // Agregado de integrantes dde una cuenta

                if (dgvAsientosContables[4, fila].Value.ToString() != "")
                {
                    debe += Convert.ToDouble(dgvAsientosContables[4, fila].Value.ToString());
                    izq = dgvAsientosContables[0, fila].Value.ToString();
                    der = "";
                }
                if (dgvAsientosContables[5, fila].Value.ToString() != "")
                {
                    haber += Convert.ToDouble(dgvAsientosContables[5, fila].Value.ToString());
                    der = dgvAsientosContables[0, fila].Value.ToString();
                    izq = "";
                }

                dgvCuentasT.Rows.Add(id, izq, dgvAsientosContables[4, fila].Value.ToString(), dgvAsientosContables[5, fila].Value.ToString(), der);

                if (fila + 1 == dgvAsientosContables.Rows.Count)
                {
                    // Cierre de cuenta T
                    if (debe >= haber)
                    {
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "Suma", debe.ToString("N2"), haber.ToString("N2"), "");
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "S Deudor", (debe - haber).ToString("N2"), "", "");
                        AgregarCuenta(id, debe - haber);
                    }
                    else
                    {
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "", debe.ToString("N2"), haber.ToString("N2"), "Suma");
                        dgvCuentasT.Rows.Add("", "", "______________", "______________", "");
                        dgvCuentasT.Rows.Add(id, "", "", (haber - debe).ToString("N2"), "S Acreedor");
                        AgregarCuenta(id, haber - debe);
                    }
                }
            }

            // Vuelta al acomodo similar al inicial
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[4], ListSortDirection.Descending);
            dgvAsientosContables.Sort(dgvAsientosContables.Columns[0], ListSortDirection.Ascending);

            //==============================================================================================================
            // Elaboracion de calculos finales y Eliminacion de cuentas vacias en documentos contables

            // Total de mercancias disponibles

            List<int> eliminacionesER = new List<int>();  // Se indica posicion de renglones a eliminar despues de calculos. Pueden no aparecer Compras, Gasto Venta, Gasto adm, Otros productos
            List<int> eliminacionesBG = new List<int>();

            double valor1 = 0, valor2 = 0, valor3=0;

            if (dgvEdoResultados[2, 0].Value.ToString() == "") // No hubo ventas
            {
                dgvEdoResultados[2, 0].Value = "0.00";   // Se conserva su lugar
                dgvEdoResultados[3, 0].Value = "100";
                valor3 = 0.001;
            }
            else
            {
                valor3 = Convert.ToDouble(dgvEdoResultados[2, 0].Value);
            }

            valor1 = Convert.ToDouble(tbInventarioInicial.Text);
            if (dgvEdoResultados[1, 2].Value.ToString() == "") // No hubo compras
            {
                eliminacionesER.Add(2);
                dgvEdoResultados[1, 4].Value = valor1.ToString("N2");   // Total mercancia disponible
            }
            else
            {
                valor1 += Convert.ToDouble(dgvEdoResultados[1, 2].Value.ToString()); // Suma las compras
                dgvEdoResultados[1, 4].Value = valor1.ToString("N2");
            }

            // Costo de lo vendido
            valor1 = valor1 - Convert.ToDouble(tbInventarioFinal.Text);   // Obtiene costo de lo vendido
            dgvEdoResultados[2, 7].Value = valor1.ToString("N2");
            dgvEdoResultados[3, 7].Value = ( Math.Round(100*valor1/valor3,0) ).ToString();

            // Utilidad bruta
            valor1 = Convert.ToDouble(dgvEdoResultados[2, 0].Value.ToString()) - valor1;  // Ventas netas menos costo de lo vendido
            dgvEdoResultados[2, 9].Value = valor1.ToString("N2");
            dgvEdoResultados[3, 9].Value = (Math.Round(100 * valor1 / valor3, 0)).ToString();
            double FutilBruta = valor1;

            valor2 = 0;
            if (dgvEdoResultados[1, 10].Value.ToString() == "")   // No hubo gasto de ventas
            {
                eliminacionesER.Add(10);
            }
            else
            {
                valor2 = Convert.ToDouble(dgvEdoResultados[1, 10].Value.ToString());
            }
            if (dgvEdoResultados[1, 11].Value.ToString() == "")   // No hubo gasto administrativo
            {
                eliminacionesER.Add(11);
                if(valor2 == 0) eliminacionesER.Add(12);  // No hubo gasto de ventas
            }
            else
            {
                valor2 += Convert.ToDouble(dgvEdoResultados[1, 11].Value.ToString());
            }
            dgvEdoResultados[2, 13].Value = valor2.ToString("N2");   // Total gasto operativo
            dgvEdoResultados[3, 13].Value = (Math.Round(100 * valor2 / valor3, 0)).ToString();

            // Utilidad operativa
            valor1 -= valor2;
            dgvEdoResultados[2, 15].Value = valor1.ToString("N2");
            dgvEdoResultados[3, 15].Value = (Math.Round(100 * valor1 / valor3, 0)).ToString();
            double FutilOperativa = valor1;

            // Gastos financieros
            try
            {
                valor2 = Convert.ToDouble(dgvEdoResultados[2, 16].Value.ToString());
            }
            catch
            {
                valor2 = 0;
            }

            // Utilidad financiera
            valor1 -= valor2;
            double FutilFinanciera = valor1;
            dgvEdoResultados[2, 18].Value = valor1.ToString("N2");
            dgvEdoResultados[3, 18].Value = (Math.Round(100 * valor1 / valor3, 0)).ToString();

            // Otros productos
            try
            {
                valor2 = Convert.ToDouble(dgvEdoResultados[2, 19].Value.ToString());
            }
            catch
            {
                valor2 = 0;
            }

            // Utilidad antes de impuestos
            valor1 += valor2;
            dgvEdoResultados[2, 21].Value = valor1.ToString("N2");
            dgvEdoResultados[3, 21].Value = (Math.Round(100 * valor1 / valor3, 0)).ToString();

            // ISR (30%)
            dgvEdoResultados[2, 22].Value = (valor1*0.3).ToString("N2");
            // PTU (10%)
            dgvEdoResultados[2, 23].Value = (valor1*0.1).ToString("N2");

            // Utilidad neta
            dgvEdoResultados[2, 25].Value = (valor1*0.6).ToString("N2");
            dgvEdoResultados[3, 25].Value = (Math.Round(100 * valor1 * 0.6 / valor3, 0)).ToString();



            // BALANCE GENERAL

            // valor1 vale utilidad antes de impuestos, se agregara al balance general
            dgvBalanceGeneral[1, 70].Value = valor1.ToString("N2");


            // Revision de eliminaciones de bloques de: Activo diferido, Pasivo Largo Plazo y Otros pasivos
            // Activo Diferido ocupa del 24 al 38, sus cuentas en 25-34
            // Pasivo Largo Plazo ocupa del 51-57, sus cuentas en 53-55
            // Otros pasivos ocupa del 58-63, sus cuentas en 60-61

            // eliminacionesBG.Add();

            valor1 = 0;
            for(int i = 2; i < 10; i++)   // Cuentas de activo circulante
            {
                if (dgvBalanceGeneral[1, i].Value.ToString() == "") eliminacionesBG.Add(i);
                else valor1 += Convert.ToDouble(dgvBalanceGeneral[1, i].Value.ToString());
            }
            dgvBalanceGeneral[2, 11].Value = valor1.ToString("N2");


            valor2 = 0;
            for (int i = 14; i < 21; i++)   // Cuentas de activo fijo
            {
                if (dgvBalanceGeneral[1, i].Value.ToString() == "") eliminacionesBG.Add(i);
                else valor2 += Convert.ToDouble(dgvBalanceGeneral[1, i].Value.ToString());
            }
            dgvBalanceGeneral[2, 22].Value = valor2.ToString("N2");

            valor3 = 0;
            bool elimSeccion = true;
            for (int i = 25; i < 35; i++)   // Cuentas de activo diferido
            {
                if (dgvBalanceGeneral[1, i].Value.ToString() != "")
                {
                    elimSeccion = false;
                    break;
                }
            }
            if (elimSeccion)
            {
                for (int i = 24; i < 39; i++) eliminacionesBG.Add(i);
            }
            else
            {
                for (int i = 25; i < 35; i++)   // Cuentas de activo diferido
                {
                    if (dgvBalanceGeneral[1, i].Value.ToString() == "") eliminacionesBG.Add(i);
                    else valor3 += Convert.ToDouble(dgvBalanceGeneral[1, i].Value.ToString());
                }
                dgvBalanceGeneral[2, 36].Value = valor3.ToString("N2");
            }
            
            // Elementos de analisis financiero
            double FctasPagar, FctasCobrar,  
                FactivoCirculante, FpasivoCirculante, FinventarioFinal, 
                Factivo, Fpasivo, Fcapital, FutilNeta, FventasNetas, 
                FcostoAnualdeloVendido, FcomprasDiarias, FventasDiarias,
                FinventarioMensual, FcapitalMensual, FventasAnuales, FactivoAnual, FdiasPeriodo;

            Factivo = valor1 + valor2 + valor3;
            FactivoCirculante = valor1;
            if (dgvBalanceGeneral[1, 6].Value.ToString() == "") FctasCobrar = 0;
            else FctasCobrar = Convert.ToDouble(dgvBalanceGeneral[1,6].Value.ToString());
            if (dgvBalanceGeneral[1, 44].Value.ToString() == "") FctasPagar = 0;
            else FctasPagar = Convert.ToDouble(dgvBalanceGeneral[1, 44].Value.ToString());
            FinventarioFinal = Convert.ToDouble(tbInventarioFinal.Text);
            FventasNetas = Convert.ToDouble(dgvEdoResultados[2, 0].Value.ToString());
            FutilNeta = Convert.ToDouble(dgvEdoResultados[2, dgvEdoResultados.Rows.Count-1].Value.ToString());
            TimeSpan tSpan = dtpFechaFin.Value.Date - dtpFechaInicio.Value.Date;
            if (tSpan.Days > 0) FdiasPeriodo = Convert.ToDouble(tSpan.Days);
            else FdiasPeriodo = 1;
            FactivoAnual = Factivo * 365 / FdiasPeriodo;
            FinventarioMensual = ((FinventarioFinal + Convert.ToDouble(tbInventarioInicial.Text)) / 2) * 30 / FdiasPeriodo;
            FventasDiarias = FventasNetas / FdiasPeriodo;
            FventasAnuales = FventasDiarias * 365;
            if (dgvEdoResultados[1, 2].Value.ToString() == "") FcomprasDiarias = 0;
            else FcomprasDiarias = Convert.ToDouble(dgvEdoResultados[1, 2].Value.ToString())/FdiasPeriodo;
            FcostoAnualdeloVendido = 365* Convert.ToDouble(dgvEdoResultados[2, 7].Value.ToString()) / FdiasPeriodo;


            dgvBalanceGeneral[2, 39].Value = Factivo.ToString("N2");  // Total Activo

            // Capital
            Fcapital = 0;
            for (int i = 69; i < 71; i++)   // Cuentas de capital
            {
                Fcapital += Convert.ToDouble(dgvBalanceGeneral[1, i].Value.ToString());
            }
            dgvBalanceGeneral[2, 73].Value = Fcapital.ToString("N2");   // Total Capital
            dgvBalanceGeneral[3, 73].Value = (Math.Round(100 * Fcapital / Factivo, 0)).ToString();

            FcapitalMensual = 30 * Fcapital / FdiasPeriodo;

            // Pasivo
            FpasivoCirculante = 0;
            for (int i = 43; i < 49; i++)   // Cuentas de pasivo circulante
            {
                if (dgvBalanceGeneral[1, i].Value.ToString() == "") eliminacionesBG.Add(i);
                else FpasivoCirculante += Convert.ToDouble(dgvBalanceGeneral[1, i].Value.ToString());
            }
            dgvBalanceGeneral[2, 50].Value = FpasivoCirculante.ToString("N2");
            dgvBalanceGeneral[3, 50].Value = (Math.Round(100 * FpasivoCirculante / Factivo, 0)).ToString();


            elimSeccion = true;
            valor1 = 0;
            for (int i = 53; i < 56; i++)   // Cuentas de pasivo largo plazo
            {
                if (dgvBalanceGeneral[1, i].Value.ToString() != "")
                {
                    elimSeccion = false;
                    break;
                }
            }
            if (elimSeccion)
            {
                for (int i = 51; i < 58; i++) eliminacionesBG.Add(i);
            }
            else
            {
                for (int i = 53; i < 56; i++)   // Cuentas de pasivo largo plazo
                {
                    if (dgvBalanceGeneral[1, i].Value.ToString() == "") eliminacionesBG.Add(i);
                    else valor1 += Convert.ToDouble(dgvBalanceGeneral[1, i].Value.ToString());
                }
                dgvBalanceGeneral[2, 57].Value = valor1.ToString("N2");
                dgvBalanceGeneral[3, 57].Value = (Math.Round(100 * valor1 / Factivo, 0)).ToString();
            }

            elimSeccion = true;
            valor2 = 0;
            for (int i = 60; i < 62; i++)   // Cuentas de pasivo diferido
            {
                if (dgvBalanceGeneral[1, i].Value.ToString() != "")
                {
                    elimSeccion = false;
                    break;
                }
            }
            if (elimSeccion)
            {
                for (int i = 58; i < 64; i++) eliminacionesBG.Add(i);
            }
            else
            {
                for (int i = 60; i < 62; i++)   // Cuentas de pasivo diferido
                {
                    if (dgvBalanceGeneral[1, i].Value.ToString() == "") eliminacionesBG.Add(i);
                    else valor2 += Convert.ToDouble(dgvBalanceGeneral[1, i].Value.ToString());
                }
                dgvBalanceGeneral[2, 63].Value = valor2.ToString("N2");
                dgvBalanceGeneral[3, 63].Value = (Math.Round(100 * valor2 / Factivo, 0)).ToString();
            }




            Fpasivo = FpasivoCirculante + valor1 + valor2;
            dgvBalanceGeneral[2, 66].Value = Fpasivo.ToString("N2");  // Total pasivo
            dgvBalanceGeneral[3, 66].Value = (Math.Round(100 * Fpasivo / Factivo, 0)).ToString();
            dgvBalanceGeneral[2, 77].Value = (Fpasivo+Fcapital).ToString("N2");  // Total pasivo y capital
            dgvBalanceGeneral[3, 77].Value = (Math.Round(100 * (Fpasivo + Fcapital) / Factivo, 0)).ToString();

            /*
                         // BALANCE GENERAL
            0 dgvBalanceGeneral.Rows.Add("============================", "ACTIVO       =====", "==================", "");
            1 dgvBalanceGeneral.Rows.Add("CIRCULANTE", "", "", "");
            2-9 for(int i=0; i < CuentasActivoCirculante.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasActivoCirculante[i], "", "", "");

            dgvBalanceGeneral[1, dgvBalanceGeneral.Rows.Count - 1].Value = tbInventarioFinal.Text;   // Este valor se pone directo, por lo que Inventario debe mantenerse como la ultima cuenta de activo circulante
            
            10 dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            11 dgvBalanceGeneral.Rows.Add("  TOTAL de Activo Circulante", "", "", "");
            12 dgvBalanceGeneral.Rows.Add("", "", "");
            13 dgvBalanceGeneral.Rows.Add("FIJO", "", "", "");
            14-20 for (int i = 0; i < CuentasActivoFijo.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasActivoFijo[i], "", "", "");
            21 dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            22 dgvBalanceGeneral.Rows.Add("  TOTAL de Activo Fijo", "", "", "");
            23 dgvBalanceGeneral.Rows.Add("", "", "");
            24 dgvBalanceGeneral.Rows.Add("DIFERIDO", "", "", "");
            25-34 for (int i = 0; i < CuentasActivoDiferido.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasActivoDiferido[i], "", "", "");
            35 dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            36 dgvBalanceGeneral.Rows.Add("  TOTAL de Activo Diferido", "", "", "");
            37 dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            38 dgvBalanceGeneral.Rows.Add("", "", "");
            39 dgvBalanceGeneral.Rows.Add("", "TOTAL ACTIVO", "", "");
            40 dgvBalanceGeneral.Rows.Add("", "", "");

            41 dgvBalanceGeneral.Rows.Add("============================", "PASIVO       =====", "==================", "");
            42 dgvBalanceGeneral.Rows.Add("A CORTO PLAZO", "", "", "");
            43-48 for (int i = 0; i < CuentasPasivoCortoPlazo.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasPasivoCortoPlazo[i], "", "", "");
            49 dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            50 dgvBalanceGeneral.Rows.Add("  TOTAL de Pasivo a Corto Plazo", "", "", "");
            51 dgvBalanceGeneral.Rows.Add("", "", "");
            52 dgvBalanceGeneral.Rows.Add("A LARGO PLAZO", "", "", "");
            53-55 for (int i = 0; i < CuentasPasivoLargoPlazo.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasPasivoLargoPlazo[i], "", "", "");
            56 dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            57 dgvBalanceGeneral.Rows.Add("  TOTAL de Pasivo a Largo Plazo", "", "", "");
            58 dgvBalanceGeneral.Rows.Add("", "", "");
            59 dgvBalanceGeneral.Rows.Add("OTROS PASIVOS", "", "", "");
            60-61 for (int i = 0; i < CuentasPasivoDiferido.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasPasivoDiferido[i], "", "", "");
            62 dgvBalanceGeneral.Rows.Add("", "_______________", "", "");
            63 dgvBalanceGeneral.Rows.Add("  TOTAL de Otros Pasivos", "", "", "");
            64 dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            65 dgvBalanceGeneral.Rows.Add("", "", "");
            66 dgvBalanceGeneral.Rows.Add("", "TOTAL PASIVO", "", "");
            67 dgvBalanceGeneral.Rows.Add("", "", "");

            68 dgvBalanceGeneral.Rows.Add("============================", "CAPITAL CONTABLE  ", "==================", "");
            69-70 for (int i = 0; i < CuentasCapital.Length; i++) dgvBalanceGeneral.Rows.Add("  " + CuentasCapital[i], "", "", "");
            71 dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            72 dgvBalanceGeneral.Rows.Add("", "", "");
            73 dgvBalanceGeneral.Rows.Add("", "TOTAL CAPITAL CONTABLE", "", "");
            74 dgvBalanceGeneral.Rows.Add("", "", "");
            75 dgvBalanceGeneral.Rows.Add("", "", "_______________", "");
            76 dgvBalanceGeneral.Rows.Add("", "", "");
            77 dgvBalanceGeneral.Rows.Add("", "TOTAL PASIVO Y CAPITAL", "", "");

             */


            //===================================================================================================
            // Porcentajes: Edo Resultados / Ventas Netas,  Bal Gral /Activo
            dgvBalanceGeneral[3, 11].Value = (Math.Round(100 * FactivoCirculante / Factivo, 0)).ToString();
            dgvBalanceGeneral[3, 22].Value = (Math.Round(100 * Convert.ToDouble(dgvBalanceGeneral[2, 22].Value.ToString()) / Factivo, 0)).ToString();
            dgvBalanceGeneral[3, 36].Value = (Math.Round(100 * Convert.ToDouble(dgvBalanceGeneral[2, 36].Value.ToString()) / Factivo, 0)).ToString();

            //===================================================================================================
            // Razones Financieras
            dgvRazonesFinancieras.Rows.Add("Liquidez", "Circulante (corriente) = ActivoCirculante / PasivoCtoPlazo", (FactivoCirculante/(FpasivoCirculante+0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Liquidez", "Rapida (del acido) = (ActivoCirculante - Inventario) / PasivoCtoPlazo", ((FactivoCirculante - FinventarioFinal) / (FpasivoCirculante + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Endeudamiento", "Endeudamiento = PasivoTotal / ActivoTotal", (Fpasivo / (Factivo + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Endeudamiento", "Apalancamiento = PasivoTotal / CapitalContable", (Fpasivo / (Fcapital + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Endeudamiento", "Inversion propia = CapitalContable / ActivoTotal", (Fcapital / (Factivo + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Rendimiento", "sobre la Inversion = UtilidadNetaAnual / CapitalPromMensual", ((FutilNeta*365/FdiasPeriodo) / (FcapitalMensual + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Rendimiento", "Margen de Utilidad Neta = UtilidadNeta / VentasNetas", (FutilNeta / (FventasNetas + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Rendimiento", "Margen de Utilidad Operativa = UtilidadOperativa / VentasNetas", (FutilOperativa / (FventasNetas + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Rendimiento", "Margen de Utilidad Financiera = UtilidadFinanciera / VentasNetas", (FutilFinanciera / (FventasNetas + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Rendimiento", "Margen de Utilidad Bruta = UtilidadBruta / VentasNetas", (FutilBruta / (FventasNetas + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Credito", "Periodo promedio de Pago = CuentasPorPagar / ComprasDiariasPromedio", (FctasPagar / (FcomprasDiarias + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Credito", "Periodo promedio de Cobro = CuentasPorCobrar / VentasDiariasPromedio", (FctasCobrar / (FventasDiarias + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Actividad", "Rotacion de Inventarios = CostoDeLoVendidoAnual / InventarioMensual", (FcostoAnualdeloVendido / (FinventarioMensual + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Actividad", "Tiempo en vender el Inventario = 365 / RotacionInventarios", (365 * FinventarioMensual / (FcostoAnualdeloVendido + 0.0000001)).ToString("N4"));
            dgvRazonesFinancieras.Rows.Add("Actividad", "Rotacion de Activos totales = VentasAnuales / TotalActivosAnuales", (FventasAnuales / (Factivo + 0.0000001)).ToString("N4"));


            //======================================================================================================
            // Eliminacion de renglones en Edo de Resultados
            for (int i = eliminacionesER.Count - 1; i >= 0; i--) dgvEdoResultados.Rows.RemoveAt(eliminacionesER[i]);
            for (int i = eliminacionesBG.Count - 1; i >= 0; i--) dgvBalanceGeneral.Rows.RemoveAt(eliminacionesBG[i]);


            btLimpiarAsiento_Click(null, null);
        }

        private void AgregarCuenta(string nombre, double saldo)
        {
            if (nombre == "InventarioMercancAlmacen") return;   // Inventario inicial y final se han puesto directamente
            
            // Revision de pertenencia a Edo de Resultados
            bool balGral = false;
            switch (nombre)
            {
                case "Ventas":
                    dgvEdoResultados[2, 0].Value = saldo.ToString("N2");
                    break;
                case "Compras":
                    dgvEdoResultados[1, 2].Value = saldo.ToString("N2");
                    break;
                case "Gastos de Venta":
                    dgvEdoResultados[1, 10].Value = saldo.ToString("N2");
                    break;
                case "Gastos de Admin.":
                    dgvEdoResultados[1, 11].Value = saldo.ToString("N2");
                    break;
                case "Gastos Financieros":
                    dgvEdoResultados[2, 16].Value = saldo.ToString("N2");
                    break;
                case "Otros Productos":
                    dgvEdoResultados[2, 19].Value = saldo.ToString("N2");
                    break;
                default:
                    balGral = true;
                    break;
            }

            if (!balGral) return;

            for(int i = 0; i < dgvBalanceGeneral.Rows.Count; i++)
            {
                if ("  " + nombre == dgvBalanceGeneral[0, i].Value.ToString())
                {
                    dgvBalanceGeneral[1, i].Value = saldo.ToString("N2");
                }
            }
        }


        public string DGVaSTRING(DataGridView dgv)
        {
            RichTextBox rtb = new RichTextBox();
            int[] maxLongitud = new int[dgv.Columns.Count];
            rtb.Font = new Font("Consolas", 7.0f);
            rtb.DeselectAll();
            rtb.SelectionAlignment = 0;
            //rtb.SelectionFont = new Font("Consolas", 9.0f);

            // Deteccion de longitudes mayores de string por columna para agregar espacios e igualar y hacer tablas
            for (int columna = 0; columna < dgv.Columns.Count ; columna++)
            {
                maxLongitud[columna] = dgv.Columns[columna].HeaderText.Length;
                for (int fila = 0; fila < dgv.Rows.Count; fila++)
                {
                    if (dgv[columna, fila].Value == null || dgv[columna, fila].Value.ToString() == "") continue;
                    if (maxLongitud[columna] < dgv[columna, fila].Value.ToString().Length)
                        maxLongitud[columna] = dgv[columna, fila].Value.ToString().Length;
                }
            }

            //string separador = "";
            for (int columna = 0; columna < dgv.Columns.Count ; columna++)
            {
                rtb.AppendText(dgv.Columns[columna].HeaderText + espacios(maxLongitud[columna], dgv.Columns[columna].HeaderText.Length));
                if (columna == dgv.Columns.Count - 1)
                    rtb.AppendText(" |\n");
                else
                    rtb.AppendText(" | ");
            }

            rtb.AppendText("\n");

            for (int fila = 0; fila < dgv.Rows.Count ; fila++)
            {
                for (int columna = 0; columna < dgv.Columns.Count; columna++)
                {
                    if (columna == 0)
                        rtb.AppendText(dgv[columna, fila].Value.ToString() + espacios(maxLongitud[columna], dgv[columna, fila].Value.ToString().Length));
                    else
                        rtb.AppendText(espacios(maxLongitud[columna], dgv[columna, fila].Value.ToString().Length) + dgv[columna, fila].Value.ToString());

                    if (columna == dgv.Columns.Count - 1)
                        rtb.AppendText(" |\n");
                    else
                        rtb.AppendText(" | ");
                }
            }
            rtb.AppendText("\n");

            return (rtb.Text);
        }

        private string espacios(int max, int longitud)
        {
            string espa = "";
            for (int i = longitud; i < max; i++)
            {
                espa += " ";
            }
            return espa;
        }

        private void btLimipiar_Click(object sender, EventArgs e)
        {
            if (DialogResult.Cancel == MessageBox.Show("Se borrara toda la informacion. \n\n¿Desea continuar?", "Limpieza total", MessageBoxButtons.OKCancel))
            {
                return;
            }

            VaciadoDGV();
            dgvAsientosContables.Rows.Clear();
            tbEmpresa.Text = "";
            tbInventarioInicial.Text = "";
            tbInventarioFinal.Text = "";
            // Ajuste de fechas por defecto a un mes hasta hoy
            dtpFechaFin.Value = DateTime.Today;
            dtpFechaInicio.Value = DateTime.Today.AddMonths(-1);

        }

        private void btLimpiarAsiento_Click(object sender, EventArgs e)
        {
            tbCantidad.Text = "";
            tbComentario.Text = "";
            tbNumAsiento.Text = "";
            cbTipoCuenta.Text = "";
            cbCuenta.Text = "";
            dgvAsientosContables.ClearSelection();
            btEliminar.Enabled = false;
            btActualizarCuenta.Enabled = false;
        }
    }
}
