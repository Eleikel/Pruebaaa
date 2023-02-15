using ExcelDataReader;
using FuncionalidadExcel.DATA;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FuncionalidadExcel
{
    public partial class Form1 : Form
    {
        //Un DataSet es un objeto que almacena n número de DataTables, estas tablas puedes estar conectadas dentro del dataset.
        private DataSet dtsTablas = new DataSet();

        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            DiseñoInicial();
        }

        private void DiseñoInicial()
        {
            btnRegistrarData.Cursor = Cursors.Hand;
            txtRuta.Enabled = true;

            dgvDatos.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvDatos.MultiSelect = false;
            dgvDatos.ReadOnly = true;
            dgvDatos.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvDatos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvDatos.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            OpenFileDialog oOpenFileDialog = new OpenFileDialog();
            oOpenFileDialog.Filter = "Excel Worbook|*.xlsx";

            if (oOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                dgvDatos.DataSource = null;

                txtRuta.Text = oOpenFileDialog.FileName;

                FileStream fsSource = new FileStream(oOpenFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);


                IExcelDataReader reader = ExcelReaderFactory.CreateReader(fsSource);

                //convierte todas las hojas a un DataSet
                //dtsTablas = reader.AsDataSet(new ExcelDataSetConfiguration()
                //{
                //    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                //    {
                //        UseHeaderRow = true
                //    }
                //});



                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(oOpenFileDialog.FileName);

                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];


                DataTable table = new DataTable();
                DataRow row = table.NewRow();


                string Cell_G7 = worksheet.Range["G7"].Value.ToString();
                table.Columns.Add("Empresa", typeof(string));
                row["Empresa"] = Cell_G7;

                string Cell_C6 = worksheet.Range["C6"].Value.ToString();
                table.Columns.Add("Nombre y Apellidos", typeof(string));
                row["Nombre y Apellidos"] = Cell_C6;

                string Cell_C7 = worksheet.Range["C7"].Value.ToString();
                table.Columns.Add("Fecha de Nacimiento");
                row["Fecha de Nacimiento"] = Cell_C7;

                string Cell_G8 = worksheet.Range["G8"].Value.ToString();
                table.Columns.Add("Área de Trabajo", typeof(string));
                row["Área de Trabajo"] = Cell_G8;

                string Cell_C9 = worksheet.Range["C9"].Value.ToString();
                table.Columns.Add("Cédula", typeof(string));
                row["Cédula"] = Cell_C9;

                string Cell_A14 = worksheet.Range["A14"].Value.ToString();
                table.Columns.Add("Patología", typeof(string));
                row["Patología"] = Cell_A14;

                string Cell_I14 = worksheet.Range["I14"].Value.ToString();
                table.Columns.Add("Tiempo de Padecerla", typeof(string));
                row["Tiempo de Padecerla"] = Cell_I14;


                string Cell_E14 = worksheet.Range["E14"].Value.ToString();
                table.Columns.Add("Tratamiento", typeof(string));
                row["Tratamiento"] = Cell_E14;

                string Cell_I46 = worksheet.Range["I46"].Value.ToString();
                table.Columns.Add("Peso (kg)");
                row["Peso (kg)"] = Cell_I46;


                table.Rows.Add(row);

                dgvDatos.DataSource = table;


                workbook.Close();
                excel.Quit();
            }


            //klk();


            //klk();
        }




        private void klk()
        {
            // Abre un cuadro de diálogo para que el usuario seleccione un archivo Excel
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Carga el archivo Excel en una aplicación de Excel
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(openFileDialog.FileName);

                // Obtiene la hoja de cálculo que quieres mostrar en el DataGrid
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                // Obtiene los datos de las columnas A y B de la hoja de cálculo
                //Microsoft.Office.Interop.Excel.Range range = worksheet.get_Range("A:B");
                //object[,] data = (object[,])range.Value2;

                // Crea un DataTable con los datos de las columnas A y B



                DataTable table = new DataTable();
                DataRow row = table.NewRow();


                string Cell_G7 = worksheet.Range["G7"].Value.ToString();
                table.Columns.Add("Empresa", typeof(string));
                row["Empresa"] = Cell_G7;

                string Cell_C6 = worksheet.Range["C6"].Value.ToString();
                table.Columns.Add("Nombre y Apellidos", typeof(string));
                row["Nombre y Apellidos"] = Cell_C6;

                string Cell_C7 = worksheet.Range["C7"].Value.ToString();
                table.Columns.Add("Fecha de Nacimiento");
                row["Fecha de Nacimiento"] = Cell_C7;

                string Cell_G8 = worksheet.Range["G8"].Value.ToString();
                table.Columns.Add("Área de Trabajo", typeof(string));
                row["Área de Trabajo"] = Cell_G8;

                string Cell_C9 = worksheet.Range["C9"].Value.ToString();
                table.Columns.Add("Cédula", typeof(string));
                row["Cédula"] = Cell_C9;

                string Cell_A14 = worksheet.Range["A14"].Value.ToString();
                table.Columns.Add("Patología", typeof(string));
                row["Patología"] = Cell_A14;

                string Cell_I14 = worksheet.Range["I14"].Value.ToString();
                table.Columns.Add("Tiempo de Padecerla", typeof(string));
                row["Tiempo de Padecerla"] = Cell_I14;


                string Cell_E14 = worksheet.Range["E14"].Value.ToString();
                table.Columns.Add("Tratamiento", typeof(string));
                row["Tratamiento"] = Cell_E14;

                string Cell_I46 = worksheet.Range["I46"].Value.ToString();
                table.Columns.Add("Peso (kg)") ;
                row["Peso (kg)"] = Cell_I46;



                table.Rows.Add(row);

                // Enlaza el DataTable al DataGrid
                dgvDatos.DataSource = table;


                // Cierra la aplicación de Excel
                workbook.Close();
                excel.Quit();
            }
        }






        private void btnRegistrarData_Click(object sender, EventArgs e)
        {
            DataTable data = (DataTable)(dgvDatos.DataSource);

            bool resultado = new Operaciones().CargarData(data);

            if (resultado)
            {
                MessageBox.Show("Se registro la data");
            }
            else
            {
                MessageBox.Show("Hubo un problema al registrar");
            }

        }

        private void dgvDatos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }




        //private void Celdas()
        //{
        //    // Abre un cuadro de diálogo para que el usuario seleccione un archivo Excel
        //    OpenFileDialog openFileDialog = new OpenFileDialog();
        //    //openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
        //    if (openFileDialog.ShowDialog() == DialogResult.OK)
        //    {
        //        // Carga el archivo Excel en una aplicación de Excel
        //        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        //        Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(openFileDialog.FileName);

        //        // Obtiene la hoja de cálculo que quieres mostrar en el DataGrid
        //        Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

        //        // Obtiene los datos de las columnas A y B de la hoja de cálculo
        //        //Microsoft.Office.Interop.Excel.Range range = worksheet.get_Range("A:B");
        //        //object[,] data = (object[,])range.Value2;

        //        // Crea un DataTable con los datos de las columnas A y B



        //        DataTable table = new DataTable();
        //        DataRow row = table.NewRow();


        //        string Cell_G7 = worksheet.Range["G7"].Value.ToString();
        //        table.Columns.Add("Empresa", typeof(string));
        //        row["Empresa"] = Cell_G7;

        //        string Cell_C6 = worksheet.Range["C6"].Value.ToString();
        //        table.Columns.Add("Nombre y Apellidos", typeof(string));
        //        row["Nombre y Apellidos"] = Cell_C6;

        //        string Cell_C7 = worksheet.Range["C7"].Value.ToString();
        //        table.Columns.Add("Fecha de Nacimiento");
        //        row["Fecha de Nacimiento"] = Cell_C7;

        //        string Cell_G8 = worksheet.Range["G8"].Value.ToString();
        //        table.Columns.Add("Área de Trabajo", typeof(string));
        //        row["Área de Trabajo"] = Cell_G8;

        //        string Cell_C9 = worksheet.Range["C9"].Value.ToString();
        //        table.Columns.Add("Cédula", typeof(string));
        //        row["Cédula"] = Cell_C9;

        //        string Cell_A14 = worksheet.Range["A14"].Value.ToString();
        //        table.Columns.Add("Patología", typeof(string));
        //        row["Patología"] = Cell_A14;

        //        string Cell_I14 = worksheet.Range["I14"].Value.ToString();
        //        table.Columns.Add("Tiempo de Padecerla", typeof(string));
        //        row["Tiempo de Padecerla"] = Cell_I14;


        //        string Cell_E14 = worksheet.Range["E14"].Value.ToString();
        //        table.Columns.Add("Tratamiento", typeof(string));
        //        row["Tratamiento"] = Cell_E14;

        //        string Cell_I46 = worksheet.Range["I46"].Value.ToString();
        //        table.Columns.Add("Peso (kg)");
        //        row["Peso (kg)"] = Cell_I46;



        //        table.Rows.Add(row);

        //        // Enlaza el DataTable al DataGrid
        //        dgvDatos.DataSource = table;


        //        // Cierra la aplicación de Excel
        //        workbook.Close();
        //        excel.Quit();
        //    }
        //}







    }
}
