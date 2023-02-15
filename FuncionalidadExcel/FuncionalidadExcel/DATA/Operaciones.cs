using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FuncionalidadExcel.DATA
{
    public class Operaciones
    {
        public bool CargarData(DataTable tbData)
        {
            bool resultado = true;
            //string consulta = " SELECT *  FROM [SERVIMED].[MyMedico].[LayoutCargaFormulario_Detalle]";

            using (SqlConnection cn = new SqlConnection(Configuracion.Conexion))
            {
                cn.Open();
                using (SqlBulkCopy s = new SqlBulkCopy(cn))
                {

                    //ingresamos COLUMNAS ORIGEN | COLUMNAS DESTINOS
                    s.ColumnMappings.Add("Área de Trabajo", "AreaTrabajo");
                    s.ColumnMappings.Add("Cédula", "Cedula");
                    s.ColumnMappings.Add("Empresa", "Empresa");
                    s.ColumnMappings.Add("Fecha de Nacimiento", "FechaNacimiento");
                    s.ColumnMappings.Add("Nombre y Apellidos", "Nombre");
                    s.ColumnMappings.Add("Patología", "Patologia1");
                    s.ColumnMappings.Add("Tiempo de Padecerla", "PatologiaTiempo1");
                    s.ColumnMappings.Add("Tratamiento", "PatologiaTratamiento1");
                    s.ColumnMappings.Add("Peso (kg)", "Peso");
                    
                    //definimos la tabla a cargar
                    s.DestinationTableName = "SERVIMED.MyMedico.VigilanciaMedica";


                    s.BulkCopyTimeout = 1500;
                    try
                    {
                        s.WriteToServer(tbData);
                    }
                    catch (Exception e)
                    {
                        string st = e.Message;
                        resultado = false;
                    }


                }
            }

            return resultado;
        }
    }
}
