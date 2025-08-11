namespace WebApplication1.Utils
{
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using System.IO;
    using OfficeOpenXml;
    using WebApplication1.Models;
    using System.Data.SqlClient;
    using OfficeOpenXml;
    public class ETLService
    {
        private readonly string _connectionString;

        public ETLService(string connectionString)
        {
            _connectionString = connectionString;
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage.License.SetNonCommercialPersonal("Tu Nombre");



            // necesario para EPPlus
        }

        public List<Persona> LeerExcel(string filePath)
        {
            var personas = new List<Persona>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // fila 1 son headers
                {
                    var nombre = worksheet.Cells[row, 1].Text;
                    var edadText = worksheet.Cells[row, 2].Text;
                    int edad = 0;
                    int.TryParse(edadText, out edad);

                    personas.Add(new Persona
                    {
                        Nombre = nombre,
                        Edad = edad
                    });
                }
            }

            return personas;
        }

        public void GuardarEnBase(List<Persona> personas)
        {
            using (var conn = new SqlConnection(_connectionString))
            {
                conn.Open();
                foreach (var persona in personas)
                {
                    var query = "INSERT INTO Personas (Nombre, Edad) VALUES (@Nombre, @Edad)";
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Nombre", persona.Nombre);
                        cmd.Parameters.AddWithValue("@Edad", persona.Edad);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }

}
