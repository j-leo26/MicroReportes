using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Threading.Tasks;
using PruebaReportes.Models;
using Microsoft.EntityFrameworkCore;

namespace ReportService.Controllers
{
    [ApiController]
    [Route("fisawards/[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly DbAb0bdeTalentseedsContext _context;

        // Constructor
        public ReportController(DbAb0bdeTalentseedsContext context)
        {
            _context = context;
        }
        // Endpoint para obtener la lista de usuarios
        // Endpoint para obtener la lista de usuarios
        [HttpGet("GetUsers")]
        public async Task<IActionResult> GetUsers()
        {
            try
            {
                var users = await _context.TfaUsers
                    .Select(u => new
                    {
                        u.UsersId,
                        u.UserName,
                        u.UserLastName,
                        u.UserEmail,
                        u.UserPoints
                    })
                    .ToListAsync();

                return Ok(users);
            }
            catch (Exception ex)
            {
                return BadRequest($"Ocurrió un error al obtener los usuarios: {ex.Message}");
            }
        }
        // Endpoint para generar el reporte en Excel
        [HttpGet("GenerateExcel")]
        public async Task<IActionResult> GenerateExcel([FromQuery] ReportRequest request)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (request == null || request.SelectedUsers == null || !request.SelectedUsers.Any())
                return BadRequest("La solicitud no contiene usuarios seleccionados.");

            try
            {
                using (var package = new ExcelPackage())
                {
                    foreach (var userEmail in request.SelectedUsers)
                    {
                        var userEntity = await _context.TfaUsers.FirstOrDefaultAsync(u => u.UserEmail == userEmail);

                        if (userEntity == null)
                            continue;

                        var worksheet = package.Workbook.Worksheets.Add($"Reporte - {userEntity.UserName}");
                        worksheet.Cells[1, 1].Value = "UsuarioId";
                        worksheet.Cells[1, 2].Value = "Nombre";
                        worksheet.Cells[1, 3].Value = "Periodo";
                        worksheet.Cells[1, 4].Value = "Tipo de Reporte"; // Cambiamos "Fecha" por "Tipo de Reporte"
                        worksheet.Cells[1, 5].Value = "Puntos Diarios";
                        worksheet.Cells[1, 6].Value = "Total Puntos";

                        // Definimos el rango de fechas según el tipo de reporte
                        DateTime startDate = DateTime.MinValue;
                        DateTime endDate = DateTime.Now; // Siempre será el día presente

                        switch (request.ReportType.ToLower())
                        {
                            case "mensual":
                                startDate = DateTime.Now.AddDays(-30); // Últimos 30 días
                                break;

                            case "trimestral":
                                startDate = DateTime.Now.AddDays(-90); // Últimos 90 días
                                break;

                            case "semestral":
                                startDate = DateTime.Now.AddDays(-180); // Últimos 180 días
                                break;

                            case "personalizado":
                                startDate = DateTime.Parse(request.StartDate);
                                endDate = DateTime.Parse(request.EndDate);
                                break;

                            default:
                                return BadRequest("Tipo de reporte no válido.");
                        }

                        // Obtener los resultados de la tabla TfaUsers
                        var results = await _context.TfaUsers
                            .Where(u => request.SelectedUsers.Contains(u.UserEmail) && u.UserEmail == userEmail)
                            .Select(u => new
                            {
                                UserId = u.UsersId,
                                UserName = u.UserName + " " + u.UserLastName,
                                Periodo = $"{startDate:yyyy-MM-dd} - {endDate:yyyy-MM-dd}",
                                PuntosDiarios = u.UserPoints,
                                TipoReporte = request.ReportType, // Agregamos el tipo de reporte
                            })
                            .ToListAsync();

                        int currentRow = 2;

                        // Si hay resultados, los agregamos al reporte
                        if (results.Any())
                        {
                            foreach (var result in results)
                            {
                                worksheet.Cells[currentRow, 1].Value = result.UserId;
                                worksheet.Cells[currentRow, 2].Value = result.UserName;
                                worksheet.Cells[currentRow, 3].Value = result.Periodo;
                                worksheet.Cells[currentRow, 4].Value = result.TipoReporte; // Mostramos el tipo de reporte
                                worksheet.Cells[currentRow, 5].Value = result.PuntosDiarios;
                                worksheet.Cells[currentRow, 6].Value = result.PuntosDiarios; // El total de puntos será igual a los puntos diarios

                                currentRow++;
                            }
                        }
                        else
                        {
                            worksheet.Cells[2, 1].Value = "No se encontraron puntos para este usuario en el período.";
                        }
                    }






                    var excelData = package.GetAsByteArray();

                    // Obtener el primer usuario antes de crear el historial
                    var firstUserEmail = request.SelectedUsers.FirstOrDefault();
                    if (string.IsNullOrEmpty(firstUserEmail))
                    {
                        return BadRequest("No se proporcionó un correo de usuario válido.");
                    }

                    // Buscar la entidad del primer usuario
                    var userEntityy = await _context.TfaUsers.FirstOrDefaultAsync(u => u.UserEmail == firstUserEmail);
                    if (userEntityy == null)
                    {
                        return BadRequest($"No se encontró el usuario con el correo: {firstUserEmail}");
                    }

                    // Registrar la descarga en la base de datos (TFA_HISTORY)
                    var historyEntry = new TfaHistory
                    {
                        HistoryEmission = DateOnly.FromDateTime(DateTime.Now),
                        UserHistoryId = userEntityy.UsersId, // Usamos el UsersId del usuario encontrado
                        UserCategoriesId = 2, // Ajusta el ID de la categoría según sea necesario
                        UserCertificateId = 1, // Ajusta el ID del certificado según sea necesario
                        ReportType = "excel" // Guarda el tipo de reporte como "excel"
                    };

                    _context.TfaHistories.Add(historyEntry);
                    await _context.SaveChangesAsync(); // Guarda los cambios en la base de datos


                    return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReporteUsuarios.xlsx");
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Ocurrió un error al generar el reporte: {ex.Message}");
            }
        }
    }
}



/*

using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ReportService.Controllers
{
    [ApiController]
    [Route("fisawards/[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly string connectionString = "Data Source=DESKTOP-FQSJ40B;Initial Catalog = PruebaMicroservicios; Integrated Security = True";  // Asegúrate de reemplazar esto con tu cadena de conexión correcta

        // Endpoint para obtener la lista de usuarios
        [HttpGet("GetUsers")]
        public IActionResult GetUsers()
        {
            var users = new List<string>();

            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    var query = "SELECT Nombre FROM dbo.Usuarios";
                    using (var command = new SqlCommand(query, connection))
                    {
                        var reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            users.Add(reader["Nombre"].ToString());
                        }
                    }
                }

                return Ok(users);  // Retorna la lista de nombres de usuario
            }
            catch (Exception ex)
            {
                return BadRequest($"Ocurrió un error al obtener los usuarios: {ex.Message}");
            }
        }

        // Endpoint para generar el reporte en Excel
        [HttpGet("GenerateExcel")]
        public IActionResult GenerateExcel([FromQuery] ReportRequest request)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var package = new ExcelPackage())
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Obtener la lista de usuarios desde la base de datos
                    List<string> users = new List<string>();
                    string query = "SELECT Nombre FROM dbo.Usuarios";

                    using (var command = new SqlCommand(query, connection))
                    {
                        var reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            users.Add(reader["Nombre"].ToString());
                        }
                        reader.Close();
                    }

                    foreach (var user in users)
                    {
                        // Obtener el UsuarioId a partir del nombre
                        int userId = 0;
                        string getUserIdQuery = "SELECT Id FROM dbo.Usuarios WHERE Nombre = @UserName";

                        using (var command = new SqlCommand(getUserIdQuery, connection))
                        {
                            command.Parameters.AddWithValue("@UserName", user);
                            var result = command.ExecuteScalar();
                            if (result != null)
                            {
                                userId = Convert.ToInt32(result);
                            }
                            else
                            {
                                return BadRequest("Usuario no encontrado.");
                            }
                        }

                        /*if (userId == 0)
                        {
                            continue;  // Si no se encontró un usuario válido, continuamos con el siguiente usuario
                        }

                        var worksheet = package.Workbook.Worksheets.Add($"Reporte - {user}");
                        worksheet.Cells[1, 1].Value = "UsuarioId";
                        worksheet.Cells[1, 2].Value = "Nombre";
                        worksheet.Cells[1, 3].Value = "Periodo";
                        worksheet.Cells[1, 4].Value = "Año";
                        worksheet.Cells[1, 5].Value = "TotalPuntos";

                        string queryReport = "";

                        // Verifica si el tipo de reporte es personalizado
                        if (request.ReportType == "personalizado")
                        {
                            // Convierte las fechas de inicio y fin a DateTime
                            DateTime startDate = DateTime.Parse(request.StartDate);
                            DateTime endDate = DateTime.Parse(request.EndDate);

                            queryReport = @"
                        SELECT UsuarioId, SUM(Puntos) AS TotalPuntos, YEAR(FechaPuntos) AS Año
                        FROM dbo.Puntos
                        WHERE FechaPuntos BETWEEN @StartDate AND @EndDate
                              AND UsuarioId = @UserId
                        GROUP BY UsuarioId, YEAR(FechaPuntos)";

                            // Establece los parámetros para las fechas de inicio y fin
                            using (var command = new SqlCommand(queryReport, connection))
                            {
                                command.Parameters.AddWithValue("@UserId", userId);
                                command.Parameters.AddWithValue("@StartDate", startDate);
                                command.Parameters.AddWithValue("@EndDate", endDate);

                                var reader = command.ExecuteReader();
                                int currentRow = 2;

                                // Verificar si se obtienen resultados
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        worksheet.Cells[currentRow, 1].Value = reader["UsuarioId"];
                                        worksheet.Cells[currentRow, 2].Value = user;
                                        worksheet.Cells[currentRow, 3].Value = $"{startDate.ToShortDateString()} - {endDate.ToShortDateString()}"; // Periodo
                                        worksheet.Cells[currentRow, 4].Value = reader["Año"];
                                        worksheet.Cells[currentRow, 5].Value = reader["TotalPuntos"];
                                        currentRow++;
                                    }
                                }
                                else
                                {
                                    worksheet.Cells[2, 1].Value = "No se encontraron datos para este usuario.";
                                }

                                reader.Close();
                            }
                        }
                        else
                        {
                            // Código para otros tipos de reporte (mensual, trimestral, etc.)
                            if (request.ReportType == "mensual")
                            {
                                queryReport = @"
                                SELECT UsuarioId, SUM(Puntos) AS TotalPuntos, MONTH(FechaPuntos) AS Mes, YEAR(FechaPuntos) AS Año
                                FROM dbo.Puntos
                                WHERE UsuarioId = @UserId
                                GROUP BY UsuarioId, MONTH(FechaPuntos), YEAR(FechaPuntos)
                                ORDER BY Año, Mes";

                                using (var command = new SqlCommand(queryReport, connection))
                                {
                                    command.Parameters.AddWithValue("@UserId", userId);

                                    var reader = command.ExecuteReader();
                                    int currentRow = 2;

                                    if (reader.HasRows)
                                    {
                                        while (reader.Read())
                                        {
                                            worksheet.Cells[currentRow, 1].Value = reader["UsuarioId"];
                                            worksheet.Cells[currentRow, 2].Value = user;
                                            worksheet.Cells[currentRow, 3].Value = $"{reader["Mes"]}/{reader["Año"]}";
                                            worksheet.Cells[currentRow, 4].Value = reader["Año"];
                                            worksheet.Cells[currentRow, 5].Value = reader["TotalPuntos"];
                                            currentRow++;
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Cells[2, 1].Value = "No se encontraron datos para este usuario.";
                                    }

                                    reader.Close();
                                }
                            }
                            else if (request.ReportType == "trimestral")
                            {
                                queryReport = @"
                                SELECT UsuarioId, SUM(Puntos) AS TotalPuntos, 
                                       (MONTH(FechaPuntos)-1)/3 + 1 AS Trimestre, YEAR(FechaPuntos) AS Año
                                FROM dbo.Puntos
                                WHERE UsuarioId = @UserId
                                GROUP BY UsuarioId, (MONTH(FechaPuntos)-1)/3 + 1, YEAR(FechaPuntos)
                                ORDER BY Año, Trimestre";

                                using (var command = new SqlCommand(queryReport, connection))
                                {
                                    command.Parameters.AddWithValue("@UserId", userId);

                                    var reader = command.ExecuteReader();
                                    int currentRow = 2;

                                    if (reader.HasRows)
                                    {
                                        while (reader.Read())
                                        {
                                            worksheet.Cells[currentRow, 1].Value = reader["UsuarioId"];
                                            worksheet.Cells[currentRow, 2].Value = user;
                                            worksheet.Cells[currentRow, 3].Value = $"Trimestre {reader["Trimestre"]}/{reader["Año"]}";
                                            worksheet.Cells[currentRow, 4].Value = reader["Año"];
                                            worksheet.Cells[currentRow, 5].Value = reader["TotalPuntos"];
                                            currentRow++;
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Cells[2, 1].Value = "No se encontraron datos para este usuario.";
                                    }

                                    reader.Close();
                                }
                            }
                            else if (request.ReportType == "semestral")
                            {
                                queryReport = @"
                                SELECT UsuarioId, SUM(Puntos) AS TotalPuntos, 
                                       CASE 
                                           WHEN MONTH(FechaPuntos) BETWEEN 1 AND 6 THEN 1
                                           WHEN MONTH(FechaPuntos) BETWEEN 7 AND 12 THEN 2
                                       END AS Semestre, YEAR(FechaPuntos) AS Año
                                FROM dbo.Puntos
                                WHERE UsuarioId = @UserId
                                GROUP BY UsuarioId, 
                                         CASE 
                                             WHEN MONTH(FechaPuntos) BETWEEN 1 AND 6 THEN 1
                                             WHEN MONTH(FechaPuntos) BETWEEN 7 AND 12 THEN 2
                                         END, YEAR(FechaPuntos)
                                ORDER BY Año, Semestre";

                                using (var command = new SqlCommand(queryReport, connection))
                                {
                                    command.Parameters.AddWithValue("@UserId", userId);

                                    var reader = command.ExecuteReader();
                                    int currentRow = 2;

                                    if (reader.HasRows)
                                    {
                                        while (reader.Read())
                                        {
                                            worksheet.Cells[currentRow, 1].Value = reader["UsuarioId"];
                                            worksheet.Cells[currentRow, 2].Value = user;
                                            worksheet.Cells[currentRow, 3].Value = $"Semestre {reader["Semestre"]}/{reader["Año"]}";
                                            worksheet.Cells[currentRow, 4].Value = reader["Año"];
                                            worksheet.Cells[currentRow, 5].Value = reader["TotalPuntos"];
                                            currentRow++;
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Cells[2, 1].Value = "No se encontraron datos para este usuario.";
                                    }

                                    reader.Close();
                                }
                            }

                        }

                        // Guardar los registros en TFA_HISTORY después de generar el reporte
                        var insertHistoryQuery = @"
                            INSERT INTO TFA_HISTORY (historyEmission, userHistoryID, userCategoriesID, userCertificateID, reportType)
                            VALUES (@historyEmission, @userHistoryID, @userCategoriesID, @userCertificateID, @reportType)";

                        using (var command = new SqlCommand(insertHistoryQuery, connection))
                        {
                            command.Parameters.AddWithValue("@historyEmission", DateTime.Now);
                            command.Parameters.AddWithValue("@userHistoryID", userId);
                            command.Parameters.AddWithValue("@userCategoriesID", 1); // Ajusta este valor según sea necesario
                            command.Parameters.AddWithValue("@userCertificateID", 1); // Ajusta este valor según sea necesario
                            command.Parameters.AddWithValue("@reportType", "Excel"); // Tipo de reporte (Excel)

                            command.ExecuteNonQuery();
                        }
                    }

                    // Obtener los datos del paquete Excel y devolverlo
                    var excelData = package.GetAsByteArray();
                    return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReporteUsuarios.xlsx");
                }

            }
            catch (Exception ex)
            {
                return BadRequest($"Ocurrió un error al generar el reporte: {ex.Message}");
            }
        }
    }

    // Clase para los parámetros de la solicitud (Request)
    public class ReportRequest
    {
        public List<string> SelectedUsers { get; set; }  // Lista de usuarios seleccionados
        public string ReportType { get; set; }           // Tipo de reporte (mensual, trimestral, etc.)
        public string StartDate { get; set; }            // Fecha de inicio (para reporte personalizado)
        public string EndDate { get; set; }              // Fecha de fin (para reporte personalizado)
    }
}

/*using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System;
using System.IO;

namespace ReportService.Controllers
{
    [ApiController]
    [Route("fisawards/[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly string connectionString = "Data Source=DESKTOP-FQSJ40B;Initial Catalog = PruebaMicroservicios; Integrated Security = True";  // Asegúrate de reemplazar esto con tu cadena de conexión correcta

        // Endpoint para generar el reporte en Excel
        [HttpGet("GenerateExcel")]
        public IActionResult GenerateExcel([FromQuery] ReportRequest request)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var package = new ExcelPackage())
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Obtener la lista de usuarios desde la base de datos
                    List<string> users = new List<string>();
                    string query = "SELECT Nombre FROM dbo.Usuarios";

                    using (var command = new SqlCommand(query, connection))
                    {
                        var reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            users.Add(reader["Nombre"].ToString());
                        }
                        reader.Close();
                    }

                    foreach (var user in users)
                    {
                        // Obtener el UsuarioId a partir del nombre
                        int userId = 0;
                        string getUserIdQuery = "SELECT Id FROM dbo.Usuarios WHERE Nombre = @UserName";

                        using (var command = new SqlCommand(getUserIdQuery, connection))
                        {
                            command.Parameters.AddWithValue("@UserName", user);
                            var result = command.ExecuteScalar();
                            if (result != null)
                            {
                                userId = Convert.ToInt32(result);
                            }
                        }

                        if (userId == 0)
                        {
                            continue;  // Si no se encontró un usuario válido, continuamos con el siguiente usuario
                        }

                        var worksheet = package.Workbook.Worksheets.Add($"Reporte - {user}");

                        // Insertar imagen en el encabezado
                        var headerImagePath = "imageEncabezado.png";  
                        if (System.IO.File.Exists(headerImagePath))
                        {
                            var headerImage = worksheet.Drawings.AddPicture("HeaderImage", headerImagePath);
                            headerImage.SetPosition(0, 0, 0, 0);  // Ajusta la posición
                            headerImage.SetSize(200, 100); // Ajusta el tamaño
                        }

                        // Insertar imagen en el pie de página
                        var footerImagePath = "imagePiePagina.png";  
                        if (System.IO.File.Exists(footerImagePath))
                        {
                            var footerImage = worksheet.Drawings.AddPicture("FooterImage", footerImagePath);
                            footerImage.SetPosition(worksheet.Dimension.End.Row + 1, 0, 0, 0);  // Ajusta la posición
                            footerImage.SetSize(200, 100); // Ajusta el tamaño
                        }

                        // Centrar el contenido de la celda
                        worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        // Añadir encabezado de columnas
                        worksheet.Cells[1, 1].Value = "UsuarioId";
                        worksheet.Cells[1, 2].Value = "Nombre";
                        worksheet.Cells[1, 3].Value = "Periodo";
                        worksheet.Cells[1, 4].Value = "Año";
                        worksheet.Cells[1, 5].Value = "TotalPuntos";

                        string queryReport = @"
                            SELECT UsuarioId, SUM(Puntos) AS TotalPuntos, YEAR(FechaPuntos) AS Año
                            FROM dbo.Puntos
                            WHERE UsuarioId = @UserId
                            GROUP BY UsuarioId, YEAR(FechaPuntos)";

                        // Obtener datos para el reporte
                        using (var command = new SqlCommand(queryReport, connection))
                        {
                            command.Parameters.AddWithValue("@UserId", userId);

                            var reader = command.ExecuteReader();
                            int currentRow = 2;

                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    worksheet.Cells[currentRow, 1].Value = reader["UsuarioId"];
                                    worksheet.Cells[currentRow, 2].Value = user;
                                    worksheet.Cells[currentRow, 3].Value = "Periodo Personalizado"; // Cambia esta parte según lo que sea necesario
                                    worksheet.Cells[currentRow, 4].Value = reader["Año"];
                                    worksheet.Cells[currentRow, 5].Value = reader["TotalPuntos"];
                                    currentRow++;
                                }
                            }
                            else
                            {
                                worksheet.Cells[2, 1].Value = "No se encontraron datos para este usuario.";
                            }

                            reader.Close();
                        }

                        // Crear bordes en la tabla
                        var range = worksheet.Cells[1, 1, 100, 5]; // Ajusta el rango según sea necesario
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    // Obtener los datos del paquete Excel y devolverlo
                    var excelData = package.GetAsByteArray();
                    return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReporteUsuarios.xlsx");
                }

            }
            catch (Exception ex)
            {
                return BadRequest($"Ocurrió un error al generar el reporte: {ex.Message}");
            }
        }
    }

    // Clase para los parámetros de la solicitud (Request)
    public class ReportRequest
    {
        public List<string> SelectedUsers { get; set; }  // Lista de usuarios seleccionados
        public string ReportType { get; set; }           // Tipo de reporte (mensual, trimestral, etc.)
        public string StartDate { get; set; }            // Fecha de inicio (para reporte personalizado)
        public string EndDate { get; set; }              // Fecha de fin (para reporte personalizado)
    }
}


*/