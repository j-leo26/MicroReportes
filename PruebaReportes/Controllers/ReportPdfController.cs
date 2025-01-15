using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Data.SqlClient;
using System.IO;
using Microsoft.EntityFrameworkCore;
using PruebaReportes.Models;

namespace PdfService.Controllers
{
    [ApiController]
    [Route("fisawards/[controller]")]
    public class PdfReportController : ControllerBase
    {
        private readonly string _connectionString = "Data Source=DESKTOP-FQSJ40B;Initial Catalog=PruebaMicroservicios;Integrated Security=True";
        private readonly DbAb0bdeTalentseedsContext _context;

        public PdfReportController(DbAb0bdeTalentseedsContext context)
        {
            _context = context;
        }

        [HttpGet("GeneratePdf")]
        public IActionResult GeneratePdf(int usuarioId = 1)
        {
            string userName = "";
            string recognitionDate = DateTime.Now.ToString("dd 'de' MMMM 'del' yyyy");
            int points = 0;
            string pointDateRange = "";
            string reportType = "PDF"; // Tipo de reporte

            // Obtener datos del usuario con Entity Framework
            var user = _context.TfaUsers.FirstOrDefault(u => u.UsersId == usuarioId);
            if (user != null)
            {
                userName = user.UserName + user.UserLastName;
            }
            else
            {
                return NotFound("Usuario no encontrado.");
            }

            // Obtener datos de puntos usando Entity Framework
            var pointsData = _context.TfaUsers
                .Where(u => u.UsersId == usuarioId)
                .Select(u => new
                {
                    TotalPuntos = u.UserPoints, // Usamos directamente UserPoints de TfaUser
                    FechaInicio = u.TfaHistories.Min(h => h.HistoryEmission), // Fecha de inicio desde TfaHistory (ajustar si es necesario)
                    FechaFin = u.TfaHistories.Max(h => h.HistoryEmission) // Fecha de fin desde TfaHistory (ajustar si es necesario)
                })
                .FirstOrDefault();


            if (pointsData != null)
            {
                points = pointsData.TotalPuntos;
                pointDateRange = $"{pointsData.FechaInicio} - {pointsData.FechaFin}";
            }

            // Registrar la descarga en TFA_HISTORY
            var categoryId = 2; // Asegúrate de que este ID exista en la tabla TfaCategory
            var certificateId = 1; // Asegúrate de que este ID exista en la tabla TfaCertificate

            // Verificar si la categoría y el certificado existen
            var categoryExists = _context.TfaCategories.Any(c => c.CategoryId == categoryId);
            var certificateExists = _context.TfaCertificates.Any(c => c.CertificatesId == certificateId);

            if (!categoryExists)
            {
                return NotFound("Categoría no encontrada.");
            }

            if (!certificateExists)
            {
                return NotFound("Certificado no encontrado.");
            }

            // Registrar el historial (sin asignar un valor para HistoryId)
            var historyEntry = new TfaHistory
            {
                HistoryEmission = DateOnly.FromDateTime(DateTime.Now), // Ajustado para DateOnly
                UserHistoryId = usuarioId,
                UserCategoriesId = categoryId,
                UserCertificateId = certificateId,
                ReportType = reportType
            };

            // Agregar el nuevo registro en la tabla TfaHistory
            _context.TfaHistories.Add(historyEntry);

            // No es necesario asignar un valor explícito para HistoryId, ya que se genera automáticamente
            _context.SaveChanges();


            // Generación del PDF
            using (var memoryStream = new MemoryStream())
            {
                var document = new Document(PageSize.LETTER);
                var writer = PdfWriter.GetInstance(document, memoryStream);

                writer.PageEvent = new PdfHeaderFooter();

                document.Open();

                // Contenido del PDF
                document.Add(new Paragraph("\n\n\n\n"));
                var Fecha = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
                var Title1 = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);
                var Otor = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14);
                var Title2 = FontFactory.GetFont(FontFactory.HELVETICA_BOLDOBLIQUE, 14);
                var Att = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 13);
                var font = FontFactory.GetFont(FontFactory.HELVETICA, 12);

                document.Add(new Paragraph($"                                                                                                             Quito, {recognitionDate}", Fecha));
                document.Add(new Paragraph("\n"));

                document.Add(new Paragraph("Reconocimiento:", Title1));
                document.Add(new Paragraph("\n"));
                document.Add(new Paragraph("Otorgado por Talento Humano y Fisa group a:", Otor));
                document.Add(new Paragraph("\n"));
                document.Add(new Paragraph($"                                           Sr(a). Interno(a) {userName}.", Title2));
                document.Add(new Paragraph("\n"));
                document.Add(new Paragraph($"Por obtener la cantidad de puntos de {points} llevados acabo en las fechas: {pointDateRange}", font));
                document.Add(new Paragraph("\n\n"));
                document.Add(new Paragraph("Atentamente,", font));
                document.Add(new Paragraph("\n"));
                document.Add(new Paragraph("Talento Humano", Att));
                document.Add(new Paragraph("Fisa Group ", Att));

                document.Close();

                var pdfBytes = memoryStream.ToArray();
                return File(pdfBytes, "application/pdf", "Reconocimiento.pdf");
            }
        }
    }

    public class PdfHeaderFooter : PdfPageEventHelper
    {
        public override void OnEndPage(PdfWriter writer, Document document)
        {
            var cb = writer.DirectContent;

            string headerImagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "imageEncabezado.png");
            string footerImagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "imagePiePagina.png");

            if (File.Exists(headerImagePath))
            {
                Image headerImage = Image.GetInstance(headerImagePath);
                headerImage.ScaleToFit(510, 550);
                headerImage.SetAbsolutePosition(50, document.PageSize.Height - 100);
                cb.AddImage(headerImage);
            }

            if (File.Exists(footerImagePath))
            {
                Image footerImage = Image.GetInstance(footerImagePath);
                footerImage.ScaleToFit(677, 316);
                footerImage.SetAbsolutePosition(0, 0);
                cb.AddImage(footerImage);
            }
        }
    }
}
