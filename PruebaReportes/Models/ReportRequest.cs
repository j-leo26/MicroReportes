namespace PruebaReportes.Models
{
    public class ReportRequest
    {
        public string ReportType { get; set; }   // Tipo de informe: "personalizado", "mensual", "semestral"
        public string StartDate { get; set; }     // Fecha de inicio del informe
        public string EndDate { get; set; }       // Fecha de fin del informe

        // Propiedad para almacenar los usuarios seleccionados
        public List<string> SelectedUsers { get; set; } = new List<string>();  // Lista de usuarios seleccionados

    }
}
