using System;
using System.Collections.Generic;

namespace PruebaReportes.Models;

public partial class TfaHistory
{
    public int HistoryId { get; set; }

    public DateOnly? HistoryEmission { get; set; }

    public int? UserHistoryId { get; set; }

    public int? UserCategoriesId { get; set; }

    public int? UserCertificateId { get; set; }

    public string? ReportType { get; set; }

    public virtual TfaCategory? UserCategories { get; set; }

    public virtual TfaCertificate? UserCertificate { get; set; }

    public virtual TfaUser? UserHistory { get; set; }
}
