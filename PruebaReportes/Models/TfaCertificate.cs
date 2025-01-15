using System;
using System.Collections.Generic;

namespace PruebaReportes.Models;

public partial class TfaCertificate
{
    public int CertificatesId { get; set; }

    public string? CertificateMessage { get; set; }

    public int? CertificateReceptor { get; set; }

    public int? CertificateLead { get; set; }

    public int? CertificateCategory { get; set; }

    public virtual TfaCategory? CertificateCategoryNavigation { get; set; }

    public virtual TfaUser? CertificateLeadNavigation { get; set; }

    public virtual TfaUser? CertificateReceptorNavigation { get; set; }

    public virtual ICollection<TfaHistory> TfaHistories { get; set; } = new List<TfaHistory>();
}
