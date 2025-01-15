using System;
using System.Collections.Generic;

namespace PruebaReportes.Models;

public partial class TfaUser
{
    public int UsersId { get; set; }

    public string UserName { get; set; } = null!;

    public string UserLastName { get; set; } = null!;

    public string UserEmail { get; set; } = null!;

    public int UserPoints { get; set; }

    public int? RolId { get; set; }

    public int? RolIdaddional { get; set; }

    public string Contrasenia { get; set; } = null!;

    public virtual TfaRol? Rol { get; set; }

    public virtual TfaRol? RolIdaddionalNavigation { get; set; }

    public virtual ICollection<TfaCertificate> TfaCertificateCertificateLeadNavigations { get; set; } = new List<TfaCertificate>();

    public virtual ICollection<TfaCertificate> TfaCertificateCertificateReceptorNavigations { get; set; } = new List<TfaCertificate>();

    public virtual ICollection<TfaHistory> TfaHistories { get; set; } = new List<TfaHistory>();

    public virtual ICollection<TfaTeam> TfaTeams { get; set; } = new List<TfaTeam>();

    public virtual ICollection<TfaUserPoint> TfaUserPoints { get; set; } = new List<TfaUserPoint>();

    public virtual ICollection<TfaUsersTask> TfaUsersTasks { get; set; } = new List<TfaUsersTask>();

    public virtual ICollection<TfaTeam> ColaboratorTeams { get; set; } = new List<TfaTeam>();
}
