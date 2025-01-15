using System;
using System.Collections.Generic;

namespace PruebaReportes.Models;

public partial class TfaRol
{
    public int RolId { get; set; }

    public string RolName { get; set; } = null!;

    public string RolDescription { get; set; } = null!;

    public virtual ICollection<TfaUser> TfaUserRolIdaddionalNavigations { get; set; } = new List<TfaUser>();

    public virtual ICollection<TfaUser> TfaUserRols { get; set; } = new List<TfaUser>();
}
