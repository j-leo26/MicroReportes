using System;
using System.Collections.Generic;

namespace PruebaReportes.Models;

public partial class TfaTeamstatus
{
    public int StatusId { get; set; }

    public string StatusName { get; set; } = null!;

    public virtual ICollection<TfaTeam> TfaTeams { get; set; } = new List<TfaTeam>();
}
