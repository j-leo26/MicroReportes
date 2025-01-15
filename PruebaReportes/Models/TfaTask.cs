using System;
using System.Collections.Generic;

namespace PruebaReportes.Models;

public partial class TfaTask
{
    public int TaskId { get; set; }

    public string TaskName { get; set; } = null!;

    public int? CategoryId { get; set; }

    public virtual TfaCategory? Category { get; set; }

    public virtual ICollection<TfaUsersTask> TfaUsersTasks { get; set; } = new List<TfaUsersTask>();
}
