using System;
using System.Collections.Generic;

namespace PruebaReportes.Models;

public partial class TfaUserPoint
{
    public int UserPointsId { get; set; }

    public DateOnly? PointsDate { get; set; }

    public int? UserId { get; set; }

    public int? PointsAmount { get; set; }

    public virtual TfaUser? User { get; set; }
}
