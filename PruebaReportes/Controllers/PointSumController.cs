using Microsoft.AspNetCore.Mvc;
using PruebaReportes.Models;

namespace PruebaReportes.Controllers
{
    public class PointSumController : Controller
    {
        private readonly DbAb0bdeTalentseedsContext _context;

        public PointSumController(DbAb0bdeTalentseedsContext context)
        {
            _context = context;
        }

        public IActionResult InsertPointsAndUpdateTotal()
        {
            // Actualizar el total de puntos de cada usuario
            var users = _context.TfaUsers.ToList();
            foreach (var user in users)
            {
                user.UserPoints = _context.TfaUserPoints
                    .Where(p => p.UserId == user.UsersId)
                    .Sum(p => (int?)p.PointsAmount) ?? 0; // Maneja valores null con ?? 0

                _context.TfaUsers.Update(user);
            }

            _context.SaveChanges();

            return RedirectToAction("Index"); // O cualquier otra vista o acción
        }
    }
}
