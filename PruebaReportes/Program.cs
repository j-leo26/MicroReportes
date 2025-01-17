using Microsoft.EntityFrameworkCore;
using PruebaReportes.Models;
 
var builder = WebApplication.CreateBuilder(args);
 
// Add services to the container.
builder.Services.AddControllers();
 
// Configure Swagger/OpenAPI
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
 
// Configure the database context
builder.Services.AddDbContext<DbAb0bdeTalentseedsContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"))
);
 
// Add CORS
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowReactApp", builder =>
    {
        builder.WithOrigins("http://localhost:5173") 
               .AllowAnyMethod()
               .AllowAnyHeader();
    });
});
 
var app = builder.Build();
 
// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
 
// Enable CORS middleware
app.UseCors("AllowReactApp");
 
// Uncomment if using HTTPS
// app.UseHttpsRedirection();
 
app.UseAuthorization();
 
app.MapControllers();
 
app.Run();