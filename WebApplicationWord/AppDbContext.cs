using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.EntityFrameworkCore;

namespace WebApplicationWord
{
    public class AppDbcontext : DbContext
    {
        public AppDbcontext(DbContextOptions<AppDbcontext> options)
            : base(options)
        { }

        public DbSet<Models.Student> Students { get; set; }
    }
}