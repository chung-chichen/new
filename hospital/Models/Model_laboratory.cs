namespace hospital.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class Model_laboratory : DbContext
    {
        public Model_laboratory()
            : base("name=Model_laboratory")
        {
        }

        public virtual DbSet<Inspection> Inspection { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}
