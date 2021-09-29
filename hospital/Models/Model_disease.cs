namespace hospital.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class Model_disease : DbContext
    {
        public Model_disease()
            : base("name=Model_disease")
        {
        }

        public virtual DbSet<ICD9> ICD9 { get; set; }
        public virtual DbSet<ICD10> ICD10 { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}
