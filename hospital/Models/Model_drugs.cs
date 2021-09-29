namespace hospital.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class Model_drugs : DbContext
    {
        public Model_drugs()
            : base("name=Model_drugs")
        {
        }

        //public virtual DbSet<New_健保用藥查詢品項> Drugs_name { get; set; }
        public virtual DbSet<New_NHI_Drugs_item> Drugs_name { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}
