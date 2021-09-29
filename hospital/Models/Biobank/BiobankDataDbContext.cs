using hospital.Migrations;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.Entity;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    public class BiobankDataDbContext : DbContext
    {
        //static private string s_migrationSqlitePath;
        //static BiobankDataDbContext()
        //{
        //    var exeDir = AppDomain.CurrentDomain.BaseDirectory;
        //    var exeDirInfo = new DirectoryInfo(exeDir);
        //    var projectDir = exeDirInfo.Parent.Parent.FullName;
        //    s_migrationSqlitePath = $@"{projectDir}\BiobankData.db";
        //}

        //public BiobankDataDbContext() : base(new SQLiteConnection($"DATA Source={s_migrationSqlitePath}"), false)
        //{
        //}

        //public BiobankDataDbContext(DbConnection connection) : base(connection, true)
        //{
        //    Migrate();
        //}
        public BiobankDataDbContext() : base("BiobankData")
        {
            Migrate();
        }

        private static readonly bool[] s_migrated = { false };

        private static void Migrate()
        {
            if (!s_migrated[0])
            {
                lock (s_migrated)
                {
                    if (!s_migrated[0])
                    {
                        Database.SetInitializer(new MigrateDatabaseToLatestVersion<BiobankDataDbContext,
                                                    Configuration>());
                        s_migrated[0] = true;
                    }
                }
            }
        }

        public virtual DbSet<TOTFAE> TOTFAEs { get; set; }
        public virtual DbSet<TOTFAO1> TOTFAO1s { get; set; }
        public virtual DbSet<TOTFAO2> TOTFAO2s { get; set; }
        public virtual DbSet<TOTFBE> TOTFBEs { get; set; }
        public virtual DbSet<TOTFBO1> TOTFBO1s { get; set; }
        public virtual DbSet<TOTFBO2> TOTFBO2s { get; set; }
        public virtual DbSet<LABD1> LABD1s { get; set; }
        public virtual DbSet<LABD2> LABD2s { get; set; }
        public virtual DbSet<LABM1> LABM1s { get; set; }
        public virtual DbSet<LABM2> LABM2s { get; set; }
        public virtual DbSet<CRLF> CRLFs { get; set; }
        public virtual DbSet<CRSF> CRSFs { get; set; }
        public virtual DbSet<DEATH> DEATHs { get; set; }
        public virtual DbSet<CASE> CASEs { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}