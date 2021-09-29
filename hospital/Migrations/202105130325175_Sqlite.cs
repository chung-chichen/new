namespace hospital.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Sqlite : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.CASE",
                c => new
                    {
                        SSN = c.String(nullable: false, maxLength: 128),
                        d3 = c.String(nullable: false, maxLength: 128),
                        m2 = c.String(maxLength: 2147483647),
                        m3 = c.String(maxLength: 2147483647),
                        m4 = c.String(maxLength: 2147483647),
                        m5 = c.String(maxLength: 2147483647),
                        m6 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.SSN, t.d3 });
            
            CreateTable(
                "dbo.CRLF",
                c => new
                    {
                        LF1_1 = c.String(nullable: false, maxLength: 128),
                        LF1_4 = c.String(nullable: false, maxLength: 128),
                        LF1_6 = c.String(nullable: false, maxLength: 128),
                        LF2_4 = c.String(nullable: false, maxLength: 128),
                        LF1_2 = c.String(maxLength: 2147483647),
                        LF1_3 = c.String(maxLength: 2147483647),
                        LF1_5 = c.String(maxLength: 2147483647),
                        LF1_7 = c.String(maxLength: 2147483647),
                        LF2_1 = c.String(maxLength: 2147483647),
                        LF2_2 = c.String(maxLength: 2147483647),
                        LF2_3 = c.String(maxLength: 2147483647),
                        LF2_3_1 = c.String(maxLength: 2147483647),
                        LF2_3_2 = c.String(maxLength: 2147483647),
                        LF2_5 = c.String(maxLength: 2147483647),
                        LF2_6 = c.String(maxLength: 2147483647),
                        LF2_7 = c.String(maxLength: 2147483647),
                        LF2_8 = c.String(maxLength: 2147483647),
                        LF2_9 = c.String(maxLength: 2147483647),
                        LF2_10_1 = c.String(maxLength: 2147483647),
                        LF2_10_2 = c.String(maxLength: 2147483647),
                        LF2_11 = c.String(maxLength: 2147483647),
                        LF2_12 = c.String(maxLength: 2147483647),
                        LF2_13 = c.String(maxLength: 2147483647),
                        LF2_13_1 = c.String(maxLength: 2147483647),
                        LF2_13_2 = c.String(maxLength: 2147483647),
                        LF2_14 = c.String(maxLength: 2147483647),
                        LF2_15 = c.String(maxLength: 2147483647),
                        LF3_1 = c.String(maxLength: 2147483647),
                        LF3_2 = c.String(maxLength: 2147483647),
                        LF3_3 = c.String(maxLength: 2147483647),
                        LF3_4 = c.String(maxLength: 2147483647),
                        LF3_5 = c.String(maxLength: 2147483647),
                        LF3_6 = c.String(maxLength: 2147483647),
                        LF3_7 = c.String(maxLength: 2147483647),
                        LF3_8 = c.String(maxLength: 2147483647),
                        LF3_10 = c.String(maxLength: 2147483647),
                        LF3_11 = c.String(maxLength: 2147483647),
                        LF3_12 = c.String(maxLength: 2147483647),
                        LF3_13 = c.String(maxLength: 2147483647),
                        LF3_14 = c.String(maxLength: 2147483647),
                        LF3_16 = c.String(maxLength: 2147483647),
                        LF3_17 = c.String(maxLength: 2147483647),
                        LF3_19 = c.String(maxLength: 2147483647),
                        LF3_21 = c.String(maxLength: 2147483647),
                        LF4_1 = c.String(maxLength: 2147483647),
                        LF4_1_1 = c.String(maxLength: 2147483647),
                        LF4_1_2 = c.String(maxLength: 2147483647),
                        LF4_1_3 = c.String(maxLength: 2147483647),
                        LF4_1_4 = c.String(maxLength: 2147483647),
                        LF4_1_4_1 = c.String(maxLength: 2147483647),
                        LF4_1_5 = c.String(maxLength: 2147483647),
                        LF4_1_5_1 = c.String(maxLength: 2147483647),
                        LF4_1_6 = c.String(maxLength: 2147483647),
                        LF4_1_7 = c.String(maxLength: 2147483647),
                        LF4_1_8 = c.String(maxLength: 2147483647),
                        LF4_1_9 = c.String(maxLength: 2147483647),
                        LF4_1_10 = c.String(maxLength: 2147483647),
                        LF4_2_1_1 = c.String(maxLength: 2147483647),
                        LF4_2_1_2 = c.String(maxLength: 2147483647),
                        LF4_2_1_3 = c.String(maxLength: 2147483647),
                        LF4_2_1_4 = c.String(maxLength: 2147483647),
                        LF4_2_1_5 = c.String(maxLength: 2147483647),
                        LF4_2_1_6 = c.String(maxLength: 2147483647),
                        LF4_2_1_8 = c.String(maxLength: 2147483647),
                        LF4_2_2_1 = c.String(maxLength: 2147483647),
                        LF4_2_2_2_1 = c.String(maxLength: 2147483647),
                        LF4_2_2_2_2 = c.String(maxLength: 2147483647),
                        LF4_2_2_2_3 = c.String(maxLength: 2147483647),
                        LF4_2_2_3_1 = c.String(maxLength: 2147483647),
                        LF4_2_2_3_2 = c.String(maxLength: 2147483647),
                        LF4_2_2_3_3 = c.String(maxLength: 2147483647),
                        LF4_2_3_1 = c.String(maxLength: 2147483647),
                        LF4_2_3_2 = c.String(maxLength: 2147483647),
                        LF4_2_3_3_1 = c.String(maxLength: 2147483647),
                        LF4_2_3_3_2 = c.String(maxLength: 2147483647),
                        LF4_2_3_3_3 = c.String(maxLength: 2147483647),
                        LF4_3_1 = c.String(maxLength: 2147483647),
                        LF4_3_2 = c.String(maxLength: 2147483647),
                        LF4_3_3 = c.String(maxLength: 2147483647),
                        LF4_3_4 = c.String(maxLength: 2147483647),
                        LF4_3_5 = c.String(maxLength: 2147483647),
                        LF4_3_6 = c.String(maxLength: 2147483647),
                        LF4_3_7 = c.String(maxLength: 2147483647),
                        LF4_3_8 = c.String(maxLength: 2147483647),
                        LF4_3_9 = c.String(maxLength: 2147483647),
                        LF4_3_10 = c.String(maxLength: 2147483647),
                        LF4_3_11 = c.String(maxLength: 2147483647),
                        LF4_3_12 = c.String(maxLength: 2147483647),
                        LF4_3_13 = c.String(maxLength: 2147483647),
                        LF4_3_14 = c.String(maxLength: 2147483647),
                        LF4_3_15 = c.String(maxLength: 2147483647),
                        LF4_4 = c.String(maxLength: 2147483647),
                        LF4_5_1 = c.String(maxLength: 2147483647),
                        LF4_5_2 = c.String(maxLength: 2147483647),
                        LF5_1 = c.String(maxLength: 2147483647),
                        LF5_2 = c.String(maxLength: 2147483647),
                        LF5_3 = c.String(maxLength: 2147483647),
                        LF5_4 = c.String(maxLength: 2147483647),
                        LF6_1 = c.String(maxLength: 2147483647),
                        LF7_1 = c.String(maxLength: 2147483647),
                        LF7_2 = c.String(maxLength: 2147483647),
                        LF7_3 = c.String(maxLength: 2147483647),
                        LF7_4 = c.String(maxLength: 2147483647),
                        LF7_5 = c.String(maxLength: 2147483647),
                        LF7_6 = c.String(maxLength: 2147483647),
                        LF8_1 = c.String(maxLength: 2147483647),
                        LF8_2 = c.String(maxLength: 2147483647),
                        LF8_3 = c.String(maxLength: 2147483647),
                        LF8_4 = c.String(maxLength: 2147483647),
                        LF8_5 = c.String(maxLength: 2147483647),
                        LF8_6 = c.String(maxLength: 2147483647),
                        LF8_7 = c.String(maxLength: 2147483647),
                        LF8_8 = c.String(maxLength: 2147483647),
                        LF8_9 = c.String(maxLength: 2147483647),
                        LF8_10 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.LF1_1, t.LF1_4, t.LF1_6, t.LF2_4 });
            
            CreateTable(
                "dbo.CRSF",
                c => new
                    {
                        SF1_1 = c.String(nullable: false, maxLength: 128),
                        SF1_4 = c.String(nullable: false, maxLength: 128),
                        SF1_6 = c.String(nullable: false, maxLength: 128),
                        SF2_4 = c.String(nullable: false, maxLength: 128),
                        SF1_2 = c.String(maxLength: 2147483647),
                        SF1_3 = c.String(maxLength: 2147483647),
                        SF1_5 = c.String(maxLength: 2147483647),
                        SF1_7 = c.String(maxLength: 2147483647),
                        SF2_1 = c.String(maxLength: 2147483647),
                        SF2_2 = c.String(maxLength: 2147483647),
                        SF2_3 = c.String(maxLength: 2147483647),
                        SF2_3_1 = c.String(maxLength: 2147483647),
                        SF2_3_2 = c.String(maxLength: 2147483647),
                        SF2_5 = c.String(maxLength: 2147483647),
                        SF2_6 = c.String(maxLength: 2147483647),
                        SF2_7 = c.String(maxLength: 2147483647),
                        SF2_8 = c.String(maxLength: 2147483647),
                        SF2_9 = c.String(maxLength: 2147483647),
                        SF2_10_1 = c.String(maxLength: 2147483647),
                        SF2_10_2 = c.String(maxLength: 2147483647),
                        SF2_11 = c.String(maxLength: 2147483647),
                        SF2_12 = c.String(maxLength: 2147483647),
                        SF4_1_1 = c.String(maxLength: 2147483647),
                        SF4_1_4 = c.String(maxLength: 2147483647),
                        SF4_2_1_3 = c.String(maxLength: 2147483647),
                        SF4_2_1_7 = c.String(maxLength: 2147483647),
                        SF4_3_3 = c.String(maxLength: 2147483647),
                        SF4_3_4 = c.String(maxLength: 2147483647),
                        SF4_3_6 = c.String(maxLength: 2147483647),
                        SF4_3_7 = c.String(maxLength: 2147483647),
                        SF4_3_9 = c.String(maxLength: 2147483647),
                        SF4_3_10 = c.String(maxLength: 2147483647),
                        SF4_3_11 = c.String(maxLength: 2147483647),
                        SF4_3_12 = c.String(maxLength: 2147483647),
                        SF4_3_14 = c.String(maxLength: 2147483647),
                        SF4_3_15 = c.String(maxLength: 2147483647),
                        SF4_4 = c.String(maxLength: 2147483647),
                        SF4_5_1 = c.String(maxLength: 2147483647),
                        SF4_5_2 = c.String(maxLength: 2147483647),
                        SF6_1 = c.String(maxLength: 2147483647),
                        SF7_1 = c.String(maxLength: 2147483647),
                        SF7_2 = c.String(maxLength: 2147483647),
                        SF7_3 = c.String(maxLength: 2147483647),
                        SF7_4 = c.String(maxLength: 2147483647),
                        SF7_5 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.SF1_1, t.SF1_4, t.SF1_6, t.SF2_4 });
            
            CreateTable(
                "dbo.DEATH",
                c => new
                    {
                        SSN = c.String(nullable: false, maxLength: 128),
                        d3 = c.String(nullable: false, maxLength: 128),
                        d2 = c.String(maxLength: 2147483647),
                        d4 = c.String(maxLength: 2147483647),
                        d5 = c.String(maxLength: 2147483647),
                        d6 = c.String(maxLength: 2147483647),
                        d7 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.SSN, t.d3 });
            
            CreateTable(
                "dbo.LABD1",
                c => new
                    {
                        h2 = c.String(nullable: false, maxLength: 128),
                        h4 = c.String(nullable: false, maxLength: 128),
                        h5 = c.String(nullable: false, maxLength: 128),
                        h6 = c.String(nullable: false, maxLength: 128),
                        h7 = c.String(nullable: false, maxLength: 128),
                        h1 = c.String(maxLength: 2147483647),
                        h3 = c.String(maxLength: 2147483647),
                        h8 = c.String(maxLength: 2147483647),
                        h9 = c.String(maxLength: 2147483647),
                        h10 = c.String(maxLength: 2147483647),
                        h11 = c.String(maxLength: 2147483647),
                        h12 = c.String(maxLength: 2147483647),
                        h13 = c.String(maxLength: 2147483647),
                        h14 = c.String(maxLength: 2147483647),
                        h15 = c.String(maxLength: 2147483647),
                        h19 = c.String(maxLength: 2147483647),
                        h20 = c.String(maxLength: 2147483647),
                        h22 = c.String(maxLength: 2147483647),
                        r1 = c.String(maxLength: 2147483647),
                        r2 = c.String(maxLength: 2147483647),
                        r3 = c.String(maxLength: 2147483647),
                        r4 = c.String(maxLength: 2147483647),
                        r5 = c.String(maxLength: 2147483647),
                        r6_1 = c.String(maxLength: 2147483647),
                        r6_2 = c.String(maxLength: 2147483647),
                        r7 = c.String(maxLength: 2147483647),
                        r8_1 = c.String(maxLength: 2147483647),
                        r10 = c.String(maxLength: 2147483647),
                        r12 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.h2, t.h4, t.h5, t.h6, t.h7 });
            
            CreateTable(
                "dbo.LABD2",
                c => new
                    {
                        h2 = c.String(nullable: false, maxLength: 128),
                        h4 = c.String(nullable: false, maxLength: 128),
                        h5 = c.String(nullable: false, maxLength: 128),
                        h6 = c.String(nullable: false, maxLength: 128),
                        h7 = c.String(nullable: false, maxLength: 128),
                        h1 = c.String(maxLength: 2147483647),
                        h3 = c.String(maxLength: 2147483647),
                        h8 = c.String(maxLength: 2147483647),
                        h9 = c.String(maxLength: 2147483647),
                        h10 = c.String(maxLength: 2147483647),
                        h11 = c.String(maxLength: 2147483647),
                        h12 = c.String(maxLength: 2147483647),
                        h13 = c.String(maxLength: 2147483647),
                        h14 = c.String(maxLength: 2147483647),
                        h15 = c.String(maxLength: 2147483647),
                        h19 = c.String(maxLength: 2147483647),
                        h20 = c.String(maxLength: 2147483647),
                        h22 = c.String(maxLength: 2147483647),
                        r1 = c.String(maxLength: 2147483647),
                        r2 = c.String(maxLength: 2147483647),
                        r3 = c.String(maxLength: 2147483647),
                        r4 = c.String(maxLength: 2147483647),
                        r5 = c.String(maxLength: 2147483647),
                        r6_1 = c.String(maxLength: 2147483647),
                        r6_2 = c.String(maxLength: 2147483647),
                        r7 = c.String(maxLength: 2147483647),
                        r8_1 = c.String(maxLength: 2147483647),
                        r10 = c.String(maxLength: 2147483647),
                        r12 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.h2, t.h4, t.h5, t.h6, t.h7 });
            
            CreateTable(
                "dbo.LABM1",
                c => new
                    {
                        h2 = c.String(nullable: false, maxLength: 128),
                        h4 = c.String(nullable: false, maxLength: 128),
                        h5 = c.String(nullable: false, maxLength: 128),
                        h6 = c.String(nullable: false, maxLength: 128),
                        h7 = c.String(nullable: false, maxLength: 128),
                        h8 = c.String(nullable: false, maxLength: 128),
                        h1 = c.String(maxLength: 2147483647),
                        h3 = c.String(maxLength: 2147483647),
                        h9 = c.String(maxLength: 2147483647),
                        h10 = c.String(maxLength: 2147483647),
                        h11 = c.String(maxLength: 2147483647),
                        h12 = c.String(maxLength: 2147483647),
                        h13 = c.String(maxLength: 2147483647),
                        h14 = c.String(maxLength: 2147483647),
                        h17 = c.String(maxLength: 2147483647),
                        h18 = c.String(maxLength: 2147483647),
                        h22 = c.String(maxLength: 2147483647),
                        h23 = c.String(maxLength: 2147483647),
                        h25 = c.String(maxLength: 2147483647),
                        r1 = c.String(maxLength: 2147483647),
                        r2 = c.String(maxLength: 2147483647),
                        r3 = c.String(maxLength: 2147483647),
                        r4 = c.String(maxLength: 2147483647),
                        r5 = c.String(maxLength: 2147483647),
                        r6_1 = c.String(maxLength: 2147483647),
                        r6_2 = c.String(maxLength: 2147483647),
                        r7 = c.String(maxLength: 2147483647),
                        r8_1 = c.String(maxLength: 2147483647),
                        r10 = c.String(maxLength: 2147483647),
                        r12 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.h2, t.h4, t.h5, t.h6, t.h7, t.h8 });
            
            CreateTable(
                "dbo.LABM2",
                c => new
                    {
                        h2 = c.String(nullable: false, maxLength: 128),
                        h4 = c.String(nullable: false, maxLength: 128),
                        h5 = c.String(nullable: false, maxLength: 128),
                        h6 = c.String(nullable: false, maxLength: 128),
                        h7 = c.String(nullable: false, maxLength: 128),
                        h8 = c.String(nullable: false, maxLength: 128),
                        h1 = c.String(maxLength: 2147483647),
                        h3 = c.String(maxLength: 2147483647),
                        h9 = c.String(maxLength: 2147483647),
                        h10 = c.String(maxLength: 2147483647),
                        h11 = c.String(maxLength: 2147483647),
                        h12 = c.String(maxLength: 2147483647),
                        h13 = c.String(maxLength: 2147483647),
                        h14 = c.String(maxLength: 2147483647),
                        h17 = c.String(maxLength: 2147483647),
                        h18 = c.String(maxLength: 2147483647),
                        h22 = c.String(maxLength: 2147483647),
                        h23 = c.String(maxLength: 2147483647),
                        h25 = c.String(maxLength: 2147483647),
                        r1 = c.String(maxLength: 2147483647),
                        r2 = c.String(maxLength: 2147483647),
                        r3 = c.String(maxLength: 2147483647),
                        r4 = c.String(maxLength: 2147483647),
                        r5 = c.String(maxLength: 2147483647),
                        r6_1 = c.String(maxLength: 2147483647),
                        r6_2 = c.String(maxLength: 2147483647),
                        r7 = c.String(maxLength: 2147483647),
                        r8_1 = c.String(maxLength: 2147483647),
                        r10 = c.String(maxLength: 2147483647),
                        r12 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.h2, t.h4, t.h5, t.h6, t.h7, t.h8 });
            
            CreateTable(
                "dbo.TOTFAE",
                c => new
                    {
                        t2 = c.String(nullable: false, maxLength: 128),
                        t3 = c.String(nullable: false, maxLength: 128),
                        t5 = c.String(nullable: false, maxLength: 128),
                        t6 = c.String(nullable: false, maxLength: 128),
                        d1 = c.String(nullable: false, maxLength: 128),
                        d2 = c.String(nullable: false, maxLength: 128),
                        d4 = c.String(maxLength: 2147483647),
                        d5 = c.String(maxLength: 2147483647),
                        d6 = c.String(maxLength: 2147483647),
                        d7 = c.String(maxLength: 2147483647),
                        d8 = c.String(maxLength: 2147483647),
                        d9 = c.String(maxLength: 2147483647),
                        d10 = c.String(maxLength: 2147483647),
                        d11 = c.String(maxLength: 2147483647),
                        d3 = c.String(maxLength: 2147483647),
                        d19 = c.String(maxLength: 2147483647),
                        d20 = c.String(maxLength: 2147483647),
                        d21 = c.String(maxLength: 2147483647),
                        d22 = c.String(maxLength: 2147483647),
                        d23 = c.String(maxLength: 2147483647),
                        d24 = c.String(maxLength: 2147483647),
                        d25 = c.String(maxLength: 2147483647),
                        d26 = c.String(maxLength: 2147483647),
                        d27 = c.String(maxLength: 2147483647),
                        d28 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 });
            
            CreateTable(
                "dbo.TOTFAO1",
                c => new
                    {
                        t2 = c.String(maxLength: 128),
                        t3 = c.String(maxLength: 128),
                        t5 = c.String(maxLength: 128),
                        t6 = c.String(maxLength: 128),
                        d1 = c.String(maxLength: 128),
                        d2 = c.String(maxLength: 128),
                        Id = c.Int(nullable: false, identity: true),
                        p1 = c.String(maxLength: 2147483647),
                        p2 = c.String(maxLength: 2147483647),
                        p3 = c.String(maxLength: 2147483647),
                        p4 = c.String(maxLength: 2147483647),
                        p5 = c.String(maxLength: 2147483647),
                        p6 = c.String(maxLength: 2147483647),
                        p7 = c.String(maxLength: 2147483647),
                        p9 = c.String(maxLength: 2147483647),
                        p10 = c.String(maxLength: 2147483647),
                        p13 = c.String(maxLength: 2147483647),
                        p14 = c.String(maxLength: 2147483647),
                        p15 = c.String(maxLength: 2147483647),
                        p17 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.TOTFAE", t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => t.Id);
            
            CreateTable(
                "dbo.TOTFAO2",
                c => new
                    {
                        t2 = c.String(maxLength: 128),
                        t3 = c.String(maxLength: 128),
                        t5 = c.String(maxLength: 128),
                        t6 = c.String(maxLength: 128),
                        d1 = c.String(maxLength: 128),
                        d2 = c.String(maxLength: 128),
                        Id = c.Int(nullable: false, identity: true),
                        p1 = c.String(maxLength: 2147483647),
                        p2 = c.String(maxLength: 2147483647),
                        p3 = c.String(maxLength: 2147483647),
                        p4 = c.String(maxLength: 2147483647),
                        p5 = c.String(maxLength: 2147483647),
                        p6 = c.String(maxLength: 2147483647),
                        p7 = c.String(maxLength: 2147483647),
                        p9 = c.String(maxLength: 2147483647),
                        p10 = c.String(maxLength: 2147483647),
                        p13 = c.String(maxLength: 2147483647),
                        p14 = c.String(maxLength: 2147483647),
                        p15 = c.String(maxLength: 2147483647),
                        p17 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.TOTFAE", t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => t.Id);
            
            CreateTable(
                "dbo.TOTFBE",
                c => new
                    {
                        t2 = c.String(nullable: false, maxLength: 128),
                        t3 = c.String(nullable: false, maxLength: 128),
                        t5 = c.String(nullable: false, maxLength: 128),
                        t6 = c.String(nullable: false, maxLength: 128),
                        d1 = c.String(nullable: false, maxLength: 128),
                        d2 = c.String(nullable: false, maxLength: 128),
                        d3 = c.String(maxLength: 2147483647),
                        d6 = c.String(maxLength: 2147483647),
                        d9 = c.String(maxLength: 2147483647),
                        d10 = c.String(maxLength: 2147483647),
                        d11 = c.String(maxLength: 2147483647),
                        d14 = c.String(maxLength: 2147483647),
                        d15 = c.String(maxLength: 2147483647),
                        d18 = c.String(maxLength: 2147483647),
                        d21 = c.String(maxLength: 2147483647),
                        d24 = c.String(maxLength: 2147483647),
                        d25 = c.String(maxLength: 2147483647),
                        d26 = c.String(maxLength: 2147483647),
                        d27 = c.String(maxLength: 2147483647),
                        d28 = c.String(maxLength: 2147483647),
                        d29 = c.String(maxLength: 2147483647),
                        d45 = c.String(maxLength: 2147483647),
                        d46 = c.String(maxLength: 2147483647),
                        d47 = c.String(maxLength: 2147483647),
                        d48 = c.String(maxLength: 2147483647),
                        d49 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 });
            
            CreateTable(
                "dbo.TOTFBO1",
                c => new
                    {
                        t2 = c.String(maxLength: 128),
                        t3 = c.String(maxLength: 128),
                        t5 = c.String(maxLength: 128),
                        t6 = c.String(maxLength: 128),
                        d1 = c.String(maxLength: 128),
                        d2 = c.String(maxLength: 128),
                        Id = c.Int(nullable: false, identity: true),
                        p1 = c.String(maxLength: 2147483647),
                        p2 = c.String(maxLength: 2147483647),
                        p3 = c.String(maxLength: 2147483647),
                        p5 = c.String(maxLength: 2147483647),
                        p6 = c.String(maxLength: 2147483647),
                        p7 = c.String(maxLength: 2147483647),
                        p8 = c.String(maxLength: 2147483647),
                        p14 = c.String(maxLength: 2147483647),
                        p15 = c.String(maxLength: 2147483647),
                        p16 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.TOTFBE", t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => t.Id);
            
            CreateTable(
                "dbo.TOTFBO2",
                c => new
                    {
                        t2 = c.String(maxLength: 128),
                        t3 = c.String(maxLength: 128),
                        t5 = c.String(maxLength: 128),
                        t6 = c.String(maxLength: 128),
                        d1 = c.String(maxLength: 128),
                        d2 = c.String(maxLength: 128),
                        Id = c.Int(nullable: false, identity: true),
                        p1 = c.String(maxLength: 2147483647),
                        p2 = c.String(maxLength: 2147483647),
                        p3 = c.String(maxLength: 2147483647),
                        p5 = c.String(maxLength: 2147483647),
                        p6 = c.String(maxLength: 2147483647),
                        p7 = c.String(maxLength: 2147483647),
                        p8 = c.String(maxLength: 2147483647),
                        p14 = c.String(maxLength: 2147483647),
                        p15 = c.String(maxLength: 2147483647),
                        p16 = c.String(maxLength: 2147483647),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.TOTFBE", t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => new { t.t2, t.t3, t.t5, t.t6, t.d1, t.d2 })
                .Index(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.TOTFBO2", new[] { "t2", "t3", "t5", "t6", "d1", "d2" }, "dbo.TOTFBE");
            DropForeignKey("dbo.TOTFBO1", new[] { "t2", "t3", "t5", "t6", "d1", "d2" }, "dbo.TOTFBE");
            DropForeignKey("dbo.TOTFAO2", new[] { "t2", "t3", "t5", "t6", "d1", "d2" }, "dbo.TOTFAE");
            DropForeignKey("dbo.TOTFAO1", new[] { "t2", "t3", "t5", "t6", "d1", "d2" }, "dbo.TOTFAE");
            DropIndex("dbo.TOTFBO2", new[] { "Id" });
            DropIndex("dbo.TOTFBO2", new[] { "t2", "t3", "t5", "t6", "d1", "d2" });
            DropIndex("dbo.TOTFBO1", new[] { "Id" });
            DropIndex("dbo.TOTFBO1", new[] { "t2", "t3", "t5", "t6", "d1", "d2" });
            DropIndex("dbo.TOTFAO2", new[] { "Id" });
            DropIndex("dbo.TOTFAO2", new[] { "t2", "t3", "t5", "t6", "d1", "d2" });
            DropIndex("dbo.TOTFAO1", new[] { "Id" });
            DropIndex("dbo.TOTFAO1", new[] { "t2", "t3", "t5", "t6", "d1", "d2" });
            DropTable("dbo.TOTFBO2");
            DropTable("dbo.TOTFBO1");
            DropTable("dbo.TOTFBE");
            DropTable("dbo.TOTFAO2");
            DropTable("dbo.TOTFAO1");
            DropTable("dbo.TOTFAE");
            DropTable("dbo.LABM2");
            DropTable("dbo.LABM1");
            DropTable("dbo.LABD2");
            DropTable("dbo.LABD1");
            DropTable("dbo.DEATH");
            DropTable("dbo.CRSF");
            DropTable("dbo.CRLF");
            DropTable("dbo.CASE");
        }
    }
}
