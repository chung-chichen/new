namespace hospital.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Inspection")]
    public partial class Inspection
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(20)]
        public string Code { get; set; }

        [Key]
        [Column(Order = 2)]
        public string ChtName { get; set; }

        public string EngName { get; set; }

        public int? LblNumber { get; set; }

        [StringLength(50)]
        public string LblDate { get; set; }

        [StringLength(20)]
        public string IsUpdate { get; set; }
    }
}
