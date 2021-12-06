using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace EDI_API.Models
{
    public class ShippingResultsHistory
    {
        [Key]
        public string cuid { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string cfacility { get; set; }

        [Column(TypeName = "varchar(12)")]
        public string ctpcode { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string cradleypn { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string cpartsnum { get; set; }

        [Column(TypeName = "int")]
        public int nshipqty { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string cshippernum { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string creferencenum { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string ctradingpartner { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string ccustpo { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string cdestination { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime dshipdate { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime dshiptime { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string dasndate { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string dasntime { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string cmovetype { get; set; }

        [Column(TypeName = "varchar(50)")]
        public string caddedby { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime ddateadded { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime dtimeadded { get; set; }
    }
}
