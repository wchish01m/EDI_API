using System;

namespace EDI_API.Models
{
    public class SalesRequirements
    {
        public string facility { get; set; }
        public string tradingpartner { get; set; }
        public string partsnum { get; set; }
        public string destination { get; set; }
        public string dockcode { get; set; }
        public string referencenum { get; set; }
        public string recordtype { get; set; }
        public string reqdate { get; set; }
        public string reqtime { get; set; }
        public string reqqty { get; set; }
        public string netqty { get; set; }
        public string releasenum { get; set; }
        public string releasedate { get; set; }
        public string totalshipped { get; set; }
        public string latestdateshipped { get; set; }
        public string asnstatus { get; set; }
        public SalesRequirements()
        {
            this.facility = "";
            this.tradingpartner = "";
            this.partsnum = "";
            this.destination = "";
            this.dockcode = "";
            this.referencenum = "";
            this.recordtype = "";
            this.reqdate = "";
            this.reqtime = "";
            this.reqqty = "";
            this.netqty = "";
            this.releasenum = "";
            this.releasedate = "";
            this.totalshipped = "";
            this.latestdateshipped = "";
            this.asnstatus = "";
        }

        public SalesRequirements(string facility, string tradingpartner, string partsnum, string destination, string dockcode, string referencenum, string recordtype, 
                                 string reqdate, string reqtime, string reqqty, string netqty, string releasenum, string releasedate, string totalshipped, 
                                 string latestdateshipped, string asnstatus)
        {
            this.facility = facility;
            this.tradingpartner = tradingpartner;
            this.partsnum = partsnum;
            this.destination = destination;
            this.dockcode = dockcode;
            this.referencenum = referencenum;
            this.recordtype = recordtype;
            this.reqdate = reqdate;
            this.reqtime = reqtime;
            this.reqqty = reqqty;
            this.netqty = netqty;
            this.releasenum = releasenum;
            this.releasedate = releasedate;
            this.totalshipped = totalshipped;
            this.latestdateshipped = latestdateshipped;
            this.asnstatus = asnstatus;
        }
    }
}
