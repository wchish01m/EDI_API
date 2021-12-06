using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EDI_API.Models
{
    public class ShippingResults
    {
        public string tpCode { get; set; }
        public string radleyPartNum { get; set; }
        public string partNum { get; set; }
        public string shipQty { get; set; }
        public string shipperNum { get; set; }
        public string referenceNum { get; set; }
        public string tradingPartner { get; set; }
        public string customerPO { get; set; }
        public string destination { get; set; }
        public string shipTime { get; set; }
        public string asnTime { get; set; }

        public ShippingResults()
        {
            this.tpCode = "";
            this.radleyPartNum = "";
            this.partNum = "";
            this.shipQty = "";
            this.shipperNum = "";
            this.referenceNum = "";
            this.tradingPartner = "";
            this.customerPO = "";
            this.destination = "";
            this.shipTime = "";
            this.asnTime = "";
        }

        public ShippingResults(string tpcode, string radleypartnum, string partnum, string shipqty, string shippernum, string referencenum, string tradingpartner, 
                           string customerpo, string destination, string shiptime, string asntime)
        {
            this.tpCode = tpcode;
            this.radleyPartNum = radleypartnum;
            this.partNum = partnum;
            this.shipQty = shipqty;
            this.shipperNum = shippernum;
            this.referenceNum = referencenum;
            this.tradingPartner = tradingpartner;
            this.customerPO = customerpo;
            this.destination = destination;
            this.shipTime = shiptime;
            this.asnTime = asntime;
        }
    }
}
