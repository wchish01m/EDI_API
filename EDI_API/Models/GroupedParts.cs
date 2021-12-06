using System.Collections.Generic;

namespace EDI_API.Models
{
    public class GroupedParts
    {
        public string tpCode { get; set; }
        public string partsNum { get; set; }
        public string sumQty { get; set; }
        public string shipDate { get; set; }
        public List<ShippingResults> shipDetails { get; set; }

        public GroupedParts()
        {
            this.tpCode = "";
            this.partsNum = "";
            this.sumQty = "";
            this.shipDate = "";
            this.shipDetails = new List<ShippingResults>();
        }

        public GroupedParts(string tpcode, string partsnum, string sumqty, string shipdate, List<ShippingResults> shipDetails)
        {
            this.tpCode = tpcode;
            this.partsNum = partsnum;
            this.sumQty = sumqty;
            this.shipDate = shipdate;
            this.shipDetails = shipDetails;
        }
    }
}
