#nullable enable
namespace EDI_API.Models
{
    public class ShippingResultDetails
    {
        public string partsNum { get; set; }
        public string serialNumber { get; set; }
        public string partLoc { get; set; }
        public string shift { get; set; }
        public string pkgQty { get; set; }
        public string printedBy { get; set; }
        public string prodDate { get; set; }
        public string dateCreated { get; set; }
        public string timestamp { get; set; }
        public string facility { get; set; }
        public string tpCode { get; set; }
        public string shipperNum { get; set; }
        public string referenceNum { get; set; }
        public string partsNum2 { get; set; }
        public string shipSerial { get; set; }
        public string serial { get; set; }

        public ShippingResultDetails()
        {
            this.partsNum = "";
            this.serialNumber = "";
            this.partLoc = "";
            this.shift = "";
            this.pkgQty = "";
            this.printedBy = "";
            this.prodDate = "";
            this.dateCreated = "";
            this.timestamp = "";
            this.facility = "";
            this.tpCode = "";
            this.shipperNum = "";
            this.referenceNum = "";
            this.partsNum2 = "";
            this.shipSerial = "";
            this.serial = "";
        }

        public ShippingResultDetails(string partsNum, string serialNumber, string partLoc, string shift, string pkgQty, string printedBy, string prodDate, string dateCreated, string timestamp, 
                         string facility, string tpCode, string shipperNum, string referenceNum, string partsNum2, string shipserial, string serial, string dateAdded, string timeAdded)
        {
            this.partsNum = partsNum;
            this.serialNumber = serialNumber;
            this.partLoc = partLoc;
            this.shift = shift;
            this.pkgQty = pkgQty;
            this.printedBy = printedBy;
            this.prodDate = prodDate;
            this.dateCreated = dateCreated;
            this.timestamp = timestamp;
            this.facility = facility;
            this.tpCode = tpCode;
            this.shipperNum = shipperNum;
            this.referenceNum = referenceNum;
            this.partsNum2 = partsNum2;
            this.shipSerial = shipserial;
            this.serial = serial;
        }
    }
}
