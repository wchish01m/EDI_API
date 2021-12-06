using EDI_API.Models;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

#nullable enable

namespace EDI_API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ShippingResultDetailsController : ControllerBase
    {
        [HttpGet("GetDetailsBySearch")]
        public List<ShippingResultDetails> GetDetailsBySearch(string? partNum = null, string? tpCode = null, string? shipperNum = null, string? referenceNum = null, string? custSerial = null,
                                                  string? topSerial = null, string? facility = null, string? startDate = null, string? endDate = null)
        {
            return Queries.GetDetailsBySearch(partNum, tpCode, shipperNum, referenceNum, custSerial, topSerial, facility, startDate, endDate);
        }
    }
}
