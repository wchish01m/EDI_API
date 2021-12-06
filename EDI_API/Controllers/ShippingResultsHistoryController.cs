using EDI_API.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace EDI_API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ShippingResultsHistoryController : ControllerBase
    {
        private readonly ShippingResultsHistoryContext _context;

        public ShippingResultsHistoryController(ShippingResultsHistoryContext context)
        {
            _context = context;
        }

        /**
         * This function will get all the data in the Shipping_Results_History table
         * within the specified range.
         */
        [HttpGet("GetByRange/{startDate}/{endDate}")]
        public async Task<dynamic> GetResultsInRange(DateTime startDate, DateTime endDate)
        {
            var resultsInRange = await _context.Shipping_Results_History
                .Where(e => e.dshipdate >= startDate && e.dshipdate <= endDate)
                .OrderBy(e => e.dshipdate)
                .ToListAsync();
            return resultsInRange;
        }
    }
}
