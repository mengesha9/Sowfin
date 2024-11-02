using Microsoft.AspNetCore.Mvc;
using Sowfin.API.Lib;
using Sowfin.API.ViewModels;
using System;
using System.Collections.Generic;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CapexDepController : ControllerBase
    {
        [HttpPost]
        [Route("[action]")]
        public ActionResult<Object> AddCapexDep([FromBody] CapexBody model)
        {
            object[] capexArray = new object[model.depreciationArray.Length];
            model.depreciationArray.CopyTo(capexArray, 0);

            var depreciation = Depreciation.CalculateDepreciation(model.duration, model.add_to_year, capexArray, model.customPercent, model.customPercenList
                 , model.add_to_year_array);
            Dictionary<string, object> result = new Dictionary<string, object>{
                {"Depreciation", depreciation}
            };
            return Ok(result);

        }

    }
}