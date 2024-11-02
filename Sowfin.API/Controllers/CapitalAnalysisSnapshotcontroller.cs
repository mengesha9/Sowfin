using System.Reflection.Emit;
using System.Data;
using Internal;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Sowfin.API.ViewModels;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;

namespace Sowfin.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CapitalAnalysisSnapshotController : ControllerBase
    {


        private readonly ICapitalAnalysisSnapshots iCapitalAnalysisSnapshots = null;
        public CapitalAnalysisSnapshotController(ICapitalAnalysisSnapshots _iCapitalAnalysisSnapshots)
        {
            iCapitalAnalysisSnapshots = _iCapitalAnalysisSnapshots;
        }

        [HttpGet]
        [Route("[action]")]
        public ActionResult<Object> AddCapitalAnalysisSnapshot(string str)
        {
            Console.WriteLine("String is " + str);
            CapitalAnalysisSnapshotViewModel model = JsonConvert.DeserializeObject<CapitalAnalysisSnapshotViewModel>(str);

            CapitalAnalysisSnapshot capitalAnalysisSnapshot = new CapitalAnalysisSnapshot
            {
                SnapShot = model.SnapShot,
                Description = model.Description,
                UserId = model.UserId
            };
            try
            {
                iCapitalAnalysisSnapshots.Add(capitalAnalysisSnapshot);
                return Ok(new { id = capitalAnalysisSnapshot.Id, result = "Saved sucesfully" });
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }

        [HttpGet]
        [Route("GetAllCASnapShots/{UserId}")]
        public ActionResult<Object> GetAllCASnapShots(long UserId)
        {
            try
            {
                var SnapShot = iCapitalAnalysisSnapshots.FindBy(s => s.UserId == UserId);
                if (SnapShot == null)
                {
                    return NotFound();
                }
                return Ok(SnapShot);
            }
            catch (Exception)
            {
                return BadRequest();

            }

        }

        [HttpGet]
        [Route("CapitalAnalysisSnapshot/{id}")]
        public ActionResult<Object> CapitalAnalysisSnapshots(long Id)
        {
            try
            {
                var SnapShot = iCapitalAnalysisSnapshots.GetSingle(s => s.Id == Id);
                if (SnapShot == null)
                {
                    return NotFound();
                }
                return Ok(SnapShot);
            }
            catch (Exception)
            {
                return BadRequest();
            }

        }

        [HttpPut]
        [Route("UpdateCASnapshot")]
        public ActionResult<Object> UpdateCASnapshot([FromBody]CapitalAnalysisSnapshotViewModel snapshot)
        {
            if (ModelState.IsValid)
            {
                Console.WriteLine("Model state is valid");

                CapitalAnalysisSnapshot capitalAnalysisSnapshot = new CapitalAnalysisSnapshot
                {
                    Id = snapshot.Id,
                    SnapShot = snapshot.SnapShot,
                    Description = snapshot.Description,
                    UserId = snapshot.UserId
                };
                try
                {
                    iCapitalAnalysisSnapshots.Update(capitalAnalysisSnapshot);

                    return Ok();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName ==
                             "Microsoft.EntityFrameworkCore.DbUpdateConcurrencyException")
                    {
                        return NotFound();
                    }

                    return BadRequest();
                }

            }
            Console.WriteLine("Model state is not valid");  

            return BadRequest();
        }


        [HttpDelete]
        [Route("DeleteCSSnapShot/{id}")]
        public ActionResult<Object> DeleteCSSnapShot(long id)
        {
            Console.WriteLine("Id is " + id);
            if (id == 0)
            {
                Console.WriteLine("Id is 0");
                return BadRequest();
            }
            try
            {
                iCapitalAnalysisSnapshots.DeleteWhere(s => s.Id == id);
                return Ok(new { result = "Deleted sucessfully" });
            }
            catch (Exception)
            {

                return BadRequest();
            }

        }


    }
}
