using Microsoft.AspNetCore.Mvc;
using Sowfin.API.ViewModels;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class BusinessUnitController : ControllerBase
    {
        private readonly IBusinessUnit iBusinessUnit = null;
        private readonly IProjectRepository iProject = null;

        public BusinessUnitController(IBusinessUnit _iBusinessUnit, IProjectRepository _iProject)
        {
            iBusinessUnit = _iBusinessUnit;
            iProject = _iProject;
        }

        [HttpPost]
        [Route("AddBusinessUnit")]
        public ActionResult<Object> AddBusinessUnit([FromBody] BusinessUnitViewModel model)
        {

            try
            {
                BusinessUnit businessUnit = new BusinessUnit
                {
                    Name = model.Name,
                    Description = model.Description,
                    Status = model.Status,
                    YearEstablished = model.YearEstablished,
                    UserId = model.UserId
                };

                if (model.Id == 0)
                {
                    iBusinessUnit.Add(businessUnit);
                }
                else
                {
                    businessUnit.Id = model.Id;
                    iBusinessUnit.Update(businessUnit);
                }

                iBusinessUnit.Commit();

                return new
                {
                    id = businessUnit.Id,
                    result = model.Id == 0 ? "Business Unit Created Sucessfully" : "Business Unit Modified Sucessfully"
                };
            }
            catch (Exception ex)
            {
                return BadRequest(new { Message = "Invalid Entry", StatusCode = 400 });
            }
        }

        [HttpGet]
        [Route("GetAllBusinessUnit/{UserId}")]
        public ActionResult<Object> GetAllBusinessUnit(long UserId)
        {
            try
            {
                var businessUnit = iBusinessUnit.FindBy(s => s.UserId == UserId && s.Status==1);
                if (businessUnit == null)
                {
                    return NotFound(new { Message = "UserId or Record not Found", StatusCode = 404 });
                }
                return Ok(new { result = businessUnit, StatusCode = 200 });
            }
            catch (Exception)
            {
                return BadRequest(new { Message = "Invalid Request", StatusCode = 400 });

            }
        }

        [HttpGet]
        [Route("GetBusinessUnits/{Id}")]
        public ActionResult<Object> GetBusinessUnits(long Id)
        {
            try
            {
                var businessUnit = iBusinessUnit.GetSingle(s => s.Id == Id && s.Status==1);
                if (businessUnit == null)
                {
                    return NotFound(new { Message = "UserId or Record not Found", StatusCode = 404 });
                }
                return Ok(new { result = businessUnit, StatusCode = 200 });
            }
            catch (Exception)
            {
                return BadRequest(new { Message = "Invalid Request", StatusCode = 400 });

            }

        }

        [HttpDelete]
        [Route("DeleteBusinessUnits/{Id}")]
        public ActionResult<Object> DeleteBusinessUnits(long Id)
        {
            try
            {
                var businessUnit = iBusinessUnit.GetSingle(s => s.Id == Id);
                if (businessUnit == null)
                {
                    return NotFound(new { Message = "Record not found", StatusCode = 0 });
                }
                else
                {
                    //List<Projects> ProjectUnit = new List<Projects>();
                    //ProjectUnit = iProject.FindBy(s => s.BusinessUnitId == businessUnit.Id).ToList();
                    //if (ProjectUnit != null && ProjectUnit.Count > 0)
                    //{
                    //    iProject.DeleteMany(ProjectUnit);
                    //    iProject.Commit();
                    //}
                    businessUnit.Status = 0;
                    iBusinessUnit.Update(businessUnit);
                    iBusinessUnit.Commit();                   
                    return Ok(new { Message = "Record deleted", StatusCode = 1 });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { Message = "Invalid Request", StatusCode = 400 });
            }
        }
    }
}
