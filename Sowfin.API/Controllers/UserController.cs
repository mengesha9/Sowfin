﻿using Sowfin.Model;
using Microsoft.AspNetCore.Mvc;
using Sowfin.API.ViewModels;
using Sowfin.API.ViewModels.User;
using Sowfin.Data.Abstract;
using System;
using System.Linq;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]

    public class UserController : Controller
    {

        private readonly IUserRepository iUser = null;
        public UserController(IUserRepository _iUser)
        {
            iUser = _iUser;

        }

        [HttpPost("Login")]
        public ActionResult<Object> Post([FromBody]UserLoginViewModel model)
        {
            if (!ModelState.IsValid) return BadRequest(ModelState);

            var user = iUser.GetSingle(u => u.UserName == model.username);

            if (user == null)
            {
                return BadRequest(new { email = "no user with this email" });
            }

            if (model.password != user.Password)
            {
                return BadRequest(new { password = "invalid password" });
            }

            return Ok(new
            {
                id = user.Id,
                userName = user.UserName,
                roleId = user.RoleId,
                token = "Test token end point",
                message = "success"
            });
        }

        [HttpPost("Register")]
        public IActionResult Register([FromBody]UserViewModel model)
        {
            if (model.UserName == null || model.Password == null)
            {
                return BadRequest(new { message = "Username or password cannot be empty", status = 404 });
            }
            try
            {
                User user = new User
                {
                    UserName = model.UserName,
                    Password = model.Password,
                    CreatedAt = DateTime.UtcNow,
                    UpdateAt = DateTime.UtcNow,
                    Deleted = 0,
                    Active = 0,
                    Email = model.Email,
                    RoleId = model.RoleId

                };
                iUser.Add(user);
                iUser.Commit();

            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message, statusCode = 400 });
            }

            return Ok(new { message = "User created sucessfulyy", statusCode = 200 });
        }

        [HttpPost("UsernameCheck")]
        public IActionResult UsernameCheck([FromBody]UserCheck userCheck)
        {
            var condition = iUser.FindBy(s => s.UserName == userCheck.username || s.Email == userCheck.email).ToArray();
            if (condition.Length == 0)
            {
                return Ok(new { message = "Username or email doesnt exists", statusCode = 200 });
            }
            else
            {
                return BadRequest(new { message = "Username or email already exists", statusCode = 400 });
            }
        }



    }


}
