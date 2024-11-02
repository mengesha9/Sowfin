using AutoMapper;
using Sowfin.API.Notifications;
using Sowfin.API.ViewModels;
using Sowfin.Data.Abstract;
using Sowfin.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using System;
using System.Linq;


namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    //[Authorize]
    [ApiController]
    public class StatementsController : ControllerBase
    {
        IStatementRepository statementRepository;
        ISynonymRepository synonymRepository;
        IHubContext<NotificationsHub> hubContext;
        IMapper mapper;

        public StatementsController(
            IStatementRepository statementRepository,
            ISynonymRepository synonymRepository,
            IHubContext<NotificationsHub> hubContext,
            IMapper mapper
        )
        {
            this.statementRepository = statementRepository;
            this.synonymRepository = synonymRepository;
            this.hubContext = hubContext;
            this.mapper = mapper;
        }

        [HttpGet()]
        public ActionResult GetStatements()
        {
            var statements = statementRepository.AllIncluding();
            return Ok(statements.Select(mapper.Map<StatementViewModel>).ToList());
        }

        /*
        [HttpGet("{id}")]
        public ActionResult<StoryDetailViewModel> GetStoryDetail(string id)
        {
            var story = statementRepository.GetSingle(s => s.Id == id, s => s.Owner, s => s.Likes);
            var userId = HttpContext.User.Identity.Name;
            var liked = story.Likes.Exists(l => l.UserId == userId);
            
            return mapper.Map<Story, StoryDetailViewModel>(
                story,
                opt => opt.AfterMap((src, dest) => dest.Liked = liked)
            );
        }
        */

        [HttpPost]
        public ActionResult<StatementCreationViewModel> Post([FromBody]UpdateStatementViewModel model)
        {
            if (!ModelState.IsValid) return BadRequest(ModelState);

            var statement = new Statement
            {
                LineItem = model.LineItem,
                Category = model.Category,
                OtherTags = model.OtherTags,
                IsMultiInstances = model.IsMultiInstances,
                Description = model.Description,

                Sequence = model.Sequence,
                Synonyms = model.Synonyms
            };

            statementRepository.Add(statement);
            statementRepository.Commit();

            return new StatementCreationViewModel
            {
                StatementId = statement.StatementId
            };
        }

        [HttpDelete("{id}")]
        public ActionResult Delete(string id)
        {
            long statementId = 0;
            try
            {
                statementId = long.Parse(id);
            }
            catch (Exception e)
            {
                return BadRequest("Invalid ID");
            }

            statementRepository.DeleteWhere(statement => statement.StatementId == statementId);
            statementRepository.Commit();

            return NoContent();
        }

        [HttpPatch("{id}")]
        public ActionResult Patch(string id, [FromBody]UpdateStatementViewModel model)
        {
            if (!ModelState.IsValid) return BadRequest(ModelState);
            long statementId = 0;
            try
            {
                statementId = long.Parse(id);
            }
            catch (Exception e)
            {
                return BadRequest("Invalid ID");
            }

            var newStatement = statementRepository.GetSingle(s => s.StatementId == statementId);

            newStatement.LineItem = model.LineItem;
            newStatement.Category = model.Category;
            newStatement.OtherTags = model.OtherTags;
            newStatement.IsMultiInstances = model.IsMultiInstances;
            newStatement.Description = model.Description;
            newStatement.Sequence = model.Sequence;
            newStatement.Synonyms = model.Synonyms;
            statementRepository.Update(newStatement);
            statementRepository.Commit();

            return NoContent();
        }

    }


}