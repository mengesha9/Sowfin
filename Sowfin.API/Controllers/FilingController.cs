using AutoMapper;
using Sowfin.API.Notifications;
using Sowfin.API.ViewModels;
using Sowfin.Data.Abstract;
using Sowfin.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Text;
using Sowfin.Model.Entities;
using Sowfin.Data.Abstract;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    //[Authorize]
    [ApiController]
    public class FilingController : ControllerBase
    {
        IFilingRepository filingRepository;
        ISynonymRepository synonymRepository;
        IStatementRepository statementRepository;
        IHubContext<NotificationsHub> hubContext;
        IFindataRepository finDataRepository;
        IYearsRepository yearsRepository;
        IMapper mapper;

        public FilingController(
            IFilingRepository filingRepository,
            IHubContext<NotificationsHub> hubContext,
            ISynonymRepository synonymRepository,
            IStatementRepository statementRepository,
            IFindataRepository finDataRepository,
             IYearsRepository yearsRepository,
            IMapper mapper
        )
        {
            this.filingRepository = filingRepository;
            this.hubContext = hubContext;
            this.mapper = mapper;
            this.synonymRepository = synonymRepository;
            this.statementRepository = statementRepository;
            this.finDataRepository = finDataRepository;
            this.yearsRepository = yearsRepository;
        }

        [HttpGet()]
        public ActionResult<FilingsViewModel> GetAllFilings()
        {
            IEnumerable<Filing> filings = filingRepository.AllIncluding();

            FilingsViewModel ret = new FilingsViewModel
            {
                Filings = filings
                .Select(mapper.Map<FilingViewModel>)
                .ToList()
            };

            return Ok(ret);
        }

        //[HttpGet("{cik}")]
        //public ActionResult GetFilings(string cik)
        //{
        //    IEnumerable<Filing> filings = filingRepository.AllIncluding();

        //    var yrs = filings
        //        .Select(o => o.YearEnd)
        //        .Distinct()
        //        .ToArray();

        //    Array.Sort(yrs);
        //    Array.Reverse(yrs);

        //    var xbrlTags = filings
        //        .Select(mapper.Map<LineItem>)
        //        .Distinct()
        //        .ToArray();

        //    var dumps = filings
        //        .Select(mapper.Map<FilingViewModel>)
        //        .Where(o => o.Cik == cik)
        //        .OrderBy(o => (o.YearEnd))
        //        .GroupBy(o => (o.YearEnd))
        //        .ToList();

        //    var tbl = new List<FinRow>();

        //    foreach (var tag in xbrlTags)
        //    {
        //        var finCells = new List<FinCell>();
        //        foreach (var y in yrs)
        //        {
        //            var f = filings
        //                .Select(mapper.Map<FilingViewModel>)
        //                .Where(o => o.YearEnd == y)
        //                .FirstOrDefault();

        //            finCells.Add(new FinCell
        //            {
        //                YearEnd = y,
        //                Value = f != null ? f.Value : null
        //            });
        //        }

        //        var finRow = new FinRow
        //        {

        //            Field = tag.Field,
        //            FinCells = finCells
        //        };


        //        tbl.Add(finRow);

        //    }

        //    return Ok(tbl);
        //}

        [HttpGet]
        [Route("intialSetup")]
        public ActionResult GetFiling()
        {
            var filings = filingRepository.AllIncluding().OrderBy(m => m.Sequence)
                .GroupBy(s => new { s.FilingDate, s.StatementType })
                .GroupBy(s => new { s.Key.FilingDate });
            IEnumerable<Years> years = yearsRepository.AllIncluding();
            return Ok(new { filings, years });

        }

        [HttpGet]
        [Route("dataProcessor")]
        public ActionResult DataProcessor()
        {
            IEnumerable<Filing> filings = filingRepository.FindBy(s => s.Cik == "0000050863" && s.StatementType == "INCOME_STATEMENT");
            IEnumerable<Statement> statements = statementRepository.FindBy(s => s.StatementType == "INCOME_STATEMENT");

            var mappedSuccessfully = (from statement in statements
                                      from filing in filings
                                      where (filing.LineItem.ToLower() == statement.LineItem.ToLower() ||
                                      statement.Synonyms.Select(s => s.ToLower()).ToArray().Contains(filing.LineItem.ToLower()))
                                      select new FindataViewModelEx
                                      {
                                          Cik = filing.Cik,
                                          LineItem = filing.LineItem,
                                          StatementLineItem = statement.LineItem,
                                          Value = RemoveSpecialCharacters(filing.Value),
                                          OtherTags = statement.OtherTags,
                                          Category = statement.Category,
                                          Sequence = statement.Sequence,                                          FilingDate = filing.FilingDate,
                                          ParentItem = filing.ParentItem
                                      }).OrderBy(s => s.Sequence);

            var filingLineItem = String.Join(";", filings.Select(c => c.LineItem + "+" + c.Value + "+" + c.Id + "+" + c.ParentItem + "+" + c.Cik + "+" + c.FilingDate).ToArray()).Split(";");
            var mappedLineItem = String.Join(";", mappedSuccessfully.Select(c => c.LineItem).ToArray()).Split(";");

            var unMappedItems = filingLineItem.Where(i => !mappedLineItem.Contains(i.Split("+")[0])).Select(j => new { Id = j.Split("+")[2], Value = RemoveSpecialCharacters(j.Split("+")[1]), LineItem = j.Split("+")[0], Category = "", OtherTags = "", Sequency = "", ParentItem = j.Split("+")[3], Cik = j.Split("+")[4], FilingDate = j.Split("+")[5] });
            List<Findata> findatas = new List<Findata>();


            foreach (var item in mappedSuccessfully)
            {
                Findata findata = new Findata
                {
                    Cik = item.Cik,
                    LineItem = item.LineItem,
                    Value = item.Value,
                    OtherTags = item.OtherTags,
                    Category = item.Category,
                    Sequence = item.Sequence,
                    FilingDate = item.FilingDate,
                    ParentItem = item.ParentItem,
                    StatementType = "INCOME_STATEMENT"
                };
                findatas.Add(findata);
            }

            finDataRepository.AddMany(findatas);



            return Ok(new { mappedSuccessfully, unMappedItems });
        }

        public static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if (c >= '0' && c <= '9')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }
    }
}