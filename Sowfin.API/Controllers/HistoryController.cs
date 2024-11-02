using AutoMapper;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using Sowfin.API.Lib;
using Sowfin.API.ViewModels;
using Sowfin.API.ViewModels.History;
using Sowfin.API.ViewModels.PayoutPolicy;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HistoryController : ControllerBase
    {
        const string HISTORYSNAPSHOT = "History_SnapShot";
        private readonly IHistory _iHistory = null;
        private readonly IWebHostEnvironment  _hostingEnvironment = null;
        private readonly ISnapshots _isnapShots = null;
        ICurrentSetup iCurrentSetup;
        IMapper mapper;
        public HistoryController(IWebHostEnvironment  hostingEnvironment, IHistory iHistory, ISnapshots isnapShots, ICurrentSetup _iCurrentSetup, IMapper mapper)
        {
            this.mapper = mapper;
            _hostingEnvironment = hostingEnvironment;
            _iHistory = iHistory;
            _isnapShots = isnapShots;
            iCurrentSetup = _iCurrentSetup;
        }

        [HttpPost("EvaluateHistory")]
        public ActionResult<Object> EvaluateHistory([FromBody] HistoryObject model)
        {

            model.payoutTable = ConvterUnits(model.payoutTable, 0);
            model.shareTable = ConvterUnits(model.shareTable, 0);
            model.currentCapitalTable = ConvterUnits(model.currentCapitalTable, 0);
            model.otherInputTable = ConvterUnits(model.otherInputTable, 0);
            model.financialTable = ConvterUnits(model.financialTable, 0);
            var result = MathHistory.SummaryOuput(model);
            HistorySummary resultEx = new HistorySummary();
            resultEx = result;
            resultEx.TotalCashReturned.Insert(0, "Total Cash Returned");
            resultEx.TotalCashReturnedFCFE.Insert(0, "Total Cash Returned /FCFE");
            resultEx.TotalPayoutQuarter.Insert(0, "Total Payout");
            resultEx.DPSQuarter.Insert(0, "DPS(Dividents Per Share - Basic)");
            resultEx.PayoutRatioQuarter.Insert(0, "Dividents Payout Ratio %");
            resultEx.DividendYieldQuarter.Insert(0, "Dividend Yield %");
            resultEx.TotalPayoutOne.Insert(0, "Total Payout");
            resultEx.DPSOne.Insert(0, "DPS(Dividents Per Share - Basic)");
            resultEx.DividendPayoutRatioOne.Insert(0, "Dividend Payout Ratio");
            resultEx.DividendYieldOne.Insert(0, "Dividend Yield");
            resultEx.TotalPayout.Insert(0, "Total Payout");
            resultEx.SharesRepurchased.Insert(0, "Shared Repurchased");
            resultEx.DebtMarketEquity.Insert(0, "Debt-to-Market Equity Ratio(%)");
            resultEx.DebtBookEquity.Insert(0, "Debt-to-Book Equity Ratio");
            resultEx.EBITInterestCoverage.Insert(0, "EBIT Interest Coverage(x)");
            resultEx.EBITDAInterestCoverage.Insert(0, "EBITDA Interest Coverage(x)");
            resultEx.CashFlow.Insert(0, "Cash Flow from Operation/Total Debt(%)");
            resultEx.TotalDebt.Insert(0, "Total Debt / EBITDA (x)");
            resultEx.DebtRating.Insert(0, "S&P Debt Rating");
            resultEx.CashEquivalent.Insert(0, "Cash & Equivalent");
            resultEx.CashNeededCapital.Insert(0, "Cash Needed for Working Capital");
            resultEx.ExcessCash.Insert(0, "Excess Cash");
            try
            {
                if (model.id == 0)
                {
                    History history = new History
                    {
                        UserId = model.userId,
                        PayoutTable = ConvterObjectToString(model.payoutTable),
                        ShareTable = ConvterObjectToString(model.shareTable),
                        CurrentCapitalTable = ConvterObjectToString(model.currentCapitalTable),
                        CurrentCostOfCapTable = ConvterObjectToString(model.currentCostOfCapTable),
                        OtherInputTable = ConvterObjectToString(model.otherInputTable),
                        FinancialTable = ConvterObjectToString(model.financialTable),
                        SummaryOutput = ConvterObjectToString(result),
                        StartYear = model.startYear,
                        EndYear = model.endYear,
                        SummaryFlag = model.summaryFlag
                    };
                    _iHistory.Add(history);
                    _iHistory.Commit();
                    return Ok(resultEx);

                }
                else
                {
                    History updateHistory = new History
                    {
                        Id = model.id,
                        UserId = model.userId,
                        PayoutTable = ConvterObjectToString(model.payoutTable),
                        ShareTable = ConvterObjectToString(model.shareTable),
                        CurrentCapitalTable = ConvterObjectToString(model.currentCapitalTable),
                        CurrentCostOfCapTable = ConvterObjectToString(model.currentCostOfCapTable),
                        OtherInputTable = ConvterObjectToString(model.otherInputTable),
                        FinancialTable = ConvterObjectToString(model.financialTable),
                        SummaryOutput = ConvterObjectToString(result),
                        StartYear = model.startYear,
                        EndYear = model.endYear,
                        SummaryFlag = model.summaryFlag
                    };
                    _iHistory.Update(updateHistory);
                    _iHistory.Commit();
                    return Ok(resultEx);

                }
            }
            catch (Exception ss)
            {
                return BadRequest();
            }
        }

        private CurrentSetupViewModel GetCurrentSetupDetailsByUserId(long UserId)
        {
            CurrentSetupViewModel currentSetupVM = new CurrentSetupViewModel();
            var tblCurrentSetupObj = iCurrentSetup.GetSingle(x => x.UserId == UserId);
            if (tblCurrentSetupObj != null)
            {
                //map currentsetup to VM	
                currentSetupVM = mapper.Map<CurrentSetup, CurrentSetupViewModel>(tblCurrentSetupObj);
            }
            return currentSetupVM;
        }



        [HttpGet("GetAllHistory/{UserId}")]
        public ActionResult<Object> GetAllHistory(long UserId)
        {

            if (UserId != 0)
            {
                List<History> history = _iHistory.FindBy(s => s.UserId == UserId).OrderByDescending(x=>x.Id).ToList();
                if (history.Count != 0)
                {
                    history[0].PayoutTable = ReturnOuput(history[0].PayoutTable);
                    history[0].ShareTable = ReturnOuput(history[0].ShareTable);
                    history[0].CurrentCapitalTable = ReturnOuput(history[0].CurrentCapitalTable);
                    history[0].OtherInputTable = ReturnOuput(history[0].OtherInputTable);
                    history[0].FinancialTable = ReturnOuput(history[0].FinancialTable);
                }
                var currentSetupObj = GetCurrentSetupDetailsByUserId(UserId);

               return Ok(new { history = history , currentSetupObj = currentSetupObj });
            }
            else
            {
                return BadRequest("User Id not Found");
            }

        }

        static string ReturnOuput(string str)
        {
            var obj = ConvterStringToObject(str);
            str = ConvterObjectToString(ConvterUnits(obj, 1));
            return str;
        }


        [HttpPost]
        [Route("AddHistorySnapshots")]
        public ActionResult<Object> AddHistorySnapshots([FromBody] SnapshotsViewSnapshots model)
        {
            
            try
            {
                Snapshots snapshots = new Snapshots
                {
                    SnapShot = model.SnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    ProjectId = model.ProjectId,
                    SnapShotType = HISTORYSNAPSHOT,
                    NPV = model.NVP,
                    CNPV = model.CNVP


                };
                _isnapShots.Add(snapshots);
                _isnapShots.Commit();
                return Ok(new { id = snapshots.Id, result = "Snapshot saved sucessfully" });

            }
            catch (Exception)
            {

                return BadRequest("Invalid Entry");
            }
        }

        [HttpGet]
        [Route("HistorySnapShots/{UserId}")]
        public ActionResult<Object> HistorySnapShots(long UserId)
        {
            try
            {
                var SnapShot = _isnapShots.FindBy(s => s.SnapShotType == HISTORYSNAPSHOT && s.UserId == UserId);
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
        [Route("ApproveHistory/{UserId}")]
        public ActionResult<Object> ApproveHistory(long UserId)
        {
            try
            {
                var history = _iHistory.GetSingle(s => s.UserId == UserId);
                if (history != null)
                {
                    history.ApprovalFlag = 1;
                    _iHistory.Update(history);
                    _iHistory.Commit();
                    return Ok("Successfully approved");
                }
                else
                {
                    return NotFound("record not foud");
                }
            }
            catch (Exception)
            {

                return BadRequest();
            }


        }

        [HttpGet]
        [Route("EditInputFlag/{UserId}")]
        public ActionResult<Object> EditIputs(long UserId)
        {
            var history = _iHistory.GetSingle(s => s.UserId == UserId);
            if (history == null)
            {
                return NotFound("No history record found for the given UserId.");
            }
            history.SummaryFlag = 0;
            _iHistory.Update(history);
            _iHistory.Commit();
            return Ok("Flag of History changed to zero");
        }


        static object[][] ConvterUnits(object[][] values, int flag)
        {
            for (int i = 0; i < values.Length; i++)
            {
                var dividend = UnitConversion.ReturnDividend(Convert.ToString(values[i][1]));
                for (int j = 2; j < values[0].Length; j++)
                {
                    if (flag == 0)
                    {
                        values[i][j] = ParseDouble(values[i][j]) * dividend;
                    }
                    else if (flag == 1)
                    {
                        values[i][j] = ParseDouble(values[i][j]) / dividend;
                    }

                }
            }

            return values;
        }



        //static double Returndividend(string strVal)
        //{
        //    double dividend = 1;
        //    if (strVal.Contains("(") && strVal.Contains(")"))
        //    {
        //        string output = strVal.Split('(', ')')[1];
        //        dividend = UnitConversion.ReturnDividend(output);
        //    }

        //    return dividend;
        //}

        [HttpGet("ExportHistory/{UserId}")]
        public ActionResult<Object> ExportHistory(long UserId)
        {
            string rootFolder = _hostingEnvironment.WebRootPath;
            string fileName = @"payout_policy.xlsx";
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            string formattedCustomObject = (string)null;
            var history = _iHistory.FindBy(s => s.UserId == UserId).ToArray();
                if (history.Length == 0)
    {
        return NotFound("No history records found for the given UserId.");
    }

            var historySummary = JsonConvert.DeserializeObject<HistorySummary>(history[0].SummaryOutput);
            var startYear = history[0].StartYear;

            using (ExcelPackage package = new ExcelPackage(file))
            {
                var wsHistory = package.Workbook.Worksheets["PayoutHistory"];
                List<List<string>> payoutTableAddr = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 7, 4, ConvterStringToObject(history[0].PayoutTable), null, out payoutTableAddr);
                List<List<string>> shareTableAddr = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 14, 4, ConvterStringToObject(history[0].ShareTable), null, out shareTableAddr);
                List<List<string>> currentCapitalTableAddr = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 20, 4, ConvterStringToObject(history[0].CurrentCapitalTable), shareTableAddr, out currentCapitalTableAddr);
                List<List<string>> currentCostOfCapTableAddr = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 25, 4, ConvterStringToObject(history[0].CurrentCostOfCapTable), null, out currentCostOfCapTableAddr);
                List<List<string>> otherInputTableAddr = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 32, 4, ConvterStringToObject(history[0].OtherInputTable), null, out otherInputTableAddr);
                List<List<string>> financialTableAddr1 = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 41, 4, ConvterStringToObject(history[0].FinancialTable).Skip(0).Take(9).ToArray(), null, out financialTableAddr1);
                List<List<string>> financialTableAddr2 = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 51, 4, ConvterStringToObject(history[0].FinancialTable).Skip(9).Take(14 - 9).ToArray(), null, out financialTableAddr2);
                List<List<string>> financialTableAddr3 = new List<List<string>>();
                wsHistory = ExecelGeneration(wsHistory, 57, 4, ConvterStringToObject(history[0].FinancialTable).Skip(14).Take(15 - 14).ToArray(), null, out financialTableAddr3);

                HistoryFormulaObj historyFormulaObj = new HistoryFormulaObj();
                historyFormulaObj.payoutTableAddr = payoutTableAddr;
                historyFormulaObj.shareTableAddr = shareTableAddr;
                historyFormulaObj.currentCapitalTableAddr = currentCapitalTableAddr;
                historyFormulaObj.otherInputTableAddr = otherInputTableAddr;
                historyFormulaObj.financialTableAddr1 = financialTableAddr1;
                historyFormulaObj.financialTableAddr2 = financialTableAddr2;
                historyFormulaObj.financialTableAddr3 = financialTableAddr3;

                wsHistory = ExecelGenerationEX(wsHistory, 64, 4, historySummary.TotalCashReturned.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 65, 4, historySummary.TotalCashReturnedFCFE.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 67, 4, historySummary.TotalPayoutQuarter.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 68, 4, historySummary.DPSQuarter.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 69, 4, historySummary.PayoutRatioQuarter.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 70, 4, historySummary.DividendYieldQuarter.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 72, 4, historySummary.TotalPayoutOne.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 73, 4, historySummary.DPSOne.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 74, 4, historySummary.DividendPayoutRatioOne.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 75, 4, historySummary.DividendYieldOne.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 77, 4, historySummary.TotalPayout.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 78, 4, historySummary.SharesRepurchased.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 81, 4, historySummary.DebtMarketEquity.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 82, 4, historySummary.DebtBookEquity.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 83, 4, historySummary.EBITInterestCoverage.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 84, 4, historySummary.EBITDAInterestCoverage.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 85, 4, historySummary.CashFlow.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 86, 4, historySummary.TotalDebt.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 87, 4, historySummary.DebtRating.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 90, 4, historySummary.CashEquivalent.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 91, 4, historySummary.CashNeededCapital.Skip(1).ToList(), historyFormulaObj);
                wsHistory = ExecelGenerationEX(wsHistory, 92, 4, historySummary.ExcessCash.Skip(1).ToList(), historyFormulaObj);

                var arraySize = historySummary.ExcessCash.Skip(1).ToList().Count();
                wsHistory = AddYears(6, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(13, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(19, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(24, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(31, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(39, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(63, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(80, 4, wsHistory, arraySize, startYear);
                wsHistory = AddYears(89, 4, wsHistory, arraySize, startYear);

                ExcelPackage excelPackage = new ExcelPackage();
                excelPackage.Workbook.Worksheets.Add("PayoutHistory", wsHistory);
                //package.Save();
                ExcelPackage epOut = excelPackage;
                byte[] myStream = epOut.GetAsByteArray();
                var inputAsString = Convert.ToBase64String(myStream);
                formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);

                return Ok(formattedCustomObject);

            }


        }


        private static ExcelWorksheet ExecelGeneration(ExcelWorksheet wsHistory, int row, int col, object[][] obj, List<List<string>> formulaAdrr, out List<List<string>> cellAddress)
        {
            List<List<string>> addLists = new List<List<string>>();
            for (int o = 0; o < obj.Length; o++)
            {
                ExcelAddress addr = null;
                List<string> addList = new List<string>();
                var cellFormat = UnitConversion.ReturnCellFormat((Convert.ToString(obj[o][1])));
                for (int p = 2; p < (obj[0].Length); p++)
                {
                    wsHistory.Cells[(row + o), (col + p - 2)].Value = obj[o][p];
                    if (formulaAdrr != null && o == 0 && p != 0)
                    {
                        wsHistory.Cells[(row + o), (col + (p - 2))].Formula = formulaAdrr[0][p - 2] + "*" + formulaAdrr[1][p - 2];
                    }
                    wsHistory.Cells[(row + o), (col + (p - 2))].Style.Numberformat.Format = cellFormat;
                    addr = new ExcelAddress((row + o), (col + (p - 2)), (row + o), (col + (p - 2)));
                    addList.Add(Convert.ToString(addr));

                }
                addLists.Add(addList);
            }
            cellAddress = addLists;
            return wsHistory;
        }


        private static ExcelWorksheet ExecelGenerationEX(ExcelWorksheet wsHistory, int row, int col, List<object> obj, HistoryFormulaObj historyFormulaObj)
        {
            for (int p = 0; p < obj.Count; p++)
            {
                wsHistory.Cells[(row), (col + p)].Value = obj[p];
                switch (row)
                {
                    case 64:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.payoutTableAddr[0][p] + "+" + historyFormulaObj.payoutTableAddr[1][p] + "+" + historyFormulaObj.payoutTableAddr[2][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 0);
                        break;
                    case 65:
                        wsHistory.Cells[(row), (col + p)].Formula = new ExcelAddress((row - 1), (col + p), (row - 1), (col + p)) + "/" + historyFormulaObj.otherInputTableAddr[1][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;


                    case 67:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.payoutTableAddr[0][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 0);
                        break;
                    case 68:
                        wsHistory.Cells[(row), (col + p)].Formula = new ExcelAddress((row - 1), (col + p), (row - 1), (col + p)) + "/" + historyFormulaObj.shareTableAddr[1][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 0);
                        break;
                    case 69:
                        wsHistory.Cells[(row), (col + p)].Formula = new ExcelAddress((row - 2), (col + p), (row - 2), (col + p)) + "/" + historyFormulaObj.financialTableAddr1[8][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;
                    case 70:
                        wsHistory.Cells[(row), (col + p)].Formula = new ExcelAddress((row - 2), (col + p), (row - 2), (col + p)) + "/" + historyFormulaObj.shareTableAddr[0][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;





                    case 72:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.payoutTableAddr[1][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 0);
                        break;
                    case 73:
                        wsHistory.Cells[(row), (col + p)].Formula = new ExcelAddress((row - 1), (col + p), (row - 1), (col + p)) + "/" + historyFormulaObj.shareTableAddr[1][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 0);
                        break;
                    case 74:
                        wsHistory.Cells[(row), (col + p)].Formula = new ExcelAddress((row - 2), (col + p), (row - 2), (col + p)) + "/" + historyFormulaObj.financialTableAddr1[8][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;
                    case 75:
                        wsHistory.Cells[(row), (col + p)].Formula = new ExcelAddress((row - 2), (col + p), (row - 2), (col + p)) + "/" + historyFormulaObj.shareTableAddr[0][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;




                    case 77:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.payoutTableAddr[2][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 0);
                        break;
                    case 78:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.payoutTableAddr[3][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 1);
                        break;





                    case 81:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.currentCapitalTableAddr[1][p] + "/" + historyFormulaObj.currentCapitalTableAddr[0][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;
                    case 82:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.financialTableAddr2[3][p] + "/" + historyFormulaObj.financialTableAddr2[4][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;
                    case 83:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.financialTableAddr1[4][p] + "/" + historyFormulaObj.financialTableAddr1[5][p];
                        break;
                    case 84:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.financialTableAddr1[1][p] + "/" + historyFormulaObj.financialTableAddr1[5][p];
                        break;
                    case 85:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.financialTableAddr3[0][p] + "/" + historyFormulaObj.financialTableAddr2[3][p];
                        wsHistory.Cells[(row), (col + p)].Style.Numberformat.Format = "0.00 % _); (0.00 %)";
                        break;
                    case 86:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.financialTableAddr2[3][p] + "/" + historyFormulaObj.financialTableAddr1[1][p];
                        break;



                    case 90:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.financialTableAddr2[0][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 1);
                        break;
                    case 91:
                        wsHistory.Cells[(row), (col + p)].Formula = historyFormulaObj.otherInputTableAddr[0][p];
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 1);
                        break;
                    case 92:
                        wsHistory.Cells[(row), (col + p)].Formula = (new ExcelAddress((row - 2), (col + p), (row - 2), (col + p))).ToString() + "-" + (new ExcelAddress((row - 1), (col + p), (row - 1), (col + p))).ToString();
                        wsHistory.Calculate();
                        wsHistory = ExcelFormatting(row, (col + p), wsHistory, 1);
                        break;
                    default:
                        Console.WriteLine("Default case");
                        break;
                }







            }

            return wsHistory;
        }


        static ExcelWorksheet ExcelFormatting(int row, int col, ExcelWorksheet wsSheet, int flag)
        {
            var output = UnitConversion.FindFomartLetter(Convert.ToDouble(wsSheet.Cells[row, col].Value));
            if (flag == 0) { wsSheet.Cells[row, col].Style.Numberformat.Format = "$" + UnitConversion.ReturnCellFormat(output); }
            else if (flag == 1) { wsSheet.Cells[row, col].Style.Numberformat.Format = UnitConversion.ReturnCellFormat(output); }
            return wsSheet;
        }

        static ExcelWorksheet AddYears(int row, int col, ExcelWorksheet wsSheet, int arraySize, int startYear)
        {
            for (int i = 0; i < arraySize; i++)
            {
                wsSheet.Cells[row, (col + i)].Value = "FY'" + (startYear + i);
            }
            return wsSheet;
        }

        private static string ConvterObjectToString(Object obj)
        {
            var str = JsonConvert.SerializeObject(obj);
            return str;
        }
        private static object[][] ConvterStringToObject(string str)
        {
            var obj = JsonConvert.DeserializeObject<Object[][]>(str);
            return obj;
        }

        private static Double ParseDouble(Object obj)
        {
            if ((obj == null) || (obj.ToString() == ""))
            {
                return 0;
            }


            return Convert.ToDouble(obj);
        }
    }
}
