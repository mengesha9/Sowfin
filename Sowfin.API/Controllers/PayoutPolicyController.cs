using AutoMapper;
using ExcelDataReader;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Sowfin.API.Lib;
using Sowfin.API.ViewModels;
using Sowfin.API.ViewModels.CapitalStructure;
using Sowfin.API.ViewModels.PayoutPolicy;
using Sowfin.Data.Abstract;
using Sowfin.Data.Common.Enum;
using Sowfin.Data.Common.Helper;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Sowfin.Model.Entities.CapitalStructure;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PayoutPolicyController : ControllerBase
    {
        const string PAYOUTPOLICY = "PayoutPolicy_SnapShot";
        const string PAYOUTSCENARIO = "PayoutScenario_SnapShot";
        private readonly IPayoutPolicy iPayoutPolicy = null;
        private readonly IWebHostEnvironment  _hostingEnvironment = null;
        private readonly ISnapshots iSnapShots = null;
        ICurrentSetupIpDatas iCurrentSetupIpDatas;
        ICurrentSetupIpValues iCurrentSetupIpValues;
        ICurrentSetupSoDatas iCurrentSetupSoDatas;
        ICurrentSetupSoValues iCurrentSetupSoValues;
        IHistory iHistory;
        ICurrentSetup iCurrentSetup;
        IMapper mapper;
        ICostOfCapital iCostOfCapital;
        ICapitalStructure iCapitalStructure;
        IIntegratedDatas iIntegratedDatas;
        IIntegrated_ExplicitValues iIntegrated_ExplicitValues;
        ITaxRates_IValuation iTaxRates_IValuation;
        ICurrentSetupSnapshot iCurrentSetupSnapshot;
        ICurrentSetupSnapshotDatas iCurrentSetupSnapshotDatas;
        ICurrentSetupSnapshotValues iCurrentSetupSnapshotValues;
        IInitialSetup_IValuation iInitialSetup_IValuation;
        IPayoutPolicy_ScenarioDatas iPayoutPolicy_ScenarioDatas;
        IPayoutPolicy_ScenarioValues iPayoutPolicy_ScenarioValues;
        IFilings iFilings;
        public PayoutPolicyController(IPayoutPolicy _iPayoutPolicy, IWebHostEnvironment  hostingEnvironment, ISnapshots _iSnapShots, ICurrentSetupIpDatas _iCurrentSetupIpDatas,
            ICurrentSetupIpValues _iCurrentSetupIpValues, ICurrentSetupSoDatas _iCurrentSetupSoDatas, ICurrentSetupSoValues _iCurrentSetupSoValues, IHistory _iHistory,
            ICurrentSetup _iCurrentSetup, IMapper mapper, ICostOfCapital _iCostOfCapital, ICapitalStructure _iCapitalStructure, IIntegratedDatas _iIntegratedDatas,
            IIntegrated_ExplicitValues _iIntegrated_ExplicitValues, ITaxRates_IValuation _iTaxRates_IValuation, ICurrentSetupSnapshot _iCurrentSetupSnapshot,
            ICurrentSetupSnapshotDatas _iCurrentSetupSnapshotDatas, ICurrentSetupSnapshotValues _iCurrentSetupSnapshotValues, IInitialSetup_IValuation _iInitialSetup_IValuation, IPayoutPolicy_ScenarioDatas _iPayoutPolicy_ScenarioDatas, IPayoutPolicy_ScenarioValues _iPayoutPolicy_ScenarioValues, IFilings _iFilings)
        {
            iFilings = _iFilings;
            iPayoutPolicy = _iPayoutPolicy;
            _hostingEnvironment = hostingEnvironment;
            iSnapShots = _iSnapShots;
            iCurrentSetupIpDatas = _iCurrentSetupIpDatas;
            iCurrentSetupIpValues = _iCurrentSetupIpValues;
            iCurrentSetupSoDatas = _iCurrentSetupSoDatas;
            iCurrentSetupSoValues = _iCurrentSetupSoValues;
            iHistory = _iHistory;
            iCurrentSetup = _iCurrentSetup;
            this.mapper = mapper;
            iCostOfCapital = _iCostOfCapital;
            iCapitalStructure = _iCapitalStructure;
            iIntegratedDatas = _iIntegratedDatas;
            iIntegrated_ExplicitValues = _iIntegrated_ExplicitValues;
            iTaxRates_IValuation = _iTaxRates_IValuation;
            iCurrentSetupSnapshot = _iCurrentSetupSnapshot;
            iCurrentSetupSnapshotDatas = _iCurrentSetupSnapshotDatas;
            iCurrentSetupSnapshotValues = _iCurrentSetupSnapshotValues;
            iInitialSetup_IValuation = _iInitialSetup_IValuation;
            iPayoutPolicy_ScenarioDatas = _iPayoutPolicy_ScenarioDatas;
            iPayoutPolicy_ScenarioValues = _iPayoutPolicy_ScenarioValues;
        }



        #region Prev code

        [HttpPost("Evaluate")]
        public ActionResult<Object> Evaluate([FromBody] PayoutPolicyViewModel data)
        {
            double[] values ={  data.currentSharePrice,data.NumberSharesBasic,data.NumberSharesDiluted,
                data.PreferredSharePrice,data.NoPreferredShares,data.PreferredDividend,
                data.MarketValueDebt,data.CashEquivalent,data.CashNeededWorkingCap};

            string[] units = { data.equityUnits.CurrentSharePriceUnit, data.equityUnits.NumberShareBasicUnit,data.equityUnits.NumberShareOutstandingUnit,
            data.prefferedEquityUnits.PrefferedSharePriceUnit,data.prefferedEquityUnits.PrefferedShareOutstandingUnit,data.prefferedEquityUnits.PrefferedDividendUnit,
            data.MarketValueDebtUnit,data.balanceSheetUnits.CashEquivalentUnit,data.CashCapitalUnit};

            var output = UnitConversion.ConvertUnits(values, units, 0);

            Equity equity = new Equity { CurrentSharePrice = output[0], NumberShareBasic = output[1], NumberShareOutstanding = output[2], CostOfEquity = data.CostOfEquity };
            PrefferedEquity prefferedEquity = new PrefferedEquity { PrefferedSharePrice = output[3], PrefferedShareOutstanding = output[4], PrefferedDividend = output[5], CostPreffEquity = data.CostOfPrefEquity };
            Debt debt = new Debt { MarketValueDebt = output[6], CostOfDebt = data.CostOfDebt };
            CapitalStructureViewModel model = new CapitalStructureViewModel
            {
                NewLeveragePolicy = "Constant Debt-to-Equity Ratio  (Target Leverage Ratio)",
                CashEquivalent = output[7],
                CashNeededCapital = output[8],
                InterestCoverage = data.InterestCoverRatio,
                MarginalTaxRate = data.MarginalTax,
                FreeCashFlow = data.FreeCashFlow,
                equity = equity,
                prefferedEquity = prefferedEquity,
                debt = debt
            };

            double[] payout = { data.TotalOngoingPayout,data.OneTimePayout,data.StockBuyBack,
            data.NetSales,data.Ebitda,data.DepreciationAmortization,data.InterestIncome,data.Ebit,data.InterestExpense,data.Ebt,
            data.Taxes,data.NetEarings,data.TotalCurrentAssets,data.TotalAssests,data.TotalDebt,data.ShareholderEquity,
            data.CashFlowOperation};
            var incomeUnit = data.incomeStatementUnits;
            var balanceUnit = data.balanceSheetUnits;
            string[] payoutUnits = { data.payoutPolicyUnits.TotalPayoutUnit,data.payoutPolicyUnits.OneTimeDividendUnit,data.payoutPolicyUnits.StockAmountUnit,
             incomeUnit.NetEarningsUnit,incomeUnit.EbitaUnit,incomeUnit.DeprecAmortizationUnit,incomeUnit.IncomeIncomeUnit,
             incomeUnit.EbitUnit,incomeUnit.InterestExpenseUnit,incomeUnit.EbtUnit,incomeUnit.TaxesUnit,incomeUnit.NetEarningsUnit,
             balanceUnit.TotalCurrentAssetUnit,balanceUnit.TotalAssetUnit,balanceUnit.TotalDebtUnit,
             balanceUnit.ShareEquitUnit,data.CashFlowUnit};

            var payoutOuput = UnitConversion.ConvertUnits(payout, payoutUnits, 0);


            data.currentSharePrice = output[0];
            data.NumberSharesBasic = output[1];
            data.NumberSharesDiluted = output[2];
            data.CostOfEquity = data.CostOfEquity;
            data.PreferredSharePrice = output[3];
            data.NoPreferredShares = output[4];
            data.PreferredDividend = output[5];
            data.MarketValueDebt = output[6];
            data.CashNeededWorkingCap = output[8];
            data.TotalOngoingPayout = payoutOuput[0];
            data.OneTimePayout = payoutOuput[1];
            data.StockBuyBack = payoutOuput[2];
            data.NetSales = payoutOuput[3];
            data.Ebitda = payoutOuput[4];
            data.DepreciationAmortization = payoutOuput[5];
            data.InterestIncome = payoutOuput[6];
            data.Ebit = payoutOuput[7];
            data.InterestExpense = payoutOuput[8];
            data.Ebt = payoutOuput[9];
            data.Taxes = payoutOuput[10];
            data.NetEarings = payoutOuput[11];
            data.CashEquivalent = output[7];
            data.TotalCurrentAssets = payoutOuput[12];
            data.TotalAssests = payoutOuput[13];
            data.TotalDebt = payoutOuput[14];
            data.ShareholderEquity = payoutOuput[15];
            data.CashFlowOperation = payoutOuput[16];


            var cap = JsonConvert.SerializeObject(MathCapStructure.SummaryOutput(model));
            var outputs = JsonConvert.DeserializeObject<SummaryOutput>(cap);
            var result = MathPayoutPolicy.SummaryOutput(data, outputs);
            data.SummaryOutput = JsonConvert.SerializeObject(result);
            try
            {
                if (data.Id == 0)
                {
                    PayoutPolicy payoutPolicy = new PayoutPolicy
                    {

                        currentSharePrice = data.currentSharePrice,
                        NumberSharesBasic = data.NumberSharesBasic,
                        NumberSharesDiluted = data.NumberSharesDiluted,
                        CostOfEquity = data.CostOfEquity,
                        PreferredSharePrice = data.PreferredSharePrice,
                        NoPreferredShares = data.NoPreferredShares,
                        PreferredDividend = data.PreferredDividend,
                        CostOfPrefEquity = data.CostOfPrefEquity,
                        MarketValueDebt = data.MarketValueDebt,
                        CostOfDebt = data.CostOfDebt,
                        CashNeededWorkingCap = data.CashNeededWorkingCap,
                        InterestCoverRatio = data.InterestCoverRatio,
                        InterestIncomeRate = data.InterestIncomeRate,
                        MarginalTax = data.MarginalTax,
                        FreeCashFlow = data.FreeCashFlow,
                        SummaryOutput = data.SummaryOutput,
                        TotalOngoingPayout = data.TotalOngoingPayout,
                        OneTimePayout = data.OneTimePayout,
                        StockBuyBack = data.StockBuyBack,
                        NetSales = data.NetSales,
                        Ebitda = data.Ebitda,
                        DepreciationAmortization = data.DepreciationAmortization,
                        InterestIncome = data.InterestIncome,
                        Ebit = data.Ebit,
                        InterestExpense = data.InterestExpense,
                        Ebt = data.Ebt,
                        Taxes = data.Taxes,
                        NetEarings = data.NetEarings,
                        CashEquivalent = data.CashEquivalent,
                        TotalCurrentAssets = data.TotalCurrentAssets,
                        TotalAssests = data.TotalAssests,
                        TotalDebt = data.TotalDebt,
                        ShareholderEquity = data.ShareholderEquity,
                        CashFlowOperation = data.CashFlowOperation,
                        SummaryFlag = data.SummaryFlag,
                        ApprovalFlag = data.ApprovalFlag,
                        ScenarioObject = data.ScenarioObject,
                        ScenarioSummary = data.ScenarioSummary,
                        ScenarioFlag = data.ScenarioFlag,
                        UserId = data.UserId,
                        MarketValueDebtUnit = data.MarketValueDebtUnit,
                        CashCapitalUnit = data.CashCapitalUnit,
                        CashFlowUnit = data.CashFlowUnit,
                        equityUnits = data.equityUnits,
                        prefferedEquityUnits = data.prefferedEquityUnits,
                        payoutPolicyUnits = data.payoutPolicyUnits,
                        incomeStatementUnits = data.incomeStatementUnits,
                        balanceSheetUnits = data.balanceSheetUnits

                    };
                    iPayoutPolicy.Add(payoutPolicy);
                    iPayoutPolicy.Commit();
                    return Ok(new Dictionary<string, object>
                     {{"result",result }});


                }
                else
                {
                    PayoutPolicy updatePayoutPolicy = new PayoutPolicy
                    {
                        Id = data.Id,
                        currentSharePrice = data.currentSharePrice,
                        NumberSharesBasic = data.NumberSharesBasic,
                        NumberSharesDiluted = data.NumberSharesDiluted,
                        CostOfEquity = data.CostOfEquity,
                        PreferredSharePrice = data.PreferredSharePrice,
                        NoPreferredShares = data.NoPreferredShares,
                        PreferredDividend = data.PreferredDividend,
                        CostOfPrefEquity = data.CostOfPrefEquity,
                        MarketValueDebt = data.MarketValueDebt,
                        CostOfDebt = data.CostOfDebt,
                        CashNeededWorkingCap = data.CashNeededWorkingCap,
                        InterestCoverRatio = data.InterestCoverRatio,
                        InterestIncomeRate = data.InterestIncomeRate,
                        MarginalTax = data.MarginalTax,
                        FreeCashFlow = data.FreeCashFlow,
                        SummaryOutput = data.SummaryOutput,
                        TotalOngoingPayout = data.TotalOngoingPayout,
                        OneTimePayout = data.OneTimePayout,
                        StockBuyBack = data.StockBuyBack,
                        NetSales = data.NetSales,
                        Ebitda = data.Ebitda,
                        DepreciationAmortization = data.DepreciationAmortization,
                        InterestIncome = data.InterestIncome,
                        Ebit = data.Ebit,
                        InterestExpense = data.InterestExpense,
                        Ebt = data.Ebt,
                        Taxes = data.Taxes,
                        NetEarings = data.NetEarings,
                        CashEquivalent = data.CashEquivalent,
                        TotalCurrentAssets = data.TotalCurrentAssets,
                        TotalAssests = data.TotalAssests,
                        TotalDebt = data.TotalDebt,
                        ShareholderEquity = data.ShareholderEquity,
                        CashFlowOperation = data.CashFlowOperation,
                        SummaryFlag = data.SummaryFlag,
                        ApprovalFlag = data.ApprovalFlag,
                        ScenarioObject = data.ScenarioObject,
                        ScenarioSummary = data.ScenarioSummary,
                        ScenarioFlag = data.ScenarioFlag,
                        UserId = data.UserId,
                        MarketValueDebtUnit = data.MarketValueDebtUnit,
                        CashCapitalUnit = data.CashCapitalUnit,
                        CashFlowUnit = data.CashFlowUnit,
                        equityUnits = data.equityUnits,
                        prefferedEquityUnits = data.prefferedEquityUnits,
                        payoutPolicyUnits = data.payoutPolicyUnits,
                        incomeStatementUnits = data.incomeStatementUnits,
                        balanceSheetUnits = data.balanceSheetUnits
                    };
                    iPayoutPolicy.Update(updatePayoutPolicy);
                    iPayoutPolicy.Commit();
                    return Ok(new Dictionary<string, object>
                     {{"result",result }});

                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex);
            }


        }

        [HttpPost("Evaluate/Scenario")]
        public ActionResult<Object> EvaluateScenario([FromBody] PayoutScenario payoutScenario)
        {

            var payoutPolicy = iPayoutPolicy.GetSingle(s => s.UserId == payoutScenario.UserId);
            double[] values =
            {
                payoutScenario.TotalPayoutAnnual,
                payoutScenario.OneTimePayout,
                payoutScenario.StockBuyBack
            };
            string[] units = {
                payoutScenario.TotalPayoutAnnualUnit,
                payoutScenario.OneTimePayoutUnit,
                payoutScenario.StockBuyBackUnit
            };
            var output = UnitConversion.ConvertUnits(values, units, 0);
            payoutScenario.TotalPayoutAnnual = output[0];
            payoutScenario.OneTimePayout = output[1];
            payoutScenario.StockBuyBack = output[2];
            var scenarioSummary = MathPayoutScenario.SummaryOutput(payoutScenario, payoutPolicy);
            payoutPolicy.ScenarioObject = JsonConvert.SerializeObject(payoutScenario);
            payoutPolicy.ScenarioSummary = JsonConvert.SerializeObject(scenarioSummary);
            iPayoutPolicy.Update(payoutPolicy);
            iPayoutPolicy.Commit();
            return Ok(new Dictionary<string, object>
                     {

                       {"result",scenarioSummary }

                     });

        }

        [HttpGet("GetAllPayoutPolicy/{UserId}")]
        public ActionResult<Object> GetAllPayoutPolicy(long UserId)
        {

            var payoutPolicy = iPayoutPolicy.FindBy(s => s.UserId == UserId).OrderByDescending(x => x.Id).ToArray();

            if (payoutPolicy.Length != 0)
            {
                double[] values ={  payoutPolicy[0].currentSharePrice,payoutPolicy[0].NumberSharesBasic,payoutPolicy[0].NumberSharesDiluted,
                payoutPolicy[0].PreferredSharePrice,payoutPolicy[0].NoPreferredShares,payoutPolicy[0].PreferredDividend,
                payoutPolicy[0].MarketValueDebt,payoutPolicy[0].CashEquivalent,payoutPolicy[0].CashNeededWorkingCap};

                string[] units = { payoutPolicy[0].equityUnits.CurrentSharePriceUnit, payoutPolicy[0].equityUnits.NumberShareBasicUnit,payoutPolicy[0].equityUnits.NumberShareOutstandingUnit,
            payoutPolicy[0].prefferedEquityUnits.PrefferedSharePriceUnit,payoutPolicy[0].prefferedEquityUnits.PrefferedShareOutstandingUnit,payoutPolicy[0].prefferedEquityUnits.PrefferedDividendUnit,
            payoutPolicy[0].MarketValueDebtUnit,payoutPolicy[0].balanceSheetUnits.CashEquivalentUnit,payoutPolicy[0].CashCapitalUnit};

                var output = UnitConversion.ConvertUnits(values, units, 1);

                double[] payout = { payoutPolicy[0].TotalOngoingPayout,payoutPolicy[0].OneTimePayout,payoutPolicy[0].StockBuyBack,
            payoutPolicy[0].NetSales,payoutPolicy[0].Ebitda,payoutPolicy[0].DepreciationAmortization,payoutPolicy[0].InterestIncome,payoutPolicy[0].Ebit,payoutPolicy[0].InterestExpense,payoutPolicy[0].Ebt,
            payoutPolicy[0].Taxes,payoutPolicy[0].NetEarings,payoutPolicy[0].TotalCurrentAssets,payoutPolicy[0].TotalAssests,payoutPolicy[0].TotalDebt,payoutPolicy[0].ShareholderEquity,
            payoutPolicy[0].CashFlowOperation};
                var incomeUnit = payoutPolicy[0].incomeStatementUnits;
                var balanceUnit = payoutPolicy[0].balanceSheetUnits;
                string[] payoutUnits = { payoutPolicy[0].payoutPolicyUnits.TotalPayoutUnit,payoutPolicy[0].payoutPolicyUnits.OneTimeDividendUnit,payoutPolicy[0].payoutPolicyUnits.StockAmountUnit,
             incomeUnit.NetEarningsUnit,incomeUnit.EbitaUnit,incomeUnit.DeprecAmortizationUnit,incomeUnit.IncomeIncomeUnit,
             incomeUnit.EbitUnit,incomeUnit.InterestExpenseUnit,incomeUnit.EbtUnit,incomeUnit.TaxesUnit,incomeUnit.NetEarningsUnit,
             balanceUnit.TotalCurrentAssetUnit,balanceUnit.TotalAssetUnit,balanceUnit.TotalDebtUnit,
             balanceUnit.ShareEquitUnit,payoutPolicy[0].CashFlowUnit};

                var payoutOuput = UnitConversion.ConvertUnits(payout, payoutUnits, 1);

                payoutPolicy[0].currentSharePrice = output[0];
                payoutPolicy[0].NumberSharesBasic = output[1];
                payoutPolicy[0].NumberSharesDiluted = output[2];
                payoutPolicy[0].CostOfEquity = payoutPolicy[0].CostOfEquity;
                payoutPolicy[0].PreferredSharePrice = output[3];
                payoutPolicy[0].NoPreferredShares = output[4];
                payoutPolicy[0].PreferredDividend = output[5];
                payoutPolicy[0].MarketValueDebt = output[6];
                payoutPolicy[0].CashNeededWorkingCap = output[8];
                payoutPolicy[0].TotalOngoingPayout = payoutOuput[0];
                payoutPolicy[0].OneTimePayout = payoutOuput[1];
                payoutPolicy[0].StockBuyBack = payoutOuput[2];
                payoutPolicy[0].NetSales = payoutOuput[3];
                payoutPolicy[0].Ebitda = payoutOuput[4];
                payoutPolicy[0].DepreciationAmortization = payoutOuput[5];
                payoutPolicy[0].InterestIncome = payoutOuput[6];
                payoutPolicy[0].Ebit = payoutOuput[7];
                payoutPolicy[0].InterestExpense = payoutOuput[8];
                payoutPolicy[0].Ebt = payoutOuput[9];
                payoutPolicy[0].Taxes = payoutOuput[10];
                payoutPolicy[0].NetEarings = payoutOuput[11];
                payoutPolicy[0].CashEquivalent = output[7];
                payoutPolicy[0].TotalCurrentAssets = payoutOuput[12];
                payoutPolicy[0].TotalAssests = payoutOuput[13];
                payoutPolicy[0].TotalDebt = payoutOuput[14];
                payoutPolicy[0].ShareholderEquity = payoutOuput[15];
                payoutPolicy[0].CashFlowOperation = payoutOuput[16];
            }

            return Ok(new Dictionary<string, object>
                     {

                       {"result",payoutPolicy }

                     });

        }

        [HttpGet("ExportPayoutPolicy/{UserId}/{Flag}")]
        public ActionResult<Object> ExportPayoutPolicy(long UserId, int Flag)
        {
            if (UserId != 0)
            {
                string rootFolder = _hostingEnvironment.WebRootPath;
                string fileName = @"payout_policy.xlsx";
                FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
                var formattedCustomObject = (String)null;
                var payoutPolicy = iPayoutPolicy.FindBy(s => s.UserId == UserId).ToArray();
                if(payoutPolicy == null || payoutPolicy.Length == 0){
                    return NotFound("no data found");
                }
                var payout = payoutPolicy[0];
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet wsPayoutPolicy = null;
                    if (Flag == 1) { wsPayoutPolicy = package.Workbook.Worksheets["PayoutPolicy"]; }
                    else if (Flag == 2) { wsPayoutPolicy = package.Workbook.Worksheets["ScenarioAnalysis"]; }

                    wsPayoutPolicy = CellFormat("D10", wsPayoutPolicy, payout.currentSharePrice, payout.equityUnits.CurrentSharePriceUnit);
                    wsPayoutPolicy = CellFormat("D11", wsPayoutPolicy, payout.NumberSharesBasic, payout.equityUnits.NumberShareBasicUnit);
                    wsPayoutPolicy = CellFormat("D12", wsPayoutPolicy, payout.NumberSharesDiluted, payout.equityUnits.NumberShareOutstandingUnit);
                    wsPayoutPolicy = CellFormat("D13", wsPayoutPolicy, payout.CostOfEquity / 100, null);
                    wsPayoutPolicy = CellFormat("D15", wsPayoutPolicy, payout.PreferredSharePrice, payout.prefferedEquityUnits.PrefferedSharePriceUnit);
                    wsPayoutPolicy = CellFormat("D16", wsPayoutPolicy, payout.NoPreferredShares, payout.prefferedEquityUnits.PrefferedShareOutstandingUnit);
                    wsPayoutPolicy = CellFormat("D17", wsPayoutPolicy, payout.PreferredDividend, payout.prefferedEquityUnits.PrefferedDividendUnit);
                    wsPayoutPolicy = CellFormat("D18", wsPayoutPolicy, payout.CostOfPrefEquity / 100, null);

                    wsPayoutPolicy = CellFormat("D20", wsPayoutPolicy, payout.MarketValueDebt, payout.MarketValueDebtUnit);
                    wsPayoutPolicy = CellFormat("D21", wsPayoutPolicy, payout.CostOfDebt / 100, null);


                    wsPayoutPolicy = CellFormat("D24", wsPayoutPolicy, payout.CashNeededWorkingCap, payout.CashCapitalUnit);
                    wsPayoutPolicy = CellFormat("D25", wsPayoutPolicy, payout.InterestCoverRatio, null);
                    wsPayoutPolicy = CellFormat("D26", wsPayoutPolicy, payout.InterestIncomeRate / 100, null);
                    wsPayoutPolicy = CellFormat("D27", wsPayoutPolicy, payout.MarginalTax / 100, null);

                    wsPayoutPolicy = ReturnPrePayout(wsPayoutPolicy, "D", payout);


                    SummaryOutput summaryOutput = JsonConvert.DeserializeObject<SummaryOutput>(payout.SummaryOutput);

                    string[] units = new string[4];
                    var output = UnitConversion.ConvertOutputUnits(out units, summaryOutput.marketValueEquity, summaryOutput.marketValuePreferredEquity,
                        summaryOutput.marketValueDebt, summaryOutput.marketValueNetDebt);

                    wsPayoutPolicy = CellFormat("D57", wsPayoutPolicy, output[0], units[0]);
                    wsPayoutPolicy = CellFormat("D58", wsPayoutPolicy, output[1], units[1]);
                    wsPayoutPolicy = CellFormat("D59", wsPayoutPolicy, output[2], units[2]);
                    wsPayoutPolicy = CellFormat("D60", wsPayoutPolicy, output[3], units[3]);

                    wsPayoutPolicy = CellFormat("D64", wsPayoutPolicy, summaryOutput.costEquity / 100, null);
                    wsPayoutPolicy = CellFormat("D65", wsPayoutPolicy, summaryOutput.costPreferredEquity / 100, null);
                    wsPayoutPolicy = CellFormat("D66", wsPayoutPolicy, summaryOutput.costDebt / 100, null);
                    wsPayoutPolicy = CellFormat("D67", wsPayoutPolicy, summaryOutput.unleveredCostCapital / 100, null);
                    wsPayoutPolicy = CellFormat("D68", wsPayoutPolicy, summaryOutput.weighedAverageCapital / 100, null);

                    wsPayoutPolicy = CellFormat("D71", wsPayoutPolicy, "D59/D57", null);
                    wsPayoutPolicy = CellFormat("D72", wsPayoutPolicy, "D59/(D57+D58+D59)", null);


                    wsPayoutPolicy = CellFormat("D75", wsPayoutPolicy, "D46-D87-D92-D97", "U");
                    wsPayoutPolicy = CellFormat("D76", wsPayoutPolicy, "D75-D24", "U");


                    string[] PayoutAnalysis = { "D30", "D87", "D88", "D31", "D92", "D93", "D32", "D97", "D98", "D101", "D102" };
                    wsPayoutPolicy = ReturnNewPayoutAnalysis(wsPayoutPolicy, "D", PayoutAnalysis);


                    wsPayoutPolicy = ReturnPostPayout(wsPayoutPolicy, "D", new string[] { "D21" });

                    wsPayoutPolicy = ReturnDebtRatios(wsPayoutPolicy, "D");

                    if (Flag == 2)
                    {
                        PayoutScenario payoutScenario = JsonConvert.DeserializeObject<PayoutScenario>(payout.ScenarioObject);
                        wsPayoutPolicy = CellFormat("K25", wsPayoutPolicy, payoutScenario.TargetDebtToEquity / 100, null);
                        wsPayoutPolicy = CellFormat("K27", wsPayoutPolicy, payoutScenario.CostOfDebt / 100, null);
                        wsPayoutPolicy = CellFormat("K30", wsPayoutPolicy, payoutScenario.TotalPayoutAnnual, payoutScenario.TotalPayoutAnnualUnit);
                        wsPayoutPolicy = CellFormat("K31", wsPayoutPolicy, payoutScenario.OneTimePayout, payoutScenario.OneTimePayoutUnit);
                        wsPayoutPolicy = CellFormat("K32", wsPayoutPolicy, payoutScenario.StockBuyBack, payoutScenario.StockBuyBackUnit);
                        package.Save();
                        wsPayoutPolicy = ReturnPrePayout(wsPayoutPolicy, "K", payout);

                        wsPayoutPolicy = CellFormat("K57", wsPayoutPolicy, "D57", "U");
                        wsPayoutPolicy = CellFormat("K58", wsPayoutPolicy, "D58", "U");
                        wsPayoutPolicy = CellFormat("K59", wsPayoutPolicy, "K25*K57", "U");
                        wsPayoutPolicy = CellFormat("K60", wsPayoutPolicy, "K59-D59", "U");
                        wsPayoutPolicy = CellFormat("K61", wsPayoutPolicy, "K59-K76", "U");

                        wsPayoutPolicy = CellFormat("K64", wsPayoutPolicy, "D67 + ((K61 / K57) * (D67 - K66)) + ((K58 / K57) * (D67 - K65))", null);
                        wsPayoutPolicy = CellFormat("K65", wsPayoutPolicy, "D65", null);
                        wsPayoutPolicy = CellFormat("K66", wsPayoutPolicy, "K27", null);
                        wsPayoutPolicy = CellFormat("K67", wsPayoutPolicy, "D67", null);
                        wsPayoutPolicy = CellFormat("K68", wsPayoutPolicy, "(K64 * (K57 / (K57 + K58 + K61))) + K65 * (K58 / (K57 + K58 + K61)) + K66 * (K61 / (K57 + K58 + K61)) * (1 -D27)", null);
                        //= G67 * (G60 / (G60 + G61 + G64)) + G68 * (G61 / (G60 + G61 + G64)) + G69 * (G64 / (G60 + G61 + G64)) * (1 -$C$30)
                        wsPayoutPolicy = CellFormat("K71", wsPayoutPolicy, "K59/ K57", null);
                        wsPayoutPolicy = CellFormat("K72", wsPayoutPolicy, "K59 /(K57+K58+K59)", null);


                        wsPayoutPolicy = CellFormat("K75", wsPayoutPolicy, "D46 + K60 - K87 - K92 - K97", "U");
                        wsPayoutPolicy = CellFormat("K76", wsPayoutPolicy, "K75 -D24", "U");

                        string[] PayoutAnalysisScenario = { "K30", "K87", "K88", "K31", "K92", "K93", "K32", "K97", "K98", "K101", "K102" };
                        wsPayoutPolicy = ReturnNewPayoutAnalysis(wsPayoutPolicy, "K", PayoutAnalysisScenario);

                        wsPayoutPolicy = ReturnPostPayout(wsPayoutPolicy, "K", new string[] { "K27" });

                        wsPayoutPolicy = ReturnDebtRatios(wsPayoutPolicy, "K");

                    }



                    ExcelPackage excelPackage = new ExcelPackage();
                    excelPackage.Workbook.Worksheets.Add("PayoutPolicy", wsPayoutPolicy);
                    //package.Save();
                    ExcelPackage epOut = excelPackage;
                    byte[] myStream = epOut.GetAsByteArray();
                    var inputAsString = Convert.ToBase64String(myStream);
                    formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);
                }
                return Ok(formattedCustomObject);

            }
            else
            {
                return BadRequest("Id Not Found ");
            }
        }

        private static ExcelWorksheet CellFormat(string cell, ExcelWorksheet wsSheet, object value, string unit)
        {
            if (Convert.ToString(value).Contains("D") || Convert.ToString(value).Contains("K"))
            {
                wsSheet.Cells[cell].Formula = Convert.ToString(value);
                wsSheet.Calculate();
                var output = UnitConversion.FindFomartLetter(Convert.ToDouble(wsSheet.Cells[cell].Value));
                if (unit == "U")
                {
                    if (cell == "D98" || cell == "D101" || cell == "D102" || cell == "K98" || cell == "K101" || cell == "K102") { wsSheet.Cells[cell].Style.Numberformat.Format = UnitConversion.ReturnCellFormat(output); }
                    else { wsSheet.Cells[cell].Style.Numberformat.Format = "$" + UnitConversion.ReturnCellFormat(output); }

                }
                else if (unit == "UX") { wsSheet.Cells[cell].Style.Numberformat.Format = "0.00"; }
                else { wsSheet.Cells[cell].Style.Numberformat.Format = "0.00 % _); (0.00 %)"; }

            }
            else
            {
                wsSheet.Cells[cell].Value = value;
                if (unit == null) { wsSheet.Cells[cell].Style.Numberformat.Format = "0.00 % _); (0.00 %)"; }
                else
                {
                    var format = UnitConversion.ReturnCellFormat(unit);
                    wsSheet.Cells[cell].Style.Numberformat.Format = format;
                }
            }

            return wsSheet;
        }

        private static ExcelWorksheet ReturnNewPayoutAnalysis(ExcelWorksheet wsSheet, string cellPrefix, string[] output)
        {
            wsSheet = CellFormat(cellPrefix + "87", wsSheet, output[0], "U");
            wsSheet = CellFormat(cellPrefix + "88", wsSheet, output[1] + "/D11", "U");
            wsSheet = CellFormat(cellPrefix + "89", wsSheet, output[1] + "/D44", null);
            wsSheet = CellFormat(cellPrefix + "90", wsSheet, output[2] + "/D10", null);

            wsSheet = CellFormat(cellPrefix + "92", wsSheet, output[3], "U");
            wsSheet = CellFormat(cellPrefix + "93", wsSheet, output[4] + "/D11", "U");
            wsSheet = CellFormat(cellPrefix + "94", wsSheet, output[4] + "/D44", null);
            wsSheet = CellFormat(cellPrefix + "95", wsSheet, output[5] + "/D10", null);

            wsSheet = CellFormat(cellPrefix + "97", wsSheet, output[6], "U");
            wsSheet = CellFormat(cellPrefix + "98", wsSheet, output[7] + "/D10", "U");

            wsSheet = CellFormat(cellPrefix + "101", wsSheet, "D11-" + output[8], "U");
            wsSheet = CellFormat(cellPrefix + "102", wsSheet, "D12-" + output[8], "U");
            wsSheet = CellFormat(cellPrefix + "103", wsSheet, cellPrefix + "116/" + output[9], "U");
            wsSheet = CellFormat(cellPrefix + "104", wsSheet, cellPrefix + "116/" + output[10], "U");

            return wsSheet;

        }

        private static ExcelWorksheet ReturnPostPayout(ExcelWorksheet wsSheet, string cellPrefix, string[] output)
        {
            wsSheet = CellFormat(cellPrefix + "108", wsSheet, cellPrefix + "36", "U");
            wsSheet = CellFormat(cellPrefix + "109", wsSheet, cellPrefix + "37", "U");
            wsSheet = CellFormat(cellPrefix + "110", wsSheet, cellPrefix + "38", "U");

            wsSheet = CellFormat(cellPrefix + "111", wsSheet, cellPrefix + "118*D26", "U");
            wsSheet = CellFormat(cellPrefix + "112", wsSheet, cellPrefix + "109-" + cellPrefix + "110+" + cellPrefix + "111", "U");
            wsSheet = CellFormat(cellPrefix + "113", wsSheet, cellPrefix + "121*" + output[0], "U");
            if (cellPrefix == "D")
            {
                wsSheet = CellFormat(cellPrefix + "114", wsSheet, cellPrefix + "112-" + cellPrefix + "113", "U");
            }
            else if (cellPrefix == "K")
            {
                wsSheet = CellFormat(cellPrefix + "114", wsSheet, cellPrefix + "112+" + cellPrefix + "113", "U");
            }


            wsSheet = CellFormat(cellPrefix + "115", wsSheet, cellPrefix + "114*D27", "U");
            wsSheet = CellFormat(cellPrefix + "116", wsSheet, cellPrefix + "114-" + cellPrefix + "115", "U");

            wsSheet = CellFormat(cellPrefix + "118", wsSheet, cellPrefix + "75", "U");
            wsSheet = CellFormat(cellPrefix + "119", wsSheet, cellPrefix + "47+" + "(" + cellPrefix + "118-" + cellPrefix + "46)", "U");

            wsSheet = CellFormat(cellPrefix + "120", wsSheet, cellPrefix + "48+" + "(" + cellPrefix + "118-" + cellPrefix + "46)", "U");
            if (cellPrefix == "D")
            {
                wsSheet = CellFormat(cellPrefix + "121", wsSheet, cellPrefix + "49", "U");
            }
            else if (cellPrefix == "K")
            {
                wsSheet = CellFormat(cellPrefix + "121", wsSheet, cellPrefix + "59", "U");
            }

            wsSheet = CellFormat(cellPrefix + "122", wsSheet, cellPrefix + "50+" + "(" + cellPrefix + "118-" + cellPrefix + "46)", "U");

            wsSheet = CellFormat(cellPrefix + "124", wsSheet, cellPrefix + "52+" + "(" + cellPrefix + "116-" + cellPrefix + "44)", "U");


            return wsSheet;

        }

        private static ExcelWorksheet ReturnDebtRatios(ExcelWorksheet wSheet, string cellPrefix)
        {
            wSheet = CellFormat(cellPrefix + "127", wSheet, cellPrefix + "59/" + cellPrefix + "57", null);
            wSheet = CellFormat(cellPrefix + "128", wSheet, cellPrefix + "121/" + cellPrefix + "122", null);
            wSheet = CellFormat(cellPrefix + "129", wSheet, cellPrefix + "112/" + cellPrefix + "113", "UX");
            wSheet = CellFormat(cellPrefix + "130", wSheet, cellPrefix + "109/" + cellPrefix + "113", "UX");
            wSheet = CellFormat(cellPrefix + "131", wSheet, cellPrefix + "124/" + cellPrefix + "121", null);
            wSheet = CellFormat(cellPrefix + "132", wSheet, cellPrefix + "121/" + cellPrefix + "109", "UX");

            return wSheet;
        }

        private static ExcelWorksheet ReturnPrePayout(ExcelWorksheet wSheet, string cellPrefix, PayoutPolicy payout)
        {
            if (cellPrefix != "K")
            {
                wSheet = CellFormat(cellPrefix + "30", wSheet, payout.TotalOngoingPayout, payout.payoutPolicyUnits.TotalPayoutUnit);
                wSheet = CellFormat(cellPrefix + "31", wSheet, payout.OneTimePayout, payout.payoutPolicyUnits.OneTimeDividendUnit);
                wSheet = CellFormat(cellPrefix + "32", wSheet, payout.StockBuyBack, payout.payoutPolicyUnits.StockAmountUnit);
            }
            wSheet = CellFormat(cellPrefix + "36", wSheet, payout.NetSales, payout.incomeStatementUnits.NetSalesUnit);
            wSheet = CellFormat(cellPrefix + "37", wSheet, payout.Ebitda, payout.incomeStatementUnits.EbitaUnit);
            wSheet = CellFormat(cellPrefix + "38", wSheet, payout.DepreciationAmortization, payout.incomeStatementUnits.DeprecAmortizationUnit);
            wSheet = CellFormat(cellPrefix + "39", wSheet, payout.InterestIncome, payout.incomeStatementUnits.IncomeIncomeUnit);
            wSheet = CellFormat(cellPrefix + "40", wSheet, payout.Ebit, payout.incomeStatementUnits.EbitUnit);
            wSheet = CellFormat(cellPrefix + "41", wSheet, payout.InterestExpense, payout.incomeStatementUnits.InterestExpenseUnit);
            wSheet = CellFormat(cellPrefix + "42", wSheet, payout.Ebt, payout.incomeStatementUnits.EbtUnit);
            wSheet = CellFormat(cellPrefix + "43", wSheet, payout.Taxes, payout.incomeStatementUnits.TaxesUnit);
            wSheet = CellFormat(cellPrefix + "44", wSheet, payout.NetEarings, payout.incomeStatementUnits.NetEarningsUnit);
            wSheet = CellFormat(cellPrefix + "46", wSheet, payout.CashEquivalent, payout.balanceSheetUnits.CashEquivalentUnit);
            wSheet = CellFormat(cellPrefix + "47", wSheet, payout.TotalCurrentAssets, payout.balanceSheetUnits.TotalCurrentAssetUnit);
            wSheet = CellFormat(cellPrefix + "48", wSheet, payout.TotalAssests, payout.balanceSheetUnits.TotalAssetUnit);

            wSheet = CellFormat(cellPrefix + "49", wSheet, payout.TotalDebt, payout.balanceSheetUnits.TotalDebtUnit);
            wSheet = CellFormat(cellPrefix + "50", wSheet, payout.ShareholderEquity, payout.balanceSheetUnits.ShareEquitUnit);
            wSheet = CellFormat(cellPrefix + "52", wSheet, payout.CashFlowOperation, payout.CashFlowUnit);

            return wSheet;
        }

        ///snapshots
        [HttpPost]
        [Route("AddPayoutSnapshots")]
        public ActionResult<Object> AddPayoutSnapshots([FromBody] SnapshotsViewSnapshots model)
        {

            try
            {
                Snapshots snapshots = new Snapshots
                {
                    SnapShot = model.SnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    SnapShotType = PAYOUTPOLICY,
                    NPV = model.NVP,
                    CNPV = model.CNVP
                };
                iSnapShots.Add(snapshots);
                Console.WriteLine("blow the snapshot");
                iSnapShots.Commit();
                Console.WriteLine("blow the snapshot commit");
                return Ok(new { id = snapshots.Id, result = "Snapshot saved sucessfully" });
            }
            catch (Exception ex)
            {

                return BadRequest(ex.Message);
            }
        }


        [HttpPost]
        [Route("AddPayoutScenarioSnapshots")]
        public ActionResult<Object> AddPayoutScenarioSnapshots([FromBody] Snapshots model)
        {

            try
            {
                Snapshots snapshots = new Snapshots
                {
                    SnapShot = model.SnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    SnapShotType = PAYOUTSCENARIO,
                    NPV = model.NPV,
                    CNPV = model.CNPV

                };
                iSnapShots.Add(snapshots);
                iSnapShots.Commit();
                return Ok(new { id = snapshots.Id, result = "Snapshot saved sucessfully" });
            }
            catch (Exception ex)
            {

                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        [Route("PayoutSnapShots/{UserId}")]
        public ActionResult<Object> PayoutSnapShots(long UserId)
        {
            try
            {
                var SnapShot = iSnapShots.FindBy(s => s.SnapShotType == PAYOUTPOLICY && s.UserId == UserId);
                if (SnapShot == null)
                {
                    return NotFound();
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest( ex.Message);

            }

        }
        [HttpGet]
        [Route("PayoutScenarioSnapShots/{UserId}")]
        public ActionResult<Object> PayoutScenarioSnapShots(long UserId)
        {
            try
            {
                var SnapShot = iSnapShots.FindBy(s => s.SnapShotType == PAYOUTSCENARIO && s.UserId == UserId);
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
        [Route("ApprovePayout/{UserId}")]
        public ActionResult<Object> ApprovePayout(long UserId)
        {
            try
            {
                PayoutPolicy payoutPolicy = iPayoutPolicy.GetSingle(s => s.UserId == UserId);
                if (payoutPolicy != null)
                {
                    payoutPolicy.ApprovalFlag = 1;
                    iPayoutPolicy.Update(payoutPolicy);
                    iPayoutPolicy.Commit();
                    return Ok("Successfully approved");
                }
                else
                {
                    return NotFound("record not foud");
                }
            }
            catch (Exception ex)
            {

                return BadRequest(ex.Message);
            }


        }

        [HttpGet]
        [Route("EditInputFlag/{UserId}")]
        public ActionResult<Object> EditIputs(long UserId)
        {

            var payoutPolicy = iPayoutPolicy.GetSingle(s => s.UserId == UserId);
            if (payoutPolicy == null)
            {
                return NotFound($"No payout policy found for this UserId");
            }
            payoutPolicy.SummaryFlag = 0;
            iPayoutPolicy.Update(payoutPolicy);
            iPayoutPolicy.Commit();
            return Ok("Flag of Payout Policy changed to zero");


        }

        #endregion          

        #region CURRENT SETUP

        double GetMillionValue(string Unit, double value)
        {
            double Mvalue;
            if (Unit == "M" || Unit == "" || Unit == null)
            {
                Mvalue = value;
            }
            else if (Unit == "K")
            {
                Mvalue = value / 1000;
            }
            else if (Unit == "B")
            {
                Mvalue = value * 1000;
            }
            else if (Unit == "T")
            {
                Mvalue = value * 1000000;
            }
            else
            {
                Mvalue = value;
            }
            return Mvalue;
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

        [HttpGet]
        [Route("GetCurrentSetup_PayoutPolicy/{UserId}")]
        public ActionResult GetCurrentSetup_PayoutPolicy(long UserId)
        {
            CurrentSetup_PayoutPolicyResult result = new CurrentSetup_PayoutPolicyResult();
            CurrentSetupViewModel currentSetupObj = new CurrentSetupViewModel();
            bool IsSaved = false;

            //Date 22-July-2020 | Created By : anonymous | Enhancement : Uniform Units Object | Start
            List<ValueTextWrapper> unitList = EnumHelper.GetEnumDescriptions<CurrencyValueEnum>();
            List<ValueTextWrapper> numberCounts = EnumHelper.GetEnumDescriptions<NumberCountEnum>();
            result.currencyValueList = unitList;
            result.numberCountList = numberCounts;
            // End

            try
            {
                currentSetupObj = GetCurrentSetupDetailsByUserId(UserId);

                // for first time	
                if (currentSetupObj != null)
                {

                    //check for the timeline	
                    result.currentSetupObj = currentSetupObj;

                    List<InitialSetup_IValuation> initialSetup_IValuationList = new List<InitialSetup_IValuation>();
                    var initialSetupList = iInitialSetup_IValuation.FindBy(x => x.CIKNumber == currentSetupObj.CIKNumber && x.UserId == UserId).ToList();
                    if (initialSetupList != null && initialSetupList.Count > 0)
                    {
                        foreach (InitialSetup_IValuation obj in initialSetupList)
                        {
                            int? yearTo = (obj.YearTo != null ? obj.YearTo : 0) + (obj.ExplicitYearCount != null ? obj.ExplicitYearCount : 0);
                            if (Convert.ToInt32(currentSetupObj.CurrentYear) >= Convert.ToInt32(obj.YearFrom) && Convert.ToInt32(currentSetupObj.EndYear) <= Convert.ToInt32(yearTo))
                            {
                                //check if data exist in integrated financial Statement then add to list
                                // IntegratedDatas integratedDatasobj = iIntegratedDatas.FindBy(x=>x.InitialSetupId==obj.Id);
                                List<IntegratedDatas> integratedDatasListobj = iIntegratedDatas.FindBy(x => x.InitialSetupId == obj.Id).ToList();

                                //Check for Explicit Count
                                List<Integrated_ExplicitValues> integratedExplicitList = integratedDatasListobj != null && integratedDatasListobj.Count > 0 ? iIntegrated_ExplicitValues.FindBy(x => integratedDatasListobj.Any(t => t.Id == x.IntegratedDatasId)).ToList() : null;

                                if (integratedExplicitList != null && integratedExplicitList.Count > 0)
                                    initialSetup_IValuationList.Add(obj);
                            }
                        }
                    }

                    //check if data exist for current User
                    CurrentSetupIpResult renderResult = new CurrentSetupIpResult();

                    List<CurrentSetupIpDatas> TblCurrentSetupIpDatasList = currentSetupObj != null && currentSetupObj.Id != null && currentSetupObj.Id != 0 ? iCurrentSetupIpDatas.FindBy(x => x.CurrentSetupId == currentSetupObj.Id).ToList() : null;
                    if (TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0)
                    {
                        IsSaved = true;
                        //Get saved code from database
                        if (initialSetup_IValuationList != null && initialSetup_IValuationList.Count > 0)
                            result.InitialSetup_IValuationList = initialSetup_IValuationList;

                        List<CurrentSetupIpValues> TblCurrentSetupIpValuesList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? iCurrentSetupIpValues.FindBy(x => TblCurrentSetupIpDatasList.Any(t => t.Id == x.CurrentSetupIpDatasId)).ToList() : null;
                        if (TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0)
                        {
                            CurrentSetupIpDatasViewModel CurrentSetupDatasVMObj;
                            List<CurrentSetupIpDatasViewModel> CurrentSetupDatasVMListObj = new List<CurrentSetupIpDatasViewModel>();
                            CurrentSetupIpValuesViewModel CurrentSetupValuesVMObj;
                            foreach (CurrentSetupIpDatas obj in TblCurrentSetupIpDatasList)
                            {
                                // map data to dataVM
                                CurrentSetupDatasVMObj = new CurrentSetupIpDatasViewModel();
                                CurrentSetupDatasVMObj = mapper.Map<CurrentSetupIpDatas, CurrentSetupIpDatasViewModel>(obj);
                                CurrentSetupDatasVMObj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                                List<CurrentSetupIpValues> tempValuesList = TblCurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == obj.Id).ToList();
                                foreach (CurrentSetupIpValues valueObj in tempValuesList)
                                {
                                    CurrentSetupValuesVMObj = new CurrentSetupIpValuesViewModel();
                                    CurrentSetupValuesVMObj = mapper.Map<CurrentSetupIpValues, CurrentSetupIpValuesViewModel>(valueObj);
                                    CurrentSetupDatasVMObj.CurrentSetupIpValuesVM.Add(CurrentSetupValuesVMObj);
                                }
                                CurrentSetupDatasVMListObj.Add(CurrentSetupDatasVMObj);
                            }

                            List<CurrentSetupIpFillingsViewModel> filingList = new List<CurrentSetupIpFillingsViewModel>();
                            CurrentSetupIpFillingsViewModel CurrentSetupIpFilings;
                            List<CurrentSetupIpDatasViewModel> CurrentSetupIpDatasList;

                            // Sources  of Financing
                            CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                            CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
                            CurrentSetupIpFilings.StatementType = "Sources of Financing";
                            CurrentSetupIpDatasList = CurrentSetupDatasVMListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing).ToList();
                            CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                            filingList.Add(CurrentSetupIpFilings);

                            // Other Inputs
                            CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                            CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
                            CurrentSetupIpFilings.StatementType = "Other Inputs";
                            CurrentSetupIpDatasList = CurrentSetupDatasVMListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.OtherInputs).ToList();
                            CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                            filingList.Add(CurrentSetupIpFilings);

                            // Current Payout Policy
                            CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                            CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
                            CurrentSetupIpFilings.StatementType = "Current Payout Policy";
                            CurrentSetupIpDatasList = CurrentSetupDatasVMListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.CurrentPayoutPolicy).ToList();
                            CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                            filingList.Add(CurrentSetupIpFilings);

                            // Financial Statements: Pre-Payout
                            CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                            CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
                            CurrentSetupIpFilings.StatementType = "Financial Statements: Pre-Payout";
                            CurrentSetupIpDatasList = CurrentSetupDatasVMListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout).ToList();
                            CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                            filingList.Add(CurrentSetupIpFilings);

                            result.FilingResult = filingList;
                        }
                    }
                    else
                    {
                        if (initialSetup_IValuationList != null && initialSetup_IValuationList.Count > 0)
                        {
                            result.InitialSetup_IValuationList = initialSetup_IValuationList;
                            //if (initialSetup_IValuationList.Count == 1)
                            //{
                            // bind directly if only single item exist	
                            long initialSetupId = 0;
                            if (currentSetupObj.InitialSetupId != null && currentSetupObj.InitialSetupId != 0)
                                initialSetupId = Convert.ToInt64(currentSetupObj.InitialSetupId);
                            else
                                initialSetupId = initialSetup_IValuationList.Find(x => x.Id != 0).Id;
                            // CurrentSetupIpResult renderResult = new CurrentSetupIpResult();
                            renderResult = GetCurrentSetupIP(UserId, initialSetupId);
                            result.FilingResult = renderResult.Result;
                            // }
                        }
                        else
                        {
                            // CurrentSetupIpResult renderResult = new CurrentSetupIpResult();
                            renderResult = GetCurrentSetupIP(UserId, 0);
                            result.FilingResult = renderResult.Result;
                        }
                    }


                    result.IsSaved = IsSaved;

                }
                else
                {
                    result.IsSaved = IsSaved;
                    return Ok(result);
                }

            }
            catch (Exception ss) { }
            return Ok(result);
        }

        [HttpGet]
        [Route("GetCurrentSetup/{UserId}")]
        public ActionResult GetCurrentSetup(long UserId)
        {

            try
            {
                CurrentSetupResultObject resultObject = new CurrentSetupResultObject();
                var tblCurrentSetupObj = iCurrentSetup.GetSingle(s => s.UserId == UserId);
                CurrentSetupViewModel CurrentSetupObj = new CurrentSetupViewModel();
                if (tblCurrentSetupObj == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    // map table to vm
                    CurrentSetupObj = mapper.Map<CurrentSetup, CurrentSetupViewModel>(tblCurrentSetupObj);
                    return Ok(CurrentSetupObj);
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost]
        [Route("SaveCurrentSetup")]
        public ActionResult<Object> SaveCurrentSetup([FromBody] CurrentSetupViewModel model)
        {
            try
            {
                if (model.Id != 0 && model.isChanged == true)
                {
                    //delte data and Values 
                    List<CurrentSetupIpDatas> TblCurrentSetupIpDatasList = iCurrentSetupIpDatas.FindBy(x => x.CurrentSetupId == model.Id).ToList();
                    if (TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0)
                    {
                        //Get saved code from database
                        List<CurrentSetupIpValues> TblCurrentSetupIpValuesList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? iCurrentSetupIpValues.FindBy(x => TblCurrentSetupIpDatasList.Any(t => t.Id == x.CurrentSetupIpDatasId)).ToList() : null;
                        if (TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0)
                        {
                            iCurrentSetupIpValues.DeleteMany(TblCurrentSetupIpValuesList);
                            iCurrentSetupIpValues.Commit();
                        }
                        iCurrentSetupIpDatas.DeleteMany(TblCurrentSetupIpDatasList);
                        iCurrentSetupIpDatas.Commit();

                        //delete data from scenario analysis as well
                        List<PayoutPolicy_ScenarioDatas> tblPayoutPolicy_ScenarioDatasList = iPayoutPolicy_ScenarioDatas.FindBy(x => x.CurrentSetupId == model.Id).ToList();
                        if (tblPayoutPolicy_ScenarioDatasList != null && tblPayoutPolicy_ScenarioDatasList.Count > 0)
                        {
                            List<PayoutPolicy_ScenarioValues> TblPayoutPolicy_ScenarioValuesList = tblPayoutPolicy_ScenarioDatasList != null && tblPayoutPolicy_ScenarioDatasList.Count > 0 ? iPayoutPolicy_ScenarioValues.FindBy(x => tblPayoutPolicy_ScenarioDatasList.Any(t => t.Id == x.PayoutPolicy_ScenarioDatasId)).ToList() : null;
                            if (TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0)
                            {
                                iPayoutPolicy_ScenarioValues.DeleteMany(TblPayoutPolicy_ScenarioValuesList);
                                iPayoutPolicy_ScenarioValues.Commit();
                            }
                            iPayoutPolicy_ScenarioDatas.DeleteMany(tblPayoutPolicy_ScenarioDatasList);
                            iPayoutPolicy_ScenarioDatas.Commit();
                        }

                    }
                }

                //CurrentSetup currentSetup = new CurrentSetup
                //{
                //    CurrentYear = model.CurrentYear,
                //    EndYear = model.EndYear,
                //    InitialSetupId = model.InitialSetupId,
                //    UserId = model.UserId
                //};

                // map VM to table
                CurrentSetup currentSetup = mapper.Map<CurrentSetupViewModel, CurrentSetup>(model);

                //check for company
                FilingsTable filingsTable = !string.IsNullOrEmpty(model.CIKNumber) ? iFilings.GetSingle(x => x.CIK == model.CIKNumber) : null;
                if (filingsTable != null)
                    currentSetup.Company = filingsTable.CompanyName;

                if (model.Id == 0)
                {
                    iCurrentSetup.Add(currentSetup);
                }
                else
                {
                    // currentSetup.Id = model.Id;
                    iCurrentSetup.Update(currentSetup);
                }
                iCurrentSetup.Commit();

                return new
                {
                    id = currentSetup.Id,
                    result = model.Id == 0 ? "Current Setup Created Sucessfully" : "Current Setup Modified Sucessfully"
                };
            }
            catch (Exception ex)
            {
                return BadRequest(new { Message = "Invalid Entry", StatusCode = 400 });
            }
        }

        [HttpGet]
        [Route("GetCurrentSetupIP/{UserId}/{initialSetupId}")]
        public CurrentSetupIpResult GetCurrentSetupIP(long UserId, long initialSetupId)
        {
            CurrentSetupIpResult renderResult = new CurrentSetupIpResult();
            List<CurrentSetupIpFillingsViewModel> CurrentSetupIpFilingsList = new List<CurrentSetupIpFillingsViewModel>();
            List<CurrentSetupIpDatasViewModel> CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
            CurrentSetupIpDatasViewModel CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
            CurrentSetupIpFillingsViewModel CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
            List<CurrentSetupIpValuesViewModel> CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
            CurrentSetupIpValuesViewModel CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();

            try
            {
                CurrentSetup tblCurrentSetup = iCurrentSetup.GetSingle(x => x.UserId == UserId);
                if (initialSetupId != 0)
                {
                    tblCurrentSetup.InitialSetupId = initialSetupId;
                    iCurrentSetup.Update(tblCurrentSetup);
                    iCurrentSetup.Commit();
                }

                List<CurrentSetupIpDatas> TblCurrentSetupIpDatasList = tblCurrentSetup != null && tblCurrentSetup.Id != null && tblCurrentSetup.Id != 0 ? iCurrentSetupIpDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList() : null;
                if (TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0)
                {
                    //Get saved code from database


                }
                else
                {
                    List<CurrentSetupIpValuesViewModel> dummyCurrentSetupIpValuesViewModelList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpValuesViewModel dummyCurrentSetupIpValuesViewModel = new CurrentSetupIpValuesViewModel();

                    if (tblCurrentSetup != null)
                    {
                        Console.WriteLine("Current Year above the first change : " + tblCurrentSetup.CurrentYear);
                        int strt = Convert.ToInt32(tblCurrentSetup.CurrentYear);
                        int End = Convert.ToInt32(tblCurrentSetup.EndYear);
                        Console.WriteLine("Current Year below  the first change : " + tblCurrentSetup.CurrentYear);
                        for (int i = strt; i <= End; i++)
                        {
                            dummyCurrentSetupIpValuesViewModel = new CurrentSetupIpValuesViewModel();
                            Console.WriteLine("Current Year above the first change  to string: " + tblCurrentSetup.CurrentYear);
                            dummyCurrentSetupIpValuesViewModel.Year = Convert.ToString(strt);
                            dummyCurrentSetupIpValuesViewModel.Value = "";
                            dummyCurrentSetupIpValuesViewModelList.Add(dummyCurrentSetupIpValuesViewModel);
                            strt = strt + 1;
                        }
                    }

                    var costOfCapital = iCostOfCapital.GetSingle(s => s.UserId == UserId);
                    var CapitalStructure = iCapitalStructure.GetSingle(s => s.UserId == UserId);
                    var IntegratedDataList = tblCurrentSetup.InitialSetupId != null && tblCurrentSetup.InitialSetupId != 0 ? iIntegratedDatas.FindBy(x => x.InitialSetupId == tblCurrentSetup.InitialSetupId).ToList() : null;
                    var Integrated_ExplicitValueList = IntegratedDataList != null && IntegratedDataList.Count > 0 ? iIntegrated_ExplicitValues.FindBy(x => IntegratedDataList.Any(m => m.Id == x.IntegratedDatasId)).ToList() : null;

                    TaxRates_IValuation tblTaxRatesObj = new TaxRates_IValuation();
                    tblTaxRatesObj = tblCurrentSetup.InitialSetupId != null && tblCurrentSetup.InitialSetupId != 0 ? iTaxRates_IValuation.GetSingle(x => x.InitialSetupId == tblCurrentSetup.InitialSetupId) : null;

                    #region Sources of Financing
                    CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                    CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();

                    CurrentSetupIpFilings.StatementType = "Sources of Financing";

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Equity";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = true;
                    CurrentSetupIpDatasViewModelobj.Sequence = 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Price per Share";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = CapitalStructure != null && CapitalStructure.equityUnits != null && CapitalStructure.equityUnits.CurrentSharePriceUnit != null ? CapitalStructure.equityUnits.CurrentSharePriceUnit : "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.equity != null && CapitalStructure.equity.CurrentSharePrice != null ? CapitalStructure.equity.CurrentSharePrice / (CapitalStructure.equityUnits != null ? UnitConversion.ReturnDividend(CapitalStructure.equityUnits.CurrentSharePriceUnit) : 1) : 0;
                        Console.WriteLine("Current Year above the first chang the hashtag : " + tblCurrentSetup.CurrentYear);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Number of Shares Outstanding - Basic";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = CapitalStructure != null && CapitalStructure.equityUnits != null && CapitalStructure.equityUnits.NumberShareBasicUnit != null ? CapitalStructure.equityUnits.NumberShareBasicUnit : "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.equity != null && CapitalStructure.equity.NumberShareBasic != null ? CapitalStructure.equity.NumberShareBasic / UnitConversion.ReturnDividend(CurrentSetupIpDatasViewModelobj.Unit) : 0;
                        Console.WriteLine("Current Year above the first change 2nd hashtag : " + tblCurrentSetup.CurrentYear);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Number of Shares Outstanding - Diluted";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = CapitalStructure != null && CapitalStructure.equityUnits != null && CapitalStructure.equityUnits.NumberShareOutstandingUnit != null ? CapitalStructure.equityUnits.NumberShareOutstandingUnit : "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.equity != null && CapitalStructure.equity.NumberShareOutstanding != null ? CapitalStructure.equity.NumberShareOutstanding / UnitConversion.ReturnDividend(CurrentSetupIpDatasViewModelobj.Unit) : 0;
                        Console.WriteLine("Current Year above the first change third hashtag : " + tblCurrentSetup.CurrentYear);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Cost of Equity (rE)";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = "%";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.equity != null && CapitalStructure.equity.CostOfEquity != null ? CapitalStructure.equity.CostOfEquity : 0;
                        Console.WriteLine("Current Year above the first change four hashtag : " + tblCurrentSetup.CurrentYear);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Preferred Equity";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = true;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Current Preferred Share Price";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = CapitalStructure != null && CapitalStructure.prefferedEquityUnit != null && CapitalStructure.prefferedEquityUnit.PrefferedSharePriceUnit != null ? CapitalStructure.prefferedEquityUnit.PrefferedSharePriceUnit : "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.prefferedEquity != null && CapitalStructure.prefferedEquity.PrefferedSharePrice != null ? CapitalStructure.prefferedEquity.PrefferedSharePrice / UnitConversion.ReturnDividend(CurrentSetupIpDatasViewModelobj.Unit) : 0;
                        Console.WriteLine("Current Year above the first change fifth hashtag: " + tblCurrentSetup.CurrentYear);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Number of Preferred Shares Outstanding";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = CapitalStructure != null && CapitalStructure.prefferedEquityUnit != null && CapitalStructure.prefferedEquityUnit.PrefferedShareOutstandingUnit != null ? CapitalStructure.prefferedEquityUnit.PrefferedShareOutstandingUnit : "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.prefferedEquity != null && CapitalStructure.prefferedEquity.PrefferedShareOutstanding != null ? CapitalStructure.prefferedEquity.PrefferedShareOutstanding / UnitConversion.ReturnDividend(CurrentSetupIpDatasViewModelobj.Unit) : 0;
                        Console.WriteLine("Current Year above the first change  six hashtag: " + tblCurrentSetup.CurrentYear);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Preferred Dividend";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = CapitalStructure != null && CapitalStructure.prefferedEquityUnit != null && CapitalStructure.prefferedEquityUnit.PrefferedDividendUnit != null ? CapitalStructure.prefferedEquityUnit.PrefferedDividendUnit : "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.prefferedEquity != null && CapitalStructure.prefferedEquity.PrefferedDividend != null ? CapitalStructure.prefferedEquity.PrefferedDividend / UnitConversion.ReturnDividend(CurrentSetupIpDatasViewModelobj.Unit) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Cost of Preferred Equity (rP)";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = "%";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.prefferedEquity != null && CapitalStructure.prefferedEquity.CostPreffEquity != null ? CapitalStructure.prefferedEquity.CostPreffEquity : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Debt";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = true;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Market Value of Interest Bearing Debt";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = CapitalStructure != null && CapitalStructure.debtUnit != null && CapitalStructure.debtUnit.MarketValueDebtUnit != null ? CapitalStructure.debtUnit.MarketValueDebtUnit : "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.debt != null && CapitalStructure.debt.MarketValueDebt != null ? CapitalStructure.debt.MarketValueDebt / UnitConversion.ReturnDividend(CurrentSetupIpDatasViewModelobj.Unit) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Cost of Debt (rD)";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    CurrentSetupIpDatasViewModelobj.Unit = "%";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.debt != null && CapitalStructure.debt.CostOfDebt != null ? CapitalStructure.debt.CostOfDebt : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                    CurrentSetupIpFilingsList.Add(CurrentSetupIpFilings);
                    #endregion

                    #region Other Inputs

                    CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                    CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
                    CurrentSetupIpFilings.StatementType = "Other Inputs";

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Cash Needed for Working Capital";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.OtherInputs;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>(); // Operating Cash
                    var OperatingCashDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.Category == "Operating" && (x.LineItem.ToLower().Contains("cash and cash equivalents") || x.LineItem.ToLower().Contains("Operating Cash"))) : null;
                    var OperatingCashValues = OperatingCashDt != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == OperatingCashDt.Id) : null;
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = OperatingCashValues != null && OperatingCashValues.Count > 0 && OperatingCashValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OperatingCashValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(OperatingCashValues.Find(x => x.Year == obj.Year).Value) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Interest Coverage Ratio (k) - If Applicable";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.OtherInputs;
                    CurrentSetupIpDatasViewModelobj.Unit = "%";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = CapitalStructure != null && CapitalStructure.InterestCoverage != null ? CapitalStructure.InterestCoverage : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Interest Income Rate";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.OtherInputs;
                    CurrentSetupIpDatasViewModelobj.Unit = "%";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Marginal Tax Rate (Tc)";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.OtherInputs;
                    CurrentSetupIpDatasViewModelobj.Unit = "%";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = tblTaxRatesObj != null && tblTaxRatesObj.Marginal != null ? Convert.ToDouble(tblTaxRatesObj.Marginal) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                    CurrentSetupIpFilingsList.Add(CurrentSetupIpFilings);
                    #endregion

                    #region Current Payout Policy


                    CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                    CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
                    CurrentSetupIpFilings.StatementType = "Current Payout Policy";

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Total Ongoing Dividend Payout -Annual";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutPolicy;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "One Time Dividend Payout";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutPolicy;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Stock Buyback Amount";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutPolicy;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                    CurrentSetupIpFilingsList.Add(CurrentSetupIpFilings);
                    #endregion

                    #region Financial Statements: Pre-Payout

                    CurrentSetupIpFilings = new CurrentSetupIpFillingsViewModel();
                    CurrentSetupIpDatasList = new List<CurrentSetupIpDatasViewModel>();
                    CurrentSetupIpFilings.StatementType = "Financial Statements: Pre-Payout";

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Income Statement";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = true;
                    CurrentSetupIpDatasViewModelobj.Sequence = 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Net Sales";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();

                    // find Revenue
                    IntegratedDatas revenueIntegratedObj = null;

                    if (IntegratedDataList != null && IntegratedDataList.Count > 0)
                    {
                        bool flag = false;
                        string revenuesynonyms = "Net Sales%Net Revenue%Revenue%Total Revenues%Sales%Total Net Revenue%Total revenue%Total net sales%Sales to customers%Total net revenues%Total revenues (Note 4)%Revenue from Contract with Customer, Excluding Assessed Tax%Revenues%Net revenues%Revenue, net";
                        List<string> synonyms = revenuesynonyms.Split('%').ToList(); // convert comma seperated values to list
                        foreach (IntegratedDatas integrateddatasObj in IntegratedDataList)
                        {

                            if (integrateddatasObj.IsParentItem != true)
                                foreach (var syn in synonyms)
                                {
                                    if (integrateddatasObj.LineItem.ToUpper() == syn.ToUpper())
                                    {
                                        revenueIntegratedObj = integrateddatasObj;
                                        flag = true;
                                        break;
                                    }

                                }

                            if (flag == true)
                                break;
                        }
                    }

                    List<Integrated_ExplicitValues> revenueExplicitvaluesList = new List<Integrated_ExplicitValues>();
                    revenueExplicitvaluesList = revenueIntegratedObj != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == revenueIntegratedObj.Id).ToList() : null;
                    ////////////////////
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        var RevenueValue = revenueExplicitvaluesList != null && revenueExplicitvaluesList.Count > 0 ? revenueExplicitvaluesList.Find(x => x.Year == obj.Year) : null;
                        double value = 0;
                        value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) ? Convert.ToDouble(RevenueValue.Value) : 0;
                        //double value = revenueExplicitvaluesList != null && revenueExplicitvaluesList.Count > 0 && revenueExplicitvaluesList.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(revenueExplicitvaluesList.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(revenueExplicitvaluesList.Find(x => x.Year == obj.Year).Value) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "EBITDA";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var EBITDADt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem.Contains("EBITDA")) : null;
                    // var EBITDAValues = EBITDADt != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == EBITDADt.Id) : null;
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        // var eBITDAValueObj = EBITDAValues != null && EBITDAValues.Count > 0 ? EBITDAValues.Find(x => x.Year == obj.Year) : null;
                        var eBITDAValueObj = EBITDADt != null && Integrated_ExplicitValueList != null && Integrated_ExplicitValueList.Count > 0 ? Integrated_ExplicitValueList.Find(x => x.IntegratedDatasId == EBITDADt.Id && x.Year == obj.Year) : null;
                        double value = eBITDAValueObj != null && !string.IsNullOrEmpty(eBITDAValueObj.Value) ? Convert.ToDouble(eBITDAValueObj.Value) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Depreciation & Amortization";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var DepreciationDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem.ToLower().Contains("depreciation") && x.Category == "Operating") : null;
                    var DepreciationValues = DepreciationDt != null && Integrated_ExplicitValueList != null && Integrated_ExplicitValueList.Count > 0 ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == DepreciationDt.Id) : null;

                    var AmortizDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem.ToLower().Contains("amortization") && x.Category == "Non-Operating") : null;
                    var AmortizValues = AmortizDt != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == AmortizDt.Id) : null;


                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = (DepreciationValues != null && DepreciationValues.Count > 0 && DepreciationValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepreciationValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(DepreciationValues.Find(x => x.Year == obj.Year).Value) : 0) +
                                       (AmortizValues != null && AmortizValues.Count > 0 && AmortizValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(AmortizValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(AmortizValues.Find(x => x.Year == obj.Year).Value) : 0);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Interest Income";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();

                    //                    
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "EBIT";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var EBITDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem == "EBIT") : null;
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Interest Expense";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var DebtDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.FindAll(x => x.LineItem.ToLower().Contains("debt")).ToList() : null;

                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = 0;
                        double tmpValue = 0;
                        if (DebtDt != null && DebtDt.Count > 0)
                        {
                            foreach (IntegratedDatas tempData in DebtDt)
                            {
                                Integrated_ExplicitValues tempExpvalue = Integrated_ExplicitValueList.Find(x => x.IntegratedDatasId == tempData.Id && x.Year == obj.Year);
                                tmpValue = tmpValue + (tempExpvalue != null && !string.IsNullOrEmpty(tempExpvalue.Value) ? Convert.ToDouble(tempExpvalue.Value) : 0);
                            }
                            value = tmpValue * (CapitalStructure != null && CapitalStructure.debt != null && CapitalStructure.debt.CostOfDebt != null ? CapitalStructure.debt.CostOfDebt : 0);
                        }

                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "EBT (Earnings Before Taxes)";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Taxes";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Net Earnings";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Balance Sheet";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = true;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Cash and Equivalents";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var ExcessCashDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem.ToLower().Contains("cash and cash equivalents") && x.Category == "Non-Operating") : null;
                    var ExcessCashValues = ExcessCashDt != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == ExcessCashDt.Id) : null;
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = (OperatingCashValues != null && OperatingCashValues.Count > 0 && OperatingCashValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OperatingCashValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(OperatingCashValues.Find(x => x.Year == obj.Year).Value) : 0) +
                                       (ExcessCashValues != null && ExcessCashValues.Count > 0 && ExcessCashValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(ExcessCashValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(ExcessCashValues.Find(x => x.Year == obj.Year).Value) : 0);
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Total Current Assets";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var TotalCurrentAssetsDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem.ToLower().Contains("total current assets")) : null;
                    var TotalCurrentAssetsValues = TotalCurrentAssetsDt != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == TotalCurrentAssetsDt.Id) : null;
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = TotalCurrentAssetsValues != null && TotalCurrentAssetsValues.Count > 0 && TotalCurrentAssetsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalCurrentAssetsValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(TotalCurrentAssetsValues.Find(x => x.Year == obj.Year).Value) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Total Assets";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var TotalAssetsDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem.ToLower().Contains("total assets")) : null;
                    var TotalAssetsValues = TotalAssetsDt != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == TotalAssetsDt.Id) : null;
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = TotalAssetsValues != null && TotalAssetsValues.Count > 0 && TotalAssetsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalAssetsValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(TotalAssetsValues.Find(x => x.Year == obj.Year).Value) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Total Debt";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();

                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = 0;
                        double tmpValue = 0;
                        if (DebtDt != null && DebtDt.Count > 0)
                        {
                            foreach (IntegratedDatas tempData in DebtDt)
                            {
                                Integrated_ExplicitValues tempExpvalue = Integrated_ExplicitValueList.Find(x => x.IntegratedDatasId == tempData.Id && x.Year == obj.Year);
                                tmpValue = tmpValue + (tempExpvalue != null && !string.IsNullOrEmpty(tempExpvalue.Value) ? Convert.ToDouble(tempExpvalue.Value) : 0);
                            }
                            value = tmpValue;
                        }
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Shareholders Equity";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    var ShareholdersDt = IntegratedDataList != null && IntegratedDataList.Count > 0 ? IntegratedDataList.Find(x => x.LineItem.ToLower() == "total stockholders' equity" || x.LineItem.ToLower() == "total stockholders’ equity" || x.LineItem.ToLower() == "total equity" || x.LineItem.ToLower() == "total shareholders equity" || x.LineItem.ToLower() == "total shareholders' investment" || x.LineItem.ToLower() == "total stockholders’ investment" || x.LineItem.ToLower() == "total shareholders’ equity" || x.LineItem.ToLower() == "total shareholders' equity" || x.LineItem.ToLower() == "total shareholders’ equity") : null;
                    var ShareholdersValues = ShareholdersDt != null ? Integrated_ExplicitValueList.FindAll(x => x.IntegratedDatasId == ShareholdersDt.Id) : null;
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        double value = ShareholdersValues != null && ShareholdersValues.Count > 0 && ShareholdersValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(ShareholdersValues.Find(x => x.Year == obj.Year).Value) ? Convert.ToDouble(ShareholdersValues.Find(x => x.Year == obj.Year).Value) : 0;
                        CurrentSetupIpValues.Value = value.ToString("0.##");
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Cash Flow Statement";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = true;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpDatasViewModelobj = new CurrentSetupIpDatasViewModel();
                    CurrentSetupIpValuesList = new List<CurrentSetupIpValuesViewModel>();
                    CurrentSetupIpDatasViewModelobj.LineItem = "Cash Flow from Operations";
                    CurrentSetupIpDatasViewModelobj.IsParentItem = false;
                    CurrentSetupIpDatasViewModelobj.Sequence = CurrentSetupIpDatasList.Count + 1;
                    CurrentSetupIpDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPrePayout;
                    CurrentSetupIpDatasViewModelobj.Unit = "M";
                    CurrentSetupIpDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = new List<CurrentSetupIpValuesViewModel>();
                    foreach (CurrentSetupIpValuesViewModel obj in dummyCurrentSetupIpValuesViewModelList)
                    {
                        CurrentSetupIpValues = new CurrentSetupIpValuesViewModel();
                        CurrentSetupIpValues.Year = obj.Year;
                        CurrentSetupIpValues.Value = "";
                        CurrentSetupIpValuesList.Add(CurrentSetupIpValues);
                    }
                    CurrentSetupIpDatasViewModelobj.CurrentSetupIpValuesVM = CurrentSetupIpValuesList;
                    CurrentSetupIpDatasList.Add(CurrentSetupIpDatasViewModelobj);

                    CurrentSetupIpFilings.CurrentSetupIpDatasViewModelVM = CurrentSetupIpDatasList;
                    CurrentSetupIpFilingsList.Add(CurrentSetupIpFilings);
                    #endregion

                }
                renderResult.StatusCode = 1;
                renderResult.Message = "No issue found";
                renderResult.Result = CurrentSetupIpFilingsList;
                // return Ok(renderResult);
                return renderResult;
            }
            catch (Exception ex)
            {
                renderResult.StatusCode = 0;
                renderResult.Message = "Exception occured " + Convert.ToString(ex.Message);
                renderResult.Result = CurrentSetupIpFilingsList;
                // return Ok(renderResult);
                return renderResult;
            }

        }

        // save Payout Policy
        [HttpPost]
        [Route("SaveCurrentSetupIP")]
        public ActionResult SaveCurrentSetupIP([FromBody] List<CurrentSetupIpFillingsViewModel> CurrentSetupIpFillingsvmList)
        {
            try
            {

                if (CurrentSetupIpFillingsvmList != null && CurrentSetupIpFillingsvmList.Count > 0)
                {
                    CurrentSetupIpFillingsViewModel OtherInputList = CurrentSetupIpFillingsvmList.Find(x => x.StatementType == "Other Inputs");
                    CurrentSetupIpFillingsViewModel FinancialStatementPPList = CurrentSetupIpFillingsvmList.Find(x => x.StatementType == "Financial Statements: Pre-Payout");

                    var C29 = OtherInputList != null && OtherInputList.CurrentSetupIpDatasViewModelVM != null && OtherInputList.CurrentSetupIpDatasViewModelVM.Count > 0 ? OtherInputList.CurrentSetupIpDatasViewModelVM.Find(x => x.LineItem.Contains("Interest Income Rate")) : null;
                    var C30 = OtherInputList != null && OtherInputList.CurrentSetupIpDatasViewModelVM != null && OtherInputList.CurrentSetupIpDatasViewModelVM.Count > 0 ? OtherInputList.CurrentSetupIpDatasViewModelVM.Find(x => x.LineItem.Contains("Marginal Tax Rate")) : null;

                    var C49 = FinancialStatementPPList != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Count > 0 ? FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Find(x => x.LineItem.Contains("Cash and Equivalents")) : null;
                    var C40 = FinancialStatementPPList != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Count > 0 ? FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Find(x => x.LineItem.Contains("EBITDA")) : null;
                    var C41 = FinancialStatementPPList != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Count > 0 ? FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Find(x => x.LineItem.Contains("Depreciation & Amortization")) : null;
                    var C44 = FinancialStatementPPList != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM != null && FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Count > 0 ? FinancialStatementPPList.CurrentSetupIpDatasViewModelVM.Find(x => x.LineItem.Contains("Interest Expense")) : null;

                    foreach (CurrentSetupIpFillingsViewModel CurrentSetupIpFillingsObj in CurrentSetupIpFillingsvmList)
                    {
                        if (CurrentSetupIpFillingsObj.CurrentSetupIpDatasViewModelVM != null && CurrentSetupIpFillingsObj.CurrentSetupIpDatasViewModelVM.Count > 0)
                        {
                            foreach (CurrentSetupIpDatasViewModel CurrentSetupIpDatasVMObj in CurrentSetupIpFillingsObj.CurrentSetupIpDatasViewModelVM)
                            {
                                if (CurrentSetupIpDatasVMObj == null)
                                {
                                    continue; 
                                }

                                CurrentSetupIpDatas tblCurrentSetupIpDatasObj = new CurrentSetupIpDatas();
                                tblCurrentSetupIpDatasObj = mapper.Map<CurrentSetupIpDatasViewModel, CurrentSetupIpDatas>(CurrentSetupIpDatasVMObj);
                                if (CurrentSetupIpDatasVMObj.CurrentSetupIpValuesVM != null && CurrentSetupIpDatasVMObj.CurrentSetupIpValuesVM.Count > 0)
                                {
                                    tblCurrentSetupIpDatasObj.CurrentSetupIpValues = new List<CurrentSetupIpValues>();

                                    foreach (CurrentSetupIpValuesViewModel CurrentSetupIpValues in CurrentSetupIpDatasVMObj.CurrentSetupIpValuesVM)
                                    {
                                        if (CurrentSetupIpValues == null)
                                        {
                                            continue; 
                                        }
                                        CurrentSetupIpValues ExplicitValue = mapper.Map<CurrentSetupIpValuesViewModel, CurrentSetupIpValues>(CurrentSetupIpValues);

                                        ////////For some specific line items change calculation
                                        if (CurrentSetupIpDatasVMObj.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && CurrentSetupIpDatasVMObj.LineItem.Contains("Interest Income"))
                                        {
                                            double value = (C29 != null && C29.CurrentSetupIpValuesVM != null && C29.CurrentSetupIpValuesVM.Count > 0 && C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C29.Unit, Convert.ToDouble(C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) *
                                                           (C49 != null && C49.CurrentSetupIpValuesVM != null && C49.CurrentSetupIpValuesVM.Count > 0 && C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C49.Unit, Convert.ToDouble(C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) / 100;
                                            ExplicitValue.Value = value.ToString("0.##");
                                        }
                                        else if (CurrentSetupIpDatasVMObj.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && CurrentSetupIpDatasVMObj.LineItem == "EBIT")
                                        {
                                            double I42 = (C29 != null && C29.CurrentSetupIpValuesVM != null && C29.CurrentSetupIpValuesVM.Count > 0 && C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C29.Unit, Convert.ToDouble(C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) *
                                                            (C49 != null && C49.CurrentSetupIpValuesVM != null && C49.CurrentSetupIpValuesVM.Count > 0 && C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C49.Unit, Convert.ToDouble(C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) / 100;

                                            double value = (C40 != null && C40.CurrentSetupIpValuesVM != null && C40.CurrentSetupIpValuesVM.Count > 0 && C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C40.Unit, Convert.ToDouble(C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) -
                                                           (C41 != null && C41.CurrentSetupIpValuesVM != null && C41.CurrentSetupIpValuesVM.Count > 0 && C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C41.Unit, Convert.ToDouble(C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) + I42;
                                            ExplicitValue.Value = value.ToString("0.##");
                                        }
                                        else if (CurrentSetupIpDatasVMObj.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && CurrentSetupIpDatasVMObj.LineItem.Contains("EBT (Earnings Before Taxes)"))
                                        {
                                            double I42 = (C29 != null && C29.CurrentSetupIpValuesVM != null && C29.CurrentSetupIpValuesVM.Count > 0 && C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C29.Unit, Convert.ToDouble(C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) *
                                                            (C49 != null && C49.CurrentSetupIpValuesVM != null && C49.CurrentSetupIpValuesVM.Count > 0 && C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C49.Unit, Convert.ToDouble(C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) / 100;

                                            double I43 = (C40 != null && C40.CurrentSetupIpValuesVM != null && C40.CurrentSetupIpValuesVM.Count > 0 && C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C40.Unit, Convert.ToDouble(C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) -
                                                           (C41 != null && C41.CurrentSetupIpValuesVM != null && C41.CurrentSetupIpValuesVM.Count > 0 && C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C41.Unit, Convert.ToDouble(C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) + I42;

                                            double value = I43 - (C44 != null && C44.CurrentSetupIpValuesVM != null && C44.CurrentSetupIpValuesVM.Count > 0 && C44.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C44.Unit, Convert.ToDouble(C44.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0);
                                            ExplicitValue.Value = value.ToString("0.##");
                                        }
                                        else if (CurrentSetupIpDatasVMObj.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && CurrentSetupIpDatasVMObj.LineItem.Contains("Taxes"))
                                        {
                                            double I42 = (C29 != null && C29.CurrentSetupIpValuesVM != null && C29.CurrentSetupIpValuesVM.Count > 0 && C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C29.Unit, Convert.ToDouble(C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) *
                                                            (C49 != null && C49.CurrentSetupIpValuesVM != null && C49.CurrentSetupIpValuesVM.Count > 0 && C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C49.Unit, Convert.ToDouble(C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) / 100;

                                            double I43 = (C40 != null && C40.CurrentSetupIpValuesVM != null && C40.CurrentSetupIpValuesVM.Count > 0 && C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C40.Unit, Convert.ToDouble(C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) -
                                                           (C41 != null && C41.CurrentSetupIpValuesVM != null && C41.CurrentSetupIpValuesVM.Count > 0 && C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C41.Unit, Convert.ToDouble(C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) + I42;

                                            double I45 = I43 - (C44 != null && C44.CurrentSetupIpValuesVM != null && C44.CurrentSetupIpValuesVM.Count > 0 && C44.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C44.Unit, Convert.ToDouble(C44.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0);
                                            double value = I45 * (C30 != null && C30.CurrentSetupIpValuesVM != null && C30.CurrentSetupIpValuesVM.Count > 0 && C30.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C30.Unit, Convert.ToDouble(C30.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) / 100;
                                            ExplicitValue.Value = value.ToString("0.##");
                                        }
                                        else if (CurrentSetupIpDatasVMObj.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && CurrentSetupIpDatasVMObj.LineItem.Contains("Net Earnings"))
                                        {
                                            double I42 = (C29 != null && C29.CurrentSetupIpValuesVM != null && C29.CurrentSetupIpValuesVM.Count > 0 && C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C29.Unit, Convert.ToDouble(C29.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) *
                                                            (C49 != null && C49.CurrentSetupIpValuesVM != null && C49.CurrentSetupIpValuesVM.Count > 0 && C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C49.Unit, Convert.ToDouble(C49.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) / 100;

                                            double I43 = (C40 != null && C40.CurrentSetupIpValuesVM != null && C40.CurrentSetupIpValuesVM.Count > 0 && C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C40.Unit, Convert.ToDouble(C40.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) -
                                                           (C41 != null && C41.CurrentSetupIpValuesVM != null && C41.CurrentSetupIpValuesVM.Count > 0 && C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C41.Unit, Convert.ToDouble(C41.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) + I42;

                                            double I45 = I43 - (C44 != null && C44.CurrentSetupIpValuesVM != null && C44.CurrentSetupIpValuesVM.Count > 0 && C44.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C44.Unit, Convert.ToDouble(C44.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0);
                                            double I46 = I45 * (C30 != null && C30.CurrentSetupIpValuesVM != null && C30.CurrentSetupIpValuesVM.Count > 0 && C30.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year) != null ? GetMillionValue(C30.Unit, Convert.ToDouble(C30.CurrentSetupIpValuesVM.Find(x => x.Year == CurrentSetupIpValues.Year).Value)) : 0) / 100;
                                            double value = I45 - I46;
                                            ExplicitValue.Value = value.ToString("0.##");
                                        }
                                        ////////////////////////////////////////////////////////
                                        tblCurrentSetupIpDatasObj.CurrentSetupIpValues.Add(ExplicitValue);
                                    }

                                }

                                try
                                {
                                    if (tblCurrentSetupIpDatasObj.Id == 0)
                                    {
                                        // Save Code
                                        iCurrentSetupIpDatas.Add(tblCurrentSetupIpDatasObj);
                                        iCurrentSetupIpDatas.Commit();
                                    }
                                    else
                                    {
                                        // Check if the entity still exists before updating
                                        var existingData = iCurrentSetupIpDatas.GetSingle(x => x.Id == tblCurrentSetupIpDatasObj.Id);
                                        if (existingData != null)
                                        {
                                            iCurrentSetupIpDatas.Update(tblCurrentSetupIpDatasObj);
                                            iCurrentSetupIpDatas.Commit();
                                        }
                                        else
                                        {
                                            return NotFound(new { message = "Entity not found for update", status = 404, result = false });
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    return BadRequest(new { message = "Data has been modified or deleted since it was loaded. Please reload and try again.", status = 409, result = false });
                                }


                                // if (tblCurrentSetupIpDatasObj.Id == 0)
                                // {
                                //     //Save Code
                                //     iCurrentSetupIpDatas.Add(tblCurrentSetupIpDatasObj);
                                //     iCurrentSetupIpDatas.Commit();
                                // }
                                // else
                                // {
                                //     //Update Code
                                //     iCurrentSetupIpDatas.Update(tblCurrentSetupIpDatasObj);
                                //     iCurrentSetupIpDatas.Commit();
                                // }
                            }
                        }

                    }
                    return Ok(new { message = "Data Saved Successfully", status = 1, result = true });
                }
                else
                {
                    return BadRequest(new { message = "No data found to save", status = 200, result = false });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(Convert.ToString(ex.Message));
            }
        }

        #endregion

        #region Summery Output & Snapshot

        [HttpGet]
        [Route("GetCurrentSetupSO/{UserId}")]
        public ActionResult GetCurrentSetupSO(long UserId)
        {
            CurrentSetuoSoResult renderResult = new CurrentSetuoSoResult();
            List<CurrentSetupSoFillingsViewModel> CurrentSetupSoFilingsList = new List<CurrentSetupSoFillingsViewModel>();
            List<CurrentSetupSoDatasViewModel> CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
            CurrentSetupSoDatasViewModel CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
            CurrentSetupSoFillingsViewModel CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
            List<CurrentSetupSoValuesViewModel> CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
            CurrentSetupSoValuesViewModel CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();

            //Date 22-July-2020 | Created By : anonymous | Enhancement : Uniform Units Object | Start
            List<ValueTextWrapper> units = EnumHelper.GetEnumDescriptions<CurrencyValueEnum>();
            List<ValueTextWrapper> numberCounts = EnumHelper.GetEnumDescriptions<NumberCountEnum>();
            renderResult.currencyValueList = units;
            renderResult.numberCountList = numberCounts;
            // End

            try
            {
                CurrentSetup tblCurrentSetup = iCurrentSetup.GetSingle(x => x.UserId == UserId);
                List<CurrentSetupIpDatas> CurrentSetupIpDatasList = tblCurrentSetup != null && tblCurrentSetup.Id != null && tblCurrentSetup.Id != 0 ? iCurrentSetupIpDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList() : null;

                List<CurrentSetupIpValues> CurrentSetupIpValuesList = CurrentSetupIpDatasList != null && CurrentSetupIpDatasList.Count > 0 ? iCurrentSetupIpValues.FindBy(x => CurrentSetupIpDatasList.Any(m => m.Id == x.CurrentSetupIpDatasId)).ToList() : null;

                if (CurrentSetupIpDatasList != null && CurrentSetupIpDatasList.Count > 0 && CurrentSetupIpValuesList != null && CurrentSetupIpValuesList.Count > 0)
                {
                    List<CurrentSetupSoValuesViewModel> dummyCurrentSetupSoValuesViewModelList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoValuesViewModel dummyCurrentSetupSoValuesViewModel = new CurrentSetupSoValuesViewModel();

                    if (tblCurrentSetup != null)
                    {
                        int strt = Convert.ToInt32(tblCurrentSetup.CurrentYear);
                        int End = Convert.ToInt32(tblCurrentSetup.EndYear);
                        for (int i = strt; i <= End; i++)
                        {
                            dummyCurrentSetupSoValuesViewModel = new CurrentSetupSoValuesViewModel();
                            dummyCurrentSetupSoValuesViewModel.Year = Convert.ToString(strt);
                            dummyCurrentSetupSoValuesViewModel.Value = "";
                            dummyCurrentSetupSoValuesViewModelList.Add(dummyCurrentSetupSoValuesViewModel);
                            strt = strt + 1;
                        }
                    }

                    #region Current Capital Structure
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();

                    CurrentSetupSoFilings.StatementType = "Current Capital Structure";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Market Value of Equity ( E)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCapitalStructure;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var PricePerShareDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("price per share"));
                    var PricePerShareValues = PricePerShareDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == PricePerShareDt.Id).ToList() : null;

                    var NoOfShareOutstandingBasicDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("number of shares outstanding - basic"));
                    var NoOfShareOutstandingBasicValues = NoOfShareOutstandingBasicDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == NoOfShareOutstandingBasicDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        value = (PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0) *
                                (NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Market Value of Preferred Equity (P)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCapitalStructure;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var CurrentPreferSharePriceDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("current preferred share price"));
                    var CurrentPreferSharePriceValues = CurrentPreferSharePriceDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CurrentPreferSharePriceDt.Id).ToList() : null;

                    var NoOfPreferShareOutstandingDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("number of preferred shares outstanding"));
                    var NoOfPreferShareOutstandingValues = NoOfPreferShareOutstandingDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == NoOfPreferShareOutstandingDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        value = (CurrentPreferSharePriceValues != null && CurrentPreferSharePriceValues.Count > 0 && CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CurrentPreferSharePriceDt.Unit, Convert.ToDouble(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value)) : 0) *
                                (NoOfPreferShareOutstandingValues != null && NoOfPreferShareOutstandingValues.Count > 0 && NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfPreferShareOutstandingDt.Unit, Convert.ToDouble(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Market Value of Debt (D)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCapitalStructure;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var MarketvalueofInterestBearingDebtDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("market value of interest bearing debt"));
                    var MarketvalueofInterestBearingDebtValues = MarketvalueofInterestBearingDebtDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == MarketvalueofInterestBearingDebtDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = MarketvalueofInterestBearingDebtValues != null && MarketvalueofInterestBearingDebtValues.Count > 0 && MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarketvalueofInterestBearingDebtDt.Unit, Convert.ToDouble(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Market Value of Net Debt (ND)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCapitalStructure;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var CashAndEquivalentsDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("cash and equivalents"));
                    var CashAndEquivalentsValues = CashAndEquivalentsDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CashAndEquivalentsDt.Id).ToList() : null;

                    var CashNeededForWorkingcapitalDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("cash needed for working capital"));
                    var CashNeededForWorkingcapitalValues = CashNeededForWorkingcapitalDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CashNeededForWorkingcapitalDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        value = (MarketvalueofInterestBearingDebtValues != null && MarketvalueofInterestBearingDebtValues.Count > 0 && MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarketvalueofInterestBearingDebtDt.Unit, Convert.ToDouble(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value)) : 0) - (CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0) + (CashNeededForWorkingcapitalValues != null && CashNeededForWorkingcapitalValues.Count > 0 && CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashNeededForWorkingcapitalDt.Unit, Convert.ToDouble(CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    #region Current Cost of Capital
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Current Cost of Capital";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cost of Equity ( rE)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCostOfCapital;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var CostOfEquityDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("cost of equity"));
                    var CostOfEquityValues = CostOfEquityDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CostOfEquityDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = CostOfEquityValues != null && CostOfEquityValues.Count > 0 && CostOfEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfEquityDt.Unit, Convert.ToDouble(CostOfEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cost of Preferred Equity (rP)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCostOfCapital;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var CostOfPreferredEquityDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("cost of preferred equity"));
                    var CostOfPreferredEquityValues = CostOfPreferredEquityDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CostOfPreferredEquityDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = CostOfPreferredEquityValues != null && CostOfPreferredEquityValues.Count > 0 && CostOfPreferredEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfPreferredEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfPreferredEquityDt.Unit, Convert.ToDouble(CostOfPreferredEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cost of Debt (rD)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCostOfCapital;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var CostOfDebtDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("cost of debt"));
                    var CostOfDebtValues = CostOfDebtDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CostOfDebtDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);


                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Unlevered Cost of Capital/Equity (rU)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCostOfCapital;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        double I1 = CostOfEquityValues != null && CostOfEquityValues.Count > 0 && CostOfEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfEquityDt.Unit, Convert.ToDouble(CostOfEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double I2 = CostOfPreferredEquityValues != null && CostOfPreferredEquityValues.Count > 0 && CostOfPreferredEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfPreferredEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfPreferredEquityDt.Unit, Convert.ToDouble(CostOfPreferredEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double I3 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s1 = (PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0) * (NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double s2 = (CurrentPreferSharePriceValues != null && CurrentPreferSharePriceValues.Count > 0 && CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CurrentPreferSharePriceDt.Unit, Convert.ToDouble(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value)) : 0) * (NoOfPreferShareOutstandingValues != null && NoOfPreferShareOutstandingValues.Count > 0 && NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfPreferShareOutstandingDt.Unit, Convert.ToDouble(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double s3 = (MarketvalueofInterestBearingDebtValues != null && MarketvalueofInterestBearingDebtValues.Count > 0 && MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarketvalueofInterestBearingDebtDt.Unit, Convert.ToDouble(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value)) : 0) - (CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0) + (CashNeededForWorkingcapitalValues != null && CashNeededForWorkingcapitalValues.Count > 0 && CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashNeededForWorkingcapitalDt.Unit, Convert.ToDouble(CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double sum = s1 + s2 + s3;
                        value = (I1 * s1 / sum) + (I2 * s2 / sum) + (I3 * s3 / sum);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Weighted Average Cost of Capital (rWACC)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentCostOfCapital;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var MarginalTaxRateDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("marginal tax rate"));
                    var MarginalTaxRateValues = MarginalTaxRateDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == MarginalTaxRateDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        double MTR = MarginalTaxRateValues != null && MarginalTaxRateValues.Count > 0 && MarginalTaxRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarginalTaxRateDt.Unit, Convert.ToDouble(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double I1 = CostOfEquityValues != null && CostOfEquityValues.Count > 0 && CostOfEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfEquityDt.Unit, Convert.ToDouble(CostOfEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double I2 = CostOfPreferredEquityValues != null && CostOfPreferredEquityValues.Count > 0 && CostOfPreferredEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfPreferredEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfPreferredEquityDt.Unit, Convert.ToDouble(CostOfPreferredEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double I3 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s1 = (PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0) * (NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double s2 = (CurrentPreferSharePriceValues != null && CurrentPreferSharePriceValues.Count > 0 && CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CurrentPreferSharePriceDt.Unit, Convert.ToDouble(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value)) : 0) * (NoOfPreferShareOutstandingValues != null && NoOfPreferShareOutstandingValues.Count > 0 && NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfPreferShareOutstandingDt.Unit, Convert.ToDouble(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double s3 = (MarketvalueofInterestBearingDebtValues != null && MarketvalueofInterestBearingDebtValues.Count > 0 && MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarketvalueofInterestBearingDebtDt.Unit, Convert.ToDouble(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value)) : 0) - (CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0) + (CashNeededForWorkingcapitalValues != null && CashNeededForWorkingcapitalValues.Count > 0 && CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashNeededForWorkingcapitalDt.Unit, Convert.ToDouble(CashNeededForWorkingcapitalValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double sum = s1 + s2 + s3;
                        value = (I1 * s1 / sum) + (I2 * s2 / sum) + ((I3 * s3 / sum) * (1 - (MTR / 100)));
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    #region Leverage Ratios
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Leverage Ratios";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Debt-to-Equity (D/E) Ratio";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.LeverageRatios;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        double s1 = MarketvalueofInterestBearingDebtValues != null && MarketvalueofInterestBearingDebtValues.Count > 0 && MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarketvalueofInterestBearingDebtDt.Unit, Convert.ToDouble(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = (PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0) * (NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        value = (s1 / s2) * 100;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Debt-to-Value (D/V) Ratio";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.LeverageRatios;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        double s1 = MarketvalueofInterestBearingDebtValues != null && MarketvalueofInterestBearingDebtValues.Count > 0 && MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarketvalueofInterestBearingDebtDt.Unit, Convert.ToDouble(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = (PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0) * (NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double s3 = (CurrentPreferSharePriceValues != null && CurrentPreferSharePriceValues.Count > 0 && CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CurrentPreferSharePriceDt.Unit, Convert.ToDouble(CurrentPreferSharePriceValues.Find(x => x.Year == obj.Year).Value)) : 0) * (NoOfPreferShareOutstandingValues != null && NoOfPreferShareOutstandingValues.Count > 0 && NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfPreferShareOutstandingDt.Unit, Convert.ToDouble(NoOfPreferShareOutstandingValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        value = (s1 / (s1 + s2 + s3)) * 100;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);


                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    #region Cash Balance: Post-Payout
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Cash Balance: Post-Payout";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cash & Equivalent";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CashBalancePostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    var TODPA_Dt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("total ongoing dividend payout -annual"));
                    var TODPA_Values = TODPA_Dt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == TODPA_Dt.Id).ToList() : null;

                    var OTDP_Dt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("one time dividend payout"));
                    var OTDP_Values = OTDP_Dt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == OTDP_Dt.Id).ToList() : null;

                    var SBA_Dt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("stock buyback amount"));
                    var SBA_Values = SBA_Dt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == SBA_Dt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        double s1 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s3 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s4 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        value = s1 - s2 - s3 - s4;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Excess Cash";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CashBalancePostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    var CNFWC_Dt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("cash needed for working capital"));
                    var CNFWC_Values = CNFWC_Dt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CNFWC_Dt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = 0;
                        double s1 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s3 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s4 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s5 = CNFWC_Values != null && CNFWC_Values.Count > 0 && CNFWC_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CNFWC_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CNFWC_Dt.Unit, Convert.ToDouble(CNFWC_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        value = s1 - s2 - s3 - s4 - s5;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);


                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);

                    #endregion

                    #region Internal Valuation: Post-Payout
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Internal Valuation: Post-Payout";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Unlevered Enterprise Value";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.InternalValuationPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Levered Enterprise Value";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.InternalValuationPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Equity Value";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.InternalValuationPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Interest Tax Shield Value";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.InternalValuationPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Stock Price";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.InternalValuationPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    #region Current Payout Analysis
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Current Payout Analysis";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Ongoing Quarterly Dividends";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = true;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Total Payout";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "DPS (Dividends per Share -Basic)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "$";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s1 / s2;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Dividend Payout Ratio";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var NetEarnings_Dt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("net earnings"));
                    var NetEarnings_Values = NetEarnings_Dt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == NetEarnings_Dt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = NetEarnings_Values != null && NetEarnings_Values.Count > 0 && NetEarnings_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NetEarnings_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NetEarnings_Dt.Unit, Convert.ToDouble(NetEarnings_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = (s1 / s2) * 100;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Dividend Yield";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s = s1 / s2;
                        double pps = PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s * 100 / pps;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "One Time Dividend";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = true;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Total Payout";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "DPS (Dividends per Share -Basic)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "$";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s1 / s2;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Dividend Payout Ratio";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = NetEarnings_Values != null && NetEarnings_Values.Count > 0 && NetEarnings_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NetEarnings_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NetEarnings_Dt.Unit, Convert.ToDouble(NetEarnings_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s1 * 100 / s2;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);


                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Dividend Yield";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s2 = NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s = s1 / s2;
                        double pps = PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s * 100 / pps;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Stock Buybacks";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = true;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Total Payout";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Shares Repurchased";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.CurrentPayoutAnalysis;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double pps = PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s1 / pps;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    #region Shares Outstanding & EPS: Post-Payout
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Shares Outstanding & EPS: Post-Payout";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Number of Shares Outstanding - Basic";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SharesOutstandingAndEPSPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double pps = PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s = NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s - (s1 / pps);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Number of Shares Outstanding - Diluted";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SharesOutstandingAndEPSPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var NoOfShareOutstandingDilutedDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("number of shares outstanding - diluted"));
                    var NoOfShareOutstandingDilutedValues = NoOfShareOutstandingDilutedDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == NoOfShareOutstandingDilutedDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s1 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double pps = PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s = NoOfShareOutstandingDilutedValues != null && NoOfShareOutstandingDilutedValues.Count > 0 && NoOfShareOutstandingDilutedValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingDilutedValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingDilutedDt.Unit, Convert.ToDouble(NoOfShareOutstandingDilutedValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s - (s1 / pps);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Net Earnings per Share - Basic";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SharesOutstandingAndEPSPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "$";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    var EBITDADt = CurrentSetupIpDatasList.Find(x => x.LineItem.Contains("EBITDA"));
                    var EBITDAValues = EBITDADt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == EBITDADt.Id).ToList() : null;

                    var DepAmtDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("depreciation & amortization"));
                    var DepAmtValues = DepAmtDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == DepAmtDt.Id).ToList() : null;

                    var IIRateDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("interest income rate"));
                    var IIRateValues = IIRateDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == IIRateDt.Id).ToList() : null;

                    var TotalDebtDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("total debt"));
                    var TotalDebtValues = TotalDebtDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == TotalDebtDt.Id).ToList() : null;


                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;


                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s30 = MarginalTaxRateValues != null && MarginalTaxRateValues.Count > 0 && MarginalTaxRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarginalTaxRateDt.Unit, Convert.ToDouble(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;

                        double s119 = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) - (((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) * s30);

                        double s1 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double pps = PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s = NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s104 = s - (s1 / pps);


                        double value = s119 / s104;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Net Earnings per Share - Diluted";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.SharesOutstandingAndEPSPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "$";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;


                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s30 = MarginalTaxRateValues != null && MarginalTaxRateValues.Count > 0 && MarginalTaxRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarginalTaxRateDt.Unit, Convert.ToDouble(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;

                        double s119 = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) - (((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) * s30);

                        double s1 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double pps = PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s = NoOfShareOutstandingDilutedValues != null && NoOfShareOutstandingDilutedValues.Count > 0 && NoOfShareOutstandingDilutedValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingDilutedValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingDilutedDt.Unit, Convert.ToDouble(NoOfShareOutstandingDilutedValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s105 = s - (s1 / pps);
                        double value = s119 / s105;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    #region Shares Outstanding & EPS: Post-Payout
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Financial Statements: Post-Payout";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Income Statement";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = true;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Net Sales";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    var NetSalesDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("net sales"));
                    var NetSalesValues = NetSalesDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == NetSalesDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = NetSalesValues != null && NetSalesValues.Count > 0 && NetSalesValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NetSalesValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NetSalesDt.Unit, Convert.ToDouble(NetSalesValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "EBITDA";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Depreciation & Amortization";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Interest Income";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;

                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = (s49 - s33 - s34 - s35) * s29;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "EBIT";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;

                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s40 - s41 + ((s49 - s33 - s34 - s35) * s29);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Interest Expense";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;

                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s24 * s52;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "EBT (Earnings Before Taxes)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;

                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;

                        double value = (s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s24 * s52);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Taxes";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;

                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s30 = MarginalTaxRateValues != null && MarginalTaxRateValues.Count > 0 && MarginalTaxRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarginalTaxRateDt.Unit, Convert.ToDouble(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value)) : 0;

                        double value = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s24 * s52)) * s30;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Net Earnings";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;

                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s30 = MarginalTaxRateValues != null && MarginalTaxRateValues.Count > 0 && MarginalTaxRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarginalTaxRateDt.Unit, Convert.ToDouble(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value)) : 0;

                        double value = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) - (((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) * s30);
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Balance Sheet";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = true;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cash and Equivalents";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s49 - s33 - s34 - s35;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Total Current Assets";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    var TotalCurrentAssetsDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("total current assets"));
                    var TotalCurrentAssetsValues = TotalCurrentAssetsDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == TotalCurrentAssetsDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s50 = TotalCurrentAssetsValues != null && TotalCurrentAssetsValues.Count > 0 && TotalCurrentAssetsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalCurrentAssetsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalCurrentAssetsDt.Unit, Convert.ToDouble(TotalCurrentAssetsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s50 - s33 - s34 - s35;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Total Assets";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    var TotalAssetsDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("total assets"));
                    var TotalAssetsValues = TotalAssetsDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == TotalAssetsDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s51 = TotalAssetsValues != null && TotalAssetsValues.Count > 0 && TotalAssetsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalAssetsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalAssetsDt.Unit, Convert.ToDouble(TotalAssetsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s51 - s33 - s34 - s35;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Total Debt";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();
                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double value = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Shareholders Equity";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    var ShareholdersEquityDt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("shareholders equity"));
                    var ShareholdersEquityValues = ShareholdersEquityDt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == ShareholdersEquityDt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s53 = ShareholdersEquityValues != null && ShareholdersEquityValues.Count > 0 && ShareholdersEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(ShareholdersEquityDt.Unit, Convert.ToDouble(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s53 - s33 - s34 - s35;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cash Flow Statement";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = true;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cash Flow from Operations";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.FinancialStatementsPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "M";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    var CFFODt = CurrentSetupIpDatasList.Find(x => x.LineItem.ToLower().Contains("cash flow from operations"));
                    var CFFOValues = CFFODt != null ? CurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == CFFODt.Id).ToList() : null;

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s30 = MarginalTaxRateValues != null && MarginalTaxRateValues.Count > 0 && MarginalTaxRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarginalTaxRateDt.Unit, Convert.ToDouble(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s55 = CFFOValues != null && CFFOValues.Count > 0 && CFFOValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CFFOValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CFFODt.Unit, Convert.ToDouble(CFFOValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s47 = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) - (((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) * s30);
                        double s119 = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) - (((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) * s30);

                        double value = s55 + s119 - s47;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    #region Shares Outstanding & EPS: Post-Payout
                    CurrentSetupSoFilings = new CurrentSetupSoFillingsViewModel();
                    CurrentSetupSoDatasList = new List<CurrentSetupSoDatasViewModel>();
                    CurrentSetupSoFilings.StatementType = "Debt Ratios & Analysis: Post-Payout";

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Debt-to-Market Equity Ratio";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s60 = (PricePerShareValues != null && PricePerShareValues.Count > 0 && PricePerShareValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(PricePerShareValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(PricePerShareDt.Unit, Convert.ToDouble(PricePerShareValues.Find(x => x.Year == obj.Year).Value)) : 0) *
                                     (NoOfShareOutstandingBasicValues != null && NoOfShareOutstandingBasicValues.Count > 0 && NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(NoOfShareOutstandingBasicDt.Unit, Convert.ToDouble(NoOfShareOutstandingBasicValues.Find(x => x.Year == obj.Year).Value)) : 0);
                        double s62 = MarketvalueofInterestBearingDebtValues != null && MarketvalueofInterestBearingDebtValues.Count > 0 && MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarketvalueofInterestBearingDebtDt.Unit, Convert.ToDouble(MarketvalueofInterestBearingDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s62 * 100 / s60;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Debt-to-Book Equity Ratio";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s124 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s53 = ShareholdersEquityValues != null && ShareholdersEquityValues.Count > 0 && ShareholdersEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(ShareholdersEquityDt.Unit, Convert.ToDouble(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s125 = s53 - s33 - s34 - s35;
                        double value = s124 * 100 / s125;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "EBIT Interest Coverage (x)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "x";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s124 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s53 = ShareholdersEquityValues != null && ShareholdersEquityValues.Count > 0 && ShareholdersEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(ShareholdersEquityDt.Unit, Convert.ToDouble(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s125 = s53 - s33 - s34 - s35;
                        double value = s124 / s125;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "EBITDA Interest Coverage (x)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "x";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s124 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s53 = ShareholdersEquityValues != null && ShareholdersEquityValues.Count > 0 && ShareholdersEquityValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(ShareholdersEquityDt.Unit, Convert.ToDouble(ShareholdersEquityValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s125 = s53 - s33 - s34 - s35;
                        double value = s124 / s125;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Cash Flow from Operations / Total Debt";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "%";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s24 = CostOfDebtValues != null && CostOfDebtValues.Count > 0 && CostOfDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CostOfDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CostOfDebtDt.Unit, Convert.ToDouble(CostOfDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s29 = IIRateValues != null && IIRateValues.Count > 0 && IIRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(IIRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(IIRateDt.Unit, Convert.ToDouble(IIRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s30 = MarginalTaxRateValues != null && MarginalTaxRateValues.Count > 0 && MarginalTaxRateValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(MarginalTaxRateDt.Unit, Convert.ToDouble(MarginalTaxRateValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s33 = TODPA_Values != null && TODPA_Values.Count > 0 && TODPA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TODPA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TODPA_Dt.Unit, Convert.ToDouble(TODPA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s34 = OTDP_Values != null && OTDP_Values.Count > 0 && OTDP_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(OTDP_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(OTDP_Dt.Unit, Convert.ToDouble(OTDP_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s35 = SBA_Values != null && SBA_Values.Count > 0 && SBA_Values.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(SBA_Values.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(SBA_Dt.Unit, Convert.ToDouble(SBA_Values.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s40 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s41 = DepAmtValues != null && DepAmtValues.Count > 0 && DepAmtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(DepAmtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(DepAmtDt.Unit, Convert.ToDouble(DepAmtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s49 = CashAndEquivalentsValues != null && CashAndEquivalentsValues.Count > 0 && CashAndEquivalentsValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CashAndEquivalentsDt.Unit, Convert.ToDouble(CashAndEquivalentsValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s52 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;

                        double s55 = CFFOValues != null && CFFOValues.Count > 0 && CFFOValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(CFFOValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(CFFODt.Unit, Convert.ToDouble(CFFOValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s47 = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) - (((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) * s30);
                        double s119 = ((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) - (((s40 - s41 + ((s49 - s33 - s34 - s35) * s29)) - (s52 * s24)) * s30);

                        double s127 = s55 + s119 - s47;
                        double s124 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s127 * 100 / s124;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Total Debt / EBITDA (x)";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "x";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        double s112 = EBITDAValues != null && EBITDAValues.Count > 0 && EBITDAValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(EBITDAValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(EBITDADt.Unit, Convert.ToDouble(EBITDAValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double s124 = TotalDebtValues != null && TotalDebtValues.Count > 0 && TotalDebtValues.Find(x => x.Year == obj.Year) != null && !string.IsNullOrEmpty(TotalDebtValues.Find(x => x.Year == obj.Year).Value) ? GetMillionValue(TotalDebtDt.Unit, Convert.ToDouble(TotalDebtValues.Find(x => x.Year == obj.Year).Value)) : 0;
                        double value = s124 / s112;
                        CurrentSetupSoValues.Value = value.ToString("0.##");
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoDatasViewModelobj = new CurrentSetupSoDatasViewModel();
                    CurrentSetupSoValuesList = new List<CurrentSetupSoValuesViewModel>();
                    CurrentSetupSoDatasViewModelobj.LineItem = "Potential S&P Debt Rating";
                    CurrentSetupSoDatasViewModelobj.IsParentItem = false;
                    CurrentSetupSoDatasViewModelobj.Sequence = CurrentSetupSoDatasList.Count + 1;
                    CurrentSetupSoDatasViewModelobj.StatementTypeId = (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout;
                    CurrentSetupSoDatasViewModelobj.Unit = "";
                    CurrentSetupSoDatasViewModelobj.CurrentSetupId = tblCurrentSetup.Id;
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = new List<CurrentSetupSoValuesViewModel>();

                    foreach (CurrentSetupSoValuesViewModel obj in dummyCurrentSetupSoValuesViewModelList)
                    {
                        CurrentSetupSoValues = new CurrentSetupSoValuesViewModel();
                        CurrentSetupSoValues.Year = obj.Year;
                        CurrentSetupSoValues.Value = "AAA";
                        CurrentSetupSoValuesList.Add(CurrentSetupSoValues);
                    }
                    CurrentSetupSoDatasViewModelobj.CurrentSetupSoValuesVM = CurrentSetupSoValuesList;
                    CurrentSetupSoDatasList.Add(CurrentSetupSoDatasViewModelobj);

                    CurrentSetupSoFilings.CurrentSetupSoDatasViewModelVM = CurrentSetupSoDatasList;
                    CurrentSetupSoFilingsList.Add(CurrentSetupSoFilings);
                    #endregion

                    renderResult.StatusCode = 1;
                    renderResult.Message = "No issue found";
                    renderResult.Result = CurrentSetupSoFilingsList;
                    return Ok(renderResult);
                }
                else
                {
                    renderResult.StatusCode = 0;
                    renderResult.Message = "No data available for this user";
                    renderResult.Result = CurrentSetupSoFilingsList;
                    return Ok(renderResult);
                }
            }
            catch (Exception ex)
            {
                renderResult.StatusCode = 0;
                renderResult.Message = "Exception occured" + Convert.ToString(ex.Message);
                renderResult.Result = CurrentSetupSoFilingsList;
                return Ok(renderResult);
            }
        }

        [HttpGet]
        [Route("GetCurrentSetupSnapshotList/{UserId}")]
        public ActionResult GetCurrentSetupSnapshotList(long UserId)
        {
            try
            {
                CurrentSetupResultObject resultObject = new CurrentSetupResultObject();
                List<CurrentSetupSnapshotViewModel> CurrentSetupSnapshotObj = new List<CurrentSetupSnapshotViewModel>();
                var tblCurrentSetupSnapshotObj = iCurrentSetupSnapshot.FindBy(s => s.UserId == UserId).OrderByDescending(x => x.Id).ToList();
                if (tblCurrentSetupSnapshotObj == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    foreach (var obj in tblCurrentSetupSnapshotObj)
                    {
                        CurrentSetupSnapshotViewModel tempCurrentSetupSnapshotObj = mapper.Map<CurrentSetupSnapshot, CurrentSetupSnapshotViewModel>(obj);
                        CurrentSetupSnapshotObj.Add(tempCurrentSetupSnapshotObj);
                    }

                    return Ok(CurrentSetupSnapshotObj);
                }

            }
            catch (Exception ex)
            {
                return BadRequest(new { StatusCode = 0, message = Convert.ToString(ex.Message) });

            }
        }

        [HttpPost]
        [Route("SaveCurrentSetupSnapshot/{UserId}")]
        public ActionResult<Object> SaveCurrentSetupSnapshot([FromBody] CurrentSetupSnapshotViewModel model, long UserId)
        {
            try
            {
                CurrentSetupSnapshot currentSetupSnapshot = new CurrentSetupSnapshot
                {
                    Description = model.Description,
                    UserId = model.UserId
                };
                if (model.Id == 0)
                {
                    iCurrentSetupSnapshot.Add(currentSetupSnapshot);
                }
                iCurrentSetupSnapshot.Commit();
                if (currentSetupSnapshot.Id != null && currentSetupSnapshot.Id != 0)
                {
                    if (model.CurrentSetupSoFillingsVM != null && model.CurrentSetupSoFillingsVM.Count > 0)
                    {
                        //Start Code from here

                        foreach (CurrentSetupSoFillingsViewModel CurrentSetupSoFilingsObj in model.CurrentSetupSoFillingsVM)
                        {


                            if (CurrentSetupSoFilingsObj.CurrentSetupSoDatasViewModelVM != null && CurrentSetupSoFilingsObj.CurrentSetupSoDatasViewModelVM.Count > 0)
                            {
                                foreach (CurrentSetupSoDatasViewModel CurrentSetupSoDatasVMObj in CurrentSetupSoFilingsObj.CurrentSetupSoDatasViewModelVM)
                                {
                                    CurrentSetupSnapshotDatas tblCurrentSetupSnapshotDatasObj = new CurrentSetupSnapshotDatas();
                                    // map CurrentSetupSoDatasViewModel to CurrentSetupSnapshotDatas
                                    tblCurrentSetupSnapshotDatasObj = mapper.Map<CurrentSetupSoDatasViewModel, CurrentSetupSnapshotDatas>(CurrentSetupSoDatasVMObj);
                                    tblCurrentSetupSnapshotDatasObj.CurrentSetupSnapshotId = currentSetupSnapshot.Id;

                                    if (CurrentSetupSoDatasVMObj.CurrentSetupSoValuesVM != null && CurrentSetupSoDatasVMObj.CurrentSetupSoValuesVM.Count > 0)
                                    {
                                        tblCurrentSetupSnapshotDatasObj.CurrentSetupSnapshotValues = new List<CurrentSetupSnapshotValues>();
                                        foreach (CurrentSetupSoValuesViewModel CurrentSetupSoValues in CurrentSetupSoDatasVMObj.CurrentSetupSoValuesVM)
                                        {
                                            // map CurrentSetupSoValuesViewModel to CurrentSetupSnapshotValues
                                            CurrentSetupSnapshotValues ValueObj = mapper.Map<CurrentSetupSoValuesViewModel, CurrentSetupSnapshotValues>(CurrentSetupSoValues);
                                            tblCurrentSetupSnapshotDatasObj.CurrentSetupSnapshotValues.Add(ValueObj);
                                        }
                                    }

                                    if (tblCurrentSetupSnapshotDatasObj.Id == 0)
                                    {
                                        //Save Code                                        
                                        iCurrentSetupSnapshotDatas.Add(tblCurrentSetupSnapshotDatasObj);
                                        iCurrentSetupSnapshotDatas.Commit();
                                    }

                                }
                            }
                        }
                    }
                    return new
                    {
                        id = currentSetupSnapshot.Id,
                        result = "Current Setup Snapshot Created Sucessfully"
                    };
                }
                else
                {
                    return new
                    {
                        id = currentSetupSnapshot.Id,
                        result = "Data is not saved"
                    };
                }

            }
            catch (Exception ex)
            {
                return BadRequest(new { Message = "Invalid Entry", StatusCode = 400 });
            }
        }

        [HttpGet]
        [Route("GetCurrentSetupSnapshot/{Id}")]
        public ActionResult GetCurrentSetupSnapshot(long Id)
        {
            CurrentSetupSnapshotResult renderResult = new CurrentSetupSnapshotResult();
            CurrentSetupSnapshotFillingsViewModel FillingObj = new CurrentSetupSnapshotFillingsViewModel();
            List<CurrentSetupSnapshotFillingsViewModel> CurrentSetupSnapshotFillingsList = new List<CurrentSetupSnapshotFillingsViewModel>();
            List<CurrentSetupSnapshotDatasViewModel> CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
            CurrentSetupSnapshotDatasViewModel CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
            //CurrentSetupSnapshot tblCurrentSetupSnapshot = iCurrentSetupSnapshot.GetSingle(x => x.Id == Id);
            try
            {
                List<CurrentSetupSnapshotDatas> tblCurrentSetupSnapshotDatas = iCurrentSetupSnapshotDatas.FindBy(x => x.CurrentSetupSnapshotId == Id).ToList();

                if (tblCurrentSetupSnapshotDatas != null && tblCurrentSetupSnapshotDatas.Count > 0)
                {
                    // List<long> snapshotDataIds = tblCurrentSetupSnapshotDatas
                    //     .Select(m => m.Id) // Directly select Id since it's not nullable
                    //     .ToList();

                    // List<CurrentSetupSnapshotValues> tblCurrentSetupSnapshotValues = iCurrentSetupSnapshotValues.FindBy(x => snapshotDataIds.Contains(x.CurrentSetupSnapshotDatasId)).ToList();

                    List<CurrentSetupSnapshotValues> tblCurrentSetupSnapshotValues = iCurrentSetupSnapshotValues.FindBy(x => tblCurrentSetupSnapshotDatas.Any(m => m.Id == x.CurrentSetupSnapshotDatasId)).ToList();
                    List<CurrentSetupSnapshotValues> tempvalueList = new List<CurrentSetupSnapshotValues>();

                    #region Current Capital Structure 
                    List<CurrentSetupSnapshotDatas> DataList20 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.CurrentCapitalStructure).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Current Capital Structure";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList20)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                       
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Current Cost of Capital                    
                    List<CurrentSetupSnapshotDatas> DataList21 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.CurrentCostOfCapital).ToList();
                    // if(DataList21.Length < 1){
                    //     return NotFound("Cost of Capital not found")
                    // }
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Current Cost of Capital";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList21)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Leverage Ratios                    
                    List<CurrentSetupSnapshotDatas> DataList22 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.LeverageRatios).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Leverage Ratios";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList22)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Cash Balance: Post-Payout                    
                    List<CurrentSetupSnapshotDatas> DataList23 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.CashBalancePostPayout).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Cash Balance: Post-Payout";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList23)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Internal Valuation: Post-Payout                    
                    List<CurrentSetupSnapshotDatas> DataList24 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.InternalValuationPostPayout).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Internal Valuation: Post-Payout";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList24)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Current Payout Analysis                    
                    List<CurrentSetupSnapshotDatas> DataList25 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.CurrentPayoutAnalysis).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Current Payout Analysis";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList25)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Shares Outstanding & EPS: Post-Payout                    
                    List<CurrentSetupSnapshotDatas> DataList26 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.SharesOutstandingAndEPSPostPayout).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Shares Outstanding & EPS: Post-Payout";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList26)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Financial Statements: Post-Payout                    
                    List<CurrentSetupSnapshotDatas> DataList27 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPostPayout).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Financial Statements: Post-Payout";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList27)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion

                    #region Debt Ratios & Analysis: Post-Payout                    
                    List<CurrentSetupSnapshotDatas> DataList28 = tblCurrentSetupSnapshotDatas.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.DebtRatiosAndAnalysisPostPayout).ToList();
                    FillingObj = new CurrentSetupSnapshotFillingsViewModel();
                    CurrentSetupSnapshotDatasList = new List<CurrentSetupSnapshotDatasViewModel>();
                    FillingObj.StatementType = "Debt Ratios & Analysis: Post-Payout";
                    foreach (CurrentSetupSnapshotDatas DataObj in DataList28)
                    {
                        CurrentSetupSnapshotDatasVM = new CurrentSetupSnapshotDatasViewModel();
                        CurrentSetupSnapshotDatasVM = mapper.Map<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>(DataObj);
                        CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM = new List<CurrentSetupSnapshotValuesViewModel>();
                        tempvalueList = new List<CurrentSetupSnapshotValues>();
                        tempvalueList = tblCurrentSetupSnapshotValues != null && tblCurrentSetupSnapshotValues.Count > 0 ? tblCurrentSetupSnapshotValues.FindAll(x => x.CurrentSetupSnapshotDatasId == DataObj.Id).ToList() : null;
                        foreach (var obj in tempvalueList)
                        {
                            CurrentSetupSnapshotValuesViewModel tempValues = mapper.Map<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>(obj);
                            CurrentSetupSnapshotDatasVM.CurrentSetupSnapshotValuesVM.Add(tempValues);
                        }
                        CurrentSetupSnapshotDatasList.Add(CurrentSetupSnapshotDatasVM);
                    }
                    FillingObj.CurrentSetupSnapshotDatasVM = CurrentSetupSnapshotDatasList;
                    CurrentSetupSnapshotFillingsList.Add(FillingObj);
                    #endregion
                }

                renderResult.StatusCode = 1;
                renderResult.Message = "No issue found";
                renderResult.Result = CurrentSetupSnapshotFillingsList;
                return Ok(renderResult);
            }
            catch (Exception ex)
            {
                renderResult.StatusCode = 0;
                renderResult.Message = ex.Message;
                renderResult.Result = CurrentSetupSnapshotFillingsList;
                return Ok(renderResult);
            }
        }

        [HttpGet]
        [Route("ExportCurrentSetup/{UserId}/{Flag}")]
        public ActionResult ExportCurrentSetup(long UserId, int Flag)
        {
            CurrentSetup tblCurrentSetup = iCurrentSetup.GetSingle(x => x.UserId == UserId);
            List<CurrentSetupIpDatas> TblCurrentSetupIpDatasList = tblCurrentSetup != null && tblCurrentSetup.Id != null && tblCurrentSetup.Id != 0 ? iCurrentSetupIpDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList() : null;
            List<CurrentSetupIpValues> TblCurrentSetupIpValuesList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? iCurrentSetupIpValues.FindBy(x => TblCurrentSetupIpDatasList.Any(m => m.Id == x.CurrentSetupIpDatasId)).ToList() : null;

            string rootFolder = _hostingEnvironment.WebRootPath;
            string fileName = @"payout_policy.xlsx";
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            string formattedCustomObject = (string)null;

            using (ExcelPackage package = new ExcelPackage(file))
            {
                var wsHistory = package.Workbook.Worksheets["CurrentSetup"];
                //wsHistory.Cells[9, 4].Value = 15;

                if (tblCurrentSetup != null)
                {
                    int start;
                    int end;

                    if (!int.TryParse(tblCurrentSetup.CurrentYear, out start))
                    {

                        return NotFound($"Invalid format for CurrentYear: {tblCurrentSetup.CurrentYear}");
                    }

                    if (!int.TryParse(tblCurrentSetup.EndYear, out end))
                    {
                        return NotFound($"Invalid format for EndYear: {tblCurrentSetup.EndYear}");
                    }

                    start = Convert.ToInt32(tblCurrentSetup.CurrentYear);
                    end = Convert.ToInt32(tblCurrentSetup.EndYear);

                    int year = start;
                    int count = end - start + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[3, i + 2].Value = year;
                        //For set BackgroundColor
                        wsHistory.Cells[3, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsHistory.Cells[3, i + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                        /////////////////////////
                        year = year + 1;
                    }

                    List<CurrentSetupIpDatas> SOFDataList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? TblCurrentSetupIpDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing).OrderBy(x => x.Sequence).ToList() : null;
                    List<CurrentSetupIpValues> TempvalueList;
                    int row = 6;
                    foreach (CurrentSetupIpDatas Data in SOFDataList)
                    {
                        TempvalueList = new List<CurrentSetupIpValues>();
                        TempvalueList = TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0 ? TblCurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == Data.Id).ToList() : null;
                        if (Data.IsParentItem != true)
                            wsHistory = WorkSheetGeneration(wsHistory, row, start, end, count, TempvalueList);
                        row = row + 1;
                    }

                    ///////                   
                    List<CurrentSetupIpDatas> OIDataList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? TblCurrentSetupIpDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.OtherInputs).OrderBy(x => x.Sequence).ToList() : null;
                    row = row + 2;
                    foreach (CurrentSetupIpDatas Data in OIDataList)
                    {
                        TempvalueList = new List<CurrentSetupIpValues>();
                        TempvalueList = TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0 ? TblCurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == Data.Id).ToList() : null;
                        if (Data.IsParentItem != true)
                            wsHistory = WorkSheetGeneration(wsHistory, row, start, end, count, TempvalueList);
                        row = row + 1;
                    }


                    List<CurrentSetupIpDatas> CPPDataList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? TblCurrentSetupIpDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.CurrentPayoutPolicy).OrderBy(x => x.Sequence).ToList() : null;
                    row = row + 2;
                    foreach (CurrentSetupIpDatas Data in CPPDataList)
                    {
                        TempvalueList = new List<CurrentSetupIpValues>();
                        TempvalueList = TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0 ? TblCurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == Data.Id).ToList() : null;
                        if (Data.IsParentItem != true)
                            wsHistory = WorkSheetGeneration(wsHistory, row, start, end, count, TempvalueList);
                        row = row + 1;
                    }


                    List<CurrentSetupIpDatas> FSPPDataList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? TblCurrentSetupIpDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout).OrderBy(x => x.Sequence).ToList() : null;
                    row = row + 2;
                    foreach (CurrentSetupIpDatas Data in FSPPDataList)
                    {
                        TempvalueList = new List<CurrentSetupIpValues>();
                        TempvalueList = TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0 ? TblCurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == Data.Id).ToList() : null;
                        if (Data.IsParentItem != true)
                            wsHistory = WorkSheetGeneration(wsHistory, row, start, end, count, TempvalueList);
                        row = row + 1;
                    }

                    //For percentage value
                    for (int i = 0; i < count; i++)
                    {
                        ///Input
                        wsHistory.Cells[10, i + 2].Value = wsHistory.Cells[10, i + 2].Value != null ? Convert.ToDouble(wsHistory.Cells[10, i + 2].Value) / 100 : 0;
                        wsHistory.Cells[10, i + 2].Style.Numberformat.Format = "#.00%";

                        wsHistory.Cells[15, i + 2].Value = wsHistory.Cells[15, i + 2].Value != null ? Convert.ToDouble(wsHistory.Cells[15, i + 2].Value) / 100 : 0;
                        wsHistory.Cells[15, i + 2].Style.Numberformat.Format = "#.00%";

                        wsHistory.Cells[18, i + 2].Value = wsHistory.Cells[18, i + 2].Value != null ? Convert.ToDouble(wsHistory.Cells[18, i + 2].Value) / 100 : 0;
                        wsHistory.Cells[18, i + 2].Style.Numberformat.Format = "#.00%";

                        wsHistory.Cells[22, i + 2].Value = wsHistory.Cells[22, i + 2].Value != null ? Convert.ToDouble(wsHistory.Cells[22, i + 2].Value) / 100 : 0;
                        wsHistory.Cells[22, i + 2].Style.Numberformat.Format = "#.00%";

                        wsHistory.Cells[23, i + 2].Value = wsHistory.Cells[23, i + 2].Value != null ? Convert.ToDouble(wsHistory.Cells[23, i + 2].Value) / 100 : 0;
                        wsHistory.Cells[23, i + 2].Style.Numberformat.Format = "#.00%";

                        wsHistory.Cells[24, i + 2].Value = wsHistory.Cells[24, i + 2].Value != null ? Convert.ToDouble(wsHistory.Cells[24, i + 2].Value) / 100 : 0;
                        wsHistory.Cells[24, i + 2].Style.Numberformat.Format = "#.00%";

                    }


                    ///////  Formula /////// /////// /////// 
                    row = row + 4;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[7, i + 2] + "*" + wsHistory.Cells[8, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[12, i + 2] + "*" + wsHistory.Cells[13, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[17, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[56, i + 2] + "-" + wsHistory.Cells[43, i + 2] + "+" + wsHistory.Cells[21, i + 2];
                    }
                    row = row + 3;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[10, i + 2].ToString();
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[15, i + 2].ToString();
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[18, i + 2].ToString();
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }
                    //=C60*(C54/(C54+C55+C57))+C61*(C55/(C54+C55+C57))+C62*(C57/(C54+C55+C57))
                    //=C67*(C60/(C60+C61+C63))+C68*(C61/(C60+C61+C63))+C69*(C63/(C60+C61+C63))
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[60, i + 2] + "*" + "(" + wsHistory.Cells[54, i + 2] + "/" + "(" + wsHistory.Cells[54, i + 2] + "+" + wsHistory.Cells[55, i + 2] + "+" + wsHistory.Cells[57, i + 2] + "))+" + wsHistory.Cells[61, i + 2] + "*" + "(" + wsHistory.Cells[55, i + 2] + "/" + "(" + wsHistory.Cells[54, i + 2] + "+" + wsHistory.Cells[55, i + 2] + "+" + wsHistory.Cells[57, i + 2] + "))+" + wsHistory.Cells[62, i + 2] + "*" + "(" + wsHistory.Cells[57, i + 2] + "/" + "(" + wsHistory.Cells[54, i + 2] + "+" + wsHistory.Cells[55, i + 2] + "+" + wsHistory.Cells[57, i + 2] + "))";
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[60, i + 2] + "*" + "(" + wsHistory.Cells[54, i + 2] + "/" + "(" + wsHistory.Cells[54, i + 2] + "+" + wsHistory.Cells[55, i + 2] + "+" + wsHistory.Cells[57, i + 2] + "))+" + wsHistory.Cells[61, i + 2] + "*" + "(" + wsHistory.Cells[55, i + 2] + "/" + "(" + wsHistory.Cells[54, i + 2] + "+" + wsHistory.Cells[55, i + 2] + "+" + wsHistory.Cells[57, i + 2] + "))+" + wsHistory.Cells[62, i + 2] + "*" + "(" + wsHistory.Cells[57, i + 2] + "/" + "(" + wsHistory.Cells[54, i + 2] + "+" + wsHistory.Cells[55, i + 2] + "+" + wsHistory.Cells[57, i + 2] + "))*(1-" + wsHistory.Cells[24, i + 2] + ")";
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }

                    row = row + 3;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[56, i + 2] + "/" + wsHistory.Cells[54, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[56, i + 2] + "/SUM(" + wsHistory.Cells[54, i + 2] + ":" + wsHistory.Cells[56, i + 2] + ")";
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }

                    row = row + 3;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[43, i + 2] + "-" + wsHistory.Cells[83, i + 2] + "-" + wsHistory.Cells[88, i + 2] + "-" + wsHistory.Cells[93, i + 2];
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[71, i + 2] + "-" + wsHistory.Cells[21, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }

                    row = row + 11;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[27, i + 2].ToString();
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[83, i + 2] + "/" + wsHistory.Cells[8, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[83, i + 2] + "/" + wsHistory.Cells[41, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[84, i + 2] + "/" + wsHistory.Cells[7, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }
                    row = row + 2;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[28, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[88, i + 2] + "/" + wsHistory.Cells[8, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[88, i + 2] + "/" + wsHistory.Cells[41, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[89, i + 2] + "/" + wsHistory.Cells[7, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }
                    row = row + 2;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[29, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[93, i + 2] + "/" + wsHistory.Cells[7, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 3;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[8, i + 2] + "-" + wsHistory.Cells[94, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[9, i + 2] + "-" + wsHistory.Cells[94, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[112, i + 2] + "/" + wsHistory.Cells[97, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[112, i + 2] + "/" + wsHistory.Cells[98, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 4;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[33, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[34, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[35, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[114, i + 2] + "*" + wsHistory.Cells[23, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[105, i + 2] + "-" + wsHistory.Cells[106, i + 2] + "+" + wsHistory.Cells[107, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[117, i + 2] + "*" + wsHistory.Cells[18, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[108, i + 2] + "-" + wsHistory.Cells[109, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[110, i + 2] + "*" + wsHistory.Cells[24, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[110, i + 2] + "-" + wsHistory.Cells[111, i + 2];
                    }
                    row = row + 2;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[71, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[44, i + 2] + "+" + wsHistory.Cells[114, i + 2] + "-" + wsHistory.Cells[43, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[45, i + 2] + "+" + wsHistory.Cells[114, i + 2] + "-" + wsHistory.Cells[43, i + 2];
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[46, i + 2].ToString();
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[47, i + 2] + "+" + wsHistory.Cells[114, i + 2] + "-" + wsHistory.Cells[43, i + 2];
                    }
                    row = row + 2;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[49, i + 2] + "+" + wsHistory.Cells[112, i + 2] + "-" + wsHistory.Cells[41, i + 2];
                    }
                    row = row + 3;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[56, i + 2] + "/" + wsHistory.Cells[54, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[117, i + 2] + "/" + wsHistory.Cells[118, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[108, i + 2] + "/" + wsHistory.Cells[109, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[105, i + 2] + "/" + wsHistory.Cells[109, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[120, i + 2] + "/" + wsHistory.Cells[117, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                    }

                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Formula = wsHistory.Cells[117, i + 2] + "/" + wsHistory.Cells[105, i + 2];
                        wsHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00";
                    }
                    row = row + 1;
                    for (int i = 0; i < count; i++)
                    {
                        wsHistory.Cells[row, i + 2].Value = "AAA";

                    }
                }

                ExcelPackage excelPackage = new ExcelPackage();
                if (Flag == 0)
                {
                    excelPackage.Workbook.Worksheets.Add("CurrentSetup", wsHistory);
                }
                else if (Flag == 1)
                {
                    tblCurrentSetup = iCurrentSetup.GetSingle(x => x.UserId == UserId);
                    List<PayoutPolicy_ScenarioDatas> DataList = tblCurrentSetup != null && tblCurrentSetup.Id != null && tblCurrentSetup.Id != 0 ? iPayoutPolicy_ScenarioDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList() : null;
                    List<PayoutPolicy_ScenarioValues> ValuesList = DataList != null && DataList.Count > 0 ? iPayoutPolicy_ScenarioValues.FindBy(x => DataList.Any(m => m.Id == x.PayoutPolicy_ScenarioDatasId)).ToList() : null;

                    var saHistory = package.Workbook.Worksheets["NewScenario"];
                    if (tblCurrentSetup != null)
                    {
                        int start;
                        int end;

                        if (!int.TryParse(tblCurrentSetup.CurrentYear, out start))
                        {
                            // Handle the case where CurrentYear is not a valid integer
                            return BadRequest($"Invalid format for CurrentYear: {tblCurrentSetup.CurrentYear}");
                        }

                        if (!int.TryParse(tblCurrentSetup.EndYear, out end))
                        {
                            // Handle the case where EndYear is not a valid integer
                            return BadRequest($"Invalid format for EndYear: {tblCurrentSetup.EndYear}");
                        }

                        start = Convert.ToInt32(tblCurrentSetup.CurrentYear);
                        end = Convert.ToInt32(tblCurrentSetup.EndYear);
                        int year = start;
                        int count = end - start + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[9, i + 2].Value = year;
                            //For set BackgroundColor
                            saHistory.Cells[9, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            saHistory.Cells[9, i + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            /////////////////////////
                            year = year + 1;
                        }

                        List<PayoutPolicy_ScenarioDatas> InputDataList = DataList != null && DataList.Count > 0 ? DataList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.Inputs).OrderBy(x => x.Sequence).ToList() : null;
                        List<PayoutPolicy_ScenarioValues> TempvalueList;
                        int row = 27;
                        foreach (PayoutPolicy_ScenarioDatas Data in InputDataList)
                        {
                            TempvalueList = new List<PayoutPolicy_ScenarioValues>();
                            TempvalueList = ValuesList != null && ValuesList.Count > 0 ? ValuesList.FindAll(x => x.PayoutPolicy_ScenarioDatasId == Data.Id).ToList() : null;
                            if (Data.IsParentItem != true)
                                saHistory = WorkSheet_SA_Generation(saHistory, row, start, end, count, TempvalueList);
                            row = row + 1;
                        }

                        List<PayoutPolicy_ScenarioDatas> NPPDataList = DataList != null && DataList.Count > 0 ? DataList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.NewPayoutPolicy).OrderBy(x => x.Sequence).ToList() : null;
                        row = row + 2;
                        foreach (PayoutPolicy_ScenarioDatas Data in NPPDataList)
                        {
                            TempvalueList = new List<PayoutPolicy_ScenarioValues>();
                            TempvalueList = ValuesList != null && ValuesList.Count > 0 ? ValuesList.FindAll(x => x.PayoutPolicy_ScenarioDatasId == Data.Id).ToList() : null;
                            if (Data.IsParentItem != true)
                                saHistory = WorkSheet_SA_Generation(saHistory, row, start, end, count, TempvalueList);
                            row = row + 1;
                        }

                        List<PayoutPolicy_ScenarioDatas> FSPPDataList = DataList != null && DataList.Count > 0 ? DataList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout).OrderBy(x => x.Sequence).ToList() : null;
                        row = row + 2;
                        foreach (PayoutPolicy_ScenarioDatas Data in FSPPDataList)
                        {
                            TempvalueList = new List<PayoutPolicy_ScenarioValues>();
                            TempvalueList = ValuesList != null && ValuesList.Count > 0 ? ValuesList.FindAll(x => x.PayoutPolicy_ScenarioDatasId == Data.Id).ToList() : null;
                            if (Data.IsParentItem != true)
                                saHistory = WorkSheet_SA_Generation(saHistory, row, start, end, count, TempvalueList);
                            row = row + 1;
                        }

                        //For percentage value
                        for (int i = 0; i < count; i++)
                        {
                            ///Input
                            saHistory.Cells[28, i + 2].Value = saHistory.Cells[28, i + 2].Value != null ? Convert.ToDouble(saHistory.Cells[28, i + 2].Value) / 100 : 0;
                            saHistory.Cells[28, i + 2].Style.Numberformat.Format = "#.00%";

                            saHistory.Cells[30, i + 2].Value = saHistory.Cells[30, i + 2].Value != null ? Convert.ToDouble(saHistory.Cells[30, i + 2].Value) / 100 : 0;
                            saHistory.Cells[30, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        //Formula
                        row = row + 4;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[54, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[55, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[28, i + 2] + "*" + saHistory.Cells[60, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[62, i + 2] + "-" + "CurrentSetup!" + wsHistory.Cells[56, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[62, i + 2] + "-" + saHistory.Cells[79, i + 2];
                        }
                        row = row + 3;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[63, i + 2] + "+((" + saHistory.Cells[64, i + 2] + "/" + saHistory.Cells[60, i + 2] + ")*(" + "CurrentSetup!" + wsHistory.Cells[63, i + 2] + "-" + saHistory.Cells[69, i + 2] + "))+((" + saHistory.Cells[61, i + 2] + "/" + saHistory.Cells[60, i + 2] + ")*(" + "CurrentSetup!" + wsHistory.Cells[63, i + 2] + "-" + saHistory.Cells[68, i + 2] + "))";
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[61, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[30, i + 2] + "";
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[63, 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[67, i + 2] + "*(" + saHistory.Cells[60, i + 2] + "/(" + saHistory.Cells[60, i + 2] + "+" + saHistory.Cells[61, i + 2] + "+" + saHistory.Cells[64, i + 2] + "))+" + saHistory.Cells[68, i + 2] + "*(" + saHistory.Cells[61, i + 2] + "/(" + saHistory.Cells[60, i + 2] + "+" + saHistory.Cells[61, i + 2] + "+" + saHistory.Cells[64, i + 2] + "))+" + saHistory.Cells[69, i + 2] + "*(" + saHistory.Cells[64, i + 2] + "/(" + saHistory.Cells[60, i + 2] + "+" + saHistory.Cells[61, i + 2] + "+" + saHistory.Cells[64, i + 2] + "))*(" + 1 + "-" + "CurrentSetup!" + wsHistory.Cells[24, i + 2] + ")";
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 3;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[62, i + 2] + "/" + saHistory.Cells[60, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[62, i + 2] + "/SUM(" + saHistory.Cells[60, i + 2] + ":" + saHistory.Cells[62, i + 2] + ")";
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 3;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[43, 2] + "+" + saHistory.Cells[63, i + 2] + "-" + saHistory.Cells[90, i + 2] + "-" + saHistory.Cells[95, i + 2] + "-" + saHistory.Cells[100, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[78, i + 2] + "-" + "CurrentSetup!" + wsHistory.Cells[21, i + 2];
                        }
                        row = row + 11;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[33, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[90, i + 2] + "/" + "CurrentSetup!" + wsHistory.Cells[8, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[90, i + 2] + "/" + saHistory.Cells[47, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[91, i + 2] + "/" + "CurrentSetup!" + wsHistory.Cells[7, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 2;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[34, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[95, i + 2] + "/" + "CurrentSetup!" + wsHistory.Cells[8, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[95, i + 2] + "/" + saHistory.Cells[47, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[96, i + 2] + "/" + "CurrentSetup!" + wsHistory.Cells[7, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 2;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[35, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[100, i + 2] + "/" + "CurrentSetup!" + wsHistory.Cells[7, i + 2];
                        }
                        row = row + 3;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[8, i + 2] + "-" + saHistory.Cells[101, i + 2];

                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = "CurrentSetup!" + wsHistory.Cells[9, i + 2] + "-" + saHistory.Cells[101, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[119, i + 2] + "/" + saHistory.Cells[104, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[119, i + 2] + "/" + saHistory.Cells[105, i + 2];
                        }
                        row = row + 4;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[39, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[40, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[41, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[121, i + 2] + "*" + "CurrentSetup!" + wsHistory.Cells[23, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[112, i + 2] + "-" + saHistory.Cells[113, i + 2] + "+" + saHistory.Cells[114, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[124, i + 2] + "*" + saHistory.Cells[30, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[115, i + 2] + "+" + saHistory.Cells[116, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[117, i + 2] + "*" + "CurrentSetup!" + wsHistory.Cells[24, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[117, i + 2] + "-" + saHistory.Cells[118, i + 2];
                        }
                        row = row + 2;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[78, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[50, i + 2] + "+" + saHistory.Cells[121, i + 2] + "-" + saHistory.Cells[49, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[51, i + 2] + "+" + saHistory.Cells[121, i + 2] + "-" + saHistory.Cells[49, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[62, i + 2] + "";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[53, i + 2] + "+" + saHistory.Cells[121, i + 2] + "-" + saHistory.Cells[49, i + 2];
                        }
                        row = row + 2;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[55, i + 2] + "+" + saHistory.Cells[119, i + 2] + "-" + saHistory.Cells[47, i + 2];
                        }
                        row = row + 3;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[62, i + 2] + "/" + saHistory.Cells[60, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[124, i + 2] + "/" + saHistory.Cells[125, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[115, i + 2] + "/" + saHistory.Cells[116, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[112, i + 2] + "/" + saHistory.Cells[116, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[127, i + 2] + "/" + saHistory.Cells[124, i + 2];
                            saHistory.Cells[row, i + 2].Style.Numberformat.Format = "#.00%";
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Formula = saHistory.Cells[124, i + 2] + "/" + saHistory.Cells[112, i + 2];
                        }
                        row = row + 1;
                        for (int i = 0; i < count; i++)
                        {
                            saHistory.Cells[row, i + 2].Value = "AA";
                        }
                    }



                    excelPackage.Workbook.Worksheets.Add("CurrentSetup", wsHistory);
                    excelPackage.Workbook.Worksheets.Add("ScenarioAnalysis", saHistory);
                }

                //package.Save();
                ExcelPackage epOut = excelPackage;
                byte[] myStream = epOut.GetAsByteArray();
                var inputAsString = Convert.ToBase64String(myStream);
                formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);
                return Ok(formattedCustomObject);
            }

        }
        
        private static ExcelWorksheet WorkSheetGeneration(ExcelWorksheet wsHistory, int r, int start, int end, int count, List<CurrentSetupIpValues> valList)
        {
            int year = start;
            for (int i = 0; i < count; i++)
            {
                wsHistory.Cells[r, i + 2].Value = valList != null && valList.Count > 0 && valList.Find(x => x.Year == year.ToString()) != null && !string.IsNullOrEmpty(valList.Find(x => x.Year == year.ToString()).Value) ? Convert.ToDouble(valList.Find(x => x.Year == year.ToString()).Value) : 0;
                year = year + 1;
            }
            return wsHistory;
        }
        private static ExcelWorksheet WorkSheet_SA_Generation(ExcelWorksheet wsHistory, int r, int start, int end, int count, List<PayoutPolicy_ScenarioValues> valList)
        {
            int year = start;
            for (int i = 0; i < count; i++)
            {
                wsHistory.Cells[r, i + 2].Value = valList != null && valList.Count > 0 && valList.Find(x => x.Year == year.ToString()) != null && !string.IsNullOrEmpty(valList.Find(x => x.Year == year.ToString()).Value) ? Convert.ToDouble(valList.Find(x => x.Year == year.ToString()).Value) : 0;
                year = year + 1;
            }
            return wsHistory;
        }
        #endregion


        #region Scenario Analysis
        [HttpGet]
        [Route("GetPayoutScenarioAnalysis/{UserId}")]
        public ActionResult GetPayoutScenarioAnalysis(long UserId)
        {
            PayoutPolicy_ScenarioResult result = new PayoutPolicy_ScenarioResult();
            result.IsSaved = false;
            try
            {
                result.ReportName = "Scenario Inputs:";
                List<PayoutPolicy_ScenarioDatas> scenarioDatasListObj = new List<PayoutPolicy_ScenarioDatas>();
                List<PayoutPolicy_ScenarioFillingsViewModel> filingList = new List<PayoutPolicy_ScenarioFillingsViewModel>();
                PayoutPolicy_ScenarioFillingsViewModel PayoutPolicy_ScenarioFillingsVM;
                List<PayoutPolicy_ScenarioDatasViewModel> ScenarioDatasVMList;
                List<PayoutPolicy_ScenarioValuesViewModel> ScenarioValuesVMList;

                //get current setup
                CurrentSetup tblCurrentSetup = iCurrentSetup.GetSingle(x => x.UserId == UserId);
                //check its own table
                scenarioDatasListObj = tblCurrentSetup != null ? iPayoutPolicy_ScenarioDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList() : null;
                if (scenarioDatasListObj != null && scenarioDatasListObj.Count > 0)
                {
                    result.IsSaved = true;
                    List<PayoutPolicy_ScenarioValues> TblCurrentSetupIpValuesList = scenarioDatasListObj != null && scenarioDatasListObj.Count > 0 ? iPayoutPolicy_ScenarioValues.FindBy(x => scenarioDatasListObj.Any(t => t.Id == x.PayoutPolicy_ScenarioDatasId)).ToList() : null;
                    if (TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0)
                    {
                        PayoutPolicy_ScenarioDatasViewModel CurrentSetupDatasVMObj;
                        List<PayoutPolicy_ScenarioDatasViewModel> CurrentSetupDatasVMListObj = new List<PayoutPolicy_ScenarioDatasViewModel>();
                        PayoutPolicy_ScenarioValuesViewModel CurrentSetupValuesVMObj;
                        foreach (PayoutPolicy_ScenarioDatas obj in scenarioDatasListObj)
                        {
                            // map data to dataVM
                            CurrentSetupDatasVMObj = new PayoutPolicy_ScenarioDatasViewModel();
                            CurrentSetupDatasVMObj = mapper.Map<PayoutPolicy_ScenarioDatas, PayoutPolicy_ScenarioDatasViewModel>(obj);
                            CurrentSetupDatasVMObj.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            List<PayoutPolicy_ScenarioValues> tempValuesList = TblCurrentSetupIpValuesList.FindAll(x => x.PayoutPolicy_ScenarioDatasId == obj.Id).ToList();
                            foreach (PayoutPolicy_ScenarioValues valueObj in tempValuesList)
                            {
                                CurrentSetupValuesVMObj = new PayoutPolicy_ScenarioValuesViewModel();
                                CurrentSetupValuesVMObj = mapper.Map<PayoutPolicy_ScenarioValues, PayoutPolicy_ScenarioValuesViewModel>(valueObj);
                                CurrentSetupDatasVMObj.PayoutPolicy_ScenarioValuesVM.Add(CurrentSetupValuesVMObj);
                            }
                            CurrentSetupDatasVMListObj.Add(CurrentSetupDatasVMObj);
                        }


                        // Inputs
                        PayoutPolicy_ScenarioFillingsVM = new PayoutPolicy_ScenarioFillingsViewModel();
                        ScenarioDatasVMList = new List<PayoutPolicy_ScenarioDatasViewModel>();
                        PayoutPolicy_ScenarioFillingsVM.StatementType = "Inputs";
                        ScenarioDatasVMList = CurrentSetupDatasVMListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.Inputs).ToList();
                        PayoutPolicy_ScenarioFillingsVM.PayoutPolicy_ScenarioDatasVM = ScenarioDatasVMList;
                        filingList.Add(PayoutPolicy_ScenarioFillingsVM);

                        // New Payout Policy
                        PayoutPolicy_ScenarioFillingsVM = new PayoutPolicy_ScenarioFillingsViewModel();
                        ScenarioDatasVMList = new List<PayoutPolicy_ScenarioDatasViewModel>();
                        PayoutPolicy_ScenarioFillingsVM.StatementType = "New Payout Policy";
                        ScenarioDatasVMList = CurrentSetupDatasVMListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.NewPayoutPolicy).ToList();
                        PayoutPolicy_ScenarioFillingsVM.PayoutPolicy_ScenarioDatasVM = ScenarioDatasVMList;
                        filingList.Add(PayoutPolicy_ScenarioFillingsVM);

                        // Financial Statements: Pre-Payout
                        PayoutPolicy_ScenarioFillingsVM = new PayoutPolicy_ScenarioFillingsViewModel();
                        ScenarioDatasVMList = new List<PayoutPolicy_ScenarioDatasViewModel>();
                        PayoutPolicy_ScenarioFillingsVM.StatementType = "Financial Statements: Pre-Payout";
                        ScenarioDatasVMList = CurrentSetupDatasVMListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout).ToList();
                        PayoutPolicy_ScenarioFillingsVM.PayoutPolicy_ScenarioDatasVM = ScenarioDatasVMList;
                        filingList.Add(PayoutPolicy_ScenarioFillingsVM);

                        result.FilingResult = filingList;
                    }
                }
                else
                {
                    List<PayoutPolicy_ScenarioValuesViewModel> dummyPayoutPolicy_ScenarioViewModelList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                    PayoutPolicy_ScenarioValuesViewModel dummyPayoutPolicy_ScenarioValuesViewModel = new PayoutPolicy_ScenarioValuesViewModel();

                    if (tblCurrentSetup != null)
                    {

                        // find if inputdatas exist or not
                        List<CurrentSetupIpDatas> TblCurrentSetupIpDatasList = tblCurrentSetup != null && tblCurrentSetup.Id != null && tblCurrentSetup.Id != 0 ? iCurrentSetupIpDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList() : null;
                        if (TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0)
                        {
                            PayoutPolicy_ScenarioDatasViewModel ScenarioDatasVM;
                            PayoutPolicy_ScenarioValuesViewModel ScenarioValuesVM;
                            //List<PayoutPolicy_ScenarioValuesViewModel> dummyPayoutPolicy_ScenarioViewModelList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            int strt = Convert.ToInt32(tblCurrentSetup.CurrentYear);
                            int End = Convert.ToInt32(tblCurrentSetup.EndYear);
                            for (int i = strt; i <= End; i++)
                            {
                                dummyPayoutPolicy_ScenarioValuesViewModel = new PayoutPolicy_ScenarioValuesViewModel();
                                dummyPayoutPolicy_ScenarioValuesViewModel.Year = Convert.ToString(strt);
                                dummyPayoutPolicy_ScenarioValuesViewModel.Value = "";
                                dummyPayoutPolicy_ScenarioViewModelList.Add(dummyPayoutPolicy_ScenarioValuesViewModel);
                                strt = strt + 1;
                            }


                            #region Inputs
                            // Inputs
                            PayoutPolicy_ScenarioFillingsVM = new PayoutPolicy_ScenarioFillingsViewModel();
                            ScenarioDatasVMList = new List<PayoutPolicy_ScenarioDatasViewModel>();
                            PayoutPolicy_ScenarioFillingsVM.StatementType = "Inputs";

                            // Target Capital Structure
                            ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                            ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            ScenarioDatasVM.LineItem = "Target Capital Structure";
                            ScenarioDatasVM.IsParentItem = true;
                            ScenarioDatasVM.Sequence = 1;
                            ScenarioDatasVM.StatementTypeId = (int)StatementTypeEnum.Inputs;
                            ScenarioDatasVM.Unit = "";
                            ScenarioDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            foreach (PayoutPolicy_ScenarioValuesViewModel obj in dummyPayoutPolicy_ScenarioViewModelList)
                            {
                                ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                ScenarioValuesVM.Year = obj.Year;
                                ScenarioValuesVM.PayoutPolicy_ScenarioDatasId = 0;
                                ScenarioValuesVM.Value = "";
                                ScenarioValuesVMList.Add(ScenarioValuesVM);
                            }
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = ScenarioValuesVMList;
                            ScenarioDatasVMList.Add(ScenarioDatasVM);

                            //Target Debt-to-Equity (D/E) Ratio
                            ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                            ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            ScenarioDatasVM.LineItem = "Target Debt-to-Equity (D/E) Ratio";
                            ScenarioDatasVM.IsParentItem = false;
                            ScenarioDatasVM.Sequence = ScenarioDatasVMList.Count + 1;
                            ScenarioDatasVM.StatementTypeId = (int)StatementTypeEnum.Inputs;
                            ScenarioDatasVM.Unit = "%";
                            ScenarioDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            foreach (PayoutPolicy_ScenarioValuesViewModel obj in dummyPayoutPolicy_ScenarioViewModelList)
                            {
                                ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                ScenarioValuesVM.Year = obj.Year;
                                ScenarioValuesVM.Value = "";
                                ScenarioValuesVMList.Add(ScenarioValuesVM);
                            }
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = ScenarioValuesVMList;
                            ScenarioDatasVMList.Add(ScenarioDatasVM);

                            //New Cost of Capital
                            ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                            ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            ScenarioDatasVM.LineItem = "New Cost of Capital";
                            ScenarioDatasVM.IsParentItem = true;
                            ScenarioDatasVM.Sequence = ScenarioDatasVMList.Count + 1;
                            ScenarioDatasVM.StatementTypeId = (int)StatementTypeEnum.Inputs;
                            ScenarioDatasVM.Unit = "";
                            ScenarioDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            foreach (PayoutPolicy_ScenarioValuesViewModel obj in dummyPayoutPolicy_ScenarioViewModelList)
                            {
                                ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                ScenarioValuesVM.Year = obj.Year;
                                ScenarioValuesVM.PayoutPolicy_ScenarioDatasId = 0;
                                ScenarioValuesVM.Value = "";
                                ScenarioValuesVMList.Add(ScenarioValuesVM);
                            }
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = ScenarioValuesVMList;
                            ScenarioDatasVMList.Add(ScenarioDatasVM);

                            //Cost of Debt (rD)
                            ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                            ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            ScenarioDatasVM.LineItem = "Cost of Debt (rD)";
                            ScenarioDatasVM.IsParentItem = false;
                            ScenarioDatasVM.Sequence = ScenarioDatasVMList.Count + 1;
                            ScenarioDatasVM.StatementTypeId = (int)StatementTypeEnum.Inputs;
                            ScenarioDatasVM.Unit = "%";
                            ScenarioDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            foreach (PayoutPolicy_ScenarioValuesViewModel obj in dummyPayoutPolicy_ScenarioViewModelList)
                            {
                                ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                ScenarioValuesVM.Year = obj.Year;
                                ScenarioValuesVM.PayoutPolicy_ScenarioDatasId = 0;
                                ScenarioValuesVM.Value = "";
                                ScenarioValuesVMList.Add(ScenarioValuesVM);
                            }
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = ScenarioValuesVMList;
                            ScenarioDatasVMList.Add(ScenarioDatasVM);


                            PayoutPolicy_ScenarioFillingsVM.PayoutPolicy_ScenarioDatasVM = ScenarioDatasVMList;
                            filingList.Add(PayoutPolicy_ScenarioFillingsVM);

                            #endregion

                            #region New Payout Policy
                            // New Payout Policy
                            PayoutPolicy_ScenarioFillingsVM = new PayoutPolicy_ScenarioFillingsViewModel();
                            ScenarioDatasVMList = new List<PayoutPolicy_ScenarioDatasViewModel>();
                            PayoutPolicy_ScenarioFillingsVM.StatementType = "New Payout Policy";

                            // Total Ongoing Dividend Payout -Annual
                            ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                            ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            ScenarioDatasVM.LineItem = "Total Ongoing Dividend Payout -Annual";
                            ScenarioDatasVM.IsParentItem = false;
                            ScenarioDatasVM.Sequence = 1;
                            ScenarioDatasVM.StatementTypeId = (int)StatementTypeEnum.NewPayoutPolicy;
                            ScenarioDatasVM.Unit = "M";
                            ScenarioDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            foreach (PayoutPolicy_ScenarioValuesViewModel obj in dummyPayoutPolicy_ScenarioViewModelList)
                            {
                                ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                ScenarioValuesVM.Year = obj.Year;
                                ScenarioValuesVM.PayoutPolicy_ScenarioDatasId = 0;
                                ScenarioValuesVM.Value = "";
                                ScenarioValuesVMList.Add(ScenarioValuesVM);
                            }
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = ScenarioValuesVMList;
                            ScenarioDatasVMList.Add(ScenarioDatasVM);

                            //One Time Dividend Payout
                            ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                            ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            ScenarioDatasVM.LineItem = "One Time Dividend Payout";
                            ScenarioDatasVM.IsParentItem = false;
                            ScenarioDatasVM.Sequence = ScenarioDatasVMList.Count + 1;
                            ScenarioDatasVM.StatementTypeId = (int)StatementTypeEnum.NewPayoutPolicy;
                            ScenarioDatasVM.Unit = "$000s";
                            ScenarioDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            foreach (PayoutPolicy_ScenarioValuesViewModel obj in dummyPayoutPolicy_ScenarioViewModelList)
                            {
                                ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                ScenarioValuesVM.Year = obj.Year;
                                ScenarioValuesVM.PayoutPolicy_ScenarioDatasId = 0;
                                ScenarioValuesVM.Value = "";
                                ScenarioValuesVMList.Add(ScenarioValuesVM);
                            }
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = ScenarioValuesVMList;
                            ScenarioDatasVMList.Add(ScenarioDatasVM);

                            //Stock Buyback Amount
                            ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                            ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            ScenarioDatasVM.LineItem = "Stock Buyback Amount";
                            ScenarioDatasVM.IsParentItem = false;
                            ScenarioDatasVM.Sequence = ScenarioDatasVMList.Count + 1;
                            ScenarioDatasVM.StatementTypeId = (int)StatementTypeEnum.NewPayoutPolicy;
                            ScenarioDatasVM.Unit = "$000s";
                            ScenarioDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                            foreach (PayoutPolicy_ScenarioValuesViewModel obj in dummyPayoutPolicy_ScenarioViewModelList)
                            {
                                ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                ScenarioValuesVM.Year = obj.Year;
                                ScenarioValuesVM.PayoutPolicy_ScenarioDatasId = 0;
                                ScenarioValuesVM.Value = "";
                                ScenarioValuesVMList.Add(ScenarioValuesVM);
                            }
                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = ScenarioValuesVMList;
                            ScenarioDatasVMList.Add(ScenarioDatasVM);

                            PayoutPolicy_ScenarioFillingsVM.PayoutPolicy_ScenarioDatasVM = ScenarioDatasVMList;
                            filingList.Add(PayoutPolicy_ScenarioFillingsVM);


                            #endregion


                            #region Financial Statements: Pre-Payout


                            PayoutPolicy_ScenarioFillingsVM = new PayoutPolicy_ScenarioFillingsViewModel();
                            ScenarioDatasVMList = new List<PayoutPolicy_ScenarioDatasViewModel>();
                            List<CurrentSetupIpDatas> tempCurrentIPDatasList = new List<CurrentSetupIpDatas>();
                            tempCurrentIPDatasList = TblCurrentSetupIpDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout).OrderBy(x => x.Sequence).ToList();
                            PayoutPolicy_ScenarioFillingsVM.StatementType = "Financial Statements: Pre-Payout";
                            if (tempCurrentIPDatasList != null && tempCurrentIPDatasList.Count > 0)
                            {
                                foreach (CurrentSetupIpDatas datasObj in tempCurrentIPDatasList)
                                {
                                    ScenarioDatasVM = new PayoutPolicy_ScenarioDatasViewModel();
                                    ScenarioValuesVMList = new List<PayoutPolicy_ScenarioValuesViewModel>();
                                    ScenarioDatasVM = mapper.Map<CurrentSetupIpDatas, PayoutPolicy_ScenarioDatasViewModel>(datasObj);
                                    ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM = new List<PayoutPolicy_ScenarioValuesViewModel>();
                                    List<CurrentSetupIpValues> TblCurrentSetupIpValuesList = TblCurrentSetupIpDatasList != null && TblCurrentSetupIpDatasList.Count > 0 ? iCurrentSetupIpValues.FindBy(x => TblCurrentSetupIpDatasList.Any(t => t.Id == x.CurrentSetupIpDatasId)).ToList() : null;
                                    ScenarioDatasVM.Id = 0;

                                    if (TblCurrentSetupIpValuesList != null && TblCurrentSetupIpValuesList.Count > 0)
                                    {
                                        List<CurrentSetupIpValues> tempValuesList = TblCurrentSetupIpValuesList.FindAll(x => x.CurrentSetupIpDatasId == datasObj.Id).ToList();
                                        foreach (CurrentSetupIpValues valueObj in tempValuesList)
                                        {
                                            ScenarioValuesVM = new PayoutPolicy_ScenarioValuesViewModel();
                                            ScenarioValuesVM = mapper.Map<CurrentSetupIpValues, PayoutPolicy_ScenarioValuesViewModel>(valueObj);
                                            ScenarioValuesVM.Id = 0;
                                            ScenarioValuesVM.PayoutPolicy_ScenarioDatasId = 0;
                                            ScenarioDatasVM.PayoutPolicy_ScenarioValuesVM.Add(ScenarioValuesVM);
                                        }
                                        ScenarioDatasVMList.Add(ScenarioDatasVM);
                                    }

                                }

                                PayoutPolicy_ScenarioFillingsVM.PayoutPolicy_ScenarioDatasVM = ScenarioDatasVMList;
                                filingList.Add(PayoutPolicy_ScenarioFillingsVM);
                            }
                            #endregion

                            result.FilingResult = filingList;
                        }
                        else
                        {
                            result.Message = "please save current setup first to get the data Available";
                            result.StatusCode = 2;
                            return BadRequest(result);
                        }
                    }


                }

            }
            catch (Exception ss)
            {
                result.Message = "There is some issue in the code";
                result.StatusCode = 0;
                return BadRequest(result);
            }

            return Ok(result);
        }


        [HttpPost]
        [Route("SavePayoutPolicy_ScenarioAnalysis")]
        public ActionResult<Object> SavePayoutPolicy_ScenarioAnalysis([FromBody] List<PayoutPolicy_ScenarioFillingsViewModel> scenariofilongList)
        {
            try
            {
                PayoutPolicy_ScenarioDatas tblDatasObj = new PayoutPolicy_ScenarioDatas();
                if (scenariofilongList != null && scenariofilongList.Count > 0)
                {

                    foreach (PayoutPolicy_ScenarioFillingsViewModel filingObj in scenariofilongList)
                    {
                        if (filingObj.PayoutPolicy_ScenarioDatasVM != null && filingObj.PayoutPolicy_ScenarioDatasVM.Count > 0)
                        {
                            //convert VM to Table obj and save
                            foreach (PayoutPolicy_ScenarioDatasViewModel dataObj in filingObj.PayoutPolicy_ScenarioDatasVM)
                            {
                                tblDatasObj = new PayoutPolicy_ScenarioDatas();
                                tblDatasObj = mapper.Map<PayoutPolicy_ScenarioDatasViewModel, PayoutPolicy_ScenarioDatas>(dataObj);

                                //save datas
                                if (tblDatasObj.Id == 0)
                                    iPayoutPolicy_ScenarioDatas.Add(tblDatasObj);
                                else
                                    iPayoutPolicy_ScenarioDatas.Update(tblDatasObj);
                                iPayoutPolicy_ScenarioDatas.Commit();

                            }

                        }
                    }
                    return Ok(new { message = "data Saved Successfully", status = 200, result = true });
                }
                else
                {
                    return BadRequest(new { message = "No data found to save", status = 200, result = false });
                }
            }
            catch (Exception ss)
            {
                return BadRequest(new { message = "some error occured while saving data", status = 0, result = false });
            }

        }

        [HttpGet]
        [Route("GetPayoutScenarioOutputAnalysis/{UserId}")]
        public ActionResult GetPayoutScenarioOutputAnalysis(long UserId)
        {
            PayoutPolicy_ScenarioOutputResult result = new PayoutPolicy_ScenarioOutputResult();
            List<PayoutPolicy_ScenarioOutputFillingsViewModel> FilingListResult = new List<PayoutPolicy_ScenarioOutputFillingsViewModel>();
            PayoutPolicy_ScenarioOutputFillingsViewModel filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
            List<PayoutPolicy_ScenarioOutputDatasViewModel> ScenarioOutputDatasVMList;
            PayoutPolicy_ScenarioOutputDatasViewModel ScenarioOutputDatasVM;
            List<PayoutPolicy_ScenarioOutputValuesViewModel> ScenarioOutputValuesVMList;
            PayoutPolicy_ScenarioOutputValuesViewModel ScenarioOutputValuesVM;

            try
            {
                result.ReportName = "Scenario Output";
                List<PayoutPolicy_ScenarioDatas> scenarioDatasListObj = new List<PayoutPolicy_ScenarioDatas>();
                List<PayoutPolicy_ScenarioValues> scenarioValuesListObj = new List<PayoutPolicy_ScenarioValues>();
                //get current setup
                CurrentSetup tblCurrentSetup = iCurrentSetup.GetSingle(x => x.UserId == UserId);
                //check its own table
                scenarioDatasListObj = tblCurrentSetup != null ? iPayoutPolicy_ScenarioDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList() : null;
                if (scenarioDatasListObj != null && scenarioDatasListObj.Count > 0)
                {
                    bool scenarioValuesCount = false;
                    scenarioValuesListObj = iPayoutPolicy_ScenarioValues.FindBy(x => scenarioDatasListObj.Any(t => t.Id == x.PayoutPolicy_ScenarioDatasId)).ToList();
                    if (scenarioValuesListObj != null && scenarioValuesListObj.Count > 0)
                        scenarioValuesCount = true;

                    //get dumy years
                    PayoutPolicy_ScenarioOutputValuesViewModel dumyvalues;
                    List<PayoutPolicy_ScenarioOutputValuesViewModel> dumyvaluesList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    int strt = Convert.ToInt32(tblCurrentSetup.CurrentYear);
                    int End = Convert.ToInt32(tblCurrentSetup.EndYear);
                    for (int i = strt; i <= End; i++)
                    {
                        dumyvalues = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        dumyvalues.Year = Convert.ToString(strt);
                        dumyvalues.Value = "";
                        dumyvaluesList.Add(dumyvalues);
                        strt = strt + 1;
                    }


                    // get currentsetup Datas and values by currentsetupId
                    List<CurrentSetupIpDatas> CurrentSetupIpDatasListObj = new List<CurrentSetupIpDatas>();
                    List<CurrentSetupIpValues> CurrentSetupIpValuesListObj = new List<CurrentSetupIpValues>();
                    bool ipdatasCount = false;
                    bool ipvaluesCount = false;

                    CurrentSetupIpDatasListObj = iCurrentSetupIpDatas.FindBy(x => x.CurrentSetupId == tblCurrentSetup.Id).ToList();
                    if (CurrentSetupIpDatasListObj != null && CurrentSetupIpDatasListObj.Count > 0)
                    {
                        ipdatasCount = true;
                        CurrentSetupIpValuesListObj = iCurrentSetupIpValues.FindBy(x => CurrentSetupIpDatasListObj.Any(m => m.Id == x.CurrentSetupIpDatasId)).ToList();
                        if (CurrentSetupIpValuesListObj != null && CurrentSetupIpValuesListObj.Count > 0)
                            ipvaluesCount = true;
                    }

                    PayoutPolicy_ScenarioDatas i39 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 2);
                    PayoutPolicy_ScenarioDatas i40 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 3);
                    PayoutPolicy_ScenarioDatas i41 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 4);
                    CurrentSetupIpDatas c29 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.OtherInputs && x.Sequence == 3) : null;

                    CurrentSetupIpDatas c30 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.OtherInputs && x.Sequence == 4) : null;

                    PayoutPolicy_ScenarioDatas i49 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 12);

                    PayoutPolicy_ScenarioDatas i50 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 13);

                    PayoutPolicy_ScenarioDatas i51 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 14);

                    PayoutPolicy_ScenarioDatas i53 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 16);

                    #region New Capital Structure  # done

                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Capital Structure";

                    // Equity Value ( E)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Equity Value ( E)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    //Price per Share ($) * Number of Shares Outstanding - Basic (Millions) => c13* c14
                    CurrentSetupIpDatas c13 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 2) : null;
                    CurrentSetupIpDatas c14 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 3) : null;


                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            value = (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";

                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Market Value of Preferred Equity (P)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Market Value of Preferred Equity (P)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    //Current Preferred Share Price ($) * Number of Preferred Shares Outstanding (Millions) => c18 * c19
                    CurrentSetupIpDatas c18 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 7) : null;
                    CurrentSetupIpDatas c19 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 8) : null;

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C18Value = c18 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c18.Id) : null;
                            CurrentSetupIpValues C19Value = c19 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c19.Id) : null;
                            value = (C18Value != null && !string.IsNullOrEmpty(C18Value.Value) ? Convert.ToDouble(C18Value.Value) : 0) * (C19Value != null && !string.IsNullOrEmpty(C19Value.Value) ? Convert.ToDouble(C19Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Debt Value (D)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Debt Value (D)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    //Target Debt-to-Equity (D/E) Ratio * Equity Value ( E) ($M) => i28* (i60 or(c13*c14))
                    PayoutPolicy_ScenarioDatas i28 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.Inputs && x.Sequence == 2);

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            value = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Change in Debt
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Change in Debt";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    //Debt Value (D) -Market Value of Interest Bearing Debt ($M) i62-c23
                    CurrentSetupIpDatas c23 = ipdatasCount == true && ipvaluesCount == true ? CurrentSetupIpDatasListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 12) : null;

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;


                            value = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) - (C23Value != null && !string.IsNullOrEmpty(C23Value.Value) ? Convert.ToDouble(C23Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Net Debt Value (ND) // skip for now
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Net Debt Value (ND)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();


                    CurrentSetupIpDatas c49 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 12) : null;

                    CurrentSetupIpValues c49Value = c49 != null && ipvaluesCount == true ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;

                    CurrentSetupIpDatas c27 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.OtherInputs && x.Sequence == 1) : null;

                    CurrentSetupIpValues c27Value = c27 != null && ipvaluesCount == true ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c27.Id) : null;

                    PayoutPolicy_ScenarioDatas i33 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.NewPayoutPolicy && x.Sequence == 1);

                    PayoutPolicy_ScenarioDatas i34 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.NewPayoutPolicy && x.Sequence == 2);

                    PayoutPolicy_ScenarioDatas i35 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.NewPayoutPolicy && x.Sequence == 3);


                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;

                            CurrentSetupIpValues C27_Value = c27 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c27.Id) : null;

                            value = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)))
                                - ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) - (C27_Value != null && !string.IsNullOrEmpty(C27_Value.Value) ? Convert.ToDouble(C27_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region New Cost of Capital
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Cost of Capital";

                    // Cost of Equity (rE)  //skip for now
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cost of Equity (rE)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    CurrentSetupIpDatas c16 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 5) : null;
                    CurrentSetupIpDatas c21 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 10) : null;
                    CurrentSetupIpDatas c24 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 13) : null;

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true && ipvaluesCount == true)
                        {
                            // for calculate A $C$27
                            CurrentSetupIpValues C13_Value = c13 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14_Value = c14 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c14.Id) : null;
                            CurrentSetupIpValues C18_Value = c18 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c18.Id) : null;
                            CurrentSetupIpValues C19_Value = c19 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c19.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C27_Value = c27 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c27.Id) : null;
                            CurrentSetupIpValues c49_Value = c49 != null && ipvaluesCount == true ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C16_Value = c16 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c16.Id) : null;
                            CurrentSetupIpValues C21_Value = c21 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c21.Id) : null;
                            CurrentSetupIpValues C24_Value = c24 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c24.Id) : null;

                            //(c13*c14)+(c18*c19)+(c23-c49+c27) => D                           
                            double DValue = (((C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? Convert.ToDouble(C13_Value.Value) : 0) * (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? Convert.ToDouble(C14_Value.Value) : 0)) + ((C18_Value != null && !string.IsNullOrEmpty(C18_Value.Value) ? Convert.ToDouble(C18_Value.Value) : 0) * (C19_Value != null && !string.IsNullOrEmpty(C19_Value.Value) ? Convert.ToDouble(C19_Value.Value) : 0)) + ((C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0) - (c49_Value != null && !string.IsNullOrEmpty(c49_Value.Value) ? Convert.ToDouble(c49_Value.Value) : 0) + (C27_Value != null && !string.IsNullOrEmpty(C27_Value.Value) ? Convert.ToDouble(C27_Value.Value) : 0)));

                            //(C16*((c13*c14)/D))+(c21*((c18*c19)/D))+(c24*((c23-c49+c27)/D)) =>A
                            double AValue = DValue != 0 ? ((C16_Value != null && !string.IsNullOrEmpty(C16_Value.Value) ? Convert.ToDouble(C16_Value.Value) : 0) * (((C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? Convert.ToDouble(C13_Value.Value) : 0) * (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? Convert.ToDouble(C14_Value.Value) : 0)) / DValue)) + ((C21_Value != null && !string.IsNullOrEmpty(C21_Value.Value) ? Convert.ToDouble(C21_Value.Value) : 0) * (((C18_Value != null && !string.IsNullOrEmpty(C18_Value.Value) ? Convert.ToDouble(C18_Value.Value) : 0) * (C19_Value != null && !string.IsNullOrEmpty(C19_Value.Value) ? Convert.ToDouble(C19_Value.Value) : 0)) / DValue)) + ((C24_Value != null && !string.IsNullOrEmpty(C24_Value.Value) ? Convert.ToDouble(C24_Value.Value) : 0) * (((C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0) - (c49_Value != null && !string.IsNullOrEmpty(c49_Value.Value) ? Convert.ToDouble(c49_Value.Value) : 0) + (C27_Value != null && !string.IsNullOrEmpty(C27_Value.Value) ? Convert.ToDouble(C27_Value.Value) : 0)) / DValue)) : 0;

                            //calculate B
                            //(i64/(c13*c14)*(A-c24))
                            // i64

                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;

                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;


                            double i64value = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)))
                                - ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) - (C27_Value != null && !string.IsNullOrEmpty(C27_Value.Value) ? Convert.ToDouble(C27_Value.Value) : 0);

                            CurrentSetupIpValues C24Value = c24 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c24.Id) : null;
                            CurrentSetupIpValues C21Value = c21 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c21.Id) : null;
                            CurrentSetupIpValues C18Value = c18 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c18.Id) : null;
                            CurrentSetupIpValues C19Value = c19 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c19.Id) : null;

                            double BValue = ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) != 0 ? ((i64value / ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (AValue - (C24Value != null && !string.IsNullOrEmpty(C24Value.Value) ? Convert.ToDouble(C24Value.Value) : 0))) : 0;

                            // (C19Value != null && !string.IsNullOrEmpty(C19Value.Value) ? Convert.ToDouble(C19Value.Value) : 0)

                            // calculate C
                            //((i61/i60)*(A-c21))
                            //((C18*c19)/(c13*c14))*(A-c21)
                            double CValue = ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) != 0 ? (((C18Value != null && !string.IsNullOrEmpty(C18Value.Value) ? Convert.ToDouble(C18Value.Value) : 0) * (C19Value != null && !string.IsNullOrEmpty(C19Value.Value) ? Convert.ToDouble(C19Value.Value) : 0)) / ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (AValue - (C21Value != null && !string.IsNullOrEmpty(C21Value.Value) ? Convert.ToDouble(C21Value.Value) : 0)) : 0;

                            //A+B+C
                            value = AValue + BValue + CValue;

                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Cost of Preferred Equity (rP)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cost of Preferred Equity (rP)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    //Cost of Preferred Equity (rP)=> c21
                    //CurrentSetupIpDatas c21 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.SourcesOfFinancing && x.Sequence == 10) : null;

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C21Value = c21 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c21.Id) : null;

                            value = C21Value != null && !string.IsNullOrEmpty(C21Value.Value) ? Convert.ToDouble(C21Value.Value) : 0;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Cost of Debt (rD)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cost of Debt (rD)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    //Cost of Debt (rD) => i30
                    PayoutPolicy_ScenarioDatas i30 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.Inputs && x.Sequence == 4);

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;
                            value = i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Unlevered Cost of Capital/Equity (rU)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Unlevered Cost of Capital/Equity (rU)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        if (ipvaluesCount == true)
                        {
                            // for calculate A $C$27
                            CurrentSetupIpValues C13_Value = c13 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14_Value = c14 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c14.Id) : null;
                            CurrentSetupIpValues C18_Value = c18 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c18.Id) : null;
                            CurrentSetupIpValues C19_Value = c19 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c19.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C27_Value = c27 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c27.Id) : null;
                            CurrentSetupIpValues c49_Value = c49 != null && ipvaluesCount == true ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C16_Value = c16 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c16.Id) : null;
                            CurrentSetupIpValues C21_Value = c21 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c21.Id) : null;
                            CurrentSetupIpValues C24_Value = c24 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c24.Id) : null;

                            //(c13*c14)+(c18*c19)+(c23-c49+c27) => D                           
                            double DValue = (((C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? Convert.ToDouble(C13_Value.Value) : 0) * (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? Convert.ToDouble(C14_Value.Value) : 0)) + ((C18_Value != null && !string.IsNullOrEmpty(C18_Value.Value) ? Convert.ToDouble(C18_Value.Value) : 0) * (C19_Value != null && !string.IsNullOrEmpty(C19_Value.Value) ? Convert.ToDouble(C19_Value.Value) : 0)) + ((C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0) - (c49_Value != null && !string.IsNullOrEmpty(c49_Value.Value) ? Convert.ToDouble(c49_Value.Value) : 0) + (C27_Value != null && !string.IsNullOrEmpty(C27_Value.Value) ? Convert.ToDouble(C27_Value.Value) : 0)));

                            //(C16*((c13*c14)/D))+(c21*((c18*c19)/D))+(c24*((c23-c49+c27)/D)) =>A
                            double value = DValue != 0 ? ((C16_Value != null && !string.IsNullOrEmpty(C16_Value.Value) ? Convert.ToDouble(C16_Value.Value) : 0) * (((C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? Convert.ToDouble(C13_Value.Value) : 0) * (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? Convert.ToDouble(C14_Value.Value) : 0)) / DValue)) + ((C21_Value != null && !string.IsNullOrEmpty(C21_Value.Value) ? Convert.ToDouble(C21_Value.Value) : 0) * (((C18_Value != null && !string.IsNullOrEmpty(C18_Value.Value) ? Convert.ToDouble(C18_Value.Value) : 0) * (C19_Value != null && !string.IsNullOrEmpty(C19_Value.Value) ? Convert.ToDouble(C19_Value.Value) : 0)) / DValue)) + ((C24_Value != null && !string.IsNullOrEmpty(C24_Value.Value) ? Convert.ToDouble(C24_Value.Value) : 0) * (((C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0) - (c49_Value != null && !string.IsNullOrEmpty(c49_Value.Value) ? Convert.ToDouble(c49_Value.Value) : 0) + (C27_Value != null && !string.IsNullOrEmpty(C27_Value.Value) ? Convert.ToDouble(C27_Value.Value) : 0)) / DValue)) : 0;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Weighted Average Cost of Capital (rWACC) //skip for now
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Weighted Average Cost of Capital (rWACC)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);



                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region New Leverage Ratios #done
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Leverage Ratios";

                    // Debt-to-Equity (D/E) Ratio %
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Debt-to-Equity (D/E) Ratio";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    //Debt Value (D) ($M) / Equity Value ( E) ($M)
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            value = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Debt-to-Value (D/V) Ratio
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Debt-to-Value (D/V) Ratio";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    // Debt Value(D) ($M)/( Equity Value ( E) ($M) + Market Value of Preferred Equity (P) ($M) +Debt Value (D) ($M))

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            CurrentSetupIpValues C18Value = c18 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c18.Id) : null;
                            CurrentSetupIpValues C19Value = c19 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c19.Id) : null;
                            double equityvalue = (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0);
                            double marketvalue = (C18Value != null && !string.IsNullOrEmpty(C18Value.Value) ? Convert.ToDouble(C18Value.Value) : 0) * (C19Value != null && !string.IsNullOrEmpty(C19Value.Value) ? Convert.ToDouble(C19Value.Value) : 0);
                            double debtvalue = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0));
                            value = (equityvalue + marketvalue + debtvalue) != 0 ? debtvalue * 100 / (equityvalue + marketvalue + debtvalue) : 0;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region New Cash Balance (after payout) #done
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Cash Balance (after payout)";

                    // Cash & Equivalent
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cash & Equivalent";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();


                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;

                            //SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35
                            value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Excess Cash
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Excess Cash";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;

                            CurrentSetupIpValues C27_Value = c27 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c27.Id) : null;


                            value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) - (C27_Value != null && !string.IsNullOrEmpty(C27_Value.Value) ? Convert.ToDouble(C27_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region New Internal Valuation: Post-Payout #done
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Internal Valuation: Post-Payout";

                    // Unlevered Enterprise Value
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Unlevered Enterprise Value";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Levered Enterprise Value
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Levered Enterprise Value";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Interest Tax Shield Value
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Interest Tax Shield Value";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Change in Interest Tax Shield Value
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Change in Interest Tax Shield Value";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Stock Price
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Stock Price";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "$";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region New Payout Analysis #done
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Payout Analysis";

                    // Ongoing Quarterly Dividends
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Ongoing Quarterly Dividends";
                    ScenarioOutputDatasVM.IsParentItem = true;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Total Payout
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Total Payout";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();



                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;

                            value = (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // DPS (Dividends per Share -Basic)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "DPS (Dividends per Share -Basic)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "$";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            CurrentSetupIpValues C14_Value = c14 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c14.Id) : null;

                            //i33/$C$14
                            value = (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) / Convert.ToDouble(C14_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Dividend Payout Ratio
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Dividend Payout Ratio";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    PayoutPolicy_ScenarioDatas i47 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 10);
                    PayoutPolicy_ScenarioDatas i55 = scenarioDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.FinancialStatementsPrePayout && x.Sequence == 18);

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i47Value = i47 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i47.Id) : null;
                            value = (i47Value != null && !string.IsNullOrEmpty(i47Value.Value) ? (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) / Convert.ToDouble(i47Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Dividend Yield
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Dividend Yield";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            CurrentSetupIpValues C14_Value = c14 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c14.Id) : null;
                            CurrentSetupIpValues C13_Value = c13 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c13.Id) : null;

                            value = (C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) / Convert.ToDouble(C14_Value.Value) : 0) / Convert.ToDouble(C13_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // One Time Dividend
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "One Time Dividend";
                    ScenarioOutputDatasVM.IsParentItem = true;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Total Payout
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Total Payout";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            value = (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // DPS (Dividends per Share -Basic)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "DPS (Dividends per Share -Basic)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "$";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            CurrentSetupIpValues C14_Value = c14 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c14.Id) : null;

                            value = (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) / Convert.ToDouble(C14_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Dividend Payout Ratio
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Dividend Payout Ratio";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i47Value = i47 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i47.Id) : null;
                            value = (i47Value != null && !string.IsNullOrEmpty(i47Value.Value) ? (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) / Convert.ToDouble(i47Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Dividend Yield
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Dividend Yield";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            CurrentSetupIpValues C14_Value = c14 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c14.Id) : null;
                            CurrentSetupIpValues C13_Value = c13 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c13.Id) : null;

                            value = (C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) / Convert.ToDouble(C14_Value.Value) : 0) / Convert.ToDouble(C13_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Stock Buybacks
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Stock Buybacks";
                    ScenarioOutputDatasVM.IsParentItem = true;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);


                    // Total Payout
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Total Payout";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            value = (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Shares Repurchased
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Shares Repurchased";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C13_Value = c13 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c13.Id) : null;

                            value = (C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0) / Convert.ToDouble(C13_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);


                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion                  

                    #region New Shares Outstanding & EPS: Post-Payout
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Shares Outstanding & EPS: Post-Payout";

                    // Number of Shares Outstanding - Basic
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Number of Shares Outstanding - Basic";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C13_Value = c13 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14_Value = c14 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c14.Id) : null;

                            value = (C14_Value != null && !string.IsNullOrEmpty(C14_Value.Value) ? Convert.ToDouble(C14_Value.Value) : 0) - (C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0) / Convert.ToDouble(C13_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Number of Shares Outstanding - Diluted
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Number of Shares Outstanding - Diluted";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();

                    CurrentSetupIpDatas c15 = ipdatasCount == true ? CurrentSetupIpDatasListObj.Find(x => x.StatementTypeId == (int)StatementTypeEnum.Inputs && x.Sequence == 4) : null;

                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C13_Value = c13 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C15_Value = c15 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c15.Id) : null;

                            value = (C15_Value != null && !string.IsNullOrEmpty(C15_Value.Value) ? Convert.ToDouble(C15_Value.Value) : 0) - (C13_Value != null && !string.IsNullOrEmpty(C13_Value.Value) ? (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0) / Convert.ToDouble(C13_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Net Earnings per Share - Basic  #skip for now
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Net Earnings per Share - Basic";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "$";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Net Earnings per Share - Diluted #skip for now
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Net Earnings per Share - Diluted";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "$";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region Proforma Financial Statements: Post-Payout #done
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "Proforma Financial Statements: Post-Payout";

                    // Income Statement
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Income Statement";
                    ScenarioOutputDatasVM.IsParentItem = true;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();



                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Net Sales
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Net Sales";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i39Value = i39 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i39.Id) : null;

                            value = (i39Value != null && !string.IsNullOrEmpty(i39Value.Value) ? Convert.ToDouble(i39Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // EBITDA i112
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "EBITDA";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;

                            value = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Depreciation & Amortization
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Depreciation & Amortization";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;

                            value = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Interest Income
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Interest Income";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {

                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");

                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // EBIT i115
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "EBIT";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;
                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;


                            //for A
                            // PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;

                            double A = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);


                            // for B
                            //PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;

                            double B = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);

                            //for C
                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            double C = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);

                            //A-B+C //i12-i113+i114
                            value = A - B + C;

                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Interest Expense  i116
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Interest Expense";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            value = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // EBT (Earnings Before Taxes)
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "EBT (Earnings Before Taxes)";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {

                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;
                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;


                            //EBIT + interest expense
                            double expense = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);

                            double A = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);


                            // for B
                            //PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;

                            double B = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);

                            //for C
                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            double C = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);

                            //A-B+C //i12-i113+i114
                            double EBIT = A - B + C;

                            //EBIT + interest expense
                            value = EBIT + expense;

                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Taxes
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Taxes";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {

                            CurrentSetupIpValues c30_Value = ipdatasCount == true && c30 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c30.Id) : null;
                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;
                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;


                            //EBIT + interest expense
                            double expense = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);

                            double A = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);


                            // for B
                            //PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;

                            double B = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);

                            //for C
                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            double C = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);

                            //A-B+C //i12-i113+i114
                            double EBIT = A - B + C;

                            //EBT=EBIT + interest expense
                            //Taxes=EBT * $C$30
                            value = (EBIT + expense) * (c30_Value != null && !string.IsNullOrEmpty(c30_Value.Value) ? Convert.ToDouble(c30_Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Net Earnings => I117-I118(EBT-taxes) //skip for now
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Net Earnings";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {

                            CurrentSetupIpValues c30_Value = ipdatasCount == true && c30 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c30.Id) : null;
                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;
                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;


                            //EBIT + interest expense
                            double expense = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);

                            double A = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);


                            // for B
                            //PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;

                            double B = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);

                            //for C
                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            double C = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);

                            //A-B+C //i12-i113+i114
                            double EBIT = A - B + C;

                            //EBT=EBIT + interest expense
                            //Taxes=EBT * $C$30
                            double Taxes = (EBIT + expense) * (c30_Value != null && !string.IsNullOrEmpty(c30_Value.Value) ? Convert.ToDouble(c30_Value.Value) : 0);
                            value = ((EBIT + expense) - ((EBIT + expense) * (c30_Value != null && !string.IsNullOrEmpty(c30_Value.Value) ? Convert.ToDouble(c30_Value.Value) : 0)));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Balance Sheet
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Balance Sheet";
                    ScenarioOutputDatasVM.IsParentItem = true;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Cash and Equivalents
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cash and Equivalents";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;

                            //SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35
                            value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Total Current Assets
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Total Current Assets";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;


                            PayoutPolicy_ScenarioValues i49Value = i49 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i49.Id) : null;
                            PayoutPolicy_ScenarioValues i50Value = i50 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i50.Id) : null;
                            //value = (i50Value != null && !string.IsNullOrEmpty(i50Value.Value) ? Convert.ToDouble(i50Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0);
                            //i50+(i121 -i49)

                            double i121value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0));


                            value = (i50Value != null && !string.IsNullOrEmpty(i50Value.Value) ? Convert.ToDouble(i50Value.Value) : 0) + ((i121value) - (i49Value != null && !string.IsNullOrEmpty(i49Value.Value) ? Convert.ToDouble(i49Value.Value) : 0));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Total Assets
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Total Assets";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;

                            PayoutPolicy_ScenarioValues i49Value = i49 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i49.Id) : null;
                            PayoutPolicy_ScenarioValues i51Value = i51 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i51.Id) : null;
                            //value = (i50Value != null && !string.IsNullOrEmpty(i50Value.Value) ? Convert.ToDouble(i50Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0);

                            double i121value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0));

                            //i51+(i121-i49)
                            value = (i51Value != null && !string.IsNullOrEmpty(i51Value.Value) ? Convert.ToDouble(i51Value.Value) : 0) + ((i121value) - (i49Value != null && !string.IsNullOrEmpty(i49Value.Value) ? Convert.ToDouble(i49Value.Value) : 0));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Total Debt  i124
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Total Debt";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            value = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Shareholders Equity  i125
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Shareholders Equity";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            PayoutPolicy_ScenarioValues i49Value = i49 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i49.Id) : null;
                            PayoutPolicy_ScenarioValues i53Value = i53 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i53.Id) : null;

                            double i121value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0));

                            //i51+(i121-i49)

                            value = (i53Value != null && !string.IsNullOrEmpty(i53Value.Value) ? Convert.ToDouble(i53Value.Value) : 0) + (i121value - (i49Value != null && !string.IsNullOrEmpty(i49Value.Value) ? Convert.ToDouble(i49Value.Value) : 0));
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Cash Flow Statement
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cash Flow Statement";
                    ScenarioOutputDatasVM.IsParentItem = true;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Cash Flow from Operations
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cash Flow from Operations";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "M";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {

                            CurrentSetupIpValues c30_Value = ipdatasCount == true && c30 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c30.Id) : null;
                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;
                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            PayoutPolicy_ScenarioValues i55Value = i55 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i55.Id) : null;
                            PayoutPolicy_ScenarioValues i47Value = i47 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i47.Id) : null;


                            //EBIT + interest expense
                            double expense = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);

                            double A = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);


                            // for B
                            //PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;

                            double B = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);

                            //for C
                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            double C = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);

                            //A-B+C //i12-i113+i114
                            double EBIT = A - B + C;
                            //EBT=EBIT + interest expense
                            //Taxes=EBT * $C$30
                            double Taxes = (EBIT + expense) * (c30_Value != null && !string.IsNullOrEmpty(c30_Value.Value) ? Convert.ToDouble(c30_Value.Value) : 0);
                            //EBT-Taxes
                            double netearningsvalue = ((EBIT + expense) - ((EBIT + expense) * (c30_Value != null && !string.IsNullOrEmpty(c30_Value.Value) ? Convert.ToDouble(c30_Value.Value) : 0)));
                            //i55+(netearningsvalue-i47)
                            value = (i55Value != null && !string.IsNullOrEmpty(i55Value.Value) ? Convert.ToDouble(i55Value.Value) : 0) + (netearningsvalue - (i47Value != null && !string.IsNullOrEmpty(i47Value.Value) ? Convert.ToDouble(i47Value.Value) : 0));

                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);


                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region Debt Ratios & Analysis: Post-Payout
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "Debt Ratios & Analysis: Post-Payout";

                    // Debt-to-Market Equity Ratio i130
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Debt-to-Market Equity Ratio";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            value = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0);
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Debt-to-Book Equity Ratio i131
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Debt-to-Book Equity Ratio";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            double i124 = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0));

                            //for i125

                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23_Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            CurrentSetupIpValues C49_Value = c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            PayoutPolicy_ScenarioValues i49Value = i49 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i49.Id) : null;
                            PayoutPolicy_ScenarioValues i53Value = i53 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i53.Id) : null;

                            double i121value = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0));

                            //i51+(i121-i49)

                            double i125 = (i53Value != null && !string.IsNullOrEmpty(i53Value.Value) ? Convert.ToDouble(i53Value.Value) : 0) + (i121value - (i49Value != null && !string.IsNullOrEmpty(i49Value.Value) ? Convert.ToDouble(i49Value.Value) : 0));
                            value = i125 != 0 ? (i124 / i125) * 100 : 0;

                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // EBIT Interest Coverage i132
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "EBIT Interest Coverage";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "x";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;
                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            double A = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);

                            double B = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);

                            //for C
                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            double C = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);

                            //A-B+C //i12-i113+i114
                            double EBIT = A - B + C;

                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;

                            double expense = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);
                            //i115/i116
                            value = EBIT / expense;


                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // EBITDA Interest Coverage i133
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "EBITDA Interest Coverage";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "x";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (scenarioValuesCount == true)
                        {
                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            double i116 = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);

                            double i112 = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);
                            value = i116 != 0 ? i112 / i116 : 0;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Cash Flow from Operations / Total Debt
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Cash Flow from Operations / Total Debt";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "%";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {

                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            double i124 = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0));

                            double i112 = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);

                            // for i127
                            CurrentSetupIpValues c30_Value = ipdatasCount == true && c30 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c30.Id) : null;
                            PayoutPolicy_ScenarioValues i30Value = i30 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i30.Id) : null;

                            PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;
                            CurrentSetupIpValues C23_Value = ipdatasCount == true && c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;

                            CurrentSetupIpValues c29_Value = ipdatasCount == true && c29 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c29.Id) : null;

                            CurrentSetupIpValues C49_Value = ipdatasCount == true && c49 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c49.Id) : null;

                            PayoutPolicy_ScenarioValues i33Value = i33 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i33.Id) : null;
                            PayoutPolicy_ScenarioValues i34Value = i34 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i34.Id) : null;
                            PayoutPolicy_ScenarioValues i35Value = i35 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i35.Id) : null;
                            CurrentSetupIpValues C23Value = c23 != null ? CurrentSetupIpValuesListObj.OrderBy(x => x.Id).FirstOrDefault(x => x.CurrentSetupIpDatasId == c23.Id) : null;
                            PayoutPolicy_ScenarioValues i55Value = i55 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i55.Id) : null;
                            PayoutPolicy_ScenarioValues i47Value = i47 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i47.Id) : null;


                            //EBIT + interest expense
                            double expense = ((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0))) * (i30Value != null && !string.IsNullOrEmpty(i30Value.Value) ? Convert.ToDouble(i30Value.Value) : 0);

                            double A = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);


                            // for B
                            //PayoutPolicy_ScenarioValues i41Value = i41 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i41.Id) : null;

                            double B = (i41Value != null && !string.IsNullOrEmpty(i41Value.Value) ? Convert.ToDouble(i41Value.Value) : 0);

                            //for C
                            //(SC$49+((I28*i13*i14)-$c$23)-i33-i34-i35) *$c$29
                            double C = ((C49_Value != null && !string.IsNullOrEmpty(C49_Value.Value) ? Convert.ToDouble(C49_Value.Value) : 0) + (((i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * (C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0)) - (C23_Value != null && !string.IsNullOrEmpty(C23_Value.Value) ? Convert.ToDouble(C23_Value.Value) : 0)) - (i33Value != null && !string.IsNullOrEmpty(i33Value.Value) ? Convert.ToDouble(i33Value.Value) : 0) - (i34Value != null && !string.IsNullOrEmpty(i34Value.Value) ? Convert.ToDouble(i34Value.Value) : 0) - (i35Value != null && !string.IsNullOrEmpty(i35Value.Value) ? Convert.ToDouble(i35Value.Value) : 0)) * (c29_Value != null && !string.IsNullOrEmpty(c29_Value.Value) ? Convert.ToDouble(c29_Value.Value) : 0);

                            //A-B+C //i12-i113+i114
                            double EBIT = A - B + C;
                            //EBT=EBIT + interest expense
                            //Taxes=EBT * $C$30
                            double Taxes = (EBIT + expense) * (c30_Value != null && !string.IsNullOrEmpty(c30_Value.Value) ? Convert.ToDouble(c30_Value.Value) : 0);
                            //EBT-Taxes
                            double netearningsvalue = ((EBIT + expense) - ((EBIT + expense) * (c30_Value != null && !string.IsNullOrEmpty(c30_Value.Value) ? Convert.ToDouble(c30_Value.Value) : 0)));
                            //i55+(netearningsvalue-i47)
                            double i127 = (i55Value != null && !string.IsNullOrEmpty(i55Value.Value) ? Convert.ToDouble(i55Value.Value) : 0) + (netearningsvalue - (i47Value != null && !string.IsNullOrEmpty(i47Value.Value) ? Convert.ToDouble(i47Value.Value) : 0));


                            //124/112
                            value = i124 != 0 ? (i127 / i124) * 100 : 0;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);


                    // Total Debt / EBITDA
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Total Debt / EBITDA";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "x";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        double value = 0;
                        if (ipvaluesCount == true && scenarioValuesCount == true)
                        {

                            PayoutPolicy_ScenarioValues i40Value = i40 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i40.Id) : null;
                            CurrentSetupIpValues C13Value = c13 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c13.Id) : null;
                            CurrentSetupIpValues C14Value = c14 != null ? CurrentSetupIpValuesListObj.Find(x => x.Year == obj.Year && x.CurrentSetupIpDatasId == c14.Id) : null;
                            PayoutPolicy_ScenarioValues i28Value = i28 != null ? scenarioValuesListObj.Find(x => x.Year == obj.Year && x.PayoutPolicy_ScenarioDatasId == i28.Id) : null;
                            double totalDebt = (i28Value != null && !string.IsNullOrEmpty(i28Value.Value) ? Convert.ToDouble(i28Value.Value) : 0) / 100 * ((C13Value != null && !string.IsNullOrEmpty(C13Value.Value) ? Convert.ToDouble(C13Value.Value) : 0) * (C14Value != null && !string.IsNullOrEmpty(C14Value.Value) ? Convert.ToDouble(C14Value.Value) : 0));

                            double i112 = (i40Value != null && !string.IsNullOrEmpty(i40Value.Value) ? Convert.ToDouble(i40Value.Value) : 0);
                            //124/112
                            value = i112 != 0 ? totalDebt / i112 : 0;
                            ScenarioOutputValuesVM.Value = value.ToString("0.##");
                        }
                        else
                            ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Potential S&P Debt Rating
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Potential S&P Debt Rating";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "x";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "AA";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);


                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    #region New Payout Policy Implications #done
                    filingVM = new PayoutPolicy_ScenarioOutputFillingsViewModel();
                    ScenarioOutputDatasVMList = new List<PayoutPolicy_ScenarioOutputDatasViewModel>();
                    filingVM.StatementType = "New Payout Policy Implications";

                    // Agency Issues
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Agency Issues";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Signaling to Markets
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Signaling to Markets";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Tax Implications
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Tax Implications";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Investor Clienteles
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Investor Clienteles";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);

                    // Risks
                    ScenarioOutputDatasVM = new PayoutPolicy_ScenarioOutputDatasViewModel();
                    ScenarioOutputValuesVMList = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    ScenarioOutputDatasVM.LineItem = "Risks";
                    ScenarioOutputDatasVM.IsParentItem = false;
                    ScenarioOutputDatasVM.Sequence = ScenarioOutputDatasVMList.Count + 1;
                    ScenarioOutputDatasVM.StatementTypeId = (int)StatementTypeEnum.SourcesOfFinancing;
                    ScenarioOutputDatasVM.Unit = "";
                    ScenarioOutputDatasVM.CurrentSetupId = tblCurrentSetup.Id;
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = new List<PayoutPolicy_ScenarioOutputValuesViewModel>();
                    foreach (PayoutPolicy_ScenarioOutputValuesViewModel obj in dumyvaluesList)
                    {
                        ScenarioOutputValuesVM = new PayoutPolicy_ScenarioOutputValuesViewModel();
                        ScenarioOutputValuesVM.Year = obj.Year;
                        ScenarioOutputValuesVM.PayoutPolicy_ScenarioOutputDatasId = 0;
                        ScenarioOutputValuesVM.Value = "";
                        ScenarioOutputValuesVMList.Add(ScenarioOutputValuesVM);
                    }
                    ScenarioOutputDatasVM.PayoutPolicy_ScenarioOutputValuesVM = ScenarioOutputValuesVMList;
                    ScenarioOutputDatasVMList.Add(ScenarioOutputDatasVM);



                    filingVM.PayoutPolicy_ScenarioOutputDatasVM = ScenarioOutputDatasVMList;
                    FilingListResult.Add(filingVM);
                    #endregion

                    result.StatusCode = 1;
                    result.Message = "no issue found";
                    result.FilingResult = FilingListResult;
                }
                else
                {
                    result.Message = "please save Scenario analysis inputs first to get the data Available";
                    result.StatusCode = 2;
                    return BadRequest(result);
                }
            }
            catch (Exception ss)
            {
                result.Message = "There is some issue in the code";
                result.StatusCode = 0;
                return BadRequest(result);
            }
            return Ok(result);
        }



        #endregion

    }

}







