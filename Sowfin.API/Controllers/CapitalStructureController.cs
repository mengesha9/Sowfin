using System.ComponentModel.DataAnnotations.Schema;
using AutoMapper;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Sowfin.API.Lib;
using Sowfin.API.ViewModels;
using Sowfin.API.ViewModels.CapitalStructure;
using Sowfin.Data.Abstract;
using Sowfin.Data.Common.Enum;
using Sowfin.Data.Common.Helper;
using Sowfin.Data.Repositories;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using static Sowfin.API.ViewModels.CapitalStructure.MathClass;
using static Sowfin.Model.Entities.CapitalAnalysis;
using static Sowfin.Model.Entities.CapitalStructure;


namespace Sowfin.API.Controllers
{



    [Route("api/[controller]")]
    [ApiController]
    public class CapitalStructureController : ControllerBase
    {
        private readonly ICapitalStructure iCapitalStructure = null;
        private readonly IWebHostEnvironment  _hostingEnvironment = null;
        private readonly ISnapshots iSnapshots = null;
        private readonly ICostOfCapital iCostOfCapital = null;
        private readonly ICalculateBeta iCalculateBeta = null;
        private const string CAPITALSTRUCTURE = "CAPITAL_STRUCTURE";
        private const string CAPITALSTRUCTUREANALYSIS = "CAPITAL_STRUCTURE_ANALYSIS";
        private const string CAPITALSTRUCTURESCENARIO = "CAPITAL_STRUCTURE_SCENARIO";
        public object MathCapitalStructure { get; private set; }
        /// new changes
        private readonly IMasterCostofCapitalNStructure iMasterCostofCapitalNStructure = null;
        private readonly ICapitalStructure_Input iCapitalStructure_Input = null;
        private readonly ICostofCapital_Input iCostofCapital_Input = null;
        private readonly ISnapshot_CostofCapitalNStructure iSnapshot_CostofCapitalNStructure = null;
        private readonly ICapitalStructure_Snapshot iCapitalStructure_Snapshot = null;
        private readonly ICostofCapital_Snapshot iCostofCapital_Snapshot = null;
        ICapitalStructureScenarioSnapshot iCapitalStructureScenarioSnapshot = null;

        IMapper mapper;

        public CapitalStructureController(ICapitalStructure _iCapitalStructure, IWebHostEnvironment  hostingEnvironment,
             ICostOfCapital _iCostCapital, ISnapshots _iSnapshots,IMasterCostofCapitalNStructure _iMasterCostofCapitalNStructure, IMapper _imapper, ICapitalStructure_Input _iCapitalStructure_Input, ICostofCapital_Input _iCostofCapital_Input, ISnapshot_CostofCapitalNStructure _iSnapshot_CostofCapitalNStructure, ICapitalStructure_Snapshot _iCapitalStructure_Snapshot, ICostofCapital_Snapshot _iCostofCapital_Snapshot,
             ICapitalStructureScenarioSnapshot _icapitalStructureScenarioSnapshot,ICalculateBeta _iCalculateBeta)
        {

            _hostingEnvironment = hostingEnvironment;
            iCapitalStructure = _iCapitalStructure;
            iCostOfCapital = _iCostCapital;
            iSnapshots = _iSnapshots;
            mapper = _imapper;

            //new changes
            iMasterCostofCapitalNStructure = _iMasterCostofCapitalNStructure;
            iCapitalStructure_Input = _iCapitalStructure_Input;
            iCostofCapital_Input = _iCostofCapital_Input;

            // snapshot
            iSnapshot_CostofCapitalNStructure = _iSnapshot_CostofCapitalNStructure;
            iCapitalStructure_Snapshot = _iCapitalStructure_Snapshot;
            iCostofCapital_Snapshot = _iCostofCapital_Snapshot;

            iCalculateBeta = _iCalculateBeta;

            iCapitalStructureScenarioSnapshot = _icapitalStructureScenarioSnapshot;

        }

        [HttpGet]
        [Route("GetCostofCapitalNStructure/{UserId}")]
        public ActionResult<Object> GetCostofCapitalNStructure(long UserId)
        {
            MasterCostofCapitalNStructureViewModel result = new MasterCostofCapitalNStructureViewModel();
            try
            {
               

                //check if master data is saved or not 
                MasterCostofCapitalNStructure tblMaster = iMasterCostofCapitalNStructure.GetSingle(x => x.UserId == UserId);
                if(tblMaster!=null)
                {
                    //bind data from Saved Values
                    result = mapper.Map<MasterCostofCapitalNStructure, MasterCostofCapitalNStructureViewModel>(tblMaster);

                    //get Capital Stryucture
                    // List<CapitalStructure_Input> tblcapitalStructureInputList = iCapitalStructure_Input.FindBy(x=>x.MasterId==tblMaster.Id).ToList();
                    List<CapitalStructure_Input> tblcapitalStructureInputList = iCapitalStructure_Input.FindBy(x => x.MasterId == tblMaster.Id && x.HeaderId != 3).ToList();
                    //Map Capital Structure
                    result.CapitalStructureInputList = new List<CapitalStructure_InputViewModel>();
                    if(tblcapitalStructureInputList != null)
                    foreach (var obj in tblcapitalStructureInputList.OrderBy(x=>x.Id))
                    {
                        CapitalStructure_InputViewModel tempVM= mapper.Map<CapitalStructure_Input, CapitalStructure_InputViewModel>(obj);
                        result.CapitalStructureInputList.Add(tempVM) ;
                    }


                    //get Cost of Capital
                    List<CostofCapital_Input> tblCostofcapitalInputList = iCostofCapital_Input.FindBy(x => x.MasterId == tblMaster.Id).ToList();
                    //Map Cost of capital
                    result.CostofCapitalInputList = new List<CostofCapital_InputViewModel>();
                    if (tblCostofcapitalInputList != null)
                        foreach (var obj in tblCostofcapitalInputList.OrderBy(x => x.Id))
                        {
                            CostofCapital_InputViewModel tempVM = mapper.Map<CostofCapital_Input, CostofCapital_InputViewModel>(obj);
                            result.CostofCapitalInputList.Add(tempVM);
                        }


                    //get output for Capital Structure & Cost of Capital
                    MasterCostofCapitalNStructureViewModel tempOutput= GetOutPutforCapitalNStructure(result);
                    
                    if(tempOutput!=null)
                    {
                        result.CapitalStructureOutputList = tempOutput.CapitalStructureOutputList != null && tempOutput.CapitalStructureOutputList.Count > 0 ? tempOutput.CapitalStructureOutputList : new List<CapitalStructure_OutputViewModel>();

                        result.CostofCapitalOutputList = tempOutput.CostofCapitalOutputList != null && tempOutput.CostofCapitalOutputList.Count > 0 ? tempOutput.CostofCapitalOutputList : new List<CostofCapital_OutputViewModel>();
                    }

                    //get snapshotList
                    List<Snapshot_CostofCapitalNStructure> tblSnapshotMasterList = iSnapshot_CostofCapitalNStructure.FindBy(x => x.MasterId == tblMaster.Id && x.Active==true).ToList();
                    //Map Cost of capital & Capital Structure Snapshot Master
                    result.SnapshotMasterList = new List<Snapshot_CostofCapitalNStructureViewModel>();
                    if (tblSnapshotMasterList != null && tblSnapshotMasterList.Count>0)
                    {
                        //get Capital Structure Snapshot Data
                        List<CapitalStructure_Snapshot> capitalStructureList = iCapitalStructure_Snapshot.FindBy(x => x.Id != 0).ToList();
                        //get Cost of Capital List
                        List<CostofCapital_Snapshot> CostofCapitalList = iCostofCapital_Snapshot.FindBy(x => x.Id != 0).ToList();

                        Snapshot_CostofCapitalNStructureViewModel tempVM;
                        foreach (var obj in tblSnapshotMasterList.OrderBy(x => x.Id))
                        {
                            tempVM = new Snapshot_CostofCapitalNStructureViewModel();
                            tempVM = mapper.Map<Snapshot_CostofCapitalNStructure, Snapshot_CostofCapitalNStructureViewModel>(obj);

                            tempVM.capitalStructure_SnapshotList = new List<CapitalStructure_SnapshotViewModel>();
                            tempVM.CostofCapital_SnapshotList = new List<CostofCapital_SnapshotViewModel>();
                            //bind Capital Structure
                            foreach (var temp in  capitalStructureList.Where(x=>x.Snapshot_CostofCapitalNStructureId == obj.Id))
                            {
                                tempVM.capitalStructure_SnapshotList.Add(mapper.Map<CapitalStructure_Snapshot, CapitalStructure_SnapshotViewModel>(temp));
                            }

                            //bind Cost of capital
                            foreach (var temp in CostofCapitalList.Where(x => x.Snapshot_CostofCapitalNStructureId == obj.Id))
                            {
                                tempVM.CostofCapital_SnapshotList.Add(mapper.Map<CostofCapital_Snapshot, CostofCapital_SnapshotViewModel>(temp));
                            }

                            result.SnapshotMasterList.Add(tempVM);
                        }
                    }





                    //Calculate Beta
                    CalculateBeta testCalculateBeta = new CalculateBeta();
                    var testCalculateBetaArray = iCalculateBeta.FindBy(x => x.MasterId == result.Id).OrderByDescending(x => x.Id).ToArray();
                    //  testCalculateBeta = iCalculateBeta.GetLatestSingle(x => x.CostOfCapitals_Id == costofCapitalVM.Id).order;
                    testCalculateBeta = testCalculateBetaArray.Length != 0 ? testCalculateBetaArray[0] : null;
                    if (testCalculateBeta != null)
                    {
                        result.CalculateBeta_Id = testCalculateBeta.Id;
                        result.CalculateBetaViewModel = new CalculateBetaViewModel();
                        result.CalculateBetaViewModel.CalculateBeta_Id = testCalculateBeta.Id;
                        result.CalculateBetaViewModel.CostOfCapitals_Id = testCalculateBeta.CostOfCapitals_Id;
                        result.CalculateBetaViewModel.Frequency_Id = testCalculateBeta.Frequency_Id;
                        //  costofCapitalVM.Frequency_Value = EnumHelper.DescriptionAttr((FrequencyEnum)testCalculateBeta.Frequency_Id);
                        result.CalculateBetaViewModel.TargetMarketIndex_Id = testCalculateBeta.TargetMarketIndex_Id;
                        //[
                        result.CalculateBetaViewModel.TargetRiskFreeRate_Id = testCalculateBeta.TargetRiskFreeRate_Id;
                        result.CalculateBetaViewModel.TargetRiskFreeRate_Value = EnumHelper.DescriptionAttr((TargetRiskFreeIndexEnum)testCalculateBeta.TargetRiskFreeRate_Id);
                        result.CalculateBetaViewModel.DataSource_Id = testCalculateBeta.DataSource_Id;
                        result.CalculateBetaViewModel.DataSource_Value = EnumHelper.DescriptionAttr((DataSourceEnum)testCalculateBeta.DataSource_Id);
                        result.CalculateBetaViewModel.Duration_FromDate = testCalculateBeta.Duration_FromDate;
                        result.CalculateBetaViewModel.Duration_toDate = testCalculateBeta.Duration_toDate;
                        result.CalculateBetaViewModel.CreatedDate = testCalculateBeta.CreatedDate;
                        result.CalculateBetaViewModel.ModifiedDate = testCalculateBeta.ModifiedDate;
                        result.CalculateBetaViewModel.Active = testCalculateBeta.Active;
                        result.CalculateBetaViewModel.BetaValue = Math.Round(Convert.ToDouble(testCalculateBeta.BetaValue), 2);
                    }
                    else
                        result.CalculateBeta_Id = 0;




                }
                else
                {
                    //bydefault first leverage policy would be selected
                    result.UserId = UserId;
                    result.LeveragePolicyId = (int)LeveragePolicyEnum.DebtToEquityRatio;
                    result.HasEquity = true;
                    result.HasPreferredEquity = true;
                    result.HasDebt = true;
                    result.BetaSourceId = null;                  
                    result.CostofDebtMethodId = 1;
                    result.CreatedDate = result.ModifiedDate = System.DateTime.Now;

                    //Create data for first time
                    result.CapitalStructureInputList = new List<CapitalStructure_InputViewModel>();
                    result.CapitalStructureInputList = getInitialCapitalStructureList();

                    //Create Cost of capital Inputs for first time
                    result.CostofCapitalInputList = new List<CostofCapital_InputViewModel>();
                    result.CostofCapitalInputList = getInitialCostofCapitalList();
                }

                // get drop down from enum
                result.CurrencyValueList = EnumHelper.GetEnumListbyName<CurrencyValueEnum>(); //get CurrencyValueList
                result.NumberCountList = EnumHelper.GetEnumListbyName<NumberCountEnum>(); //get Number countList
                result.ValueTypeList = EnumHelper.GetEnumListbyName<ValueTypeEnum>(); //get Value Type List
                result.LeveragePolicyList = EnumHelper.GetEnumListbyName<LeveragePolicyEnum>();  //get leverage Policy List
                result.HeaderList = EnumHelper.GetEnumListbyName<HeadersEnum>(); //get leverage Policy List
                result.BetasourceList = EnumHelper.GetEnumListbyName<BetaSourceEnum>();  //get Beta Source List

                
               

            }
            catch (Exception ss)
            {
                Console.WriteLine(ss.Message);

            }
            return result;
        }

        private MasterCostofCapitalNStructureViewModel GetOutPutforScenarioAnalysis(MasterCostofCapitalNStructureViewModel master)
        {
            MasterCostofCapitalNStructureViewModel tempoutput = new MasterCostofCapitalNStructureViewModel();

            List<ScenarioAnalysis_OutputViewModel> ScenarioAnalysisOutputList = new List<ScenarioAnalysis_OutputViewModel>();
            ScenarioAnalysis_OutputViewModel scenarioAnalysisitem;

            double? marketValueofCommonEquityCSOutputvalue = 0;
            double? marketValueofPreferredEquityCSOutputvalue = 0;
            double? targetDebttoEquitySAInputvalue = 0;
            double? marketValueofDebtCSOutputvalue = 0;
            double? excessCashCSOutputvalue = 0;
            double? cashEquivalentCSOutputValue = 0;
            double? UnleveredCostofCapitalEquityCSOutputValue = 0;
            double? costofPreferredEquityCSInputValue = 0;
            double? marginalTaxRateCSInputValue = 0;


            double? equityValue = 0;
            double? preferredEquityValue = 0;
            double? debtValue = 0;
            double? changeInDebt = 0;
            double? NetInDebtValue = 0;
            double? cashEquivalentSAOutputValue = 0;
            double? excessCashSAOutputvalue = 0;
            double? CostofDebtSAInputValue = 0;
            double? costofEquitySAOutputValue = 0;
            double? costofPreferredEquitySAOutputValue = 0;
            double? costofDebtSAOutputValue = 0;
            double? UnleveredCostofCapitalEquitySAOutputValue = 0;
            double? WeightedAverageCostofCapitalSAOutputValue = 0;
            double? DebtToEquityRatioSAOutputValue = 0;
            double? DebtToValueRatioSAOutputValue = 0;
            double? valueofPermanentDebtSAInputValue = 0;
            double? interestCoverageRatioSAInputValue = 0;
            double? freeCashFlowNextYearSAInputValue = 0;


            int? equityValueUnit = 0;
            int? preferredEquityValueUnit = 0;
            int? debtValueUnit = 0;
            int? changeInDebtUnit = 0;
            int? NetInDebtValueUnit = 0;

            if (master.SALeveragePolicyID == 1)
            {

                //Unlevered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Levered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Levered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);               

                //Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Change in Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Stock Price
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Stock Price";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.Dollar;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Equity Value (E)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Equity Value (E)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofCommonEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") : null;
                marketValueofCommonEquityCSOutputvalue = marketValueofCommonEquityObj != null ? marketValueofCommonEquityObj.BasicValue : 0;
                equityValue = scenarioAnalysisitem.BasicValue =  marketValueofCommonEquityCSOutputvalue;
                equityValueUnit = scenarioAnalysisitem.UnitId = marketValueofCommonEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Preferred Equity Value (P)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Preferred Equity Value (P)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofpreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") : null;
                marketValueofPreferredEquityCSOutputvalue = marketValueofpreferredEquityObj != null ? marketValueofpreferredEquityObj.BasicValue : 0;
                preferredEquityValue = scenarioAnalysisitem.BasicValue = marketValueofPreferredEquityCSOutputvalue;
                preferredEquityValueUnit = scenarioAnalysisitem.UnitId = marketValueofpreferredEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt Value (D)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt Value (D)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var targetDebttoEquityObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Target Debt-to-Equity (D/E) Ratio") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Target Debt-to-Equity (D/E) Ratio") : null;
                targetDebttoEquitySAInputvalue = targetDebttoEquityObj != null ? targetDebttoEquityObj.BasicValue : 0;
                debtValue = (targetDebttoEquitySAInputvalue / 100) * equityValue ;
                scenarioAnalysisitem.BasicValue = debtValue;
                //  scenarioAnalysisitem.UnitId = marketValueofCommonEquityObj.UnitId;
                debtValueUnit = scenarioAnalysisitem.UnitId = equityValueUnit;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Change in Debt
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Debt";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofDebtObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") : null;
                marketValueofDebtCSOutputvalue = marketValueofDebtObj != null ? marketValueofDebtObj.BasicValue : 0;
                changeInDebt = debtValue - marketValueofDebtCSOutputvalue;
                scenarioAnalysisitem.BasicValue = changeInDebt;
                // changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, marketValueofDebtObj.UnitId);
                changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, marketValueofDebtObj.UnitId);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Net Debt Value (ND)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Net Debt Value (ND)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var excessCashObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") : null;
                excessCashCSOutputvalue = excessCashObj != null ? excessCashObj.BasicValue : 0;
                NetInDebtValue = debtValue - (excessCashCSOutputvalue + changeInDebt);
                scenarioAnalysisitem.BasicValue = NetInDebtValue;
                // NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, excessCashObj.UnitId);
               int ? NetInDebtValueUnit1 = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, NetInDebtValueUnit1);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cash & Equivalent
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cash & Equivalent";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var cashEquivalentObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") : null;
                cashEquivalentCSOutputValue = cashEquivalentObj != null ? cashEquivalentObj.BasicValue : 0;
                cashEquivalentSAOutputValue = cashEquivalentCSOutputValue + changeInDebt;
                scenarioAnalysisitem.BasicValue = cashEquivalentSAOutputValue;
                int? cashEquivalentSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(cashEquivalentObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Excess Cash
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Excess Cash";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                excessCashSAOutputvalue = excessCashCSOutputvalue + changeInDebt;
                scenarioAnalysisitem.BasicValue = excessCashSAOutputvalue;
                int ? excessCashSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Equity (rE)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Equity (rE)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                var unleveredCostofCapitalEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") : null;
                UnleveredCostofCapitalEquityCSOutputValue = unleveredCostofCapitalEquityObj != null ? unleveredCostofCapitalEquityObj.BasicValue : 0;

                var costofDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) : null;
                CostofDebtSAInputValue = costofDebtObj != null ? costofDebtObj.BasicValue : 0;

                // Take Cost of Preferred Equity (rP) value from Capital Structure output list because it is not available in Capital Structure input list

                var costofPreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") : null;
                costofPreferredEquityCSInputValue = costofPreferredEquityObj != null ? costofPreferredEquityObj.BasicValue : 0;

                costofEquitySAOutputValue = ((UnleveredCostofCapitalEquityCSOutputValue / 100) + ((NetInDebtValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (CostofDebtSAInputValue / 100))) + ((preferredEquityValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (costofPreferredEquityCSInputValue / 100)))) * 100;

                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofEquitySAOutputValue;                
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Preferred Equity (rP)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Preferred Equity (rP)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                costofPreferredEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofPreferredEquityCSInputValue;               
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Debt (rD)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Debt (rD)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                costofDebtSAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = CostofDebtSAInputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Unlevered Cost of Capital/Equity (rU)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Cost of Capital/Equity (rU)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                UnleveredCostofCapitalEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = UnleveredCostofCapitalEquityCSOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Weighted Average Cost of Capital (rWACC)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Weighted Average Cost of Capital (rWACC)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                var marginalTaxRateObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") : null;
                marginalTaxRateCSInputValue = marginalTaxRateObj != null ? marginalTaxRateObj.BasicValue : 0;
              
                WeightedAverageCostofCapitalSAOutputValue = ((costofEquitySAOutputValue / 100) * (equityValue / (equityValue + preferredEquityValue + NetInDebtValue)) + (costofPreferredEquitySAOutputValue / 100) * (preferredEquityValue / (equityValue + preferredEquityValue + NetInDebtValue)) + (costofDebtSAOutputValue / 100) * (NetInDebtValue / (equityValue + preferredEquityValue + NetInDebtValue)) * (1 - (marginalTaxRateCSInputValue / 100))) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = WeightedAverageCostofCapitalSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Equity (D/E) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Equity (D/E) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToEquityRatioSAOutputValue = (debtValue / equityValue) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToEquityRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Value (D/V) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Value (D/V) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToValueRatioSAOutputValue = (debtValue / (equityValue  + preferredEquityValue + debtValue)) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToValueRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

            }  
            
            else if (master.SALeveragePolicyID == 2)
            {

                //Unlevered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Levered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Levered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Change in Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Stock Price
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Stock Price";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.Dollar;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Equity Value (E)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Equity Value (E)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofCommonEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") : null;
                marketValueofCommonEquityCSOutputvalue = marketValueofCommonEquityObj != null ? marketValueofCommonEquityObj.BasicValue : 0;
                equityValue = scenarioAnalysisitem.BasicValue = marketValueofCommonEquityCSOutputvalue;
                equityValueUnit = scenarioAnalysisitem.UnitId = marketValueofCommonEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Preferred Equity Value (P)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Preferred Equity Value (P)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofpreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") : null;
                marketValueofPreferredEquityCSOutputvalue = marketValueofpreferredEquityObj != null ? marketValueofpreferredEquityObj.BasicValue : 0;
                preferredEquityValue = scenarioAnalysisitem.BasicValue = marketValueofPreferredEquityCSOutputvalue;
                preferredEquityValueUnit = scenarioAnalysisitem.UnitId = marketValueofpreferredEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt Value (D)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt Value (D)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var targetDebttoEquityObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Target Debt-to-Equity (D/E) Ratio") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Target Debt-to-Equity (D/E) Ratio") : null;
                targetDebttoEquitySAInputvalue = targetDebttoEquityObj != null ? targetDebttoEquityObj.BasicValue : 0;
                debtValue = (targetDebttoEquitySAInputvalue / 100) * equityValue;
                scenarioAnalysisitem.BasicValue = debtValue;
                //  scenarioAnalysisitem.UnitId = marketValueofCommonEquityObj.UnitId;
                debtValueUnit = scenarioAnalysisitem.UnitId = equityValueUnit;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Change in Debt
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Debt";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofDebtObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") : null;
                marketValueofDebtCSOutputvalue = marketValueofDebtObj != null ? marketValueofDebtObj.BasicValue : 0;
                changeInDebt = debtValue - marketValueofDebtCSOutputvalue;
                scenarioAnalysisitem.BasicValue = changeInDebt;
                // changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, marketValueofDebtObj.UnitId);
                changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, marketValueofDebtObj.UnitId);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Net Debt Value (ND)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Net Debt Value (ND)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var excessCashObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") : null;
                excessCashCSOutputvalue = excessCashObj != null ? excessCashObj.BasicValue : 0;
                NetInDebtValue = debtValue - (excessCashCSOutputvalue + changeInDebt);
                scenarioAnalysisitem.BasicValue = NetInDebtValue;
                // NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, excessCashObj.UnitId);
                int? NetInDebtValueUnit1 = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, NetInDebtValueUnit1);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cash & Equivalent
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cash & Equivalent";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var cashEquivalentObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") : null;
                cashEquivalentCSOutputValue = cashEquivalentObj != null ? cashEquivalentObj.BasicValue : 0;
                cashEquivalentSAOutputValue = cashEquivalentCSOutputValue + changeInDebt;
                scenarioAnalysisitem.BasicValue = cashEquivalentSAOutputValue;
                int? cashEquivalentSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(cashEquivalentObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Excess Cash
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Excess Cash";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                excessCashSAOutputvalue = excessCashCSOutputvalue + changeInDebt;
                scenarioAnalysisitem.BasicValue = excessCashSAOutputvalue;
                int? excessCashSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Equity (rE)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Equity (rE)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                var unleveredCostofCapitalEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") : null;
                UnleveredCostofCapitalEquityCSOutputValue = unleveredCostofCapitalEquityObj != null ? unleveredCostofCapitalEquityObj.BasicValue : 0;

                var costofDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) : null;
                CostofDebtSAInputValue = costofDebtObj != null ? costofDebtObj.BasicValue : 0;

                // Take Cost of Preferred Equity (rP) value from Capital Structure output list because it is not available in Capital Structure input list

                var costofPreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") : null;
                costofPreferredEquityCSInputValue = costofPreferredEquityObj != null ? costofPreferredEquityObj.BasicValue : 0;

                costofEquitySAOutputValue = ((UnleveredCostofCapitalEquityCSOutputValue / 100) + ((NetInDebtValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (CostofDebtSAInputValue / 100))) + ((preferredEquityValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (costofPreferredEquityCSInputValue / 100)))) * 100;

                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofEquitySAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Preferred Equity (rP)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Preferred Equity (rP)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                costofPreferredEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofPreferredEquityCSInputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Debt (rD)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Debt (rD)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                costofDebtSAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = CostofDebtSAInputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Unlevered Cost of Capital/Equity (rU)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Cost of Capital/Equity (rU)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                UnleveredCostofCapitalEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = UnleveredCostofCapitalEquityCSOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Weighted Average Cost of Capital (rWACC)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Weighted Average Cost of Capital (rWACC)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                var marginalTaxRateObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") : null;
                marginalTaxRateCSInputValue = marginalTaxRateObj != null ? marginalTaxRateObj.BasicValue : 0;
               
                WeightedAverageCostofCapitalSAOutputValue = ((UnleveredCostofCapitalEquityCSOutputValue / 100) - ((NetInDebtValue / (equityValue + preferredEquityValue + NetInDebtValue)) * (marginalTaxRateCSInputValue / 100) * (costofDebtSAOutputValue / 100) * ((1 + (UnleveredCostofCapitalEquityCSOutputValue / 100)) / (1 + (costofDebtSAOutputValue / 100))))) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = WeightedAverageCostofCapitalSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Equity (D/E) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Equity (D/E) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToEquityRatioSAOutputValue = (debtValue / equityValue) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToEquityRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Value (D/V) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Value (D/V) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToValueRatioSAOutputValue = (debtValue / (equityValue + preferredEquityValue + debtValue)) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToValueRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

            }

            else if (master.SALeveragePolicyID == 3)
            {

                //Unlevered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Levered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Levered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Change in Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Stock Price
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Stock Price";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.Dollar;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Equity Value (E)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Equity Value (E)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofCommonEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") : null;
                marketValueofCommonEquityCSOutputvalue = marketValueofCommonEquityObj != null ? marketValueofCommonEquityObj.BasicValue : 0;
                equityValue = scenarioAnalysisitem.BasicValue = marketValueofCommonEquityCSOutputvalue;
                equityValueUnit = scenarioAnalysisitem.UnitId = marketValueofCommonEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Preferred Equity Value (P)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Preferred Equity Value (P)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofpreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") : null;
                marketValueofPreferredEquityCSOutputvalue = marketValueofpreferredEquityObj != null ? marketValueofpreferredEquityObj.BasicValue : 0;
                preferredEquityValue = scenarioAnalysisitem.BasicValue = marketValueofPreferredEquityCSOutputvalue;
                preferredEquityValueUnit = scenarioAnalysisitem.UnitId = marketValueofpreferredEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt Value (D)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt Value (D)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var valueofPermanentDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Value of Permanent Debt") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Value of Permanent Debt") : null;
                valueofPermanentDebtSAInputValue = valueofPermanentDebtObj != null ? valueofPermanentDebtObj.BasicValue : 0;
                debtValue = valueofPermanentDebtSAInputValue;
                scenarioAnalysisitem.BasicValue = debtValue;
                //  scenarioAnalysisitem.UnitId = marketValueofCommonEquityObj.UnitId;
                debtValueUnit = scenarioAnalysisitem.UnitId = valueofPermanentDebtObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Change in Debt
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Debt";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofDebtObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") : null;
                marketValueofDebtCSOutputvalue = marketValueofDebtObj != null ? marketValueofDebtObj.BasicValue : 0;
                changeInDebt = debtValue - marketValueofDebtCSOutputvalue;
                scenarioAnalysisitem.BasicValue = changeInDebt;
                // changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, marketValueofDebtObj.UnitId);
                changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, marketValueofDebtObj.UnitId);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Net Debt Value (ND)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Net Debt Value (ND)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var excessCashObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") : null;
                excessCashCSOutputvalue = excessCashObj != null ? excessCashObj.BasicValue : 0;
                NetInDebtValue = debtValue - (excessCashCSOutputvalue + changeInDebt);
                scenarioAnalysisitem.BasicValue = NetInDebtValue;
                // NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, excessCashObj.UnitId);
                int? NetInDebtValueUnit1 = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, NetInDebtValueUnit1);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cash & Equivalent
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cash & Equivalent";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var cashEquivalentObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") : null;
                cashEquivalentCSOutputValue = cashEquivalentObj != null ? cashEquivalentObj.BasicValue : 0;
                cashEquivalentSAOutputValue = cashEquivalentCSOutputValue + changeInDebt;
                scenarioAnalysisitem.BasicValue = cashEquivalentSAOutputValue;
                int? cashEquivalentSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(cashEquivalentObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Excess Cash
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Excess Cash";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                excessCashSAOutputvalue = excessCashCSOutputvalue + changeInDebt;
                scenarioAnalysisitem.BasicValue = excessCashSAOutputvalue;
                int? excessCashSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                ////Cost of Equity (rE)
                //scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                //scenarioAnalysisitem.LineItem = "Cost of Equity (rE)";
                //scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                //scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                //scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                //var unleveredCostofCapitalEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") : null;
                //UnleveredCostofCapitalEquityCSOutputValue = unleveredCostofCapitalEquityObj != null ? unleveredCostofCapitalEquityObj.BasicValue : 0;

                //var costofDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) : null;
                //CostofDebtSAInputValue = costofDebtObj != null ? costofDebtObj.BasicValue : 0;

                //// Take Cost of Preferred Equity (rP) value from Capital Structure output list because it is not available in Capital Structure input list

                //var costofPreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") : null;
                //costofPreferredEquityCSInputValue = costofPreferredEquityObj != null ? costofPreferredEquityObj.BasicValue : 0;

                //costofEquitySAOutputValue = ((UnleveredCostofCapitalEquityCSOutputValue / 100) + ((NetInDebtValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (CostofDebtSAInputValue / 100))) + ((preferredEquityValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (costofPreferredEquityCSInputValue / 100)))) * 100;

                //scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofEquitySAOutputValue;
                //ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                ////Cost of Preferred Equity (rP)
                //scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                //scenarioAnalysisitem.LineItem = "Cost of Preferred Equity (rP)";
                //scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                //scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                //scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                //costofPreferredEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofPreferredEquityCSInputValue;
                //ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Debt (rD)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Debt (rD)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                var costofDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) : null;
                CostofDebtSAInputValue = costofDebtObj != null ? costofDebtObj.BasicValue : 0;

                costofDebtSAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = CostofDebtSAInputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Unlevered Cost of Capital/Equity (rU)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Cost of Capital/Equity (rU)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                var unleveredCostofCapitalEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") : null;
                UnleveredCostofCapitalEquityCSOutputValue = unleveredCostofCapitalEquityObj != null ? unleveredCostofCapitalEquityObj.BasicValue : 0;

                UnleveredCostofCapitalEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = UnleveredCostofCapitalEquityCSOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                ////Weighted Average Cost of Capital (rWACC)
                //scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                //scenarioAnalysisitem.LineItem = "Weighted Average Cost of Capital (rWACC)";
                //scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                //scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                //scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                //var marginalTaxRateObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") : null;
                //marginalTaxRateCSInputValue = marginalTaxRateObj != null ? marginalTaxRateObj.BasicValue : 0;

                //WeightedAverageCostofCapitalSAOutputValue = ((UnleveredCostofCapitalEquityCSOutputValue / 100) - ((NetInDebtValue / (equityValue + preferredEquityValue + NetInDebtValue)) * (marginalTaxRateCSInputValue / 100) * (costofDebtSAOutputValue / 100) * ((1 + (UnleveredCostofCapitalEquityCSOutputValue / 100)) / (1 + (costofDebtSAOutputValue / 100))))) * 100;
                //scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = WeightedAverageCostofCapitalSAOutputValue;
                //ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Equity (D/E) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Equity (D/E) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToEquityRatioSAOutputValue = (debtValue / equityValue) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToEquityRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Value (D/V) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Value (D/V) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToValueRatioSAOutputValue = (debtValue / (equityValue + preferredEquityValue + debtValue)) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToValueRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

            }

            else if (master.SALeveragePolicyID == 4)
            {

                //Unlevered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Levered Enterprise Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Levered Enterprise Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Change in Interest Tax Shield Value
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Interest Tax Shield Value";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.DollarM;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

                //Stock Price
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Stock Price";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewInternalValuation;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = (int)CurrencyValueEnum.Dollar;
                scenarioAnalysisitem.BasicValue = scenarioAnalysisitem.Value = null;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Equity Value (E)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Equity Value (E)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofCommonEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Common Equity (E)") : null;
                marketValueofCommonEquityCSOutputvalue = marketValueofCommonEquityObj != null ? marketValueofCommonEquityObj.BasicValue : 0;
                equityValue = scenarioAnalysisitem.BasicValue = marketValueofCommonEquityCSOutputvalue;
                equityValueUnit = scenarioAnalysisitem.UnitId = marketValueofCommonEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Preferred Equity Value (P)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Preferred Equity Value (P)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofpreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Preferred Equity (P)") : null;
                marketValueofPreferredEquityCSOutputvalue = marketValueofpreferredEquityObj != null ? marketValueofpreferredEquityObj.BasicValue : 0;
                preferredEquityValue = scenarioAnalysisitem.BasicValue = marketValueofPreferredEquityCSOutputvalue;
                preferredEquityValueUnit = scenarioAnalysisitem.UnitId = marketValueofpreferredEquityObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt Value (D)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt Value (D)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var interestCoverageRatioObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Interest Coverage Ratio") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Interest Coverage Ratio") : null;
                interestCoverageRatioSAInputValue = interestCoverageRatioObj != null ? interestCoverageRatioObj.BasicValue : 0;

                var freeCashFlowNextYearObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Free Cash Flow- Next Year") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Free Cash Flow- Next Year") : null;
                freeCashFlowNextYearSAInputValue = freeCashFlowNextYearObj != null ? freeCashFlowNextYearObj.BasicValue : 0;

                var costofDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)") : null;
                CostofDebtSAInputValue = costofDebtObj != null ? costofDebtObj.BasicValue : 0;

                debtValue = (interestCoverageRatioSAInputValue / 100) * freeCashFlowNextYearSAInputValue / (CostofDebtSAInputValue / 100);
                scenarioAnalysisitem.BasicValue = debtValue;
                debtValueUnit = scenarioAnalysisitem.UnitId = freeCashFlowNextYearObj.UnitId;
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Change in Debt
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Change in Debt";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var marketValueofDebtObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Market Value of Debt") : null;
                marketValueofDebtCSOutputvalue = marketValueofDebtObj != null ? marketValueofDebtObj.BasicValue : 0;
                changeInDebt = debtValue - marketValueofDebtCSOutputvalue;
                scenarioAnalysisitem.BasicValue = changeInDebt;
                // changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, marketValueofDebtObj.UnitId);
                changeInDebtUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, marketValueofDebtObj.UnitId);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Net Debt Value (ND)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Net Debt Value (ND)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCapitalStructure;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var excessCashObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Excess Cash") : null;
                excessCashCSOutputvalue = excessCashObj != null ? excessCashObj.BasicValue : 0;
                NetInDebtValue = debtValue - (excessCashCSOutputvalue + changeInDebt);
                scenarioAnalysisitem.BasicValue = NetInDebtValue;
                // NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(marketValueofCommonEquityObj.UnitId, excessCashObj.UnitId);
                int? NetInDebtValueUnit1 = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                NetInDebtValueUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(debtValueUnit, NetInDebtValueUnit1);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cash & Equivalent
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cash & Equivalent";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                var cashEquivalentObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cash & Equivalent") : null;
                cashEquivalentCSOutputValue = cashEquivalentObj != null ? cashEquivalentObj.BasicValue : 0;
                cashEquivalentSAOutputValue = cashEquivalentCSOutputValue + changeInDebt;
                scenarioAnalysisitem.BasicValue = cashEquivalentSAOutputValue;
                int? cashEquivalentSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(cashEquivalentObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Excess Cash
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Excess Cash";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCashBalance;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                scenarioAnalysisitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;

                excessCashSAOutputvalue = excessCashCSOutputvalue + changeInDebt;
                scenarioAnalysisitem.BasicValue = excessCashSAOutputvalue;
                int? excessCashSAUnit = scenarioAnalysisitem.UnitId = UnitConversion.getHigherDenominationUnit(excessCashObj.UnitId, changeInDebtUnit);
                scenarioAnalysisitem.Value = UnitConversion.getConvertedValueforCurrency(scenarioAnalysisitem.UnitId, scenarioAnalysisitem.BasicValue);
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                ////Cost of Equity (rE)
                //scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                //scenarioAnalysisitem.LineItem = "Cost of Equity (rE)";
                //scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                //scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                //scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                //var unleveredCostofCapitalEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") : null;
                //UnleveredCostofCapitalEquityCSOutputValue = unleveredCostofCapitalEquityObj != null ? unleveredCostofCapitalEquityObj.BasicValue : 0;

                //var costofDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) : null;
                //CostofDebtSAInputValue = costofDebtObj != null ? costofDebtObj.BasicValue : 0;

                //// Take Cost of Preferred Equity (rP) value from Capital Structure output list because it is not available in Capital Structure input list

                //var costofPreferredEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Cost of Preferred Equity (rP)") : null;
                //costofPreferredEquityCSInputValue = costofPreferredEquityObj != null ? costofPreferredEquityObj.BasicValue : 0;

                //costofEquitySAOutputValue = ((UnleveredCostofCapitalEquityCSOutputValue / 100) + ((NetInDebtValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (CostofDebtSAInputValue / 100))) + ((preferredEquityValue / equityValue) * ((UnleveredCostofCapitalEquityCSOutputValue / 100) - (costofPreferredEquityCSInputValue / 100)))) * 100;

                //scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofEquitySAOutputValue;
                //ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                ////Cost of Preferred Equity (rP)
                //scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                //scenarioAnalysisitem.LineItem = "Cost of Preferred Equity (rP)";
                //scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                //scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                //scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                //costofPreferredEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = costofPreferredEquityCSInputValue;
                //ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Cost of Debt (rD)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Cost of Debt (rD)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                //var costofDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cost of Debt (rD)" && x.HeaderId == 3) : null;
                //CostofDebtSAInputValue = costofDebtObj != null ? costofDebtObj.BasicValue : 0;

                costofDebtSAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = CostofDebtSAInputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Unlevered Cost of Capital/Equity (rU)
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Unlevered Cost of Capital/Equity (rU)";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                var unleveredCostofCapitalEquityObj = master.CapitalStructureOutputList != null && master.CapitalStructureOutputList.Count > 0 && master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") != null ? master.CapitalStructureOutputList.Find(x => x.LineItem == "Unlevered Cost of Capital/Equity (rU)") : null;
                UnleveredCostofCapitalEquityCSOutputValue = unleveredCostofCapitalEquityObj != null ? unleveredCostofCapitalEquityObj.BasicValue : 0;

                UnleveredCostofCapitalEquitySAOutputValue = scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = UnleveredCostofCapitalEquityCSOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                ////Weighted Average Cost of Capital (rWACC)
                //scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                //scenarioAnalysisitem.LineItem = "Weighted Average Cost of Capital (rWACC)";
                //scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewCostofCapital;
                //scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                //scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                //var marginalTaxRateObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)") : null;
                //marginalTaxRateCSInputValue = marginalTaxRateObj != null ? marginalTaxRateObj.BasicValue : 0;

                //WeightedAverageCostofCapitalSAOutputValue = ((UnleveredCostofCapitalEquityCSOutputValue / 100) - ((NetInDebtValue / (equityValue + preferredEquityValue + NetInDebtValue)) * (marginalTaxRateCSInputValue / 100) * (costofDebtSAOutputValue / 100) * ((1 + (UnleveredCostofCapitalEquityCSOutputValue / 100)) / (1 + (costofDebtSAOutputValue / 100))))) * 100;
                //scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = WeightedAverageCostofCapitalSAOutputValue;
                //ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Equity (D/E) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Equity (D/E) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToEquityRatioSAOutputValue = (debtValue / equityValue) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToEquityRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);


                //Debt-to-Value (D/V) Ratio %
                scenarioAnalysisitem = new ScenarioAnalysis_OutputViewModel();
                scenarioAnalysisitem.LineItem = "Debt-to-Value (D/V) Ratio";
                scenarioAnalysisitem.HeaderId = (int)ScenarioAnalysisHeadersEnum.NewLeverageRatios;
                scenarioAnalysisitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                scenarioAnalysisitem.DefaultUnitId = scenarioAnalysisitem.UnitId = null;

                DebtToValueRatioSAOutputValue = (debtValue / (equityValue + preferredEquityValue + debtValue)) * 100;
                scenarioAnalysisitem.Value = scenarioAnalysisitem.BasicValue = DebtToValueRatioSAOutputValue;
                ScenarioAnalysisOutputList.Add(scenarioAnalysisitem);

            }

            tempoutput.ScenarioAnalysisOutputList = ScenarioAnalysisOutputList;

            return tempoutput;

        }


        private MasterCostofCapitalNStructureViewModel GetOutPutforCapitalNStructure(MasterCostofCapitalNStructureViewModel master)
        {
            MasterCostofCapitalNStructureViewModel tempoutput = new MasterCostofCapitalNStructureViewModel();

            //capital Structure output
            #region Capital Structure
            List<CapitalStructure_OutputViewModel> CapitalStructureOutputList = new List<CapitalStructure_OutputViewModel>();
            CapitalStructure_OutputViewModel capitalStructureitem;

            #region Current Capital Structure
            
            double? currentSharePrice = 0;
            double? NumberofSharesOutstanding_Basic = 0;
            double? marketValueofCommonEquity = 0;
            int? MarketvalueofEquityunit = 0; ;
            int? MarketvalueofPreferredEquityUnit = 0;
            int? MarketvalueofDebtUnit = 0;
            if (master.HasEquity==true)
            {
                //Market Value of Common Equity (E)
                capitalStructureitem = new CapitalStructure_OutputViewModel();
                capitalStructureitem.LineItem = "Market Value of Common Equity (E)";
                capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCapitalStructure;
                capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                capitalStructureitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;
                // formula = Current Share Price ($) * Number of Shares Outstanding - Basic (Millions)
                // find Current Share price from Capital Structure input
                var currentSharePriceObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Current Share Price") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Current Share Price") : null;
                currentSharePrice = currentSharePriceObj != null ? currentSharePriceObj.BasicValue : 0;

                //find Number of Shares Outstanding - Basic
                var NumberofSharesOutstanding_BasicObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Number of Shares Outstanding - Basic") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Number of Shares Outstanding - Basic") : null;
                NumberofSharesOutstanding_Basic = NumberofSharesOutstanding_BasicObj != null ? NumberofSharesOutstanding_BasicObj.BasicValue : 0;
                marketValueofCommonEquity = (currentSharePrice != null ? currentSharePrice : 0) * (NumberofSharesOutstanding_Basic != null ? NumberofSharesOutstanding_Basic : 0);
                capitalStructureitem.BasicValue = marketValueofCommonEquity;

                MarketvalueofEquityunit = capitalStructureitem.UnitId =UnitConversion.getHigherDenominationUnit(NumberofSharesOutstanding_BasicObj.UnitId, currentSharePriceObj.UnitId);
               
                capitalStructureitem.Value = UnitConversion.getConvertedValueforCurrency(capitalStructureitem.UnitId, capitalStructureitem.BasicValue);
                CapitalStructureOutputList.Add(capitalStructureitem);
            }


            double? currentpreferredSharePrice = 0;
            double? NumberofreferredSharesOutstanding = 0;
            double? MarketValueofPreferredEquity = 0;
            if (master.HasPreferredEquity==true)
            {
                //Market Value of Preferred Equity (P)
                capitalStructureitem = new CapitalStructure_OutputViewModel();
                capitalStructureitem.LineItem = "Market Value of Preferred Equity (P)";
                capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCapitalStructure;
                capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                capitalStructureitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;
                //formula = Current Preferred Share Price($) * Number of Preferred Shares Outstanding
                // find Current Preferred Share Price
                var currentpreferredSharePriceObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Current Preferred Share Price") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Current Preferred Share Price") : null;
                currentpreferredSharePrice = currentpreferredSharePriceObj != null ? currentpreferredSharePriceObj.BasicValue : 0;
                //find Number of Preferred Shares Outstanding
               var  NumberofreferredSharesOutstandingObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Number of Preferred Shares Outstanding") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Number of Preferred Shares Outstanding") : null;
                NumberofreferredSharesOutstanding = NumberofreferredSharesOutstandingObj != null ? NumberofreferredSharesOutstandingObj.BasicValue : 0;

                MarketValueofPreferredEquity = (currentpreferredSharePrice != null ? currentpreferredSharePrice : 0) * (NumberofreferredSharesOutstanding != null ? NumberofreferredSharesOutstanding : 0);
                capitalStructureitem.BasicValue=MarketValueofPreferredEquity;

                MarketvalueofPreferredEquityUnit = capitalStructureitem.UnitId = UnitConversion.getHigherDenominationUnit(currentpreferredSharePriceObj.UnitId, NumberofreferredSharesOutstandingObj.UnitId);

                capitalStructureitem.Value = UnitConversion.getConvertedValueforCurrency(capitalStructureitem.UnitId, capitalStructureitem.BasicValue);
                CapitalStructureOutputList.Add(capitalStructureitem);
            }

            //cosr of debt
            double? CostofDebt = 0;
            if(master.HasDebt==true)
            {
                //for Method1
                //formula=  Risk Free Rate + Default Spread
                if (master.CostofDebtMethodId == 1)
                {
                    double? RiskFreeRate = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Risk Free Rate" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Risk Free Rate" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital).BasicValue : 0;

                    double? DefaultSpread = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Default Spread" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Default Spread" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital).BasicValue : 0;
                    CostofDebt = (RiskFreeRate != null ? RiskFreeRate : 0) + (DefaultSpread != null ? DefaultSpread : 0);

                }
                else if (master.CostofDebtMethodId == 2) //Yield to Maturity (y) - Probability of Default (p) * Expected Loss Rate / Default rate (L)
                {
                    //Yield to Maturity (y)
                    double? YieldtoMaturity = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Yield to Maturity (y)" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Yield to Maturity (y)" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital).BasicValue : 0;
                    //Probability of Default (p)
                    double? ProbabilityofDefault = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Probability of Default (p)" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Probability of Default (p)" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital).BasicValue : 0;
                    //Expected Loss Rate / Default rate (L)
                    double? ExpectedLossRate = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Expected Loss Rate / Default rate (L)" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Expected Loss Rate / Default rate (L)" && x.HeaderId == (int)HeadersEnum.CostofDebtCapital).BasicValue : 0;
                    CostofDebt = (YieldtoMaturity != null ? YieldtoMaturity : 0) - ((ProbabilityofDefault != null ? ProbabilityofDefault : 0) * (ExpectedLossRate != null ? ExpectedLossRate : 0))/100;
                }

            }


            //Market Value of Debt
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Market Value of Debt";
            capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCapitalStructure;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId =  (int)CurrencyValueEnum.DollarM;
            //formula= Market Value of Interest Bearing Debt           
            

            // find Market Value of Interest Bearing Debt
            double? MarketValue_InterestBearingDebt = 0;
            if (master.LeveragePolicyId == (int)LeveragePolicyEnum.ConstantInterestRatio)
            {
                // formula= =(Interest Coverage Ratio (k)*Free Cash Flow- Next Year)/Cost of Debt (rD)
                double? InterestCoverageratio = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Interest Coverage Ratio (k) - If Applicable") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Interest Coverage Ratio (k) - If Applicable").BasicValue : 0;
                var FreeCashFlowObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Free Cash Flow- Next Year-if Applicable") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Free Cash Flow- Next Year-if Applicable") : null;
                double? FreeCashFlow = FreeCashFlowObj != null ? FreeCashFlowObj.BasicValue : 0;


                MarketValue_InterestBearingDebt = ((InterestCoverageratio != null ? InterestCoverageratio : 0) * (FreeCashFlow != null ? FreeCashFlow : 0)) / (CostofDebt != null ? CostofDebt : 0);

                MarketvalueofDebtUnit = capitalStructureitem.UnitId = FreeCashFlowObj.UnitId;

            }
            else
            {
              var  MarketValue_InterestBearingDebtObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Market Value of Interest Bearing Debt") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Market Value of Interest Bearing Debt") : null;
                MarketValue_InterestBearingDebt = MarketValue_InterestBearingDebtObj != null ? MarketValue_InterestBearingDebtObj.BasicValue : 0;
                MarketvalueofDebtUnit = capitalStructureitem.UnitId = MarketValue_InterestBearingDebtObj.UnitId;                
            }

            capitalStructureitem.BasicValue = (MarketValue_InterestBearingDebt != null ? MarketValue_InterestBearingDebt : 0);
            capitalStructureitem.Value = UnitConversion.getConvertedValueforCurrency(capitalStructureitem.UnitId, capitalStructureitem.BasicValue);
            CapitalStructureOutputList.Add(capitalStructureitem);

            double? MarketValueofNetDebt = 0;
            ////Market Value of Net Debt (ND)
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Market Value of Net Debt (ND)";
            capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCapitalStructure;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId =  (int)CurrencyValueEnum.DollarM;
            //formula= Market Value of Interest Bearing Debt-(Cash & Equivalent-Cash Needed for Working Capital)

            // find Cash & Equivalent
            var cashNEquivalentObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cash & Equivalent") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cash & Equivalent") : null;
            // find Cash & Equivalent
            double? cashNEquivalent = cashNEquivalentObj != null ? cashNEquivalentObj.BasicValue : 0;

              var workingCapitalObj = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Cash Needed for Working Capital") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Cash Needed for Working Capital") : null;
            double? workingCapital = workingCapitalObj != null ? workingCapitalObj.BasicValue : 0;

            MarketValueofNetDebt = (MarketValue_InterestBearingDebt != null ? MarketValue_InterestBearingDebt : 0) - ((cashNEquivalent != null ? cashNEquivalent : 0) - (workingCapital != null ? workingCapital : 0));
            capitalStructureitem.BasicValue = MarketValueofNetDebt;
            int? CheckUnit = UnitConversion.getHigherDenominationUnit(cashNEquivalentObj.UnitId, workingCapitalObj.UnitId);

            capitalStructureitem.UnitId = UnitConversion.getHigherDenominationUnit(MarketvalueofDebtUnit, CheckUnit);
            capitalStructureitem.Value = UnitConversion.getConvertedValueforCurrency(capitalStructureitem.UnitId, capitalStructureitem.BasicValue);
            CapitalStructureOutputList.Add(capitalStructureitem);

            #endregion

            #region Current Cost of Capital

            // find Risk Free Rate (Future)
            double? RiskfreRateFuture = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Risk Free Rate (Future)" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Risk Free Rate (Future)" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital).BasicValue : 0;

            double? AdjustedBeta = 0;
            double? RawBeta = 0;
            double? CostofEquity = 0;
            //calculate Adjusted Beta (1/3 + 2/3*raw beta)  =1/3+2/3* Raw Beta
            if (master.BetaSourceId == (int)BetaSourceEnum.ManualEntry)
                //find Manual Beta
                RawBeta = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Manual Beta" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Manual Beta" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital).BasicValue : 0;
            else if (master.BetaSourceId == (int)BetaSourceEnum.ExternalSource)
                //find External Beta
                RawBeta = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "External Beta" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "External Beta" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital).BasicValue : 0;
            else if (master.BetaSourceId == (int)BetaSourceEnum.CalculateBeta)
                //find Calculated Beta
                RawBeta = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Calculated Beta" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Calculated Beta" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital).BasicValue : 0;

            // AdjustedBeta = (1 / 3) + (2 / 3) * (RawBeta!=null ? RawBeta: 0);
            AdjustedBeta = (0.3333333333333333) + (0.6666666666666667) * (RawBeta != null ? RawBeta : 0);

            //Calculate Market Risk Premium = ( Historical Market Return -Historical Risk Free Return) + Small Stock Premium
            double? HistoricalMarketReturn = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Historical Market Return" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Historical Market Return" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital).BasicValue : 0;

            double? HistoricalRiskFreeReturn = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Historical Risk Free Return" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Historical Risk Free Return" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital).BasicValue : 0;

            double? SmallStockPremium = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Small Stock Premium" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Small Stock Premium" && x.HeaderId == (int)HeadersEnum.CostofEquityCapital).BasicValue : 0;

            double? MarketRiskPremium = ((HistoricalMarketReturn != null ? HistoricalMarketReturn : 0) - (HistoricalRiskFreeReturn != null ? HistoricalRiskFreeReturn : 0)) + (SmallStockPremium != null ? SmallStockPremium : 0); 
            if (master.HasEquity==true)
            { 
            ////Cost of Equity ( rE)
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Cost of Equity ( rE)";
            capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.percentage;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = null;
            //formula= Risk Free Rate (Future) (%) + (Adjusted Beta (1/3 + 2/3*raw beta)) * Market Risk Premium           

           CostofEquity = (RiskfreRateFuture != null ? RiskfreRateFuture : 0) + (AdjustedBeta != null ? AdjustedBeta : 0) * (MarketRiskPremium != null ? MarketRiskPremium : 0);
            capitalStructureitem.Value = capitalStructureitem.BasicValue = CostofEquity;
            CapitalStructureOutputList.Add(capitalStructureitem);
        }

            double? PreferredDividend = 0;
            double? CurrentPreferredSharePrice = 0;
            double? CostofPreferredEquity = 0;
            if (master.HasPreferredEquity==true)
            {

                ////Cost of Preferred Equity (rP)
                capitalStructureitem = new CapitalStructure_OutputViewModel();
                capitalStructureitem.LineItem = "Cost of Preferred Equity (rP)";
                capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                //formula=  Preferred Dividend / Current Preferred Share Price

                // find Preferred Dividend
                PreferredDividend = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Preferred Dividend") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Preferred Dividend").BasicValue : 0;

                // find Current Preferred Share Price
                CurrentPreferredSharePrice = master.CapitalStructureInputList != null && master.CapitalStructureInputList.Count > 0 && master.CapitalStructureInputList.Find(x => x.LineItem == "Current Preferred Share Price") != null ? master.CapitalStructureInputList.Find(x => x.LineItem == "Current Preferred Share Price").BasicValue : 0;

                CostofPreferredEquity = CurrentPreferredSharePrice != null ? (PreferredDividend != null ? PreferredDividend : 0) * 100 / CurrentPreferredSharePrice : 0;
                capitalStructureitem.Value = capitalStructureitem.BasicValue = CostofPreferredEquity;
                CapitalStructureOutputList.Add(capitalStructureitem);
            }

            if(master.HasDebt==true)
            {
                ////////Cost of Debt (rD)
                capitalStructureitem = new CapitalStructure_OutputViewModel();
                capitalStructureitem.LineItem = "Cost of Debt (rD)";
                capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                capitalStructureitem.Value = capitalStructureitem.BasicValue = CostofDebt;
                CapitalStructureOutputList.Add(capitalStructureitem);
            }

            double? UnleveredCostofCapital = 0;
            ////////Unlevered Cost of Capital/Equity (rU)
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Unlevered Cost of Capital/Equity (rU)";
            capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.percentage;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = null;
           
            //((Market Value of common stock*cost of equity)+(Market Value of Preferred Equity* cst of preferred equity)+(Market Value of Net Debt * cost of Debt))/(Market Value of common stock + Market Value of Preferred Equity +Market Value of Net Debt  )

            UnleveredCostofCapital = (((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0)*(CostofEquity!=null ? CostofEquity :0 )/100) +((MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) *(CostofPreferredEquity!=null ? CostofPreferredEquity :0) /100) +((MarketValueofNetDebt != null ? MarketValueofNetDebt : 0)*(CostofDebt!=null ? CostofDebt :0)/100))*100 / ((marketValueofCommonEquity!=null ? marketValueofCommonEquity :0) +(MarketValueofPreferredEquity!=null ? MarketValueofPreferredEquity :0) + (MarketValueofNetDebt!=null ? MarketValueofNetDebt :0));


            //
            capitalStructureitem.Value = capitalStructureitem.BasicValue = UnleveredCostofCapital;
            CapitalStructureOutputList.Add(capitalStructureitem);

            double? WACC = 0;
            
            if (master.LeveragePolicyId==(int)LeveragePolicyEnum.DebtToEquityRatio || master.LeveragePolicyId == (int)LeveragePolicyEnum.AnnuallyAdjustDebt)
            {

                ////////Weighted Average Cost of Capital (rWACC)
                capitalStructureitem = new CapitalStructure_OutputViewModel();
                capitalStructureitem.LineItem = "Weighted Average Cost of Capital (rWACC)";
                capitalStructureitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                //"Marginal Tax Rate (Tc)"
                double? MarginalTaxRate = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)" && x.HeaderId==(int)HeadersEnum.OtherInputs) != null  ? master.CostofCapitalInputList.Find(x => x.LineItem == "Marginal Tax Rate (Tc)" && x.HeaderId == (int)HeadersEnum.OtherInputs).BasicValue : 0;


                //WACC = (((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) * (CostofEquity != null ? CostofEquity : 0) / 100) + ((MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) * (CostofPreferredEquity != null ? CostofPreferredEquity : 0) / 100) + ((MarketValueofNetDebt != null ? MarketValueofNetDebt : 0) * (CostofDebt != null ? CostofDebt : 0) / 100))*(1- ((MarginalTaxRate!=null ? MarginalTaxRate :0) /100)) / ((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) + (MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) + (MarketValueofNetDebt != null ? MarketValueofNetDebt : 0));
                //WACC= ((((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) * (CostofEquity != null ? CostofEquity : 0) / 100) + ((MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) * (CostofPreferredEquity != null ? CostofPreferredEquity : 0) / 100) + ((MarketValueofNetDebt != null ? MarketValueofNetDebt : 0) * (CostofDebt != null ? CostofDebt : 0) / 100)) *100/ ((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) + (MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) + (MarketValueofNetDebt != null ? MarketValueofNetDebt : 0)))*(1-(MarginalTaxRate!=null ? MarginalTaxRate/100 : 0 ));

                WACC = (CostofEquity != null ? CostofEquity : 0) * ((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) / ((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) + (MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) + (MarketValueofNetDebt != null ? MarketValueofNetDebt : 0))) + (CostofPreferredEquity != null ? CostofPreferredEquity : 0) * ((MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) / ((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) + (MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) + (MarketValueofNetDebt != null ? MarketValueofNetDebt : 0))) + (CostofDebt != null ? CostofDebt : 0) * ((MarketValueofNetDebt != null ? MarketValueofNetDebt : 0) / ((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0) + (MarketValueofPreferredEquity != null ? MarketValueofPreferredEquity : 0) + (MarketValueofNetDebt != null ? MarketValueofNetDebt : 0))) * (1 - (MarginalTaxRate!=null ? MarginalTaxRate :0) / 100);

                capitalStructureitem.Value = capitalStructureitem.BasicValue = WACC;
                CapitalStructureOutputList.Add(capitalStructureitem);

            }
            #endregion

            #region Cash Balance

            ////Cash & Equivalent
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Cash & Equivalent";
            capitalStructureitem.HeaderId = (int)HeadersEnum.CashBalance;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;
            //formula= Cash & Equivalent
            capitalStructureitem.BasicValue = cashNEquivalent;
            capitalStructureitem.UnitId = cashNEquivalentObj.UnitId;
            capitalStructureitem.Value = UnitConversion.getConvertedValueforCurrency(capitalStructureitem.UnitId, capitalStructureitem.BasicValue);
            CapitalStructureOutputList.Add(capitalStructureitem);
            
            ////Excess Cash
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Excess Cash";
            capitalStructureitem.HeaderId = (int)HeadersEnum.CashBalance;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId =  (int)CurrencyValueEnum.DollarM;
            //formula= (Cash & Equivalent-Cash Needed for Working Capital)
            capitalStructureitem.BasicValue =(cashNEquivalent != null ? cashNEquivalent : 0) - (workingCapital != null ? workingCapital : 0);
            capitalStructureitem.UnitId = cashNEquivalentObj.UnitId;
            capitalStructureitem.Value = UnitConversion.getConvertedValueforCurrency(capitalStructureitem.UnitId, capitalStructureitem.BasicValue);
            CapitalStructureOutputList.Add(capitalStructureitem);
            #endregion

            #region Leverage Ratios

            ////Debt-to-Equity (D/E) Ratio
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Debt-to-Equity (D/E) Ratio";
            capitalStructureitem.HeaderId = (int)HeadersEnum.LeverageRatios;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.percentage;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = null;
            //formula=  Market Value of Debt (D) / Market Value of Common Equity (E)
            //marketValueofCommonEquity
            capitalStructureitem.Value = capitalStructureitem.BasicValue = (MarketValue_InterestBearingDebt != null ? MarketValue_InterestBearingDebt : 0) * 100 / (marketValueofCommonEquity != null ? marketValueofCommonEquity : 0);
            CapitalStructureOutputList.Add(capitalStructureitem);

            ////Debt-to-Value (D/V) Ratio
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Debt-to-Value (D/V) Ratio";
            capitalStructureitem.HeaderId = (int)HeadersEnum.LeverageRatios;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.percentage;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = null;
            //formula=  Market Value of Debt (D) / Market Value of Common Equity (E) + Market Value of Preferred Equity (P) + Market Value of Debt (D)
            //marketValueofCommonEquity
            capitalStructureitem.Value = capitalStructureitem.BasicValue = (MarketValue_InterestBearingDebt != null ? MarketValue_InterestBearingDebt : 0) * 100 / ((marketValueofCommonEquity != null ? marketValueofCommonEquity : 0)+(MarketValue_InterestBearingDebt != null ? MarketValue_InterestBearingDebt : 0) +(MarketValueofPreferredEquity!=null ? MarketValueofPreferredEquity :0));
            CapitalStructureOutputList.Add(capitalStructureitem);
            #endregion

            #region Internal Valuation

            //Unlevered Enterprise Value
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Unlevered Enterprise Value";
            capitalStructureitem.HeaderId = (int)HeadersEnum.InternalValuation;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = (int)CurrencyValueEnum.DollarM;            
            capitalStructureitem.BasicValue = capitalStructureitem.Value = null;
            CapitalStructureOutputList.Add(capitalStructureitem);

            //Levered Enterprise Value
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Levered Enterprise Value";
            capitalStructureitem.HeaderId = (int)HeadersEnum.InternalValuation;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = (int)CurrencyValueEnum.DollarM;
            capitalStructureitem.BasicValue = capitalStructureitem.Value = null;
            CapitalStructureOutputList.Add(capitalStructureitem);


            //Equity Value
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Equity Value";
            capitalStructureitem.HeaderId = (int)HeadersEnum.InternalValuation;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = (int)CurrencyValueEnum.DollarM;
            capitalStructureitem.BasicValue = capitalStructureitem.Value = null;
            CapitalStructureOutputList.Add(capitalStructureitem);

            //Interest Tax Shield Value
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Interest Tax Shield Value";
            capitalStructureitem.HeaderId = (int)HeadersEnum.InternalValuation;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = (int)CurrencyValueEnum.DollarM;
            capitalStructureitem.BasicValue = capitalStructureitem.Value = null;
            CapitalStructureOutputList.Add(capitalStructureitem);

            //Stock Price
            capitalStructureitem = new CapitalStructure_OutputViewModel();
            capitalStructureitem.LineItem = "Stock Price";
            capitalStructureitem.HeaderId = (int)HeadersEnum.InternalValuation;
            capitalStructureitem.ValueTypeId = (int)ValueTypeEnum.Currency;
            capitalStructureitem.DefaultUnitId = capitalStructureitem.UnitId = (int)CurrencyValueEnum.Dollar;
            capitalStructureitem.BasicValue = capitalStructureitem.Value = null;
            CapitalStructureOutputList.Add(capitalStructureitem);

            #endregion

            tempoutput.CapitalStructureOutputList = CapitalStructureOutputList;
            #endregion


            #region cost of capital

            List<CostofCapital_OutputViewModel> CostofCapitalOutputList = new List<CostofCapital_OutputViewModel>();
            CostofCapital_OutputViewModel CostofCapitalitem;

            #region Current Cost of Capital

            //Adjusted Beta (1/3 + 2/3*raw beta)
            CostofCapitalitem = new CostofCapital_OutputViewModel();
            CostofCapitalitem.LineItem = "Adjusted Beta (1/3 + 2/3*raw beta)";
            CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
            CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.Other;
            CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;           
            CostofCapitalitem.Value = CostofCapitalitem.BasicValue= AdjustedBeta;
            CostofCapitalOutputList.Add(CostofCapitalitem);

            //Market Risk Premium
            CostofCapitalitem = new CostofCapital_OutputViewModel();
            CostofCapitalitem.LineItem = "Market Risk Premium";
            CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
            CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.percentage;
            CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;
            CostofCapitalitem.Value = CostofCapitalitem.BasicValue = MarketRiskPremium;
            CostofCapitalOutputList.Add(CostofCapitalitem);

            if(master.HasEquity==true)
            {
                // Cost of Equity ( rE)
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Cost of Equity ( rE)";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                CostofCapitalitem.Value = CostofCapitalitem.BasicValue = CostofEquity;
                CostofCapitalOutputList.Add(CostofCapitalitem);
            }
           
            if(master.HasPreferredEquity)
            {

                // Cost of Preferred Equity (rP)
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Cost of Preferred Equity (rP)";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                CostofCapitalitem.Value = CostofCapitalitem.BasicValue = CostofPreferredEquity;
                CostofCapitalOutputList.Add(CostofCapitalitem);
            }

            if(master.HasDebt)
            {

                // Cost of Debt (rD)
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Cost of Debt (rD)";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                CostofCapitalitem.Value = CostofCapitalitem.BasicValue = CostofDebt;
                CostofCapitalOutputList.Add(CostofCapitalitem);

            }


            // Unlevered Cost of Capital/Equity (rU)
            CostofCapitalitem = new CostofCapital_OutputViewModel();
            CostofCapitalitem.LineItem = "Unlevered Cost of Capital/Equity (rU)";
            CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
            CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.percentage;
            CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;
            CostofCapitalitem.Value = CostofCapitalitem.BasicValue = UnleveredCostofCapital;
            CostofCapitalOutputList.Add(CostofCapitalitem);

            if(master.LeveragePolicyId==(int)LeveragePolicyEnum.DebtToEquityRatio || master.LeveragePolicyId==(int)LeveragePolicyEnum.AnnuallyAdjustDebt)
            {

                // Weighted Average Cost of Capital (rWACC)
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Weighted Average Cost of Capital (rWACC)";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                CostofCapitalitem.Value = CostofCapitalitem.BasicValue = WACC;
                CostofCapitalOutputList.Add(CostofCapitalitem);

                //  //"Project Risk Adjustment"
                double? ProjectRiskAdjustment = master.CostofCapitalInputList != null && master.CostofCapitalInputList.Count > 0 && master.CostofCapitalInputList.Find(x => x.LineItem == "Project Risk Adjustment" && x.HeaderId == (int)HeadersEnum.OtherInputs) != null ? master.CostofCapitalInputList.Find(x => x.LineItem == "Project Risk Adjustment" && x.HeaderId == (int)HeadersEnum.OtherInputs).BasicValue : 0;

                // Adjusted WACC
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Adjusted WACC";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCostofCapital;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.percentage;
                CostofCapitalitem.DefaultUnitId = capitalStructureitem.UnitId = null;
                CostofCapitalitem.Value = CostofCapitalitem.BasicValue = WACC +(ProjectRiskAdjustment!=null ? ProjectRiskAdjustment :0);
                CostofCapitalOutputList.Add(CostofCapitalitem);

            }

            #endregion

            #region Current Capital Structure

            if(master.HasEquity==true)
            {

                // Market Value Common Stock (E)
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Market Value Common Stock (E)";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCapitalStructure;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                CostofCapitalitem.DefaultUnitId =  (int)CurrencyValueEnum.DollarM;
                CostofCapitalitem.BasicValue = marketValueofCommonEquity;
                CostofCapitalitem.UnitId = MarketvalueofEquityunit;
                CostofCapitalitem.Value = UnitConversion.getConvertedValueforCurrency(CostofCapitalitem.UnitId, CostofCapitalitem.BasicValue);
                CostofCapitalOutputList.Add(CostofCapitalitem);

            }

            if(master.HasPreferredEquity==true)
            {
                // Total Value Preferred Stock (P)
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Total Value Preferred Stock (P)";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCapitalStructure;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                CostofCapitalitem.DefaultUnitId =  (int)CurrencyValueEnum.DollarM;
                CostofCapitalitem.BasicValue = MarketValueofPreferredEquity;
                CostofCapitalitem.UnitId = MarketvalueofPreferredEquityUnit;
                CostofCapitalitem.Value = UnitConversion.getConvertedValueforCurrency(CostofCapitalitem.UnitId, CostofCapitalitem.BasicValue);
                CostofCapitalOutputList.Add(CostofCapitalitem);
            }
            if(master.HasDebt==true)
            {

                // Market Value of Net Debt (D)
                CostofCapitalitem = new CostofCapital_OutputViewModel();
                CostofCapitalitem.LineItem = "Market Value of Net Debt (D)";
                CostofCapitalitem.HeaderId = (int)HeadersEnum.CurrentCapitalStructure;
                CostofCapitalitem.ValueTypeId = (int)ValueTypeEnum.Currency;
                CostofCapitalitem.DefaultUnitId = (int)CurrencyValueEnum.DollarM;
                CostofCapitalitem.BasicValue = MarketValueofNetDebt;
                // CostofCapitalitem.BasicValue = (MarketValue_InterestBearingDebt != null ? MarketValue_InterestBearingDebt : 0) - ((cashNEquivalent != null ? cashNEquivalent : 0) - (workingCapital != null ? workingCapital : 0));
                CostofCapitalitem.UnitId = UnitConversion.getHigherDenominationUnit(MarketvalueofDebtUnit, CheckUnit);
                CostofCapitalitem.Value = UnitConversion.getConvertedValueforCurrency(CostofCapitalitem.UnitId, CostofCapitalitem.BasicValue);
                CostofCapitalOutputList.Add(CostofCapitalitem);
            }


            #endregion

            tempoutput.CostofCapitalOutputList = CostofCapitalOutputList;

            #endregion

            return tempoutput;

        }

      

        [HttpPost]
        [Route("SaveCostofCapitalNStructure")]
        public ActionResult<Object> SaveCostofCapitalNStructure([FromBody] MasterCostofCapitalNStructureViewModel model)
        {
            try
            {               
                    //map MasterVM to master table
                    MasterCostofCapitalNStructure tblMaster = mapper.Map<MasterCostofCapitalNStructureViewModel, MasterCostofCapitalNStructure>(model);

                if (tblMaster != null)
                {
                    if(tblMaster.Id==0)
                        iMasterCostofCapitalNStructure.Add(tblMaster);
                    else
                    {
                        tblMaster.ModifiedDate = System.DateTime.Now;
                        iMasterCostofCapitalNStructure.Update(tblMaster);
                    }

                    iMasterCostofCapitalNStructure.Commit();

                    // Capital Structure
                    if (model.CapitalStructureInputList != null && model.CapitalStructureInputList.Count > 0)
                    {
                        List<CapitalStructure_Input> tblcapitalStructureListObj = new List<CapitalStructure_Input>();
                        CapitalStructure_Input tblCapitalStructure = new CapitalStructure_Input();
                        foreach (CapitalStructure_InputViewModel obj in model.CapitalStructureInputList)
                        {
                            obj.MasterId = tblMaster.Id;
                            if (obj.UnitId != null && obj.UnitId != 0)
                            {
                                if (obj.ValueTypeId == (int)ValueTypeEnum.Number)
                                    obj.BasicValue = UnitConversion.getBasicValueforNumbers(obj.UnitId, obj.Value);

                                if (obj.ValueTypeId == (int)ValueTypeEnum.Currency)
                                    obj.BasicValue = UnitConversion.getBasicValueforCurrency(obj.UnitId, obj.Value);
                            }
                            else
                            {
                                obj.BasicValue = obj.Value;
                            }
                            //map CapitalStructure ViewModel to table
                            tblCapitalStructure = mapper.Map<CapitalStructure_InputViewModel, CapitalStructure_Input>(obj);
                            tblcapitalStructureListObj.Add(tblCapitalStructure);
                        }
                        if (tblcapitalStructureListObj != null && tblcapitalStructureListObj.Count > 0)
                        {
                            if (model.Id == 0)
                                iCapitalStructure_Input.AddMany(tblcapitalStructureListObj);
                            else
                                iCapitalStructure_Input.UpdatedMany(tblcapitalStructureListObj) ;
                            iCapitalStructure_Input.Commit();
                        }
                    }

                    //Cost of Capital 
                    if (model.CostofCapitalInputList != null && model.CostofCapitalInputList.Count > 0)
                    {
                        List<CostofCapital_Input> tblCostofcapitalListObj = new List<CostofCapital_Input>();
                        CostofCapital_Input tblCostofCapital = new CostofCapital_Input();
                        foreach (CostofCapital_InputViewModel obj in model.CostofCapitalInputList)
                        {
                            obj.MasterId = tblMaster.Id;
                            if (obj.UnitId != null && obj.UnitId != 0)
                            {
                                if (obj.ValueTypeId == (int)ValueTypeEnum.Number)
                                    obj.BasicValue = UnitConversion.getBasicValueforNumbers(obj.UnitId, obj.Value);

                                if (obj.ValueTypeId == (int)ValueTypeEnum.Currency)
                                    obj.BasicValue = UnitConversion.getBasicValueforCurrency(obj.UnitId, obj.Value);
                            }
                            else
                            {
                                obj.BasicValue = obj.Value;
                            }
                            //map CapitalStructure ViewModel to table
                            tblCostofCapital = mapper.Map<CostofCapital_InputViewModel, CostofCapital_Input>(obj);
                            tblCostofcapitalListObj.Add(tblCostofCapital);
                        }
                        if (tblCostofcapitalListObj != null && tblCostofcapitalListObj.Count > 0)
                        {
                            if (model.Id == 0)
                                iCostofCapital_Input.AddMany(tblCostofcapitalListObj);
                            else
                                iCostofCapital_Input.UpdatedMany(tblCostofcapitalListObj);

                            iCostofCapital_Input.Commit();
                        }
                    }
                }
                    return Ok(new { StatusCode = 1, message = "Saved Successfully" });
              
            }catch(Exception ss)
            {
                return BadRequest(ss);
            }
        }
        
        private List<CostofCapital_InputViewModel> getInitialCostofCapitalList()
        {
            List<CostofCapital_InputViewModel> CostofCapitalInputList = new List<CostofCapital_InputViewModel>();
            CostofCapital_InputViewModel item;

            #region Cost of Equity Capital
            //Risk Free Rate (Future)
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Risk Free Rate (Future)";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Historical Market Return
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Historical Market Return";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Historical Risk Free Return
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Historical Risk Free Return";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Small Stock Premium
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Small Stock Premium";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Raw Beta
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Raw Beta Source";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.Other;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);


           

            //External Beta Source
            item = new CostofCapital_InputViewModel();
            item.LineItem = "External Source";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.Other;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);
            
            //Manual Beta
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Manual Beta";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.Other;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //External Source Beta
            item = new CostofCapital_InputViewModel();
            item.LineItem = "External Beta";
            item.SubHeader = "External Source";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.Other;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //External Source Beta
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Calculated Beta";
            item.HeaderId = (int)HeadersEnum.CostofEquityCapital;
            item.ValueTypeId = (int)ValueTypeEnum.Other;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            #endregion

            #region Cost of Debt Capital

            //Method 1:  Forecasted Rate on New Debt Issuance
            //Risk Free Rate
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Risk Free Rate";
            item.HeaderId = (int)HeadersEnum.CostofDebtCapital;
            item.SubHeader = "Method 1";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Default Spread
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Default Spread";
            item.HeaderId = (int)HeadersEnum.CostofDebtCapital;
            item.SubHeader = "Method 1";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            // Method 2:  Current Average Rate on Outstanding Debt
            //Yield to Maturity (y)
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Yield to Maturity (y)";
            item.HeaderId = (int)HeadersEnum.CostofDebtCapital;
            item.SubHeader = "Method 2";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Probability of Default (p)
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Probability of Default (p)";
            item.HeaderId = (int)HeadersEnum.CostofDebtCapital;
            item.SubHeader = "Method 2";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Expected Loss Rate / Default rate (L)
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Expected Loss Rate / Default rate (L)";
            item.HeaderId = (int)HeadersEnum.CostofDebtCapital;
            item.SubHeader = "Method 2";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            #endregion

            #region Other Inouts

            //Marginal Tax Rate (Tc)
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Marginal Tax Rate (Tc)";
            item.HeaderId = (int)HeadersEnum.OtherInputs;
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            //Project Risk Adjustment
            item = new CostofCapital_InputViewModel();
            item.LineItem = "Project Risk Adjustment";
            item.HeaderId = (int)HeadersEnum.OtherInputs;
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            CostofCapitalInputList.Add(item);

            #endregion

            return CostofCapitalInputList;
        }

        private List<CapitalStructure_InputViewModel> getInitialCapitalStructureList()
        {
            List<CapitalStructure_InputViewModel> capitalStructureInputList = new List<CapitalStructure_InputViewModel>();
            CapitalStructure_InputViewModel item;

            #region Source of Financing

            //Current Share Price 
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Current Share Price";
            item.HeaderId = (int)HeadersEnum.SourceOfFinancing;
            item.SubHeader = "Equity";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.Dollar;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            //Number of Shares Outstanding - Basic (Millions)
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Number of Shares Outstanding - Basic";
            item.HeaderId = (int)HeadersEnum.SourceOfFinancing;
            item.SubHeader = "Equity";
            item.ValueTypeId = (int)ValueTypeEnum.Number;
            item.DefaultUnitId = item.UnitId = (int)NumberCountEnum.EachM;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            //Number of Shares Outstanding - Diluted (Millions)
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Number of Shares Outstanding - Diluted";
            item.HeaderId = (int)HeadersEnum.SourceOfFinancing;
            item.SubHeader = "Equity";
            item.ValueTypeId = (int)ValueTypeEnum.Number;
            item.DefaultUnitId = item.UnitId = (int)NumberCountEnum.EachM;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            //Current Preferred Share Price
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Current Preferred Share Price";
            item.HeaderId = (int)HeadersEnum.SourceOfFinancing;
            item.SubHeader = "Preferred Equity";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.Dollar;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            //Number of Preferred Shares Outstanding
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Number of Preferred Shares Outstanding";
            item.HeaderId = (int)HeadersEnum.SourceOfFinancing;
            item.SubHeader = "Preferred Equity";
            item.ValueTypeId = (int)ValueTypeEnum.Number;
            item.DefaultUnitId = item.UnitId = (int)NumberCountEnum.EachM;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            //Preferred Dividend
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Preferred Dividend";
            item.HeaderId = (int)HeadersEnum.SourceOfFinancing;
            item.SubHeader = "Preferred Equity";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.Dollar;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            ////Market Value of Interest Bearing Debt
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Market Value of Interest Bearing Debt";
            item.HeaderId = (int)HeadersEnum.SourceOfFinancing;
            item.SubHeader = "Debt";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.DollarM;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            #endregion

            #region Other Inputs

            //Cash & Equivalent
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Cash & Equivalent";
            item.HeaderId = (int)HeadersEnum.OtherInputs;
            item.SubHeader = "";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.DollarM;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            //Cash Needed for Working Capital
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Cash Needed for Working Capital";
            item.HeaderId = (int)HeadersEnum.OtherInputs;
            item.SubHeader = "";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.DollarM;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            //Interest Coverage Ratio (k) - If Applicable
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Interest Coverage Ratio (k) - If Applicable";
            item.HeaderId = (int)HeadersEnum.OtherInputs;
            item.SubHeader = "";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            ////Marginal Tax Rate (Tc)
            //item = new CapitalStructure_InputViewModel();
            //item.LineItem = "Marginal Tax Rate (Tc)";
            //item.HeaderId = (int)HeadersEnum.OtherInputs;
            //item.SubHeader = "";
            //item.ValueTypeId = (int)ValueTypeEnum.percentage;
            //item.DefaultUnitId = item.UnitId = null;
            //item.BasicValue = item.Value = null;
            //capitalStructureInputList.Add(item);

            //Free Cash Flow- Next Year
            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Free Cash Flow- Next Year-if Applicable";
            item.HeaderId = (int)HeadersEnum.OtherInputs;
            item.SubHeader = "";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.DollarM;
            item.BasicValue = item.Value = null;
            capitalStructureInputList.Add(item);

            #endregion

            #region Rahi || Scenario Analysis line item || 16-11-2020

            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Target Debt-to-Equity (D/E) Ratio";
            item.HeaderId = 3;
            item.SubHeader = "Target Capital Structure";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            item.ListType = "Scenario Analysis";
            item.LeverageType = "1";
            capitalStructureInputList.Add(item);

            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Cost of Debt (rD)";
            item.HeaderId = 3;
            item.SubHeader = "New Cost to Capital";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            item.ListType = "Scenario Analysis";
            item.LeverageType = "0";
            capitalStructureInputList.Add(item);

            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Value of Permanent Debt";
            item.HeaderId = 3;
            item.SubHeader = "New Permanent Debt";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.DollarM;
            item.BasicValue = item.Value = null;
            item.ListType = "Scenario Analysis";
            item.LeverageType = "3";
            capitalStructureInputList.Add(item);

            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Interest Coverage Ratio";
            item.HeaderId = 3;
            item.SubHeader = "Coverage Ratio";
            item.ValueTypeId = (int)ValueTypeEnum.percentage;
            item.DefaultUnitId = item.UnitId = null;
            item.BasicValue = item.Value = null;
            item.ListType = "Scenario Analysis";
            item.LeverageType = "4";
            capitalStructureInputList.Add(item);

            item = new CapitalStructure_InputViewModel();
            item.LineItem = "Free Cash Flow- Next Year";
            item.HeaderId = 3;
            item.SubHeader = "Free Cash Flow";
            item.ValueTypeId = (int)ValueTypeEnum.Currency;
            item.DefaultUnitId = item.UnitId = (int)CurrencyValueEnum.DollarM;
            item.BasicValue = item.Value = null;
            item.ListType = "Scenario Analysis";
            item.LeverageType = "4";
            capitalStructureInputList.Add(item);

            #endregion

            return capitalStructureInputList;
        }


        [HttpPost]
        [Route("SaveSnapshotCostofCapitalNStructure")]
        public ActionResult<Object> SaveSnapshotCostofCapitalNStructure([FromBody] Snapshot_CostofCapitalNStructureViewModel model)
        {
            try
            {
                if (model != null)
                {
                    model.CreatedDate = DateTime.UtcNow;

                    //map MasterVM to master table
                    Snapshot_CostofCapitalNStructure tblMaster = mapper.Map<Snapshot_CostofCapitalNStructureViewModel, Snapshot_CostofCapitalNStructure>(model);

                    if(tblMaster!=null)
                    {
                        // save Master Snapshot
                        iSnapshot_CostofCapitalNStructure.Add(tblMaster);
                        iSnapshot_CostofCapitalNStructure.Commit();


                        // save Capiral Structure Snapshot
                        if(model.capitalStructure_SnapshotList!=null && model.capitalStructure_SnapshotList.Count>0)
                        {
                            List<CapitalStructure_Snapshot> tblCapitalStructureList = new List<CapitalStructure_Snapshot>();
                            CapitalStructure_Snapshot tempCapitalStructure;
                            foreach (CapitalStructure_SnapshotViewModel capitalStructureObj in model.capitalStructure_SnapshotList)
                            {
                                tempCapitalStructure = new CapitalStructure_Snapshot();
                                capitalStructureObj.Snapshot_CostofCapitalNStructureId = tblMaster.Id;
                                //map Capital Structure Snapshot VM to  table
                                tempCapitalStructure = mapper.Map<CapitalStructure_SnapshotViewModel, CapitalStructure_Snapshot>(capitalStructureObj);
                                if(tempCapitalStructure!=null)
                                tblCapitalStructureList.Add(tempCapitalStructure);
                            }
                            iCapitalStructure_Snapshot.AddMany(tblCapitalStructureList);
                            iCapitalStructure_Snapshot.Commit();
                        }


                    }


                    // save Cost of capital Snapshot
                    if (model.CostofCapital_SnapshotList != null && model.CostofCapital_SnapshotList.Count > 0)
                    {
                        List<CostofCapital_Snapshot> tblCostofCapitalList = new List<CostofCapital_Snapshot>();
                        CostofCapital_Snapshot tempCostofCapital;
                        foreach (CostofCapital_SnapshotViewModel capitalStructureObj in model.CostofCapital_SnapshotList)
                        {
                            tempCostofCapital = new CostofCapital_Snapshot();
                            capitalStructureObj.Snapshot_CostofCapitalNStructureId = tblMaster.Id;
                            //map Capital Structure Snapshot VM to  table
                            tempCostofCapital = mapper.Map<CostofCapital_SnapshotViewModel, CostofCapital_Snapshot>(capitalStructureObj);
                            if (tempCostofCapital != null)
                                tblCostofCapitalList.Add(tempCostofCapital);
                        }
                        iCostofCapital_Snapshot.AddMany(tblCostofCapitalList);
                        iCostofCapital_Snapshot.Commit();
                    }



                }

            }
            catch(Exception ss)
            {
                return BadRequest(new { StatusCode = 2, message = "error Occurred" + ss });
            }
            return Ok(new { StatusCode = 1, message = "Saved Successfully" });
        }


        #region Rahi Gadiya || Get API on Scenario Analysis || 05-11-2020

        [HttpGet]
        [Route("GetCapitalStructures_new/{UserId}")]
        public ActionResult<Object> GetCapitalStructures_new(long UserId)
        {

            MasterCostofCapitalNStructureViewModel result = new MasterCostofCapitalNStructureViewModel();
            try
            {
                //check if master data is saved or not
                MasterCostofCapitalNStructure tblMaster = iMasterCostofCapitalNStructure.GetSingle(x => x.UserId == UserId);

                if (tblMaster != null)
                {
                    result = mapper.Map<MasterCostofCapitalNStructure, MasterCostofCapitalNStructureViewModel>(tblMaster);
                   
                    if (tblMaster != null && tblMaster.SALeveragePolicyID != null && tblMaster.SALeveragePolicyID != 0)
                    {

                        //bind data from Saved Values
                        // result = mapper.Map<MasterCostofCapitalNStructure, MasterCostofCapitalNStructureViewModel>(tblMaster);

                        //get Capital Stryucture
                        List<CapitalStructure_Input> tblcapitalStructureInputList = iCapitalStructure_Input.FindBy(x => x.MasterId == tblMaster.Id).ToList();
                        //Map Capital Structure
                        result.CapitalStructureInputList = new List<CapitalStructure_InputViewModel>();
                        if (tblcapitalStructureInputList != null)
                            foreach (var obj in tblcapitalStructureInputList.OrderBy(x => x.Id))
                            {
                                CapitalStructure_InputViewModel tempVM = mapper.Map<CapitalStructure_Input, CapitalStructure_InputViewModel>(obj);
                                if (tempVM.SubHeader != "Equity" && tempVM.SubHeader != "Debt" && tempVM.SubHeader != "Preferred Equity" && tempVM.SubHeader != null && tempVM.SubHeader != "")
                                {
                                    tempVM.ListType = "Scenario Analysis";

                                }
                                else
                                {
                                    tempVM.ListType = null;
                                }
                                result.CapitalStructureInputList.Add(tempVM);
                            }

                        //get Cost of Capital
                        List<CostofCapital_Input> tblCostofcapitalInputList = iCostofCapital_Input.FindBy(x => x.MasterId == tblMaster.Id).ToList();
                        //Map Cost of capital
                        result.CostofCapitalInputList = new List<CostofCapital_InputViewModel>();
                        if (tblCostofcapitalInputList != null)
                            foreach (var obj in tblCostofcapitalInputList.OrderBy(x => x.Id))
                            {
                                CostofCapital_InputViewModel tempVM = mapper.Map<CostofCapital_Input, CostofCapital_InputViewModel>(obj);
                                result.CostofCapitalInputList.Add(tempVM);
                            }

                        //get output for Capital Structure & Cost of Capital
                        MasterCostofCapitalNStructureViewModel tempOutput = GetOutPutforCapitalNStructure(result);

                        if (tempOutput != null)
                        {
                            result.CapitalStructureOutputList = tempOutput.CapitalStructureOutputList != null && tempOutput.CapitalStructureOutputList.Count > 0 ? tempOutput.CapitalStructureOutputList : new List<CapitalStructure_OutputViewModel>();
                        }

                            //Get Output for Scenario Analysis
                            MasterCostofCapitalNStructureViewModel tempOutputofScenarioAnalysis = GetOutPutforScenarioAnalysis(result);

                            if (tempOutputofScenarioAnalysis != null)
                            {
                                result.ScenarioAnalysisOutputList = tempOutputofScenarioAnalysis.ScenarioAnalysisOutputList != null && tempOutputofScenarioAnalysis.ScenarioAnalysisOutputList.Count > 0 ? tempOutputofScenarioAnalysis.ScenarioAnalysisOutputList : new List<ScenarioAnalysis_OutputViewModel>();
                            }

                        //result.CapitalStructureInputList = null;
                        //result.CapitalStructureOutputList = null;

                    }
                    else
                    {
                        //bydefault first leverage policy would be selected
                       // result.Id = tblMaster.Id;
                        result.UserId = UserId;
                       // result.LeveragePolicyId = (int)LeveragePolicyEnum.DebtToEquityRatio;
                        result.SALeveragePolicyID = (int)ScenarioAnalysisLeveragePolicyEnum.SADebtToEquityRatio;
                      //  result.HasEquity = true;
                      //  result.HasPreferredEquity = true;
                      //   result.HasDebt = true;
                      //  result.BetaSourceId = null;
                      //  result.CostofDebtMethodId = 1;
                        result.CreatedDate = result.ModifiedDate = System.DateTime.Now;

                        //Create data for first time
                        result.CapitalStructureInputList = new List<CapitalStructure_InputViewModel>();
                        result.CapitalStructureInputList = getInitialCapitalStructureList();

                        //Create Cost of capital Inputs for first time
                        result.CostofCapitalInputList = new List<CostofCapital_InputViewModel>();
                        result.CostofCapitalInputList = getInitialCostofCapitalList();                       
                    }

                    // get drop down from enum
                    result.CurrencyValueList = EnumHelper.GetEnumListbyName<CurrencyValueEnum>(); //get CurrencyValueList
                    result.NumberCountList = EnumHelper.GetEnumListbyName<NumberCountEnum>(); //get Number countList
                    result.ValueTypeList = EnumHelper.GetEnumListbyName<ValueTypeEnum>(); //get Value Type List
                    result.LeveragePolicyList = EnumHelper.GetEnumListbyName<LeveragePolicyEnum>(); //get leverage Policy List
                    result.SALeveragePolicyList = EnumHelper.GetEnumListbyName<ScenarioAnalysisLeveragePolicyEnum>(); //get leverage Policy List
                    result.HeaderList = EnumHelper.GetEnumListbyName<HeadersEnum>(); //get leverage Policy List
                    result.ScenarioAnalysisHeaderList = EnumHelper.GetEnumListbyName<ScenarioAnalysisHeadersEnum>(); //get Scenario Analysis leverage Policy List
                    result.BetasourceList = EnumHelper.GetEnumListbyName<BetaSourceEnum>(); //get Beta Source List

                }
              
            }
            catch (Exception ss)
            {
                Console.WriteLine(ss.Message);

            }
            return result;
        }

        #endregion


        [HttpGet]
        [Route("GetCapitalStructures/{UserId}")]
        public ActionResult<Object> GetCapitalStructures(long UserId)
        {
            try
            {

                var CapitalStructure = iCapitalStructure.FindBy(s => s.UserId == UserId).OrderByDescending(x=>x.Id).ToArray();

                if (CapitalStructure.Length != 0)
                {
                    if(CapitalStructure[0].equity!=null)
                    {
                        var currentShare =!string.IsNullOrEmpty(CapitalStructure[0].equityUnits.CurrentSharePriceUnit) ? CapitalStructure[0].equity.CurrentSharePrice / UnitConversion.ReturnDividend(CapitalStructure[0].equityUnits.CurrentSharePriceUnit) : CapitalStructure[0].equity.CurrentSharePrice;
                        var numberBasic =!string.IsNullOrEmpty(CapitalStructure[0].equityUnits.NumberShareBasicUnit) ? CapitalStructure[0].equity.NumberShareBasic / UnitConversion.ReturnDividend(CapitalStructure[0].equityUnits.NumberShareBasicUnit) : CapitalStructure[0].equity.NumberShareBasic;
                        var numberOut =!string.IsNullOrEmpty(CapitalStructure[0].equityUnits.NumberShareOutstandingUnit) ? CapitalStructure[0].equity.NumberShareOutstanding / UnitConversion.ReturnDividend(CapitalStructure[0].equityUnits.NumberShareOutstandingUnit) : CapitalStructure[0].equity.NumberShareOutstanding;

                        Equity equity = new Equity
                        {
                            CurrentSharePrice = currentShare,
                            NumberShareBasic = numberBasic,
                            NumberShareOutstanding = numberOut,
                            CostOfEquity = Convert.ToDouble(decimal.Round(Convert.ToDecimal(CapitalStructure[0].equity.CostOfEquity), 2))
                        };
                        CapitalStructure[0]._Equity = JsonConvert.SerializeObject(equity);
                    }
                   
                    
                    if(CapitalStructure[0].prefferedEquity!=null)
                    {
                        var preffShare = !string.IsNullOrEmpty(CapitalStructure[0].prefferedEquityUnit.PrefferedDividendUnit) ? CapitalStructure[0].prefferedEquity.PrefferedSharePrice / UnitConversion.ReturnDividend(CapitalStructure[0].prefferedEquityUnit.PrefferedDividendUnit) : CapitalStructure[0].prefferedEquity.PrefferedSharePrice;
                        var preffShareOut = !string.IsNullOrEmpty(CapitalStructure[0].prefferedEquityUnit.PrefferedShareOutstandingUnit) ? CapitalStructure[0].prefferedEquity.PrefferedShareOutstanding / UnitConversion.ReturnDividend(CapitalStructure[0].prefferedEquityUnit.PrefferedShareOutstandingUnit) : CapitalStructure[0].prefferedEquity.PrefferedShareOutstanding;
                        var preffDivi = !string.IsNullOrEmpty(CapitalStructure[0].prefferedEquityUnit.PrefferedDividendUnit) ? CapitalStructure[0].prefferedEquity.PrefferedDividend / UnitConversion.ReturnDividend(CapitalStructure[0].prefferedEquityUnit.PrefferedDividendUnit) : CapitalStructure[0].prefferedEquity.PrefferedDividend;

                        PrefferedEquity preffEq = new PrefferedEquity
                        {
                            PrefferedSharePrice = preffShare,
                            PrefferedShareOutstanding = preffShareOut,
                            PrefferedDividend = preffDivi,
                            CostPreffEquity = CapitalStructure[0].prefferedEquity.CostPreffEquity
                        };
                        CapitalStructure[0]._PrefferedEquity = JsonConvert.SerializeObject(preffEq);
                    }

                    if(CapitalStructure[0].debt!=null)
                    {
                        var marketValue =!string.IsNullOrEmpty(CapitalStructure[0].debtUnit.MarketValueDebtUnit) ? CapitalStructure[0].debt.MarketValueDebt / UnitConversion.ReturnDividend(CapitalStructure[0].debtUnit.MarketValueDebtUnit) : CapitalStructure[0].debt.MarketValueDebt;
                        Debt debt = new Debt
                        {
                            MarketValueDebt = marketValue,
                            CostOfDebt = CapitalStructure[0].debt.CostOfDebt
                        };
                        CapitalStructure[0]._Debt = JsonConvert.SerializeObject(debt);

                    }

                    CapitalStructure[0].CashEquivalent =!string.IsNullOrEmpty(CapitalStructure[0].CashEquivalentUnit) ? CapitalStructure[0].CashEquivalent / UnitConversion.ReturnDividend(CapitalStructure[0].CashEquivalentUnit) : CapitalStructure[0].CashEquivalent;
                    CapitalStructure[0].CashNeededCapital =!string.IsNullOrEmpty(CapitalStructure[0].CashNeededCapitalUnit) ? CapitalStructure[0].CashNeededCapital / UnitConversion.ReturnDividend(CapitalStructure[0].CashNeededCapitalUnit) : CapitalStructure[0].CashNeededCapital;
                    if (CapitalStructure[0].FreeCashFlow != 0)
                    {
                        CapitalStructure[0].FreeCashFlow =!string.IsNullOrEmpty(CapitalStructure[0].FreeCashFlowUnit) ? CapitalStructure[0].FreeCashFlow / UnitConversion.ReturnDividend(CapitalStructure[0].FreeCashFlowUnit) : CapitalStructure[0].FreeCashFlow;
                    }
                    
                }
                if (CapitalStructure == null)
                {
                    return NotFound("No Data Found");
                }

                return Ok(CapitalStructure);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        [Route("GetCapitalStructure/{Id}")]
        public ActionResult<Object> GetCapitalStructure(long Id)
        {
            // if (Id == 0)
            // {
            //     return BadRequest();
            // }
            try
            {
                var CapitalStructure = iCapitalStructure.GetSingle(s => s.Id == Id);
                Console.WriteLine("CapitalStructure");
                if (CapitalStructure == null)
                {
                    Console.WriteLine("CapitalStructure 2");
                    return NotFound("No Data Found");
                }
                return Ok(CapitalStructure);
            }
            catch (Exception ex)
            {
                Console.WriteLine("CapitalStructure 3");
                return BadRequest(ex.Message);
            }
        }

        [HttpPost]
        [Route("AddCapitalStructure")]
        public ActionResult<Object> AddCapitalStructure([FromBody] CapitalStructureViewModel model)
        {

            try
            {
                if (model != null)
                { 
                double[] values = { (model.equity !=null ? model.equity.CurrentSharePrice: 0),
                        (model.equity !=null ?model.equity.NumberShareBasic :0),
                        (model.equity !=null ?model.equity.NumberShareOutstanding :0),
                 (model.prefferedEquity !=null ?model.prefferedEquity.PrefferedSharePrice:0),
                        (model.prefferedEquity !=null ?model.prefferedEquity.PrefferedShareOutstanding :0),
                        (model.prefferedEquity !=null ?model.prefferedEquity.PrefferedDividend :0),
                 (model.debt !=null ? model.debt.MarketValueDebt :0),
                        model.CashEquivalent,model.CashNeededCapital , model.FreeCashFlow };
                string[] units = { (model.equityUnits!=null ? model.equityUnits.CurrentSharePriceUnit : ""),
                        (model.equityUnits!=null ?model.equityUnits.NumberShareBasicUnit :""),
                        (model.equityUnits!=null ? model.equityUnits.NumberShareOutstandingUnit :""),
                   (model.prefferedEquityUnit!=null ? model.prefferedEquityUnit.PrefferedSharePriceUnit:""),
                        (model.prefferedEquityUnit!=null ? model.prefferedEquityUnit.PrefferedShareOutstandingUnit :""),
                        (model.prefferedEquityUnit!=null ? model.prefferedEquityUnit.PrefferedDividendUnit :""),
                 (model.debtUnit!=null ? model.debtUnit.MarketValueDebtUnit :""),model.CashEquivalentUnit, model.CashNeededCapitalUnit, model.FreeCashFlowUnit };
                var output = UnitConversion.ConvertUnits(values, units, 0);

                model.equity = new Equity { CurrentSharePrice = output[0], NumberShareBasic = output[1], NumberShareOutstanding = output[2], CostOfEquity = model.equity !=null ?model.equity.CostOfEquity : 0 };
                if(model.prefferedEquity!=null)
                model.prefferedEquity = new PrefferedEquity { PrefferedSharePrice = output[3], PrefferedShareOutstanding = output[4], PrefferedDividend = output[5], CostPreffEquity = model.prefferedEquity!= null ? model.prefferedEquity.CostPreffEquity :0 };
                if(model.debt !=null)
                model.debt = new Debt { MarketValueDebt = output[6], CostOfDebt = model.debt != null ? model.debt.CostOfDebt :0 };
                model.CashEquivalent = output[7];
                model.CashNeededCapital = output[8];
                model.FreeCashFlow = output[9];


                var summaryOutput = MathCapStructure.SummaryOutput(model);
                model.SummaryOutput = JsonConvert.SerializeObject(summaryOutput);
                SummaryOutput summOut = JsonConvert.DeserializeObject<SummaryOutput>(model.SummaryOutput);

                var outputs = UnitConversion.ConvertOutputUnits(out _, summOut.marketValueEquity, summOut.marketValuePreferredEquity, summOut.marketValueDebt);
                Dictionary<string, object> results = new Dictionary<string, object>();
                results.Add("equityVal", outputs[0]);
                results.Add("preffEquityVal", outputs[1]);
                results.Add("debtVal", outputs[2]);

                Dictionary<string, object> result = new Dictionary<string, object>();
                result.Add("summaryOutput", summaryOutput);
                result.Add("piechartvalues", results);
                model.CapitalPieChart = JsonConvert.SerializeObject(results);

                if (model.Id == 0)
                {
                        //CapitalStructure capitalStructure = new CapitalStructure();

                        //capitalStructure.NewLeveragePolicy = model.NewLeveragePolicy;
                        //capitalStructure.CashEquivalentUnit = model.CashEquivalentUnit;
                        //capitalStructure.CashNeededCapitalUnit = model.CashNeededCapitalUnit;
                        //capitalStructure.InterestCoverage = model.InterestCoverage;
                        //capitalStructure.MarginalTaxRate = model.MarginalTaxRate;
                        //capitalStructure.FreeCashFlowUnit = model.FreeCashFlowUnit;
                        //capitalStructure.ApprovalFlag = model.ApprovalFlag;
                        //capitalStructure.UserId = model.UserId;
                        //capitalStructure.SummaryOutput = model.SummaryOutput;
                        //capitalStructure.ScenarioObject = model.ScenarioObject;
                        //capitalStructure.ScenarioOutput = model.ScenarioOutput;
                        //capitalStructure.ScenarioPolicy = model.ScenarioPolicy;
                        //capitalStructure.PermanentDebtUnit = model.PermanentDebtUnit;
                        //capitalStructure.ScenarioFreeUnit = model.ScenarioFreeUnit;
                        //capitalStructure.SummaryFlag = model.SummaryFlag;
                        //capitalStructure.CapitalPieChart = model.CapitalPieChart;
                        //capitalStructure.ScenarioPieChart = model.ScenarioPieChart;
                        //capitalStructure.equity = new Equity { CurrentSharePrice = output[0], NumberShareBasic = output[1], NumberShareOutstanding = output[2], CostOfEquity = model.equity.CostOfEquity };
                        //capitalStructure.prefferedEquity = new PrefferedEquity { PrefferedSharePrice = output[3], PrefferedShareOutstanding = output[4], PrefferedDividend = output[5], CostPreffEquity = model.prefferedEquity.CostPreffEquity };
                        //capitalStructure.debt = new Debt { MarketValueDebt = output[6], CostOfDebt = model.debt.CostOfDebt };
                        //capitalStructure.CashEquivalent = output[7];
                        //capitalStructure.CashNeededCapital = output[8];
                        //capitalStructure.FreeCashFlow = output[9];
                        //capitalStructure.debtUnit = model.debtUnit;
                        //capitalStructure.equityUnits = model.equityUnits;
                        //capitalStructure.prefferedEquityUnit = model.prefferedEquityUnit;

                        CapitalStructure capitalStructure = new CapitalStructure
                        {
                            NewLeveragePolicy = model.NewLeveragePolicy,
                            CashEquivalentUnit = model.CashEquivalentUnit,
                            CashNeededCapitalUnit = model.CashNeededCapitalUnit,
                            InterestCoverage = model.InterestCoverage,
                            MarginalTaxRate = model.MarginalTaxRate,
                            FreeCashFlowUnit = model.FreeCashFlowUnit,
                            ApprovalFlag = model.ApprovalFlag,
                            UserId = model.UserId,
                            SummaryOutput = model.SummaryOutput,
                            ScenarioObject = model.ScenarioObject,
                            ScenarioOutput = model.ScenarioOutput,
                            ScenarioPolicy = model.ScenarioPolicy,
                            PermanentDebtUnit = model.PermanentDebtUnit,
                            ScenarioFreeUnit = model.ScenarioFreeUnit,
                            SummaryFlag = model.SummaryFlag,
                            CapitalPieChart = model.CapitalPieChart,
                            ScenarioPieChart = model.ScenarioPieChart,
                            equity = new Equity { CurrentSharePrice = output[0], NumberShareBasic = output[1], NumberShareOutstanding = output[2], CostOfEquity = (model.equity != null ? model.equity.CostOfEquity : 0) },
                            prefferedEquity = new PrefferedEquity { PrefferedSharePrice = output[3], PrefferedShareOutstanding = output[4], PrefferedDividend = output[5], CostPreffEquity = (model.prefferedEquity!=null ? model.prefferedEquity.CostPreffEquity :0) },
                            debt = new Debt { MarketValueDebt = output[6], CostOfDebt = (model.debt != null ? model.debt.CostOfDebt : 0) },
                            CashEquivalent = output[7],
                            CashNeededCapital = output[8],
                            FreeCashFlow = output[9],
                            debtUnit = (model.debtUnit != null ? model.debtUnit : new DebtUnit()),
                            equityUnits = (model.equityUnits != null ? model.equityUnits : new EquityUnit()),
                            prefferedEquityUnit = (model.prefferedEquityUnit != null ? model.prefferedEquityUnit : new PrefferedEquityUnit())
                        };
                        iCapitalStructure.Add(capitalStructure);
                    iCapitalStructure.Commit();
                    return Ok(result);


                }
                else
                {
                    CapitalStructure capitalStructure = new CapitalStructure
                    {
                        Id = model.Id,
                        NewLeveragePolicy = model.NewLeveragePolicy,
                        CashEquivalentUnit = model.CashEquivalentUnit,
                        CashNeededCapitalUnit = model.CashNeededCapitalUnit,
                        InterestCoverage = model.InterestCoverage,
                        MarginalTaxRate = model.MarginalTaxRate,
                        FreeCashFlowUnit = model.FreeCashFlowUnit,
                        ApprovalFlag = model.ApprovalFlag,
                        UserId = model.UserId,
                        SummaryOutput = model.SummaryOutput,
                        ScenarioObject = model.ScenarioObject,
                        ScenarioOutput = model.ScenarioOutput,
                        ScenarioPolicy = model.ScenarioPolicy,
                        PermanentDebtUnit = model.PermanentDebtUnit,
                        ScenarioFreeUnit = model.ScenarioFreeUnit,
                        SummaryFlag = model.SummaryFlag,
                        CapitalPieChart = model.CapitalPieChart,
                        ScenarioPieChart = model.ScenarioPieChart,
                        equity = new Equity { CurrentSharePrice = output[0], NumberShareBasic = output[1], NumberShareOutstanding = output[2], CostOfEquity = (model.equity!=null ? model.equity.CostOfEquity :0 ) },
                        prefferedEquity = new PrefferedEquity { PrefferedSharePrice = output[3], PrefferedShareOutstanding = output[4], PrefferedDividend = output[5], CostPreffEquity = (model.prefferedEquity!=null  ? model.prefferedEquity.CostPreffEquity :0) },
                        debt = new Debt { MarketValueDebt = output[6], CostOfDebt = (model.debt!=null ? model.debt.CostOfDebt :0) },
                        CashEquivalent = output[7],
                        CashNeededCapital = output[8],
                        FreeCashFlow = output[9],
                        debtUnit = (model.debtUnit!=null ? model.debtUnit :new DebtUnit()),
                        equityUnits = (model.equityUnits != null ? model.equityUnits : new EquityUnit()) ,
                        prefferedEquityUnit = (model.prefferedEquityUnit != null ? model.prefferedEquityUnit : new PrefferedEquityUnit()) 
                    };
                    iCapitalStructure.Update(capitalStructure);
                    iCapitalStructure.Commit();
                    return Ok(result);

                }
                }
                else
                {
                    return BadRequest(new { StatusCode0 = 0 , message= "model is empty" }) ;
                }
            }
            catch (Exception ex)
            {

                return BadRequest(ex);
            }

        }

        [HttpDelete]
        [Route("DeleteCapitalStructure/{id}")]
        public ActionResult<Object> DeleteCapitalStructure(int id)
        {
            // int result = 0;

            if (id == 0)
            {
                return BadRequest();
            }

            try
            {
                iCapitalStructure.DeleteWhere(s => s.Id == id);
                // if (result == 0)
                // {
                //     return NotFound("Record Not Found");
                // }
                return Ok("Successfully Deleted");
            }
            catch (Exception ex)
            {

                return BadRequest(ex.Message);
            }
        }


        [HttpGet]
        [Route("[action]")]
        public ActionResult<Object> ApproveCapitalStructure(int ApprovalFlag, long UserId)
        {
            // ApprovalBody model = JsonConvert.DeserializeObject<ApprovalBody>(obj);
            try
            {
                if (ApprovalFlag == 1)
                {
                    List<CapitalStructure> capitalStructure = iCapitalStructure.FindBy(s => s.UserId == UserId).ToList();
                    if (capitalStructure != null)
                    {
                        capitalStructure[0].ApprovalFlag = 1;
                        iCapitalStructure.Update(capitalStructure[0]);
                        iCapitalStructure.Commit();
                        return Ok("Successfully approved");
                    }
                    else
                    {
                        return NotFound("record not foud");
                    }
                }
                else
                {
                    return BadRequest("Approval flag is not valid");
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // public class ApprovalBody
        // {
        //     public long UserId;
        //     public int ApprovalFlag;
        // }

        [HttpPost]
        [Route("[action]")]
        public ActionResult<Object> AddCapitalStructureSnapshot([FromBody]CapitalAnalysisSnapshots model)
        {
            // CapitalAnalysisSnapshots model = JsonConvert.DeserializeObject<CapitalAnalysisSnapshots>(str);
            try
            {
                Snapshots snapshots = new Snapshots
                {

                    SnapShot = model.SnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    SnapShotType = CAPITALSTRUCTURE
                };
                iSnapshots.Add(snapshots);
                iSnapshots.Commit();
                return Ok(new { id = snapshots.Id, result = "Snapshot saved sucessfully" });

            }
            catch (Exception ex)
            {

                return BadRequest(ex.Message);
            }




        }

        [HttpGet]
        [Route("GetAllCSSnapShots/{UserId}")]
        public ActionResult<Object> GetAllCSSnapShots(long UserId)
        {
            try
            {
                var SnapShot = iSnapshots.FindBy(s => s.UserId == UserId && s.SnapShotType == CAPITALSTRUCTURE);
                if (SnapShot == null)
                {
                    return NotFound("Snapshot not found.");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);

            }

        }

        [HttpGet]
        [Route("CapitalStructureSnapShot/{Id}")]
        public ActionResult<Object> CapitalStructureSnapShot(long Id)
        {

            try
            {

                var SnapShot = iSnapshots.FindBy(s => s.Id == Id && s.SnapShotType == CAPITALSTRUCTURE);
                if (SnapShot == null)
                {
                    return NotFound("Snapshot not found.");

                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "An error occurred.", details = ex.Message });
            }

        }

        [HttpGet]
        [Route("ExportCapitalStructureNew/{UserId}")]
        public ActionResult<Object> ExportCapitalStructureNew(long UserId)
        {
            if (UserId != 0)
            {
                string rootFolder = _hostingEnvironment.WebRootPath;
                string fileName = @"capital_structure_new1.xlsx";
                FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
                var formattedCustomObject = (String)null;
             
                MasterCostofCapitalNStructureViewModel result = new MasterCostofCapitalNStructureViewModel();
                MasterCostofCapitalNStructure tblMaster = iMasterCostofCapitalNStructure.GetSingle(x => x.UserId == UserId);
                if (tblMaster != null)
                {
                    result = mapper.Map<MasterCostofCapitalNStructure, MasterCostofCapitalNStructureViewModel>(tblMaster);
                    //get Capital Stryucture
                    List<CapitalStructure_Input> tblcapitalStructureInputList = iCapitalStructure_Input.FindBy(x => x.MasterId == tblMaster.Id).ToList();
                    //Map Capital Structure
                    result.CapitalStructureInputList = new List<CapitalStructure_InputViewModel>();

                    if (tblcapitalStructureInputList != null)
                        foreach (var obj in tblcapitalStructureInputList.OrderBy(x => x.Id))
                        {
                            CapitalStructure_InputViewModel tempVM = mapper.Map<CapitalStructure_Input, CapitalStructure_InputViewModel>(obj);
                            result.CapitalStructureInputList.Add(tempVM);
                        }

                    //get Cost of Capital
                    List<CostofCapital_Input> tblCostofcapitalInputList = iCostofCapital_Input.FindBy(x => x.MasterId == tblMaster.Id).ToList();
                    //Map Cost of capital
                    result.CostofCapitalInputList = new List<CostofCapital_InputViewModel>();
                    if (tblCostofcapitalInputList != null)
                        foreach (var obj in tblCostofcapitalInputList.OrderBy(x => x.Id))
                        {
                            CostofCapital_InputViewModel tempVM = mapper.Map<CostofCapital_Input, CostofCapital_InputViewModel>(obj);
                            result.CostofCapitalInputList.Add(tempVM);
                        }

                    //get output for Capital Structure & Cost of Capital
                    MasterCostofCapitalNStructureViewModel tempOutput = GetOutPutforCapitalNStructure(result);

                    if (tempOutput != null)
                    {
                        result.CapitalStructureOutputList = tempOutput.CapitalStructureOutputList != null && tempOutput.CapitalStructureOutputList.Count > 0 ? tempOutput.CapitalStructureOutputList : new List<CapitalStructure_OutputViewModel>();

                        result.CostofCapitalOutputList = tempOutput.CostofCapitalOutputList != null && tempOutput.CostofCapitalOutputList.Count > 0 ? tempOutput.CostofCapitalOutputList : new List<CostofCapital_OutputViewModel>();
                    }

                }

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet wsCapitalStructure = null;
                   
                        wsCapitalStructure = package.Workbook.Worksheets["CapitalStructure"];
                   
                    List<MasterCostofCapitalNStructureViewModel> resultData = new List<MasterCostofCapitalNStructureViewModel>();
                    resultData.Add(result);


                    //  CapitalStructureExportRequest capitalStructureExportRequest = new CapitalStructureExportRequest{
                    //     CapitalStructure = resultData,
                    //     WorksheetName = wsCapitalStructure
                    //  };
                    wsCapitalStructure = CapitalStructureExcelExportNew(resultData,wsCapitalStructure );

                    wsCapitalStructure.Cells.AutoFitColumns();

                    ExcelPackage excelPackage = new ExcelPackage();
                    excelPackage.Workbook.Worksheets.Add("CapitalStructure", wsCapitalStructure);
                    package.Save();
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


        [HttpGet]
        [Route("ExportCapitalStructure/{UserId}/{Flag}")]
        public ActionResult<Object> ExportCapitalStructure(long UserId, int Flag)
        {
            if (UserId != 0)
            {
                string rootFolder = _hostingEnvironment.WebRootPath;
                Console.WriteLine("Hello World!", rootFolder);

                string fileName = @"capital_structure.xlsx";
                FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
                Console.WriteLine("Hello World! 2", file);
                var formattedCustomObject = (String)null;
                List<CapitalStructure> capitalStructure = iCapitalStructure.FindBy(s => s.UserId == UserId).ToList();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet wsCapitalStructure = null;
                    if (Flag == 1)
                    {
                        wsCapitalStructure = package.Workbook.Worksheets["CapitalStructure"];
                    }
                    else
                    {
                        wsCapitalStructure = package.Workbook.Worksheets["ScenarioAnalysis"];
                    }

                    if( capitalStructure != null && capitalStructure.Count > 0)
                    {
                         wsCapitalStructure = CapitalStructureExcelExport(capitalStructure, wsCapitalStructure, "D");
                    }


                    if (Flag == 2)
                    {
                        wsCapitalStructure = ExcelOutputScenario(capitalStructure, wsCapitalStructure);
                    }
                    ExcelPackage excelPackage = new ExcelPackage();
                    excelPackage.Workbook.Worksheets.Add("CapitalStructure", wsCapitalStructure);
                    package.Save();
                    ExcelPackage epOut = excelPackage;
                    byte[] myStream = epOut.GetAsByteArray();
                    var inputAsString = Convert.ToBase64String(myStream);
                    formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);
                    excelPackage.Dispose();
                }
                return Ok(formattedCustomObject);

            }
            else
            {
                return BadRequest("Id Not Found ");
            }

        }

        private static string ReturnCellUnit(int ValueTypeId , int UnitId)
        {
            string UnitValue = null;

            if (ValueTypeId == 3)
            {
                if (UnitId == 1)
                {
                    UnitValue = "$";
                }
                else if (UnitId == 2)
                {
                    UnitValue = "$K";
                }
                else if (UnitId == 3)
                {
                    UnitValue = "$M";
                }
                else if (UnitId == 4)
                {
                    UnitValue = "$B";
                }
                else if (UnitId == 5)
                {
                    UnitValue = "$T";
                }
            }
            else if (ValueTypeId == 2)
            {
                if (UnitId == 1)
                {
                    UnitValue = "Each";
                }
                else if (UnitId == 2)
                {
                    UnitValue = "K";
                }
                else if (UnitId == 3)
                {
                    UnitValue = "M";
                }
                else if (UnitId == 4)
                {
                    UnitValue = "B";
                }
                else if (UnitId == 5)
                {
                    UnitValue = "T";
                }
            }

            return UnitValue;
        }

        private static ExcelWorksheet ReturnCellStyle(string cellName, int row, int col, ExcelWorksheet wsSheetName, int flag)
        {
            wsSheetName.Cells[row, col].Value = cellName;
            wsSheetName.Cells[row, col].Style.Font.Size = 11;
            if (flag == 0)
            {
                wsSheetName.Cells[row, col].Style.Font.Bold = true;
            }
            else if (flag == 1)
            {
                wsSheetName.Cells[row, col].Style.Font.Bold = false;
            }
            return wsSheetName;
        }

        private static ExcelWorksheet ExcelGenerationOutput(List<CapitalStructure_OutputViewModel> capitalStructureOutputVM, string str, ExcelWorksheet wsSheetName, int col_count,
                                                         int vertical_count, out int row_count, out List<string> listOfFormula ,
                                                         List<string> listOfCellsEquity , List<string> listOfCellsPreferredEquity , List<string> listOfCellsDebt)
        {

            List<List<List<string>>> outer = new List<List<List<string>>>();
            List<List<string>> middle = new List<List<string>>();
            List<string> inner = new List<string>();

            foreach (CapitalStructure_OutputViewModel TempValue in capitalStructureOutputVM)
            {
                ExcelAddress addr = null;
                string UnitValue = "";

                if (str == "Current Capital Structure" || str == "Current Cost of Capital" || str == "Cash Balance" || str == "Leverage Ratios" || str == "Internal Valuation")
                {
                    wsSheetName.Cells[(6 + vertical_count), 3].Value = str;
                    wsSheetName.Cells[(6 + vertical_count), 3].Style.Font.Size = 12;
                    wsSheetName.Cells[(6 + vertical_count), 3].Style.Font.Bold = true;
                }
                else
                {
                    wsSheetName = ReturnCellStyle(str, (6 + vertical_count), 3, wsSheetName, 0);
                }

                if (TempValue.UnitId != null)
                {
                    UnitValue = ReturnCellUnit(TempValue.ValueTypeId, (int)TempValue.UnitId);
                }

                if (TempValue.LineItem == "Market Value of Common Equity (E)")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Market Value of Preferred Equity (P)")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Market Value of Debt (D)")
                {
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 2 + vertical_count), (col_count + 1), (7 + 2 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Market Value of Net Debt (ND)")
                {
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 3 + vertical_count), (col_count + 1), (7 + 3 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Cost of Equity (rE)")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem);
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }
                else if (TempValue.LineItem == "Cost of Preferred Equity (rP)")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem) ;
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }
                else if (TempValue.LineItem == "Cost of Debt (rD)")
                {
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count)].Value = (TempValue.LineItem);
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 2 + vertical_count), (col_count + 1), (7 + 2 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }
                else if (TempValue.LineItem == "Unlevered Cost of Capital/Equity (rU)")
                {
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count)].Value = (TempValue.LineItem) ;
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 3 + vertical_count), (col_count + 1), (7 + 3 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }
                else if (TempValue.LineItem == "Weighted Average Cost of Capital (rWACC)")
                {
                    wsSheetName.Cells[(7 + 4 + vertical_count), (col_count)].Value = (TempValue.LineItem) ;
                    wsSheetName.Cells[(7 + 4 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 4 + vertical_count), (col_count + 1), (7 + 4 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 4 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }
                else if (TempValue.LineItem == "Cash & Equivalent")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = "NA";
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Excess Cash")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Debt-to-Equity (D/E) Ratio")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem);
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }
                else if (TempValue.LineItem == "Debt-to-Value (D/V) Ratio")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem);
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }
                else if (TempValue.LineItem == "Unlevered Enterprise Value")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                }
                else if (TempValue.LineItem == "Levered Enterprise Value")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                }
                else if (TempValue.LineItem == "Equity Value")
                {
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 2 + vertical_count), (col_count + 1), (7 + 2 + vertical_count), (col_count + 1));
                }
                else if (TempValue.LineItem == "Interest Tax Shield Value")
                {
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 3 + vertical_count), (col_count + 1), (7 + 3 + vertical_count), (col_count + 1));
                }
                else if (TempValue.LineItem == "Stock Price")
                {
                    wsSheetName.Cells[(7 + 4 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 4 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 4 + vertical_count), (col_count + 1), (7 + 4 + vertical_count), (col_count + 1));
                }

                if (addr != null)
                {
                    inner.Add(addr.ToString());
                }
                //middle.Add(inner.Take(inner.Count()).ToList());
            }

            vertical_count = vertical_count + capitalStructureOutputVM.Count + 1;
            // outer.Add(middle);
            // middle.Add(inner);

            row_count = vertical_count;
            listOfFormula = inner;

            return wsSheetName;

        }

        private static ExcelWorksheet ExcelGenerationInput(List<CapitalStructure_InputViewModel> capitalStructureInputVM, string str, ExcelWorksheet wsSheetName, int col_count, 
                                                         int vertical_count, out int row_count, out List<string> listOfFormula)
        {

            List<List<List<string>>> outer = new List<List<List<string>>>();
            List<List<string>> middle = new List<List<string>>();
            List<string> inner = new List<string>();

            foreach (CapitalStructure_InputViewModel TempValue in capitalStructureInputVM)
            {
                ExcelAddress addr = null;
                string UnitValue = "";

                if (str == "Other Inputs")
                {
                    wsSheetName.Cells[(6 + vertical_count), 3].Value = str;
                    wsSheetName.Cells[(6 + vertical_count), 3].Style.Font.Size = 12;
                    wsSheetName.Cells[(6 + vertical_count), 3].Style.Font.Bold = true;
                }
                else
                {
                    wsSheetName = ReturnCellStyle(str, (6 + vertical_count), 3, wsSheetName, 0);
                }
               
                if (TempValue.UnitId != null)
                {
                    UnitValue = ReturnCellUnit(TempValue.ValueTypeId, (int)TempValue.UnitId);
                }
                
                if (TempValue.LineItem == "Current Share Price")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Number of Shares Outstanding - Basic")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                }
                else if (TempValue.LineItem == "Number of Shares Outstanding - Diluted")
                {
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 2 + vertical_count), (col_count + 1), (7 + 2 + vertical_count), (col_count + 1));
                }
                else if (TempValue.LineItem == "Current Preferred Share Price")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Number of Preferred Shares Outstanding")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Preferred Dividend")
                {
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 2 + vertical_count), (col_count + 1), (7 + 2 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Market Value of Interest Bearing Debt")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Cash & Equivalent")
                {
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 0 + vertical_count), (col_count + 1), (7 + 0 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 0 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Cash Needed for Working Capital")
                {
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count)].Value = (TempValue.LineItem) + " " + "(" + UnitValue + ")";
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Value = (TempValue.Value);
                    addr = new ExcelAddress((7 + 1 + vertical_count), (col_count + 1), (7 + 1 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 1 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "$#,##0.00";
                }
                else if (TempValue.LineItem == "Interest Coverage Ratio (k) - If Applicable")
                {
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count)].Value = (TempValue.LineItem);
                    wsSheetName.Cells[(7 + 2 + vertical_count), (col_count + 1)].Value = "NA";
                    addr = new ExcelAddress((7 + 2 + vertical_count), (col_count + 1), (7 + 2 + vertical_count), (col_count + 1));
                }
                else if (TempValue.LineItem == "Marginal Tax Rate (Tc)")
                {
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count)].Value = (TempValue.LineItem);
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count + 1)].Value = (TempValue.Value) / 100;
                    addr = new ExcelAddress((7 + 3 + vertical_count), (col_count + 1), (7 + 3 + vertical_count), (col_count + 1));
                    wsSheetName.Cells[(7 + 3 + vertical_count), (col_count + 1)].Style.Numberformat.Format = "0.00%";
                }

                if (addr != null)
                {
                    inner.Add(addr.ToString());
                }             
                //middle.Add(inner.Take(inner.Count()).ToList());
            }          

            vertical_count = vertical_count + capitalStructureInputVM.Count + 1;
            // outer.Add(middle);
           // middle.Add(inner);

            row_count = vertical_count;
            listOfFormula = inner;

            return wsSheetName;

        }


        private ExcelWorksheet CapitalStructureExcelExportNew(List<MasterCostofCapitalNStructureViewModel> capitalStructure, ExcelWorksheet wsCapitalStructure)
        {

        

            var capitalStructureList = capitalStructure[0];

            List<CapitalStructure_InputViewModel> CapitalStructureInputDatasList = new List<CapitalStructure_InputViewModel>();
            CapitalStructureInputDatasList = capitalStructureList.CapitalStructureInputList;

            List<CostofCapital_InputViewModel> CostofCapitalInputDatasList = new List<CostofCapital_InputViewModel>();
            CostofCapitalInputDatasList = capitalStructureList.CostofCapitalInputList;
            
            if (capitalStructureList.LeveragePolicyId == 1)
            {
                wsCapitalStructure.Cells["C2"].Value = "Initial Set Up 1: Constant Debt-to-Equity Ratio (Target Leverage Ratio)";              
            }
            else if (capitalStructureList.LeveragePolicyId == 2)
            {
                wsCapitalStructure.Cells["C2"].Value = "Initial Set Up 2: Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio";
            }
            else if (capitalStructureList.LeveragePolicyId == 3)
            {
                wsCapitalStructure.Cells["C2"].Value = "Initial Set Up 3: Constant Permanent Debt";
            }
            else if (capitalStructureList.LeveragePolicyId == 4)
            {
                wsCapitalStructure.Cells["C2"].Value = "Initial Set Up 4: Constant Interest Coverage Ratio (% of Free Cash Flow)";
            }

            wsCapitalStructure.Cells["C2"].Style.Font.Size = 14;

            wsCapitalStructure.Cells["C4"].Value = "Inputs:";
            wsCapitalStructure.Cells["C4"].Style.Font.Bold = true;
            wsCapitalStructure.Cells["C5"].Value = "Sources of Financing";
            wsCapitalStructure.Cells["C5"].Style.Font.Bold = true;
            wsCapitalStructure.Cells["C5"].Style.Font.Size = 12;

            int row_count = 0;
            // List<List<List<string>>> listOfCellsEquity = new List<List<List<string>>>();
            List<string> listOfCellsEquity = new List<string>();
            List<string> listOfCellsPreferredEquity = new List<string>();
            List<string> listOfCellsDebt = new List<string>();
            List<string> listOfCellsOtherInputs = new List<string>();

            var equity = CapitalStructureInputDatasList.FindAll(x => x.SubHeader == "Equity" && x.HeaderId == 1).ToList();
            var preferredEquity = CapitalStructureInputDatasList.FindAll(x => x.SubHeader == "Preferred Equity" && x.HeaderId == 1 && x.Value != null).ToList();
            var debt = CapitalStructureInputDatasList.FindAll(x => x.SubHeader == "Debt" && x.HeaderId == 1).ToList();
            var otherInputs = CapitalStructureInputDatasList.FindAll(x => x.SubHeader == "" && x.HeaderId == 2).ToList();
           
            // var costofEquity = CostofCapitalOutputDatasList.FindAll(x => x.LineItem.Contains("Cost of Equity (rE)"));

            if (equity != null && equity.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationInput(equity, "Equity", wsCapitalStructure, 3, row_count, out row_count, out listOfCellsEquity);

                row_count = row_count - 1;

                wsCapitalStructure.Cells[(7 + row_count), (3)].Value = "Cost of Equity (rE)";
                wsCapitalStructure.Cells[(7 + row_count), (4)].Value = "";
                // wsCapitalStructure.Cells[(7 + row_count), (4)].Style.Numberformat.Format = "0.00%";
                ExcelAddress addr = null;
                addr = new ExcelAddress((7 + row_count), (4), (7 + row_count), (4));
                listOfCellsEquity.Add(addr.ToString());
                row_count = row_count + 2;                         
            }

            if (preferredEquity != null && preferredEquity.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationInput(preferredEquity, "Preferred Equity", wsCapitalStructure, 3, row_count, out row_count, out listOfCellsPreferredEquity);

                row_count = row_count - 1;

                wsCapitalStructure.Cells[(7 + row_count), (3)].Value = "Cost of Preferred Equity (rP)";
                wsCapitalStructure.Cells[(7 + row_count), (4)].Value = "";
                wsCapitalStructure.Cells[(7 + row_count), (4)].Style.Font.Size = 11;
                // wsCapitalStructure.Cells[(7 + row_count), (4)].Style.Numberformat.Format = "0.00%";
                ExcelAddress addr = null;
                addr = new ExcelAddress((7 + row_count), (4), (7 + row_count), (4));
                listOfCellsPreferredEquity.Add(addr.ToString());
                row_count = row_count + 2;
            }

            if (debt != null && debt.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationInput(debt, "Debt", wsCapitalStructure, 3, row_count, out row_count, out listOfCellsDebt);

                row_count = row_count - 1;

                wsCapitalStructure.Cells[(7 + row_count), (3)].Value = "Cost of Debt (rD)";
                wsCapitalStructure.Cells[(7 + row_count), (4)].Value = "";
               // wsCapitalStructure.Cells[(7 + row_count), (4)].Style.Numberformat.Format = "0.00%";
                ExcelAddress addr = null;
                addr = new ExcelAddress((7 + row_count), (4), (7 + row_count), (4));
                listOfCellsDebt.Add(addr.ToString());
                row_count = row_count + 3; // Gap of one row for Other Inputs.
            }

            if (otherInputs != null && otherInputs.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationInput(otherInputs, "Other Inputs", wsCapitalStructure, 3, row_count, out row_count, out listOfCellsOtherInputs);

                row_count = row_count + 2;                
            }

            // Output

            List<CostofCapital_OutputViewModel> CostofCapitalOutputDatasList = new List<CostofCapital_OutputViewModel>();
            CostofCapitalOutputDatasList = capitalStructureList.CostofCapitalOutputList;

            List<CapitalStructure_OutputViewModel> CapitalStructureOutputDatasList = new List<CapitalStructure_OutputViewModel>();
            CapitalStructureOutputDatasList = capitalStructureList.CapitalStructureOutputList;

            var currentCapitalStructure = CapitalStructureOutputDatasList.FindAll(x => x.HeaderId == 3).ToList();
            var currentCostofCapital = CapitalStructureOutputDatasList.FindAll(x => x.HeaderId == 4).ToList();
            var cashBalance = CapitalStructureOutputDatasList.FindAll(x => x.HeaderId == 5).ToList();
            var leverageRatios = CapitalStructureOutputDatasList.FindAll(x => x.HeaderId == 6).ToList();
            var internalValuation = CapitalStructureOutputDatasList.FindAll(x => x.HeaderId == 7).ToList();

            List<string> listOfCellsCurrentCapitalStructure = new List<string>();
            List<string> listOfCellsCurrentCostofCapital = new List<string>();
            List<string> listOfCellsCashBalance = new List<string>();
            List<string> listOfCellsLeverageRatios = new List<string>();
            List<string> listOfCellsInternalValuation = new List<string>();

            wsCapitalStructure.Cells[(7 + row_count), (3)].Value = "Summay Output:";
            wsCapitalStructure.Cells[(7 + row_count), (3)].Style.Font.Bold = true;

            if (currentCapitalStructure != null && currentCapitalStructure.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationOutput(currentCapitalStructure, "Current Capital Structure", wsCapitalStructure, 3, row_count, out row_count, 
                                                           out listOfCellsCurrentCapitalStructure , listOfCellsEquity , listOfCellsPreferredEquity , listOfCellsDebt);
                row_count = row_count + 2;             
            }

            if (currentCostofCapital != null && currentCostofCapital.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationOutput(currentCostofCapital, "Current Cost of Capital", wsCapitalStructure, 3, row_count, out row_count,
                                                           out listOfCellsCurrentCostofCapital, listOfCellsEquity, listOfCellsPreferredEquity, listOfCellsDebt);
                row_count = row_count + 2;
            }

            if (cashBalance != null && cashBalance.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationOutput(cashBalance, "Cash Balance", wsCapitalStructure, 3, row_count, out row_count,
                                                           out listOfCellsCashBalance, listOfCellsEquity, listOfCellsPreferredEquity, listOfCellsDebt);
                row_count = row_count + 2;
            }

            if (leverageRatios != null && leverageRatios.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationOutput(leverageRatios, "Leverage Ratios", wsCapitalStructure, 3, row_count, out row_count,
                                                           out listOfCellsLeverageRatios, listOfCellsEquity, listOfCellsPreferredEquity, listOfCellsDebt);
                row_count = row_count + 2;
            }

            if (internalValuation != null && internalValuation.Count > 0)
            {
                wsCapitalStructure = ExcelGenerationOutput(internalValuation, "Internal Valuation", wsCapitalStructure, 3, row_count, out row_count,
                                                           out listOfCellsInternalValuation, listOfCellsEquity, listOfCellsPreferredEquity, listOfCellsDebt);
                row_count = row_count + 2;
            }

            return wsCapitalStructure;
        }

        private ExcelWorksheet CapitalStructureExcelExport(List<CapitalStructure> capitalStructure, ExcelWorksheet wsCapitalStructure, string cellPrefix)
        {

            SummaryOutput sumaryOutput = JsonConvert.DeserializeObject<SummaryOutput>(capitalStructure[0].SummaryOutput);

            var capitalStructureOuput = capitalStructure[0];

            wsCapitalStructure.Cells["D3"].Value = capitalStructureOuput.NewLeveragePolicy;

            var equity = capitalStructureOuput.equity;
            var equityUnit = capitalStructureOuput.equityUnits;
            wsCapitalStructure = CellValueFormatting("C8", "Current Share Price", "D8", equity.CurrentSharePrice, equityUnit.CurrentSharePriceUnit, 0, 0, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C9", "Number of Shares Outstanding - Basic", "D9", equity.NumberShareBasic, equityUnit.NumberShareBasicUnit, 1, 0, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C10", "Number of Shares Outstanding - Diluted", "D10", equity.NumberShareOutstanding, equityUnit.NumberShareOutstandingUnit, 1, 0, wsCapitalStructure);


            if (equity != null && equity.CostOfEquity != null)
            {
                wsCapitalStructure.Cells["D11"].Value = (equity.CostOfEquity / 100);
            }
            else
            {
                wsCapitalStructure.Cells["D11"].Value = 0; 
            }

            // wsCapitalStructure.Cells["D11"].Value = (equity.CostOfEquity / 100);

            var preffEquity = capitalStructureOuput.prefferedEquity;
            var preffEquityUnit = capitalStructureOuput.prefferedEquityUnit;
            wsCapitalStructure = CellValueFormatting("C13", "Current Preferred Share Price", "D13", preffEquity.PrefferedSharePrice, preffEquityUnit.PrefferedSharePriceUnit, 0, 0, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C14", "Number of Preferred Shares Outstanding", "D14", preffEquity.PrefferedShareOutstanding, preffEquityUnit.PrefferedShareOutstandingUnit, 1, 0, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C15", "Preferred Dividend", "D15", preffEquity.PrefferedDividend, preffEquityUnit.PrefferedDividendUnit, 0, 0, wsCapitalStructure);
            wsCapitalStructure.Cells["D16"].Value = (preffEquity.CostPreffEquity / 100);

            var debt = capitalStructureOuput.debt;
            var debtUnit = capitalStructureOuput.debtUnit;
            wsCapitalStructure = CellValueFormatting("C18", "Market Value of Interest Bearing Debt", "D18", debt.MarketValueDebt, debtUnit.MarketValueDebtUnit, 0, 0, wsCapitalStructure);
            wsCapitalStructure.Cells["D19"].Value = (debt.CostOfDebt / 100);

            wsCapitalStructure = CellValueFormatting("C22", "Cash & Equivalent", "D22", capitalStructureOuput.CashEquivalent, capitalStructureOuput.CashEquivalentUnit, 0, 0, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C23", "Cash Needed for Working Capital", "D23", capitalStructureOuput.CashNeededCapital, capitalStructureOuput.CashNeededCapitalUnit, 0, 0, wsCapitalStructure);
            wsCapitalStructure.Cells["D24"].Value = (capitalStructure[0].InterestCoverage / 100);
            wsCapitalStructure.Cells["D25"].Value = (capitalStructure[0].MarginalTaxRate / 100);
            wsCapitalStructure = CellValueFormatting("C26", "Free Cash Flow- Next Year", "D26", capitalStructureOuput.FreeCashFlow, capitalStructureOuput.FreeCashFlowUnit, 0, 0, wsCapitalStructure);

            string[] outputUnits = new string[6];
            var outputFormat = UnitConversion.ConvertOutputUnits(out outputUnits, sumaryOutput.marketValueEquity, sumaryOutput.marketValuePreferredEquity,
              sumaryOutput.marketValueDebt, sumaryOutput.cashEquivalent, sumaryOutput.excessCash, sumaryOutput.marketValueNetDebt);

            wsCapitalStructure = CellValueFormatting("C30", "Market Value of Equity (E)", "D30", "=D8*D9", outputUnits[0], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C31", "Market Value of Preferred Equity (P)", "D31", "=D13*D14", outputUnits[1], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C32", "Market Value of Debt (D)", "D32", "=D18", outputUnits[2], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C43", "Cash & Equivalent", "D43", "=D22", outputUnits[3], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C44", "Excess Cash", "D44", "=D22-D23", outputUnits[4], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("C33", "Market Value of Net Debt (ND)", "D33", "=D32-D44", outputUnits[5], 0, 1, wsCapitalStructure);
            wsCapitalStructure.Cells["D36"].Value = (sumaryOutput.costEquity / 100);
            wsCapitalStructure.Cells["D36"].Formula = "=D11";
            wsCapitalStructure.Cells["D37"].Value = (sumaryOutput.costPreferredEquity / 100);
            wsCapitalStructure.Cells["D37"].Formula = "=D16";
            wsCapitalStructure.Cells["D38"].Value = (sumaryOutput.costDebt / 100);
            wsCapitalStructure.Cells["D38"].Formula = "=D19";
            wsCapitalStructure.Cells["D39"].Value = (sumaryOutput.unleveredCostCapital / 100);
            wsCapitalStructure.Cells["D40"].Value = (sumaryOutput.weighedAverageCapital / 100);
            if (capitalStructure[0].NewLeveragePolicy == "Constant Debt-to-Equity Ratio  (Target Leverage Ratio)")
            {
                wsCapitalStructure.Cells["D40"].Style.Font.Color.SetColor(Color.Orange);
            }
            else if (capitalStructure[0].NewLeveragePolicy == "Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")
            {
                wsCapitalStructure.Cells["D40"].Formula = "=D39-(D33/(D30+D31+D33))*D25*D38*((1+D39)/(1+D38))";
            }
            else
            {
                wsCapitalStructure.Cells["D40"].Value = "NA";
            }
            if (capitalStructure[0].NewLeveragePolicy == "Constant Interest Coverage Ratio (% of Free Cash Flow)")
            {
                wsCapitalStructure.Cells["D47"].Value = sumaryOutput.debtEquityRatio;
                wsCapitalStructure.Cells["D47"].Formula = "=(D33/D30)";
                wsCapitalStructure.Cells["D48"].Value = sumaryOutput.debtValueRatio;
                wsCapitalStructure.Cells["D48"].Formula = "=D33/(D30+D31+D32+D33)";
            }
            else
            {
                wsCapitalStructure.Cells["D47"].Value = sumaryOutput.debtEquityRatio;
                wsCapitalStructure.Cells["D47"].Formula = "=(D32/D30)";
                wsCapitalStructure.Cells["D48"].Value = sumaryOutput.debtValueRatio;
                wsCapitalStructure.Cells["D48"].Formula = "=D32/(D30+D31+D32)";
            }
            wsCapitalStructure.Cells["D51"].Value = sumaryOutput.unleveredCostCapital;
            wsCapitalStructure.Cells["D52"].Value = sumaryOutput.leverateEnterpriseValue;
            wsCapitalStructure.Cells["D53"].Value = sumaryOutput.equityValue;
            wsCapitalStructure.Cells["D54"].Value = sumaryOutput.interestTaxShield;
            wsCapitalStructure.Cells["D55"].Value = sumaryOutput.stockPrice;
            return wsCapitalStructure;
        }


#region //Rahi || POST leverage policy || 04-11-2020

        [HttpPost]
        [Route("CapitalStructureAnalysis_new")]
        public ActionResult<Object> Analysis_new([FromBody] MasterCostofCapitalNStructureViewModel model)
        {
            try
            {

                //map MasterVM to master table
                MasterCostofCapitalNStructure tblMaster = mapper.Map<MasterCostofCapitalNStructureViewModel, MasterCostofCapitalNStructure>(model);


                if (tblMaster != null)
                {

                    tblMaster.ModifiedDate = System.DateTime.UtcNow;
                    iMasterCostofCapitalNStructure.Update(tblMaster);
                    iMasterCostofCapitalNStructure.Commit();

                }

                if (model.CapitalStructureInputList != null && model.CapitalStructureInputList.Count > 0)
                {
                    List<CapitalStructure_Input> tblcapitalStructureListObj = new List<CapitalStructure_Input>();
                    CapitalStructure_Input tblCapitalStructure = new CapitalStructure_Input();

                    foreach (CapitalStructure_InputViewModel obj in model.CapitalStructureInputList)
                    {

                        if (obj.ListType != null)
                        {
                           // obj.Id = 0;

                            obj.MasterId = tblMaster.Id;
                            if (obj.UnitId != null && obj.UnitId != 0)
                            {
                                if (obj.ValueTypeId == (int)ValueTypeEnum.Number)
                                    obj.BasicValue = UnitConversion.getBasicValueforNumbers(obj.UnitId, obj.Value);

                                if (obj.ValueTypeId == (int)ValueTypeEnum.Currency)
                                    obj.BasicValue = UnitConversion.getBasicValueforCurrency(obj.UnitId, obj.Value);
                            }
                            else
                            {
                                obj.BasicValue = obj.Value;
                            }

                            //map CapitalStructure ViewModel to table
                            tblCapitalStructure = mapper.Map<CapitalStructure_InputViewModel, CapitalStructure_Input>(obj);
                            tblcapitalStructureListObj.Add(tblCapitalStructure);
                        }
                    }
                    if (tblcapitalStructureListObj != null && tblcapitalStructureListObj.Count > 0)
                    {
                        if (model.Id == 0)
                            iCapitalStructure_Input.AddMany(tblcapitalStructureListObj);
                        else
                            iCapitalStructure_Input.UpdatedMany(tblcapitalStructureListObj);
                        iCapitalStructure_Input.Commit();
                    }
                }


            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }

            return Ok(new { message = "Succesfully updated scenario analysis leverage policy", code = 200 });
        }

        #endregion


        [HttpPost]
        [Route("CapitalStructureAnalysis")]
        public ActionResult<Object> Analysis([FromBody]CapitalAnalysisViewModel scenarioAnalysis)
        {
            var capitalStruct = iCapitalStructure.FindBy(s => s.UserId == scenarioAnalysis.UserId).ToArray();
            SummaryOutput summaryOutput = JsonConvert.DeserializeObject<SummaryOutput>(capitalStruct[0].SummaryOutput);

            var output = UnitConversion.ConvertOutputUnits(out _, summaryOutput.marketValueEquity, summaryOutput.marketValuePreferredEquity,
                 summaryOutput.marketValueDebt, summaryOutput.marketValueNetDebt, summaryOutput.cashEquivalent, summaryOutput.excessCash);

            BaseSummaryOutput baseSummaryOutput = new BaseSummaryOutput(
                 output[0], output[1], output[2], output[3], summaryOutput.costEquity, summaryOutput.costPreferredEquity, summaryOutput.costDebt,
                 summaryOutput.unleveredCostCapital, summaryOutput.weighedAverageCapital, output[4], output[5], summaryOutput.debtEquityRatio,
                 summaryOutput.debtValueRatio, summaryOutput.unleveredEnterpriseValue, summaryOutput.leverateEnterpriseValue, summaryOutput.equityValue,
                 summaryOutput.interestTaxShield, summaryOutput.stockPrice);

            var result = MathCapitalAnalysis.ScenarioAnalysis(scenarioAnalysis, baseSummaryOutput, capitalStruct[0].MarginalTaxRate);
            capitalStruct[0].ScenarioOutput = JsonConvert.SerializeObject(result);
            capitalStruct[0].ScenarioObject = scenarioAnalysis.AnalysisObject;
            capitalStruct[0].ScenarioPolicy = scenarioAnalysis.LeveragePolicy;
            capitalStruct[0].PermanentDebtUnit = scenarioAnalysis.PermanentDebtUnit;
            capitalStruct[0].ScenarioFreeUnit = scenarioAnalysis.FreeCashFlowUnit;
            ScenarioSummaryOutput summOut = JsonConvert.DeserializeObject<ScenarioSummaryOutput>(capitalStruct[0].ScenarioOutput);
            var outputs = UnitConversion.ConvertOutputUnits(out _, summOut.equityValue, summOut.preferredEquityValue, summOut.debtValue);
            Dictionary<string, object> results = new Dictionary<string, object>();
            results.Add("equityVal", outputs[0]);
            results.Add("preffEquityVal", outputs[1]);
            results.Add("debtVal", outputs[2]);

            Dictionary<string, object> resultsss = new Dictionary<string, object>();
            resultsss.Add("summaryOutput", result);
            resultsss.Add("piechartvalues", results);

            capitalStruct[0].ScenarioPieChart = JsonConvert.SerializeObject(results);

            iCapitalStructure.Update(capitalStruct[0]);
            iCapitalStructure.Commit();
            return Ok(resultsss);
        }

        [HttpGet]
        [Route("EditInputFlag/{UserId}/{Flag}")]
        public ActionResult<Object> EditIputs(long UserId, int Flag)
        {
            if (Flag == 1)
            {
                var capitalStruct = iCapitalStructure.FindBy(s => s.UserId == UserId).ToArray();
                if(capitalStruct.Length == 0)
                {
                    return NotFound("No record found");
                }
                capitalStruct[0].SummaryFlag = 0;
                iCapitalStructure.Update(capitalStruct[0]);
                iCapitalStructure.Commit();
                return Ok("Flag of Capital Structure changed to zero");
            }

            else if (Flag == 2)
            {
                var costOfCapital = iCostOfCapital.GetSingle(s => s.UserId == UserId);
                costOfCapital.SummaryFlag = 0;
                iCostOfCapital.Update(costOfCapital);
                iCapitalStructure.Commit();
                return Ok("Flag changed to zero");
            }
            else
            {
                return BadRequest("Invalid flag");
            }


        }


        private static ExcelWorksheet ExcelOutputScenario(List<CapitalStructure> capitalStructure, ExcelWorksheet wsCapitalStructure)
        {

            wsCapitalStructure.Cells["I3"].Value = capitalStructure[0].ScenarioPolicy;
            if (capitalStructure[0].ScenarioPolicy == "Constant Debt-to-Equity Ratio  (Target Leverage Ratio)")
            {
                ConstantDebtEquity constantDebtEquity = JsonConvert.DeserializeObject<ConstantDebtEquity>(capitalStructure[0].ScenarioObject);
                var cell1 = "Target Capital Structure";
                var cell2 = "Target Debt-to-Equity (D/E) Ratio";
                var cell3 = "New Cost of Capital";
                var cell4 = "Cost of Debt(rD)";
                wsCapitalStructure = ScenarioInputExcel(wsCapitalStructure, cell1, cell2, cell3, cell4, ((constantDebtEquity.TargetDebtEquity).ToString() + "%"),
                    (constantDebtEquity.CostOfDebt).ToString() + "%");
            }
            else if (capitalStructure[0].ScenarioPolicy == "Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")
            {
                ConstantDebtEquity constantDebtEquity = JsonConvert.DeserializeObject<ConstantDebtEquity>(capitalStructure[0].ScenarioObject);
                var cell1 = "Target Capital Structure";
                var cell2 = "Target Debt-to-Equity (D/E) Ratio";
                var cell3 = "New Cost of Capital";
                var cell4 = "Cost of Debt(rD)";
                wsCapitalStructure = ScenarioInputExcel(wsCapitalStructure, cell1, cell2, cell3, cell4, ((constantDebtEquity.TargetDebtEquity).ToString() + "%"),
                    (constantDebtEquity.CostOfDebt).ToString() + "%");
            }
            else if (capitalStructure[0].ScenarioPolicy == "Constant Permanent Debt")
            {
                ConstPermanentDebt constPermanentDebt = JsonConvert.DeserializeObject<ConstPermanentDebt>(capitalStructure[0].ScenarioObject);
                var cell1 = "New Permanent Debt";
                var cell2 = "Value of Permanent Debt";
                var cell3 = "New Cost of Capital";
                var cell4 = "Cost of Debt(rD)";
                wsCapitalStructure = ScenarioInputExcel(wsCapitalStructure, cell1, null, cell3, cell4, null,
                    (constPermanentDebt.CostOfDebt).ToString() + "%");
                var value = UnitConversion.ConvertUnits(new[] { constPermanentDebt.ValueOfPermDebt }, new string[] { capitalStructure[0].PermanentDebtUnit }, 0)[0];
                wsCapitalStructure = CellValueFormatting("H8", cell2, "I8", value, capitalStructure[0].PermanentDebtUnit, 0, 0, wsCapitalStructure);
            }
            else if (capitalStructure[0].ScenarioPolicy == "Constant Interest Coverage Ratio (% of Free Cash Flow)")
            {
                ConstInterestCovRatio constInterestCovRatio = JsonConvert.DeserializeObject<ConstInterestCovRatio>(capitalStructure[0].ScenarioObject);
                var cell1 = "Coverage Ratio";
                var cell2 = "Interest Coverage Ratio";
                var cell3 = "Free Cash Flow";
                var cell4 = "Free Cash Flow- Next Year";
                wsCapitalStructure = ScenarioInputExcel(wsCapitalStructure, cell1, cell2, cell3, null, (constInterestCovRatio.InterestCovRatio).ToString() + "%",
                   null);
                var value = UnitConversion.ConvertUnits(new[] { constInterestCovRatio.FreeCashFlow }, new[] { capitalStructure[0].ScenarioFreeUnit }, 0)[0];
                wsCapitalStructure = CellValueFormatting("H10", cell4, "I10", value, capitalStructure[0].ScenarioFreeUnit, 0, 0, wsCapitalStructure);
                wsCapitalStructure.Cells["H11"].Value = "New Cost of Capital";
                wsCapitalStructure.Cells["H12"].Value = "Cost of Debt (rD)";
                wsCapitalStructure.Cells["I12"].Value = constInterestCovRatio.CostOfDebt + "%";
            }
            ScenarioSummaryOutput scenarioSummaryOutput = JsonConvert.DeserializeObject<ScenarioSummaryOutput>(capitalStructure[0].ScenarioOutput);
            wsCapitalStructure.Cells["I22"].Value = scenarioSummaryOutput.equityValue;
            wsCapitalStructure.Cells["I22"].Formula = "=D30";

            string[] outputUnit = new string[7];
            var output = UnitConversion.ConvertOutputUnits(out outputUnit, scenarioSummaryOutput.equityValue, scenarioSummaryOutput.preferredEquityValue
                , scenarioSummaryOutput.debtValue, scenarioSummaryOutput.changeInDebt, scenarioSummaryOutput.netDebtValue,
                scenarioSummaryOutput.cashEquivalent, scenarioSummaryOutput.excessCash);
            wsCapitalStructure = CellValueFormatting("H22", "Equity Value (E)", "I22", "=D30", outputUnit[0], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("H23", "Preferred Equity Value (P)", "I23", "=D31", outputUnit[1], 0, 1, wsCapitalStructure);

            if (capitalStructure[0].ScenarioPolicy == "Constant Permanent Debt")
            {
                wsCapitalStructure = CellValueFormatting("H24", "Debt Value (D)", "I24", "=I8", outputUnit[2], 0, 1, wsCapitalStructure);
            }
            else if (capitalStructure[0].ScenarioPolicy == "Constant Interest Coverage Ratio (% of Free Cash Flow)")
            {
                wsCapitalStructure = CellValueFormatting("H24", "Debt Value (D)", "I24", "=(I8)*(I10/I12)", outputUnit[2], 0, 1, wsCapitalStructure);
            }
            else
            {
                wsCapitalStructure = CellValueFormatting("H24", "Debt Value (D)", "I24", "=I8*I22", outputUnit[2], 0, 1, wsCapitalStructure);
            }
            wsCapitalStructure = CellValueFormatting("H25", "Change in Debt", "I25", "=I24-D32", outputUnit[3], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("H26", "Net Debt Value (ND)", "I26", "=I24-I30", outputUnit[4], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("H29", "Cash & Equivalent", "I29", "=D43+I25", outputUnit[5], 0, 1, wsCapitalStructure);
            wsCapitalStructure = CellValueFormatting("H30", "Excess Cash", "I30", "=I25+D44", outputUnit[6], 0, 1, wsCapitalStructure);

            wsCapitalStructure.Cells["I33"].Value = scenarioSummaryOutput.costOfEquity;
            if (capitalStructure[0].ScenarioPolicy == "Constant Debt-to-Equity Ratio  (Target Leverage Ratio)")
            {
                wsCapitalStructure.Cells["I33"].Formula = "=(D39)+((I26/I22)*((D39)-(I10)))+((I23/I22)*((D39)-I34))";
            }
            else if (capitalStructure[0].ScenarioPolicy == "Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")
            {
                wsCapitalStructure.Cells["I33"].Formula = "=D39 + ((I26 / I22) * (D39 - I10))+((I23 / I22) * (D39 - I34))";
            }

            wsCapitalStructure.Cells["I34"].Value = scenarioSummaryOutput.preferredEquityValue;
            wsCapitalStructure.Cells["I34"].Formula = "=D16";

            wsCapitalStructure.Cells["I35"].Value = scenarioSummaryOutput.costOfDebt;
            if (capitalStructure[0].ScenarioPolicy != "Constant Interest Coverage Ratio (% of Free Cash Flow)")
            {
                wsCapitalStructure.Cells["I35"].Formula = "=I10";
            }
            else
            {
                wsCapitalStructure.Cells["I35"].Formula = "=I12";
            }

            wsCapitalStructure.Cells["I36"].Value = scenarioSummaryOutput.unleveredCostOfCapital;
            wsCapitalStructure.Cells["I36"].Formula = "=D39";

            wsCapitalStructure.Cells["I37"].Value = scenarioSummaryOutput.weightedAverage;
            if (capitalStructure[0].ScenarioPolicy == "Constant Debt-to-Equity Ratio  (Target Leverage Ratio)")
            {
                wsCapitalStructure.Cells["I37"].Formula = "=((I33)*(I22/(I22 + I23 + I26))) + ((I34)*(I23/(I22 + I23 + I26))) + ((I35)*(I26/(I22 + I23 + I26)))*(1-(D25))";
            }
            else if (capitalStructure[0].ScenarioPolicy == "Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")
            {
                wsCapitalStructure.Cells["I37"].Formula = "=(D39) - ((I26 / (I22 + I23 + I26)) *(D25) * (I35) * ((1 +(D39))/ (1 + (I35))))";
            }

            wsCapitalStructure.Cells["I40"].Value = scenarioSummaryOutput.debtToEquity;
            wsCapitalStructure.Cells["I40"].Formula = "=I24/I22";

            wsCapitalStructure.Cells["I41"].Value = scenarioSummaryOutput.debtToValue;
            wsCapitalStructure.Cells["I41"].Formula = "=I24/(I22+I23+I24)";
            return wsCapitalStructure;
        }

        static ExcelWorksheet ScenarioInputExcel(ExcelWorksheet wsCapitalStructure
            , string cell1, string cell2, string cell3, string cell4, string value1, string value2)
        {
            wsCapitalStructure.Cells["H7"].Value = cell1;
            wsCapitalStructure.Cells["H8"].Value = cell2;
            wsCapitalStructure.Cells["H9"].Value = cell3;
            wsCapitalStructure.Cells["H10"].Value = cell4;
            wsCapitalStructure.Cells["I8"].Value = value1;
            wsCapitalStructure.Cells["I10"].Value = value2;
            return wsCapitalStructure;
        }

        static ExcelWorksheet CellValueFormatting(string cell, string cellHeading, string valueCell, object value, string unitForFormat,
            int flag, int formulaRValue, ExcelWorksheet wsCapitalStructure)
        {    ///use another for flag check type for using for loop appeding for multiple cells
            if (flag == 0) { wsCapitalStructure.Cells[cell].Value = cellHeading + "($" + unitForFormat.ToUpper() + ")"; }
            else if (flag == 1) { wsCapitalStructure.Cells[cell].Value = cellHeading + "(" + unitForFormat.ToUpper() + ")"; }
            if (formulaRValue == 0) { wsCapitalStructure.Cells[valueCell].Value = value; }
            else if (formulaRValue == 1) { wsCapitalStructure.Cells[valueCell].Formula = Convert.ToString(value); }
            switch (unitForFormat.ToLower())
            {
                case "k":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#.00" + "\" K\"";
                    break;
                case "m":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#,,.00" + "\" M\"";
                    break;
                case "b":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#,,,.00" + "\" B\"";
                    break;
                case "t":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#,,,,.00" + "\" T\"";
                    break;
            }
            return wsCapitalStructure;
        }

        #region ScenarioAnalysis Snapshot
        [HttpPost]
        [Route("ScenarioSnapshot")]
        public ActionResult<Object> AddScenarioAnalysisSnapShot([FromBody]CapitalStructureScenarioSnapshotsViewModel model)
        {
            try
            {
                CapitalStructureScenarioSnapshot snapshot = new CapitalStructureScenarioSnapshot
                {
                    UserId = model.UserId,
                    MasterId = model.MasterId,
                    Description = model.Description,
                    LeveragePolicyId = model.LeveragePolicyId,
                    scenarioOutput = model.scenarioOutput.ToString(),
                    scenarioPieChart = model.scenarioPieChart.ToString(),
                    SnapshotType = CAPITALSTRUCTURESCENARIO,
                    Active = true,
                    CreatedAt = DateTime.UtcNow
                };
                iCapitalStructureScenarioSnapshot.Add(snapshot);
                iCapitalStructureScenarioSnapshot.Commit();

                return Ok(new { result = snapshot.Id, message = "Succesfully added Snapshots", code = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

        [HttpGet]
        [Route("ScenarioSnapshots/{MasterId}")]
        public ActionResult<Object> GetScenarioSnapshots(long MasterId)
        {
            try
            {
                var SnapShot = iCapitalStructureScenarioSnapshot.FindBy(s => s.MasterId == MasterId && s.SnapshotType== CAPITALSTRUCTURESCENARIO);
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
        [Route("ScenarioSnapshot/{Id}")]
        public ActionResult<Object> GetScenarioSnapshot(long Id)
        {
            try
            {
                var SnapShot = iCapitalStructureScenarioSnapshot.GetSingle(s => s.Id == Id && s.SnapshotType == CAPITALSTRUCTURESCENARIO);
                if (SnapShot == null)
                {
                    return NotFound("No record found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        #endregion

    }
}
