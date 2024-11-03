using Microsoft.AspNetCore.Hosting;
using Newtonsoft.Json;
using OfficeOpenXml;
using Sofcan.Lib;
using Sowfin.API.Lib;
using Sowfin.API.Lib.ErrorHandling;
using Sowfin.API.ViewModels;
using Sowfin.API.ViewModels.CapitalBudgeting;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using AutoMapper;
using Sowfin.Data.Common.Enum;
using Excel.FinancialFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
//using System.Xml.Linq;
//using Aspose.Cells;

/// <summary>
/// Capital Budgeting Controller
/// </summary>
namespace Sowfin.API.Controllers
{




    [Route("api/[controller]")]
    [ApiController]
    public class CapitalBudgetingController : ControllerBase
    {
        private const string CAPITALBUDGETINGSNAPSHOT = "CapitalBudgeting_snapshot";
        private const string SENSITIVITYSNAPSHOT = "Sensitivity_snapshot";
        private const string SCENARIOSNAPSHOT = "Scenario_snapshot";
        private readonly double toMill = 1000000;
        private readonly ICapitalBudgeting iCapitalBudgeting = null;
        private readonly IProjectInputDatas iProjectInputDatas = null; 
        private readonly IProjectInputValues iProjectInputValues = null;
        IProjectInputComparables iProjectInputComparables;
        private readonly IProject iProject = null; 
        private readonly IMapper mapper = null; 
        private readonly IWebHostEnvironment  _hostingEnvironment = null;
        private readonly ISowfinCache iSowfinCache = null;
        private readonly ISnapshots iSnapShots = null;
        private readonly ICapitalBugetingTables iCapitalBugetingTables = null;
        public CapitalBudgetingController(IWebHostEnvironment  hostingEnvironment, ICapitalBudgeting _iCapitalBudgeting,
            ISowfinCache _sowfinCache, ISnapshots _iSnapShots, ICapitalBugetingTables _capitalBugetingTables, 
            IProjectInputDatas _iProjectInputDatas, IProjectInputComparables _iProjectInputComparables,IProject _iProject, IMapper _imapper, 
            IProjectInputValues _iProjectInputValues)
        {
            _hostingEnvironment = hostingEnvironment;
            iCapitalBudgeting = _iCapitalBudgeting;
            iProjectInputComparables = _iProjectInputComparables;
            iProjectInputDatas = _iProjectInputDatas;
            iProjectInputValues = _iProjectInputValues;
            iProject = _iProject; 
            mapper = _imapper; 
            iSowfinCache = _sowfinCache;
            iSnapShots = _iSnapShots;
            iCapitalBugetingTables = _capitalBugetingTables;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost("Evaluate")]
        public ActionResult<Object> Evaluate([FromBody] Evaluation data)
        {
            //Evaluation data = JsonConvert.DeserializeObject<Evaluation>(str);
            var table = data.tableData.OtherFixedCost;

            List<List<double>> volumeMultiList = ProvideListOfList(0, data.tableData.RevenueVariableCostTier, data.noOfYears, 0);
            List<List<double>> unitPriceMultiList = ProvideListOfList(1, data.tableData.RevenueVariableCostTier, data.noOfYears, 0);
            List<List<double>> unitCostMultiList = ProvideListOfList(2, data.tableData.RevenueVariableCostTier, data.noOfYears, 0);
            List<List<double>> fixedcostMultiList = CombineLists(data.tableData.OtherFixedCost, data.noOfYears); ///This is to combine all the array's of fixed cost
            List<List<double>> capexMultiList = ProvideListOfList(0, data.tableData.CapexDepreciation, data.noOfYears, 0);
            
            var obj = data.tableData.CapexDepreciation;
            List<List<double>> totalDeprciationMultiList = ProvideListOfList(1, data.tableData.CapexDepreciation, data.noOfYears, 0);
            Console.WriteLine("above provideaggreate");
            if (volumeMultiList != null || unitPriceMultiList != null || unitCostMultiList != null || fixedcostMultiList != null ||
              totalDeprciationMultiList != null)
            {
                List<double> volumeSum = ProvideAggregate(volumeMultiList, 0);
                List<double> unitPriceSum = ProvideAggregate(unitPriceMultiList, 0);
                List<double> untiCostSum = ProvideAggregate(unitCostMultiList, 0);
                List<double> fixedcostSum = ProvideAggregate(fixedcostMultiList, 0);
                List<double> capexSum = ProvideAggregate(capexMultiList, 0);
                List<double> totalDeprciationSum = ProvideAggregate(totalDeprciationMultiList, 0);

                var objectDictionary = new Dictionary<string, object>();
                objectDictionary.Add("Volume", volumeSum);
                objectDictionary.Add("UnitPrice", unitPriceSum);
                objectDictionary.Add("UnitCost", untiCostSum);
                objectDictionary.Add("Fixed", fixedcostSum);
                objectDictionary.Add("NWC", ExtractValues(data.tableData.WorkingCapital));
                objectDictionary.Add("Capex", capexSum);
                objectDictionary.Add("TotalDepreciation", totalDeprciationSum);
                objectDictionary.Add("MarginalTax", data.marginalTaxRate);
                objectDictionary.Add("DiscountRate", data.discountRate);

                var summaryOutput = MathCapitalBudgeting.SummaryOutput(objectDictionary);
                string tableString = JsonConvert.SerializeObject(objectDictionary);

                string rawTableData = JsonConvert.SerializeObject(data.tableData);
                List<string> keys = new List<string>();
                List<List<double>> values = new List<List<double>>();
                var yourDictionary = FixedCost(data.tableData.OtherFixedCost, out keys, out values);
                List<string> depKey = new List<string>();
                List<List<double>> depValue = new List<List<double>>();
                List<string> capKeys = new List<string>();
                List<List<double>> capVaule = new List<List<double>>();
                Dictionary<string, object> result = new Dictionary<string, object>();

                result.Add("Sales ($M)", summaryOutput["sales"]);
                result.Add("COGS ($M)", summaryOutput["cogs"]);
                result.Add("Gross Margin ($M)", summaryOutput["grossMargins"]);

                if (data.tableData.OtherFixedCost != null)
                {
                    for (int i = 0; i < keys.Count; i++)
                    {
                       result.Add(keys[i], yourDictionary[keys[i]].Select(x => x = x / toMill));
                    }
                }

                if (data.tableData.CapexDepreciation != null)
                {
                    depKey = GetDep(data.tableData.CapexDepreciation, out depValue, out capKeys, out capVaule);
                    for (int i = 0; i < depKey.Count; i++)
                    {
                        result.Add(depKey[i] + Convert.ToString(i + 1), depValue[i].Select(x => x = x / toMill));
                    }
                }

                result.Add("Operating Income ($M)", summaryOutput["operatingIncomes"]);
                result.Add("Income Tax ($M)", summaryOutput["incomeTaxs"]);
                result.Add("Unlevered Net Income ($M)", summaryOutput["unleveredNetIcomes"]);
                result.Add("NWC(Net Working Capital)($M)", summaryOutput["nwcs"]);

                if (data.tableData.CapexDepreciation != null)
                {
                    for (int i = 0; i < depKey.Count; i++)
                    {
                        result.Add("Plus:" + " " + depKey[i] + Convert.ToString(i + 1), depValue[i].Select(x => x = x / toMill));
                    }
                    for (int i = 0; i < capKeys.Count; i++)
                    {
                        result.Add("Less:" + " " + capKeys[i] + Convert.ToString(i + 1), capVaule[i].Select(x => x = x / toMill));
                    }
                }

                result.Add("Less: Increases in NWC", summaryOutput["increaseNwcs"]);
                result.Add("Free Cash Flow", summaryOutput["freeCashFlows"]);
                result.Add("Levered Value (VL) ($M)", summaryOutput["leveredValues"]);
                result.Add("Discount Factor", summaryOutput["discountFactors"]);
                result.Add("Discount Cash Flow ($M)", summaryOutput["discountCashFlow"]);
                result.Add("Net Present Value ($M)", summaryOutput["npv"]);
                //result.Add("npv", summaryOutput["npv"]);

                var summaryString = JsonConvert.SerializeObject(DictionaryToArray(result));
                List<object> res = new List<object>();
                res.Add(summaryString);
                res.Add(summaryOutput["npv"]);

                var npvList = (List<double>)summaryOutput["npv"];

                try
                {
                    if (data.id == 0)
                    {
                        CapitalBudgeting model = new CapitalBudgeting
                        {
                            StartingYear = data.startingYear,
                            NoOfYears = data.noOfYears,
                            DiscountRate = data.discountRate,
                            MarginalTaxRate = data.marginalTaxRate,
                            TableData = tableString,
                            UserId = data.userId,
                            ApprovalFlag = data.approvalFlag,
                            SummaryOutput = summaryString,
                            RawTableData = rawTableData,
                            ProjectId = data.projectId,
                            RevenueCount = data.revenueCount,
                            CapexCount = data.capexCount,
                            FixedCostCount = data.fixedCostCount,
                            NPV = npvList[0].ToString()
                        };

                        iCapitalBudgeting.Add(model);
                        iCapitalBudgeting.Commit();

                        //var key = "Evaluation_" + data.userId.ToString() + "::" + data.projectId.ToString();
                        //iSowfinCache.Set(key, model);

                        return Ok(new { message = "Saved Sucessfully", result = res, model.Id });
                    }
                    else
                    {
                        CapitalBudgeting updateModel = new CapitalBudgeting
                        {
                            Id = data.id,
                            StartingYear = data.startingYear,
                            NoOfYears = data.noOfYears,
                            DiscountRate = data.discountRate,
                            MarginalTaxRate = data.marginalTaxRate,
                            TableData = tableString,
                            UserId = data.userId,
                            ApprovalFlag = data.approvalFlag,
                            SummaryOutput = summaryString,
                            RawTableData = rawTableData,
                            ProjectId = data.projectId,
                            RevenueCount = data.revenueCount,
                            CapexCount = data.capexCount,
                            FixedCostCount = data.fixedCostCount,
                            NPV = npvList[0].ToString()
                        };

                        iCapitalBudgeting.Update(updateModel);
                        iCapitalBudgeting.Commit();

                        //var key = "Evaluation_" + data.userId.ToString() + "::" + data.projectId.ToString();
                        //var check = iSowfinCache.IsInCache(key);
                        //if (check == true)
                        //{
                        //    iSowfinCache.Remove(key);
                        //}
                        //iSowfinCache.Set(key, updateModel);

                        return Ok(new { message = "Saved Successfully", result = res, data.id });
                    }
                }
                catch (Exception)
                {
                    
                ErrorResponse errorResponse = new ErrorResponse(HttpStatusCode.BadRequest, "Invalid data");
                return BadRequest(errorResponse);
                }
            }
            else
            {
                Console.WriteLine("Error in the data");
                ErrorResponse errorResponse = new ErrorResponse(HttpStatusCode.BadRequest, "Invalid lenght of entires");
                return BadRequest(errorResponse);
            }
        }

        //public List<ProjectOutputDatasViewModel> GetProjectOutput_Old(ProjectsViewModel projectVM)
        //{
        //    // List<ProjectOutputDatasViewModel> SummaryList = new List<ProjectOutputDatasViewModel>()
        //    List<ProjectOutputDatasViewModel> OutputDatasVMList = new List<ProjectOutputDatasViewModel>();
        //    try
        //    {
        //        //Calculate Output by Project Inputs
        //        if (projectVM != null && projectVM.ProjectInputDatasVM != null && projectVM.ProjectInputDatasVM.Count > 0)
        //        {

        //            //create dumy year Value List
        //            List<ProjectOutputValuesViewModel> dumyValueList = new List<ProjectOutputValuesViewModel>();
        //            if (projectVM.StartingYear != null && projectVM.StartingYear != 0 && projectVM.NoOfYears != null && projectVM.NoOfYears > 0)
        //            {
        //                ProjectOutputValuesViewModel dumyValue = new ProjectOutputValuesViewModel();
        //                for (int i = 0; i < projectVM.NoOfYears; i++)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    dumyValue.Id = 0;
        //                    dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = Convert.ToInt32(projectVM.StartingYear) + i;
        //                    dumyValue.Value = null;
        //                    dumyValueList.Add(dumyValue);
        //                }

        //                ProjectOutputDatasViewModel OutputDatasVM = new ProjectOutputDatasViewModel();

        //                List<ProjectInputDatasViewModel> InputDatasList = new List<ProjectInputDatasViewModel>();
        //                InputDatasList = projectVM.ProjectInputDatasVM;

        //                //create data //for Sales 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Sales";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                // ($O$6*C6)*($O$7*C7)/1000000
        //                List<ProjectInputDatasViewModel> VolumeList = new List<ProjectInputDatasViewModel>();
        //                VolumeList = InputDatasList.FindAll(x => x.LineItem == "Volume" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

        //                List<ProjectInputDatasViewModel> UnitPriceList = new List<ProjectInputDatasViewModel>();
        //                UnitPriceList = InputDatasList.FindAll(x => x.LineItem == "Unit Price" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();
        //                int? highunit = UnitConversion.getHigherDenominationUnit(VolumeList.FirstOrDefault().UnitId, UnitPriceList.FirstOrDefault().UnitId);

        //                //sales=volume*unitcost;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;

        //                    double value = 0;
        //                    //calculate sales for multiple years
        //                    if (VolumeList != null && VolumeList.Count > 0 && UnitPriceList != null && UnitPriceList.Count > 0)
        //                        foreach (var item in VolumeList)
        //                        {
        //                            var UnitPriceDatas = UnitPriceList.Find(x => x.SubHeader == item.SubHeader);
        //                            ProjectInputValuesViewModel UnitCostValue = UnitPriceDatas != null && UnitPriceDatas.ProjectInputValuesVM != null && UnitPriceDatas.ProjectInputValuesVM.Count > 0 ? UnitPriceDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectInputValuesViewModel VolumeValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((UnitPriceDatas != null ? UnitConversion.getBasicValueforCurrency(UnitPriceDatas.UnitId, UnitCostValue.Value) : 0) * (item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, VolumeValue.Value) : 0));
        //                        }


        //                    //get sum of all the Vlaues

        //                    // value = UnitConversion.getHigherDenominationUnit(UnitCostDatas.UnitId,);

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign sum of all sales value
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);


        //                //create data //for COGS 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "COGS";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                List<ProjectInputDatasViewModel> UnitCostList = new List<ProjectInputDatasViewModel>();
        //                UnitCostList = InputDatasList.FindAll(x => x.LineItem == "Unit Cost" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

        //                highunit = UnitConversion.getHigherDenominationUnit(VolumeList.FirstOrDefault().UnitId, UnitCostList.FirstOrDefault().UnitId);
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    //calculate COGS for multiple years

        //                    //get sum of all the Vlaues
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    //calculate sales for multiple years
        //                    if (VolumeList != null && VolumeList.Count > 0 && UnitCostList != null && UnitCostList.Count > 0)
        //                        foreach (var item in VolumeList)
        //                        {
        //                            var UnitCostDatas = UnitCostList.Find(x => x.SubHeader == item.SubHeader);
        //                            ProjectInputValuesViewModel UnitCostValue = UnitCostDatas != null && UnitCostDatas.ProjectInputValuesVM != null && UnitCostDatas.ProjectInputValuesVM.Count > 0 ? UnitCostDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectInputValuesViewModel VolumeValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((UnitCostDatas != null ? UnitConversion.getBasicValueforCurrency(UnitCostDatas.UnitId, UnitCostValue.Value) : 0) * (item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, VolumeValue.Value) : 0));
        //                        }


        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    //value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign sum of all sales value
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Gross Margin
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Gross Margin";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                var SalesDatas = OutputDatasVMList.Find(x => x.LineItem == "Sales");
        //                var COGSDatas = OutputDatasVMList.Find(x => x.LineItem == "COGS");
        //                highunit = UnitConversion.getHigherDenominationUnit(SalesDatas.UnitId, COGSDatas.UnitId);
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel salesValue = SalesDatas.ProjectOutputValuesVM != null && SalesDatas.ProjectOutputValuesVM.Count > 0 ? SalesDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel COGSValue = COGSDatas.ProjectOutputValuesVM != null && COGSDatas.ProjectOutputValuesVM.Count > 0 ? COGSDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    value = (salesValue != null ? UnitConversion.getBasicValueforNumbers(SalesDatas.UnitId, salesValue.Value) : 0) - (COGSValue != null ? UnitConversion.getBasicValueforNumbers(COGSDatas.UnitId, COGSValue.Value) : 0);
        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                    // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //all items of fixed cost with the same formula
        //                //find all items of Other Fixed Cost
        //                List<ProjectInputDatasViewModel> OtherFixedCostList = new List<ProjectInputDatasViewModel>();
        //                OtherFixedCostList = InputDatasList.FindAll(x => x.SubHeader.Contains("Other Fixed Cost")).ToList();
        //                if (OtherFixedCostList != null && OtherFixedCostList.Count > 0)
        //                    foreach (ProjectInputDatasViewModel fixedCostObj in OtherFixedCostList)
        //                    {
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = fixedCostObj.LineItem;
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = fixedCostObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectInputValuesViewModel Value = fixedCostObj.ProjectInputValuesVM != null && fixedCostObj.ProjectInputValuesVM.Count > 0 ? fixedCostObj.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((fixedCostObj != null ? UnitConversion.getBasicValueforNumbers(fixedCostObj.UnitId, Value.Value) : 0));
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));//assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                    }

        //                //formula=Startyear+1 value * year Value

        //                //Depreciation // Bind Depreciation Values directly from Input
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Depreciation";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                // find depreciation
        //                List<ProjectInputDatasViewModel> DepreciationList = new List<ProjectInputDatasViewModel>();
        //                DepreciationList = InputDatasList.FindAll(x => x.LineItem == "Depreciation" && x.SubHeader.Contains("Capex & Depreciation")).ToList();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationList != null && DepreciationList.Count > 0 ? DepreciationList.FirstOrDefault().UnitId : null;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    //Sales-COGS
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    if (DepreciationList != null && DepreciationList.Count > 0)
        //                        foreach (var item in DepreciationList)
        //                        {
        //                            ProjectInputValuesViewModel depreciationValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, depreciationValue.Value) : 0));
        //                        }
        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));//assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);



        //                //Operating Income 
        //                //formula = Gross Margin - (sum of all items after Gross Margin)
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Operating Income";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                var GrossMarginDatas = OutputDatasVMList.Find(x => x.LineItem == "Gross Margin");
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = GrossMarginDatas.UnitId;

        //                var DepreciationDatas = OutputDatasVMList.Find(x => x.LineItem == "Depreciation");
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationDatas.UnitId;

        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    double sum = 0;
        //                    ProjectOutputValuesViewModel GrossValue = GrossMarginDatas.ProjectOutputValuesVM != null && GrossMarginDatas.ProjectOutputValuesVM.Count > 0 ? GrossMarginDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                    ProjectOutputValuesViewModel DepreciationValue = DepreciationDatas.ProjectOutputValuesVM != null && DepreciationDatas.ProjectOutputValuesVM.Count > 0 ? DepreciationDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                    foreach (ProjectInputDatasViewModel fixedCostObj in OtherFixedCostList)
        //                    {
        //                        var otherFixedValue = fixedCostObj.ProjectInputValuesVM != null && fixedCostObj.ProjectInputValuesVM.Count > 0 ? fixedCostObj.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                        sum = sum + (otherFixedValue != null ? UnitConversion.getBasicValueforNumbers(fixedCostObj.UnitId, otherFixedValue.Value) : 0);
        //                    }
        //                    //value = (GrossMarginDatas != null ? UnitConversion.getBasicValueforNumbers(GrossMarginDatas.UnitId, GrossValue.Value) : 0)- sum;//-()sum of all items;

        //                    value = (GrossMarginDatas != null ? UnitConversion.getBasicValueforNumbers(GrossMarginDatas.UnitId, GrossValue.Value) : 0) - sum - (DepreciationDatas != null ? UnitConversion.getBasicValueforNumbers(DepreciationDatas.UnitId, DepreciationValue.Value) : 0);//-()sum of all items;

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    //  value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    //  dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);


        //                //Income Tax
        //                //formula = Operating Income (for diff year) * Marginal Tax
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Income Tax";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                var OperatingIncomeDatas = OutputDatasVMList.Find(x => x.LineItem == "Operating Income");

        //                var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));

        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = OperatingIncomeDatas.UnitId;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    value = ((OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0)) / 100;

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);

        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Unlevered Net Income
        //                //formula = Operating Income (for diff year) -Income Tax
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Unlevered Net Income";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                var IncomeTaxDatas = OutputDatasVMList.Find(x => x.LineItem == "Income Tax");
        //                highunit = UnitConversion.getHigherDenominationUnit(OperatingIncomeDatas.UnitId, IncomeTaxDatas.UnitId);
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel IncomeTaxValue = IncomeTaxDatas.ProjectOutputValuesVM != null && IncomeTaxDatas.ProjectOutputValuesVM.Count > 0 ? IncomeTaxDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    value = (OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) - (IncomeTaxValue != null ? UnitConversion.getBasicValueforNumbers(IncomeTaxDatas.UnitId, IncomeTaxValue.Value) : 0);

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //NWC (Net Working Capital)
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "NWC (Net Working Capital)";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = SalesDatas.UnitId;
        //                var NWCDatas = InputDatasList.Find(x => x.LineItem.Contains("NWC"));
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;

        //                    ProjectOutputValuesViewModel salesValue = SalesDatas.ProjectOutputValuesVM != null && SalesDatas.ProjectOutputValuesVM.Count > 0 ? SalesDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectInputValuesViewModel NWCValue = NWCDatas.ProjectInputValuesVM != null && NWCDatas.ProjectInputValuesVM.Count > 0 ? NWCDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    value = ((salesValue != null ? UnitConversion.getBasicValueforNumbers(SalesDatas.UnitId, salesValue.Value) : 0) * (NWCValue != null ? Convert.ToDouble(NWCValue.Value) : 0)) / 100;

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Plus: Depreciation
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Plus: Depreciation";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                var DepreciationOutputDatas = OutputDatasVMList.Find(x => x.LineItem == "Depreciation");
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationOutputDatas != null ? DepreciationOutputDatas.UnitId : null;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.ProjectOutputValuesVM = DepreciationOutputDatas != null ? DepreciationOutputDatas.ProjectOutputValuesVM : null;

        //                //foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                //{
        //                //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                //    //Sales-COGS
        //                //    dumyValue = new ProjectOutputValuesViewModel();
        //                //    dumyValue = TempValue;
        //                //    double value = 0;


        //                //    dumyValue.BasicValue = value;
        //                //    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                //   // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                //    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                //}
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Less: Capital Expenditures
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Less: Capital Expenditures";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                // find Capex
        //                List<ProjectInputDatasViewModel> CapexList = new List<ProjectInputDatasViewModel>();
        //                CapexList = InputDatasList.FindAll(x => x.LineItem == "Capex" && x.SubHeader.Contains("Capex & Depreciation")).ToList();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = CapexList != null && CapexList.Count > 0 ? CapexList.FirstOrDefault().UnitId : null;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    if (CapexList != null && CapexList.Count > 0)
        //                        foreach (var item in CapexList)
        //                        {
        //                            ProjectInputValuesViewModel Value = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, Value.Value) : 0));
        //                        }
        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Less: Increases in NWC
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Less: Increases in NWC";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                ProjectOutputDatasViewModel NWCObj = OutputDatasVMList.Find(x => x.LineItem.Contains("NWC"));
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = NWCObj.UnitId;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel sameValue = NWCObj != null && NWCObj.ProjectOutputValuesVM != null && NWCObj.ProjectOutputValuesVM.Count > 0 ? NWCObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel PrevValue = NWCObj != null && NWCObj.ProjectOutputValuesVM != null && NWCObj.ProjectOutputValuesVM.Count > 0 ? NWCObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

        //                    value = ((sameValue != null ? UnitConversion.getBasicValueforNumbers(NWCObj.UnitId, sameValue.Value) : 0) - (PrevValue != null ? UnitConversion.getBasicValueforNumbers(NWCObj.UnitId, PrevValue.Value) : 0));
        //                    dumyValue.BasicValue = value;

        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                List<double> FCFarray = new List<double>();

        //                //Free Cash Flow
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Free Cash Flow";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                ProjectOutputDatasViewModel UnleaveredDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Net Income"));
        //                ProjectOutputDatasViewModel PlusDepreciationDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Plus: Depreciation"));
        //                ProjectOutputDatasViewModel LessCapexDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Less: Capital Expenditures"));
        //                ProjectOutputDatasViewModel LessNWCDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Less: Increases in NWC"));
        //                highunit = UnitConversion.getHigherDenominationUnit(UnleaveredDatasObj.UnitId, PlusDepreciationDatasObj.UnitId);
        //                //Plus: Depreciation
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    //formula = C38+C40-C41-C42
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel UnleaveredValue = UnleaveredDatasObj != null && UnleaveredDatasObj.ProjectOutputValuesVM != null && UnleaveredDatasObj.ProjectOutputValuesVM.Count > 0 ? UnleaveredDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel PlusDepreciationValue = PlusDepreciationDatasObj != null && PlusDepreciationDatasObj.ProjectOutputValuesVM != null && PlusDepreciationDatasObj.ProjectOutputValuesVM.Count > 0 ? PlusDepreciationDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel LessCapexValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel LessNWCValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    value = ((UnleaveredValue != null ? UnitConversion.getBasicValueforNumbers(UnleaveredDatasObj.UnitId, UnleaveredValue.Value) : 0) + (PlusDepreciationValue != null ? UnitConversion.getBasicValueforNumbers(PlusDepreciationDatasObj.UnitId, PlusDepreciationValue.Value) : 0) - (LessCapexValue != null ? UnitConversion.getBasicValueforNumbers(LessCapexDatasObj.UnitId, LessCapexValue.Value) : 0) - (LessNWCValue != null ? UnitConversion.getBasicValueforNumbers(LessNWCDatasObj.UnitId, LessNWCValue.Value) : 0));

        //                    dumyValue.BasicValue = value;
        //                    FCFarray.Add(value);
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                if (projectVM.ValuationTechniqueId != null)
        //                {
        //                    ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                    var UCCInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
        //                    List<ProjectOutputValuesViewModel> reverseList = new List<ProjectOutputValuesViewModel>();
        //                    var WACC = InputDatasList.Find(x => x.LineItem.Contains("Weighted average cost of capital") || x.LineItem.Contains("WACC"));
        //                    var InterestCoverageDatas = InputDatasList.Find(x => x.LineItem.Contains("Interest coverage ratio =Interest expense/FCF"));
        //                    var TFCFValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                    if (projectVM.ValuationTechniqueId == 1 || projectVM.ValuationTechniqueId == 4)
        //                    {
        //                        #region Valuation1

        //                        //Levered Value (VL)
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();


        //                        //  ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                        //ProjectOutputValuesViewModel lastRecord = dumyValueList.LastOrDefault();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));


        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Discount Factor
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discount Factor";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = null;
        //                        double TValue = 1 + (WACC != null && WACC.Value != null ? Convert.ToDouble(WACC.Value) / 100 : 0);
        //                        int i = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            value = Math.Pow(TValue, i);
        //                            i++;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = dumyValue.BasicValue;
        //                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Discounted Cash Flow
        //                        //formula = Free Cash flow/Discount Factor
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discounted Cash Flow";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        ProjectOutputDatasViewModel FreeCashFlowDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                        ProjectOutputDatasViewModel DiscountFactorDatas = OutputDatasVMList.Find(x => x.LineItem == "Discount Factor");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;
        //                        double NPv = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowDatas != null && FreeCashFlowDatas.ProjectOutputValuesVM != null && FreeCashFlowDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel DiscountFactorValues = DiscountFactorDatas != null && DiscountFactorDatas.ProjectOutputValuesVM != null && DiscountFactorDatas.ProjectOutputValuesVM.Count > 0 ? DiscountFactorDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = (FreeCashValue != null ? UnitConversion.getBasicValueforNumbers(FreeCashFlowDatas.UnitId, FreeCashValue.Value) : 0) / (DiscountFactorValues != null ? UnitConversion.getBasicValueforNumbers(DiscountFactorDatas.UnitId, DiscountFactorValues.Value) : 0);


        //                            NPv = NPv + value;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPv);

        //                        OutputDatasVMList.Add(OutputDatasVM);



        //                        //IRR (Internal Rate of Return) of Free Cash Flows
        //                        //formula = 



        //                        #endregion
        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 2)
        //                    {
        //                        #region Valuation2
        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU ($M) - VU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));


        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);
        //                        #endregion
        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 3)
        //                    {
        //                        #region Valuation3
        //                        //Free Cash Flow
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Free Cash Flow to Equity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        highunit = UnitConversion.getHigherDenominationUnit(UnleaveredDatasObj.UnitId, PlusDepreciationDatasObj.UnitId);
        //                        //Plus: Depreciation
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            //formula = C38+C40-C41-C42
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel UnleaveredValue = UnleaveredDatasObj != null && UnleaveredDatasObj.ProjectOutputValuesVM != null && UnleaveredDatasObj.ProjectOutputValuesVM.Count > 0 ? UnleaveredDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel PlusDepreciationValue = PlusDepreciationDatasObj != null && PlusDepreciationDatasObj.ProjectOutputValuesVM != null && PlusDepreciationDatasObj.ProjectOutputValuesVM.Count > 0 ? PlusDepreciationDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel LessCapexValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel LessNWCValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            value = ((UnleaveredValue != null ? UnitConversion.getBasicValueforNumbers(UnleaveredDatasObj.UnitId, UnleaveredValue.Value) : 0) + (PlusDepreciationValue != null ? UnitConversion.getBasicValueforNumbers(PlusDepreciationDatasObj.UnitId, PlusDepreciationValue.Value) : 0) - (LessCapexValue != null ? UnitConversion.getBasicValueforNumbers(LessCapexDatasObj.UnitId, LessCapexValue.Value) : 0) - (LessNWCValue != null ? UnitConversion.getBasicValueforNumbers(LessNWCDatasObj.UnitId, LessNWCValue.Value) : 0));

        //                            dumyValue.BasicValue = value;
        //                            FCFarray.Add(value);
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        //Levered Value (VL)
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //  ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                        //ProjectOutputValuesViewModel lastRecord = dumyValueList.LastOrDefault();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));


        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            reverseList.Add(dumyValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //calculate Debt Capacity
        //                        //formula=levered Value * D/V Ratio
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Debt Capacity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        ProjectOutputDatasViewModel LeveredValueDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Levered Value (VL)"));
        //                        var DVRatio = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = LeveredValueDatas != null ? LeveredValueDatas.UnitId : null;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            //formula = C38+C40-C41-C42
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel LeveredValue = LeveredValueDatas != null && LeveredValueDatas.ProjectOutputValuesVM != null && LeveredValueDatas.ProjectOutputValuesVM.Count > 0 ? LeveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = ((LeveredValue != null ? UnitConversion.getBasicValueforNumbers(LeveredValueDatas.UnitId, LeveredValue.Value) : 0) * (DVRatio != null && DVRatio.Value != null ? Convert.ToDouble(DVRatio.Value) : 0) / 100);

        //                            dumyValue.BasicValue = value;
        //                            FCFarray.Add(value);
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Discount Factor
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discount Factor";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = null;
        //                        var Costofequity = InputDatasList.Find(x => x.LineItem.Contains("Cost of Equity"));

        //                        double TValue = 1 + (Costofequity != null && Costofequity.Value != null ? Convert.ToDouble(Costofequity.Value) / 100 : 0);
        //                        int i = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            value = Math.Pow(TValue, i);
        //                            i++;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = dumyValue.BasicValue;
        //                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Discounted Cash Flow
        //                        //formula = Free Cash flow/Discount Factor
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discounted Cash Flow";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        ProjectOutputDatasViewModel FreeCashFlowtoEquityDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow to Equity"));
        //                        ProjectOutputDatasViewModel DiscountFactorDatas = OutputDatasVMList.Find(x => x.LineItem == "Discount Factor");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowtoEquityDatas != null ? FreeCashFlowtoEquityDatas.UnitId : null;
        //                        double NPv = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowtoEquityDatas != null && FreeCashFlowtoEquityDatas.ProjectOutputValuesVM != null && FreeCashFlowtoEquityDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowtoEquityDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel DiscountFactorValues = DiscountFactorDatas != null && DiscountFactorDatas.ProjectOutputValuesVM != null && DiscountFactorDatas.ProjectOutputValuesVM.Count > 0 ? DiscountFactorDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = (FreeCashValue != null ? UnitConversion.getBasicValueforNumbers(FreeCashFlowtoEquityDatas.UnitId, FreeCashValue.Value) : 0) / (DiscountFactorValues != null ? UnitConversion.getBasicValueforNumbers(DiscountFactorDatas.UnitId, DiscountFactorValues.Value) : 0);


        //                            NPv = NPv + value;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowtoEquityDatas != null ? FreeCashFlowtoEquityDatas.UnitId : null;
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPv);

        //                        OutputDatasVMList.Add(OutputDatasVM);



        //                        #endregion
        //                    }
        //                    else if (projectVM.ValuationMethodId == 5)
        //                    {
        //                        #region Valuation 5
        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //formula=Interest coverage ratio =Interest expense/FCF (C187)* marginal Tax * Unleavered value of first year

        //                        var UnleaveredShortDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Value @ rU") && x.SubHeader == "APV Based Valuation - Short Approach");

        //                        // formula==$C$15(Marginal Tax)*$C$187*C208(Unleavered Value)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "PV of Interest Tax Shield";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleaveredShortDatas.UnitId;
        //                        double ULvalue = (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (lastYearValue);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, ULvalue);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Levered Value (VL = VU + T)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

        //                        double LValue = ((MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (lastYearValue)) + lastYearValue;
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, LValue);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;


        //                        double NPV = ((MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (lastYearValue)) + lastYearValue + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = UnleaveredShortDatas != null && UnleaveredShortDatas.ProjectOutputValuesVM != null ? UnleaveredShortDatas.ProjectOutputValuesVM : null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Interest Paid @ rD";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            value = (FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) * (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0);

        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Debt Capacity
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Debt Capacity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //Interest paid
        //                        var InterestpaidLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Paid @ rD" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        var CostofDebt = InputDatasList.Find(x => x.LineItem == "Cost of Debt");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year + 1) : null;
        //                            value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0);
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Interest Tax Shield @ Tc
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
        //                        first = 0;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                dumyValue.BasicValue = dumyValue.Value = value;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0);
        //                                dumyValue.BasicValue = value;
        //                                dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            }
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Tax Shield Value @ rU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Tax Shield Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;
        //                        var InterestTaxShieldLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Tax Shield @ Tc" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestTaxShieldLongDatas.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of InterestTaxShieldLongDatas L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel InterestTaxShieldValue = InterestTaxShieldLongDatas != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestTaxShieldLongDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((InterestTaxShieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestTaxShieldLongDatas.UnitId, InterestTaxShieldValue.Value) : 0) + lastYearValue) / (1 + (UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) / 100 : 0));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        //formula = Unlevered Value @ rU+ Tax Shield Value @ rU
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;
        //                        var UleaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        var InterestShieldValueLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Tax Shield Value @ rU" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        highunit = UnitConversion.getHigherDenominationUnit((UleaveredLongDatas != null ? UleaveredLongDatas.UnitId : null), (InterestShieldValueLongDatas != null ? InterestShieldValueLongDatas.UnitId : null));
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel LUnleaveredValue = UleaveredLongDatas != null && UleaveredLongDatas.ProjectOutputValuesVM != null && UleaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? UleaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                ProjectOutputValuesViewModel InterestDhieldValue = InterestShieldValueLongDatas != null && InterestShieldValueLongDatas.ProjectOutputValuesVM != null && InterestShieldValueLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestShieldValueLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (LUnleaveredValue != null ? UnitConversion.getBasicValueforCurrency(UleaveredLongDatas.UnitId, LUnleaveredValue.Value) : 0) + (InterestDhieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestShieldValueLongDatas.UnitId, InterestDhieldValue.Value) : 0);
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        var leaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredLongDatas.UnitId);
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;


        //                        var levLongValue = leaveredLongDatas != null && leaveredLongDatas.ProjectOutputValuesVM != null && leaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? leaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        #endregion
        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 6)
        //                    {
        //                        #region Valuation 6
        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Debt Capacity
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Debt Capacity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        var InputDebtCapacity = InputDatasList.Find(x => x.LineItem == "Fixed Schedule & Predetermined Debt Level, Dt");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InputDebtCapacity.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            var InputDebtCapacityValue = InputDebtCapacity != null && InputDebtCapacity.ProjectInputValuesVM != null && InputDebtCapacity.ProjectInputValuesVM.Count > 0 ? InputDebtCapacity.ProjectInputValuesVM.Find(x => x.Year == (TempValue.Year)) : null;
        //                            value = InputDebtCapacityValue != null ? UnitConversion.getBasicValueforCurrency(InputDebtCapacity.UnitId, InputDebtCapacityValue.Value) : 0;
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        // Interest Paid
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Interest Paid @ rD";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        //  OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

        //                        //  OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

        //                        var DebtCapacityDatas = OutputDatasVMList.Find(x => x.LineItem == "Debt Capacity");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DebtCapacityDatas.UnitId;

        //                        var CostOfDebtInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));

        //                        first = 0;
        //                        double FirstYearValue = 0;

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderBy(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            //change formula
        //                            //ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            //value = (FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) * (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0);

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel DebtCapacityValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;
        //                                value = (DebtCapacityValue != null ? UnitConversion.getBasicValueforCurrency(DebtCapacityDatas.UnitId, DebtCapacityValue.Value) : 0) * ((CostOfDebtInputDatas != null && CostOfDebtInputDatas.Value != null ? Convert.ToDouble(CostOfDebtInputDatas.Value) : 0) / 100);
        //                            }

        //                            dumyValue.BasicValue = value;
        //                            FirstYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Interest Tax Shield @ Tc
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        var InterestpaidLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Paid @ rD");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
        //                        first = 0;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                dumyValue.BasicValue = dumyValue.Value = value;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0);
        //                                dumyValue.BasicValue = value;
        //                                dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            }
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Tax Shield Value @ rU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Tax Shield Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;
        //                        var CostofDebt = InputDatasList.Find(x => x.LineItem == "Cost of Debt");

        //                        var InterestTaxShieldLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Tax Shield @ Tc");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestTaxShieldLongDatas.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of InterestTaxShieldLongDatas L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel InterestTaxShieldValue = InterestTaxShieldLongDatas != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestTaxShieldLongDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((InterestTaxShieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestTaxShieldLongDatas.UnitId, InterestTaxShieldValue.Value) : 0) + lastYearValue) / (1 + (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //formula = Unlevered Value @ rU+ Tax Shield Value @ rU
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;
        //                        var UleaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU");
        //                        var InterestShieldValueLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Tax Shield Value @ rU");
        //                        highunit = UnitConversion.getHigherDenominationUnit((UleaveredLongDatas != null ? UleaveredLongDatas.UnitId : null), (InterestShieldValueLongDatas != null ? InterestShieldValueLongDatas.UnitId : null));
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel LUnleaveredValue = UleaveredLongDatas != null && UleaveredLongDatas.ProjectOutputValuesVM != null && UleaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? UleaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                ProjectOutputValuesViewModel InterestDhieldValue = InterestShieldValueLongDatas != null && InterestShieldValueLongDatas.ProjectOutputValuesVM != null && InterestShieldValueLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestShieldValueLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (LUnleaveredValue != null ? UnitConversion.getBasicValueforCurrency(UleaveredLongDatas.UnitId, LUnleaveredValue.Value) : 0) + (InterestDhieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestShieldValueLongDatas.UnitId, InterestDhieldValue.Value) : 0);
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        // formula==$C$15(Marginal Tax)*$C$187*C208(Unleavered Value)
        //                        //OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //OutputDatasVM.Id = 0;
        //                        //OutputDatasVM.ProjectId = projectVM.Id;
        //                        //OutputDatasVM.HeaderId = 0;
        //                        //OutputDatasVM.SubHeader = "";
        //                        //OutputDatasVM.LineItem = "PV of Interest Tax Shield";
        //                        //OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        //OutputDatasVM.HasMultiYear = false;
        //                        //OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleaveredDatasObj.UnitId;
        //                        //double ULvalue = (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (lastYearValue);
        //                        //OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, ULvalue);
        //                        //OutputDatasVM.ProjectOutputValuesVM = null;
        //                        //OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        var leaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)");

        //                        highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredLongDatas.UnitId);
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                        var levLongValue = leaveredLongDatas != null && leaveredLongDatas.ProjectOutputValuesVM != null && leaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? leaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        double NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        #endregion
        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 7)
        //                    {
        //                        #region Valuation 7

        //                        //formula = Free Cash flow+ Next Year Value of Levered Value/(1+WACC)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "WACC Based Valuation";
        //                        OutputDatasVM.LineItem = "Levered Value @ rWACC";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "WACC Based Valuation";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        var leaveredDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value @ rWACC" && x.SubHeader == "WACC Based Valuation");

        //                        highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredDatas.UnitId);
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                        var levLongValue = leaveredDatas != null && leaveredDatas.ProjectOutputValuesVM != null && leaveredDatas.ProjectOutputValuesVM.Count > 0 ? leaveredDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        double NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //formula = Free Cash flow+ Next Year Value of Levered Value/(1+WACC)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);
        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        var leaveredAPVDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value @ rWACC" && x.SubHeader == "APV Based Valuation");

        //                        highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredDatas.UnitId);
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                        levLongValue = leaveredAPVDatas != null && leaveredAPVDatas.ProjectOutputValuesVM != null && leaveredAPVDatas.ProjectOutputValuesVM.Count > 0 ? leaveredAPVDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        #endregion

        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 8)
        //                    {
        //                        #region Valuation8

        //                        //Unlevered Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L346+L349)/(1+$C$328); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        // PV of Interest Tax Shield
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "PV of Interest Tax Shield";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        // OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;

        //                        var ICRInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Intetrest coverage ratio =Interest expense/FCF"));
        //                        var CBInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
        //                        ProjectOutputDatasViewModel UnleveredValueDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU");

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleveredValueDatas != null ? UnleveredValueDatas.UnitId : null;

        //                        // $C$15 $C$327  ((1 + C328) / (1 + C329)) * C349

        //                        ProjectOutputValuesViewModel UnLeveredValue = UnleveredValueDatas != null && UnleveredValueDatas.ProjectOutputValuesVM != null && UnleveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

        //                        OutputDatasVM.Value = ((MarginalTaxDatas != null && MarginalTaxDatas.Value != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0) / 100) * ((ICRInputDatas != null && ICRInputDatas.Value != null ? Convert.ToDouble(ICRInputDatas.Value) : 0) / 100) * ((1 + ((UCCInputDatas != null && UCCInputDatas.Value != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100)) / (1 + ((CBInputDatas != null && CBInputDatas.Value != null ? Convert.ToDouble(CBInputDatas.Value) : 0) / 100))) * (UnLeveredValue != null ? UnLeveredValue.Value : 0);
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Levered Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;

        //                        ProjectOutputDatasViewModel PVInterestDatas = OutputDatasVMList.Find(x => x.LineItem == "PV of Interest Tax Shield");

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleveredValueDatas != null ? UnleveredValueDatas.UnitId : null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = PVInterestDatas != null ? PVInterestDatas.UnitId : null;

        //                        // C349 + C350

        //                        //ProjectOutputValuesViewModel PVInterestValue = PVInterestDatas != null && PVInterestDatas.ProjectOutputValuesVM != null && PVInterestDatas.ProjectOutputValuesVM.Count > 0 ? PVInterestDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

        //                        // OutputDatasVM.Value = (UnLeveredValue != null ? UnLeveredValue.Value : 0) + (PVInterestValue != null ? PVInterestValue.Value : 0);
        //                        OutputDatasVM.Value = (UnLeveredValue != null ? UnLeveredValue.Value : 0) + (PVInterestDatas != null ? PVInterestDatas.Value : 0);

        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        //Net Present Value ($M)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;

        //                        ProjectOutputDatasViewModel LeveredValueDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)");
        //                        ProjectOutputDatasViewModel FreeCashFlowDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = LeveredValueDatas != null ? LeveredValueDatas.UnitId : null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;

        //                        // C351 + C346

        //                        //  ProjectOutputValuesViewModel LeveredValue = LeveredValueDatas != null && LeveredValueDatas.ProjectOutputValuesVM != null && LeveredValueDatas.ProjectOutputValuesVM.Count > 0 ? LeveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowDatas != null && FreeCashFlowDatas.ProjectOutputValuesVM != null && FreeCashFlowDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

        //                        //  OutputDatasVM.Value = (LeveredValue != null ? LeveredValue.Value : 0) + (FreeCashValue != null ? FreeCashValue.Value : 0);
        //                        OutputDatasVM.Value = (LeveredValueDatas != null ? LeveredValueDatas.Value : 0) + (FreeCashValue != null ? FreeCashValue.Value : 0);

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        #endregion
        //                    }

        //                    if (projectVM.ValuationTechniqueId != 8)
        //                    {
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "IRR (Internal Rate of Return) of Free Cash Flows";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        try
        //                        {
        //                            OutputDatasVM.Value = Financial.Irr(FCFarray);
        //                        }
        //                        catch (Exception ss)
        //                        {
        //                            OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(3, FCFarray.Sum());
        //                        }
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);
        //                    }

        //                }



        //            }

        //        }



        //    }
        //    catch (Exception ss)
        //    {
        //        // return OutputDatasVMList;
        //    }
        //    return OutputDatasVMList;
        //}

        // GetProjectOutput1
        //public List<ProjectOutputDatasViewModel> GetProjectOutput(ProjectsViewModel projectVM)
        //{

        //    // List<ProjectOutputDatasViewModel> SummaryList = new List<ProjectOutputDatasViewModel>()
        //    List<ProjectOutputDatasViewModel> OutputDatasVMList = new List<ProjectOutputDatasViewModel>();
        //    try
        //    {
        //        //Calculate Output by Project Inputs

        //        if (projectVM != null && projectVM.ProjectInputDatasVM != null && projectVM.ProjectInputDatasVM.Count > 0)
        //        {

        //            //create dumy year Value List
        //            List<ProjectOutputValuesViewModel> dumyValueList = new List<ProjectOutputValuesViewModel>();
        //            if (projectVM.StartingYear != null && projectVM.StartingYear != 0 && projectVM.NoOfYears != null && projectVM.NoOfYears > 0)
        //            {
        //                ProjectOutputValuesViewModel dumyValue = new ProjectOutputValuesViewModel();

        //                for (int i = 0; i < projectVM.NoOfYears; i++)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    dumyValue.Id = 0;
        //                    dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = Convert.ToInt32(projectVM.StartingYear) + i;
        //                    dumyValue.Value = null;
        //                    dumyValueList.Add(dumyValue);
        //                }

        //                ProjectOutputDatasViewModel OutputDatasVM = new ProjectOutputDatasViewModel();

        //                List<ProjectInputDatasViewModel> InputDatasList = new List<ProjectInputDatasViewModel>();
        //                InputDatasList = projectVM.ProjectInputDatasVM;

        //                //create data //for Sales 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Sales";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                // ($O$6*C6)*($O$7*C7)/1000000
        //                List<ProjectInputDatasViewModel> VolumeList = new List<ProjectInputDatasViewModel>();
        //                VolumeList = InputDatasList.FindAll(x => x.LineItem == "Volume" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

        //                List<ProjectInputDatasViewModel> UnitPriceList = new List<ProjectInputDatasViewModel>();
        //                UnitPriceList = InputDatasList.FindAll(x => x.LineItem == "Unit Price" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();
        //                int? highunit = UnitConversion.getHigherDenominationUnit(VolumeList.FirstOrDefault().UnitId, UnitPriceList.FirstOrDefault().UnitId);

        //                //sales=volume*unitcost;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();

        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;

        //                    double value = 0;
        //                    //calculate sales for multiple years

        //                    if (VolumeList != null && VolumeList.Count > 0 && UnitPriceList != null && UnitPriceList.Count > 0)
        //                        foreach (var item in VolumeList)
        //                        {
        //                            var UnitPriceDatas = UnitPriceList.Find(x => x.SubHeader == item.SubHeader);
        //                            ProjectInputValuesViewModel UnitCostValue = UnitPriceDatas != null && UnitPriceDatas.ProjectInputValuesVM != null && UnitPriceDatas.ProjectInputValuesVM.Count > 0 ? UnitPriceDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectInputValuesViewModel VolumeValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((UnitPriceDatas != null ? UnitConversion.getBasicValueforCurrency(UnitPriceDatas.UnitId, UnitCostValue.Value) : 0) * (item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, VolumeValue.Value) : 0));
        //                        }




        //                    //get sum of all the Vlaues

        //                    // value = UnitConversion.getHigherDenominationUnit(UnitCostDatas.UnitId,);

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign sum of all sales value
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);


        //                //create data //for COGS 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "COGS";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                List<ProjectInputDatasViewModel> UnitCostList = new List<ProjectInputDatasViewModel>();
        //                UnitCostList = InputDatasList.FindAll(x => x.LineItem == "Unit Cost" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

        //                highunit = UnitConversion.getHigherDenominationUnit(VolumeList.FirstOrDefault().UnitId, UnitCostList.FirstOrDefault().UnitId);
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    //calculate COGS for multiple years

        //                    //get sum of all the Vlaues
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    //calculate sales for multiple years
        //                    if (VolumeList != null && VolumeList.Count > 0 && UnitCostList != null && UnitCostList.Count > 0)
        //                        foreach (var item in VolumeList)
        //                        {
        //                            var UnitCostDatas = UnitCostList.Find(x => x.SubHeader == item.SubHeader);
        //                            ProjectInputValuesViewModel UnitCostValue = UnitCostDatas != null && UnitCostDatas.ProjectInputValuesVM != null && UnitCostDatas.ProjectInputValuesVM.Count > 0 ? UnitCostDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectInputValuesViewModel VolumeValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((UnitCostDatas != null ? UnitConversion.getBasicValueforCurrency(UnitCostDatas.UnitId, UnitCostValue.Value) : 0) * (item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, VolumeValue.Value) : 0));
        //                        }


        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    //value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign sum of all sales value
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Gross Margin
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Gross Margin";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                var SalesDatas = OutputDatasVMList.Find(x => x.LineItem == "Sales");
        //                var COGSDatas = OutputDatasVMList.Find(x => x.LineItem == "COGS");
        //                highunit = UnitConversion.getHigherDenominationUnit(SalesDatas.UnitId, COGSDatas.UnitId);
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel salesValue = SalesDatas.ProjectOutputValuesVM != null && SalesDatas.ProjectOutputValuesVM.Count > 0 ? SalesDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel COGSValue = COGSDatas.ProjectOutputValuesVM != null && COGSDatas.ProjectOutputValuesVM.Count > 0 ? COGSDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                    value = (salesValue != null ? UnitConversion.getBasicValueforNumbers(SalesDatas.UnitId, salesValue.Value) : 0) - (COGSValue != null ? UnitConversion.getBasicValueforNumbers(COGSDatas.UnitId, COGSValue.Value) : 0);
        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);


        //                    // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //all items of fixed cost with the same formula
        //                //find all items of Other Fixed Cost
        //                List<ProjectInputDatasViewModel> OtherFixedCostList = new List<ProjectInputDatasViewModel>();
        //                OtherFixedCostList = InputDatasList.FindAll(x => x.SubHeader.Contains("Other Fixed Cost")).ToList();

        //                if (OtherFixedCostList != null && OtherFixedCostList.Count > 0)
        //                    foreach (ProjectInputDatasViewModel fixedCostObj in OtherFixedCostList)
        //                    {
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = fixedCostObj.LineItem;
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = fixedCostObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectInputValuesViewModel Value = fixedCostObj.ProjectInputValuesVM != null && fixedCostObj.ProjectInputValuesVM.Count > 0 ? fixedCostObj.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((fixedCostObj != null ? UnitConversion.getBasicValueforNumbers(fixedCostObj.UnitId, Value.Value) : 0));
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));//assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                    }

        //                //formula=Startyear+1 value * year Value

        //                //Depreciation // Bind Depreciation Values directly from Input
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Depreciation";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                // find depreciation
        //                List<ProjectInputDatasViewModel> DepreciationList = new List<ProjectInputDatasViewModel>();
        //                DepreciationList = InputDatasList.FindAll(x => x.LineItem == "Depreciation" && x.SubHeader.Contains("Capex & Depreciation")).ToList();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationList != null && DepreciationList.Count > 0 ? DepreciationList.FirstOrDefault().UnitId : null;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    //Sales-COGS
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    if (DepreciationList != null && DepreciationList.Count > 0)
        //                        foreach (var item in DepreciationList)
        //                        {
        //                            ProjectInputValuesViewModel depreciationValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, depreciationValue.Value) : 0));
        //                        }
        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));//assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);



        //                //Operating Income 
        //                //formula = Gross Margin - (sum of all items after Gross Margin)
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Operating Income";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                var GrossMarginDatas = OutputDatasVMList.Find(x => x.LineItem == "Gross Margin");
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = GrossMarginDatas.UnitId;

        //                var DepreciationDatas = OutputDatasVMList.Find(x => x.LineItem == "Depreciation");
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationDatas.UnitId;

        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    double sum = 0;
        //                    ProjectOutputValuesViewModel GrossValue = GrossMarginDatas.ProjectOutputValuesVM != null && GrossMarginDatas.ProjectOutputValuesVM.Count > 0 ? GrossMarginDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                    ProjectOutputValuesViewModel DepreciationValue = DepreciationDatas.ProjectOutputValuesVM != null && DepreciationDatas.ProjectOutputValuesVM.Count > 0 ? DepreciationDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                    foreach (ProjectInputDatasViewModel fixedCostObj in OtherFixedCostList)
        //                    {
        //                        var otherFixedValue = fixedCostObj.ProjectInputValuesVM != null && fixedCostObj.ProjectInputValuesVM.Count > 0 ? fixedCostObj.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                        sum = sum + (otherFixedValue != null ? UnitConversion.getBasicValueforNumbers(fixedCostObj.UnitId, otherFixedValue.Value) : 0);
        //                    }

        //                    // value = (GrossMarginDatas != null ? UnitConversion.getBasicValueforNumbers(GrossMarginDatas.UnitId, GrossValue.Value) : 0) - sum;//-()sum of all items;

        //                    value = (GrossMarginDatas != null ? UnitConversion.getBasicValueforNumbers(GrossMarginDatas.UnitId, GrossValue.Value) : 0) - sum - (DepreciationDatas != null ? UnitConversion.getBasicValueforNumbers(DepreciationDatas.UnitId, DepreciationValue.Value) : 0);//-()sum of all items;

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                    //  value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    //  dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);


        //                //Income Tax
        //                //formula = Operating Income (for diff year) * Marginal Tax
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Income Tax";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                var OperatingIncomeDatas = OutputDatasVMList.Find(x => x.LineItem == "Operating Income");

        //                var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));

        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = OperatingIncomeDatas.UnitId;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                    value = ((OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0)) / 100;

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                    // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);

        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Unlevered Net Income
        //                //formula = Operating Income (for diff year) -Income Tax
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Unlevered Net Income";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;

        //                var IncomeTaxDatas = OutputDatasVMList.Find(x => x.LineItem == "Income Tax");
        //                highunit = UnitConversion.getHigherDenominationUnit(OperatingIncomeDatas.UnitId, IncomeTaxDatas.UnitId);
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel IncomeTaxValue = IncomeTaxDatas.ProjectOutputValuesVM != null && IncomeTaxDatas.ProjectOutputValuesVM.Count > 0 ? IncomeTaxDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    value = (OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) - (IncomeTaxValue != null ? UnitConversion.getBasicValueforNumbers(IncomeTaxDatas.UnitId, IncomeTaxValue.Value) : 0);

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                    // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //NWC (Net Working Capital)
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "NWC (Net Working Capital)";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = SalesDatas.UnitId;
        //                var NWCDatas = InputDatasList.Find(x => x.LineItem.Contains("NWC"));
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;

        //                    ProjectOutputValuesViewModel salesValue = SalesDatas.ProjectOutputValuesVM != null && SalesDatas.ProjectOutputValuesVM.Count > 0 ? SalesDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectInputValuesViewModel NWCValue = NWCDatas.ProjectInputValuesVM != null && NWCDatas.ProjectInputValuesVM.Count > 0 ? NWCDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                    value = ((salesValue != null ? UnitConversion.getBasicValueforNumbers(SalesDatas.UnitId, salesValue.Value) : 0) * (NWCValue != null ? Convert.ToDouble(NWCValue.Value) : 0)) / 100;

        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Plus: Depreciation
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Plus: Depreciation";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                var DepreciationOutputDatas = OutputDatasVMList.Find(x => x.LineItem == "Depreciation");

        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationOutputDatas != null ? DepreciationOutputDatas.UnitId : null;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.ProjectOutputValuesVM = DepreciationOutputDatas != null ? DepreciationOutputDatas.ProjectOutputValuesVM : null;

        //                //foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                //{
        //                //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                //    //Sales-COGS
        //                //    dumyValue = new ProjectOutputValuesViewModel();
        //                //    dumyValue = TempValue;
        //                //    double value = 0;


        //                //    dumyValue.BasicValue = value;
        //                //    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                //   // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                //    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                //}
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Less: Capital Expenditures
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Less: Capital Expenditures";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                // find Capex
        //                List<ProjectInputDatasViewModel> CapexList = new List<ProjectInputDatasViewModel>();
        //                CapexList = InputDatasList.FindAll(x => x.LineItem == "Capex" && x.SubHeader.Contains("Capex & Depreciation")).ToList();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = CapexList != null && CapexList.Count > 0 ? CapexList.FirstOrDefault().UnitId : null;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    if (CapexList != null && CapexList.Count > 0)
        //                        foreach (var item in CapexList)
        //                        {
        //                            ProjectInputValuesViewModel Value = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = value + ((item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, Value.Value) : 0));
        //                        }
        //                    dumyValue.BasicValue = value;
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                //Less: Increases in NWC
        //                //formula = 
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Less: Increases in NWC";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                ProjectOutputDatasViewModel NWCObj = OutputDatasVMList.Find(x => x.LineItem.Contains("NWC"));
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = NWCObj.UnitId;
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {

        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel sameValue = NWCObj != null && NWCObj.ProjectOutputValuesVM != null && NWCObj.ProjectOutputValuesVM.Count > 0 ? NWCObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel PrevValue = NWCObj != null && NWCObj.ProjectOutputValuesVM != null && NWCObj.ProjectOutputValuesVM.Count > 0 ? NWCObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

        //                    value = ((sameValue != null ? UnitConversion.getBasicValueforNumbers(NWCObj.UnitId, sameValue.Value) : 0) - (PrevValue != null ? UnitConversion.getBasicValueforNumbers(NWCObj.UnitId, PrevValue.Value) : 0));
        //                    dumyValue.BasicValue = value;

        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                List<double> FCFarray = new List<double>();

        //                //Free Cash Flow
        //                OutputDatasVM = new ProjectOutputDatasViewModel();
        //                OutputDatasVM.Id = 0;
        //                OutputDatasVM.ProjectId = projectVM.Id;
        //                OutputDatasVM.HeaderId = 0;
        //                OutputDatasVM.SubHeader = "";
        //                OutputDatasVM.LineItem = "Free Cash Flow";
        //                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                OutputDatasVM.HasMultiYear = true;
        //                OutputDatasVM.Value = null;
        //                ProjectOutputDatasViewModel UnleaveredIncomeDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Net Income"));
        //                ProjectOutputDatasViewModel PlusDepreciationDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Plus: Depreciation"));
        //                ProjectOutputDatasViewModel LessCapexDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Less: Capital Expenditures"));
        //                ProjectOutputDatasViewModel LessNWCDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Less: Increases in NWC"));
        //                highunit = UnitConversion.getHigherDenominationUnit(UnleaveredIncomeDatasObj.UnitId, PlusDepreciationDatasObj.UnitId);
        //                //Plus: Depreciation
        //                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                {
        //                    //formula = C38+C40-C41-C42
        //                    dumyValue = new ProjectOutputValuesViewModel();
        //                    // dumyValue = TempValue;
        //                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                    dumyValue.Year = TempValue.Year;
        //                    double value = 0;
        //                    ProjectOutputValuesViewModel UnleaveredIncomeValue = UnleaveredIncomeDatasObj != null && UnleaveredIncomeDatasObj.ProjectOutputValuesVM != null && UnleaveredIncomeDatasObj.ProjectOutputValuesVM.Count > 0 ? UnleaveredIncomeDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel PlusDepreciationValue = PlusDepreciationDatasObj != null && PlusDepreciationDatasObj.ProjectOutputValuesVM != null && PlusDepreciationDatasObj.ProjectOutputValuesVM.Count > 0 ? PlusDepreciationDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel LessCapexValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    ProjectOutputValuesViewModel LessNWCValue = LessNWCDatasObj != null && LessNWCDatasObj.ProjectOutputValuesVM != null && LessNWCDatasObj.ProjectOutputValuesVM.Count > 0 ? LessNWCDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                    value = ((UnleaveredIncomeValue != null ? UnitConversion.getBasicValueforNumbers(UnleaveredIncomeDatasObj.UnitId, UnleaveredIncomeValue.Value) : 0) + (PlusDepreciationValue != null ? UnitConversion.getBasicValueforNumbers(PlusDepreciationDatasObj.UnitId, PlusDepreciationValue.Value) : 0) - (LessCapexValue != null ? UnitConversion.getBasicValueforNumbers(LessCapexDatasObj.UnitId, LessCapexValue.Value) : 0) - (LessNWCValue != null ? UnitConversion.getBasicValueforNumbers(LessNWCDatasObj.UnitId, LessNWCValue.Value) : 0));

        //                    dumyValue.BasicValue = value;
        //                    FCFarray.Add(value);
        //                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                }
        //                OutputDatasVMList.Add(OutputDatasVM);

        //                if (projectVM.ValuationTechniqueId != null)

        //                {
        //                    ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                    var UCCInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
        //                    List<ProjectOutputValuesViewModel> reverseList = new List<ProjectOutputValuesViewModel>();
        //                    var WACC = InputDatasList.Find(x => x.LineItem.Contains("Weighted average cost of capital") || x.LineItem.Contains("WACC"));
        //                    var InterestCoverageDatas = InputDatasList.Find(x => x.LineItem.Contains("Interest coverage ratio =Interest expense/FCF") || x.LineItem.Contains("Intetrest coverage ratio"));
        //                    var TFCFValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                    var DVRatio = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));
        //                    var CostofDebt = InputDatasList.Find(x => x.LineItem == "Cost of Debt");
        //                    if (projectVM.ValuationTechniqueId == 1 || projectVM.ValuationTechniqueId == 4)
        //                    {
        //                        #region Valuation1

        //                        //Levered Value (VL)
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();


        //                        //  ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                        //ProjectOutputValuesViewModel lastRecord = dumyValueList.LastOrDefault();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));


        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Discount Factor
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discount Factor";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = null;
        //                        double TValue = 1 + (WACC != null && WACC.Value != null ? Convert.ToDouble(WACC.Value) / 100 : 0);
        //                        int i = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            value = Math.Pow(TValue, i);
        //                            i++;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = dumyValue.BasicValue;
        //                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Discounted Cash Flow
        //                        //formula = Free Cash flow/Discount Factor
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discounted Cash Flow";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        ProjectOutputDatasViewModel FreeCashFlowDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                        ProjectOutputDatasViewModel DiscountFactorDatas = OutputDatasVMList.Find(x => x.LineItem == "Discount Factor");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;
        //                        double NPv = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowDatas != null && FreeCashFlowDatas.ProjectOutputValuesVM != null && FreeCashFlowDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel DiscountFactorValues = DiscountFactorDatas != null && DiscountFactorDatas.ProjectOutputValuesVM != null && DiscountFactorDatas.ProjectOutputValuesVM.Count > 0 ? DiscountFactorDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = (FreeCashValue != null ? UnitConversion.getBasicValueforNumbers(FreeCashFlowDatas.UnitId, FreeCashValue.Value) : 0) / (DiscountFactorValues != null ? UnitConversion.getBasicValueforNumbers(DiscountFactorDatas.UnitId, DiscountFactorValues.Value) : 0);


        //                            NPv = NPv + value;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;

        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPv);

        //                        OutputDatasVMList.Add(OutputDatasVM);



        //                        //IRR (Internal Rate of Return) of Free Cash Flows
        //                        //formula = 



        //                        #endregion
        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 2 || projectVM.ValuationTechniqueId == 7)
        //                    {
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        if (projectVM.ValuationTechniqueId == 7)
        //                        {
        //                            #region Valuation 7

        //                            //formula = Free Cash flow+ Next Year Value of Levered Value/(1+WACC)
        //                            OutputDatasVM = new ProjectOutputDatasViewModel();
        //                            OutputDatasVM.Id = 0;
        //                            OutputDatasVM.ProjectId = projectVM.Id;
        //                            OutputDatasVM.HeaderId = 0;
        //                            OutputDatasVM.SubHeader = "WACC Based Valuation";
        //                            OutputDatasVM.LineItem = "Levered Value @ rWACC";
        //                            OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                            OutputDatasVM.HasMultiYear = true;
        //                            OutputDatasVM.Value = null;
        //                            OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                            OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                            foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                            {
        //                                dumyValue = new ProjectOutputValuesViewModel();
        //                                dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                                dumyValue.Year = TempValue.Year;
        //                                double value = 0;
        //                                if (first == 0)
        //                                {
        //                                    value = 0;
        //                                    first++;
        //                                }
        //                                else
        //                                {
        //                                    ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                    value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));
        //                                }
        //                                dumyValue.BasicValue = value;
        //                                lastYearValue = value;
        //                                dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                                OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                            }
        //                            OutputDatasVMList.Add(OutputDatasVM);

        //                            //Net Present Value
        //                            OutputDatasVM = new ProjectOutputDatasViewModel();
        //                            OutputDatasVM.Id = 0;
        //                            OutputDatasVM.ProjectId = projectVM.Id;
        //                            OutputDatasVM.HeaderId = 0;
        //                            OutputDatasVM.SubHeader = "WACC Based Valuation";
        //                            OutputDatasVM.LineItem = "Net Present Value";
        //                            OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                            OutputDatasVM.HasMultiYear = false;

        //                            var leaveredDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value @ rWACC" && x.SubHeader == "WACC Based Valuation");

        //                            highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredDatas.UnitId);
        //                            OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                            var levLongValue = leaveredDatas != null && leaveredDatas.ProjectOutputValuesVM != null && leaveredDatas.ProjectOutputValuesVM.Count > 0 ? leaveredDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                            double NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                            OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                            OutputDatasVM.ProjectOutputValuesVM = null;
        //                            OutputDatasVMList.Add(OutputDatasVM);


        //                            #endregion
        //                        }
        //                        #region Valuation2
        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));


        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        #region New approach



        //                        List<ProjectOutputValuesViewModel> DebtCapacityValueList = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;

        //                        double MarginalTaxValue = MarginalTaxDatas != null && MarginalTaxDatas.Value != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0;
        //                        double DVRatioValue = DVRatio != null && DVRatio.Value != null ? Convert.ToDouble
        //                            (DVRatio.Value) / 100 : 0;
        //                        double UCCValue = UCCInputDatas != null && UCCInputDatas.Value != null ? Convert.ToDouble(UCCInputDatas.Value) / 100 : 0;
        //                        double CostofDebtValue = CostofDebt != null && CostofDebt.Value != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0;
        //                        var UnleaveredValueDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Value @ rU"));
        //                        double unLeaveredValue = 0;

        //                        //DC List
        //                        List<ProjectOutputValuesViewModel> DCValueList = new List<ProjectOutputValuesViewModel>();

        //                        // Interest paid IP value List
        //                        List<ProjectOutputValuesViewModel> IPValueList = new List<ProjectOutputValuesViewModel>();

        //                        // Interest Tax Shield ITS Value List
        //                        List<ProjectOutputValuesViewModel> ITSValueList = new List<ProjectOutputValuesViewModel>();

        //                        // Tax Shield Value TSV Value List
        //                        List<ProjectOutputValuesViewModel> TSVValueList = new List<ProjectOutputValuesViewModel>();

        //                        //Leveared Value LV List
        //                        List<ProjectOutputValuesViewModel> LVValueList = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            // DC
        //                            var DCValue = new ProjectOutputValuesViewModel();
        //                            DCValue.Id = 0; DCValue.ProjectOutputDatasId = 0; DCValue.Year = TempValue.Year;
        //                            // IP
        //                            var IPValue = new ProjectOutputValuesViewModel();
        //                            IPValue.Id = 0; IPValue.ProjectOutputDatasId = 0; IPValue.Year = TempValue.Year;
        //                            // ITS
        //                            var ITSValue = new ProjectOutputValuesViewModel();
        //                            ITSValue.Id = 0; ITSValue.ProjectOutputDatasId = 0; ITSValue.Year = TempValue.Year;
        //                            // TSV
        //                            var TSVValue = new ProjectOutputValuesViewModel();
        //                            TSVValue.Id = 0; TSVValue.ProjectOutputDatasId = 0; TSVValue.Year = TempValue.Year;
        //                            // LV
        //                            var LVValue = new ProjectOutputValuesViewModel();
        //                            LVValue.Id = 0; LVValue.ProjectOutputDatasId = 0; LVValue.Year = TempValue.Year;


        //                            double DC = 0;
        //                            double IP = 0;
        //                            double ITS = 0;
        //                            double TSV = 0;
        //                            double LV = 0;

        //                            if (first == 0)
        //                            {
        //                                //Find Last Year Value of UnLevered Value
        //                                var UnleaveredValueGVM = UnleaveredValueDatas != null && UnleaveredValueDatas.ProjectOutputValuesVM != null && UnleaveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;
        //                                unLeaveredValue = UnleaveredValueGVM != null ? Convert.ToDouble(UnleaveredValueGVM.Value) : 0;

        //                                //for Last Year

        //                                //Interest Paid
        //                                IP = (DVRatioValue * CostofDebtValue * unLeaveredValue) / (1 - ((DVRatioValue * CostofDebtValue * MarginalTaxValue) / (1 + UCCValue)));

        //                                //ITS= Interest Paid * Marginal Tax
        //                                ITS = IP * MarginalTaxValue;
        //                                first++;
        //                            }
        //                            else
        //                            {

        //                                //Find Same Year Value of UnLevered Value
        //                                var UnleaveredValueGVM = UnleaveredValueDatas != null && UnleaveredValueDatas.ProjectOutputValuesVM != null && UnleaveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year)) : null;
        //                                unLeaveredValue = UnleaveredValueGVM != null ? Convert.ToDouble(UnleaveredValueGVM.Value) : 0;

        //                                //find Next year ITS
        //                                var NextYearITS = ITSValueList != null && ITSValueList.Count > 0 ? ITSValueList.Find(x => x.Year == (TempValue.Year + 1)) : null;

        //                                //find Next year TSV
        //                                var NextYearTSV = TSVValueList != null && TSVValueList.Count > 0 ? TSVValueList.Find(x => x.Year == (TempValue.Year + 1)) : null;

        //                                //Calculate TSV=(Next year TSV+Next year ITS)/(1+UCCValue);
        //                                TSV = (((NextYearITS != null && NextYearITS.Value != null ? Convert.ToDouble(NextYearITS.Value) : 0) + (NextYearTSV != null && NextYearTSV.Value != null ? Convert.ToDouble(NextYearTSV.Value) : 0)) / (1 + UCCValue));

        //                                //LV =Same year UV+ same year TSV
        //                                LV = unLeaveredValue + TSV;

        //                                //DC=DVRatio * LV
        //                                DC = DVRatioValue * LV;

        //                                //Find Last Year Value of UnLevered Value
        //                                var LastyearLeaveredValue = UnleaveredValueDatas != null && UnleaveredValueDatas.ProjectOutputValuesVM != null && UnleaveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

        //                                if (TempValue.Year != projectVM.StartingYear)
        //                                {

        //                                    //IP= DVRatio * Cost of Debt*(Last year UV+((Same year ITS+same year TSV)/(1+ Unleavered Cost of capital)))
        //                                    // IP = DVRatioValue * CostofDebtValue * ((LastyearLeaveredValue != null ? Convert.ToDouble(LastyearLeaveredValue.Value) : 0) + ((TSV + ITS) / (1 + UCCValue)));

        //                                    IP = ((LastyearLeaveredValue != null ? Convert.ToDouble(LastyearLeaveredValue.Value) : 0) * (1 + UCCValue) + TSV) / (((1 + UCCValue) / (DVRatioValue * CostofDebtValue)) - MarginalTaxValue);

        //                                    //ITS= Interest Paid * Marginal Tax
        //                                    ITS = IP * MarginalTaxValue;

        //                                }


        //                            }

        //                            //DC
        //                            DCValue.Value = DC;
        //                            DCValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, DCValue.Value); DCValueList.Add(DCValue);

        //                            //IP
        //                            IPValue.Value = IP;
        //                            IPValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, IPValue.Value); IPValueList.Add(IPValue);

        //                            //ITS
        //                            ITSValue.Value = ITS;
        //                            ITSValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, ITSValue.Value); ITSValueList.Add(ITSValue);

        //                            //TSV
        //                            TSVValue.Value = TSV;
        //                            TSVValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, TSVValue.Value); TSVValueList.Add(TSVValue);

        //                            //LV
        //                            LVValue.Value = LV;
        //                            LVValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, LVValue.Value); LVValueList.Add(LVValue);


        //                        }


        //                        //Debt Capacity
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Debt Capacity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.ProjectOutputValuesVM = DCValueList.Count > 0 ? DCValueList : null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Interest Paid @ rD
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Interest Paid @ rD";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.ProjectOutputValuesVM = IPValueList.Count > 0 ? IPValueList : null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Interest Tax Shield @ Tc
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.ProjectOutputValuesVM = ITSValueList.Count > 0 ? ITSValueList : null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Tax Shield Value @ rU
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Tax Shield Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.ProjectOutputValuesVM = TSVValueList.Count > 0 ? TSVValueList : null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Levered Value (VL = VU + T)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.ProjectOutputValuesVM = LVValueList.Count > 0 ? LVValueList : null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;



        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                        var lev2Value = LVValueList != null && LVValueList.Count > 0 ? LVValueList.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        double NPV2 = (lev2Value != null ? Convert.ToDouble(lev2Value.Value) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.Value) : 0);
        //                        OutputDatasVM.Value = NPV2;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);
        //                        #endregion

        //                        #region Excel Calculation

        //                        //string rootFolder = hostingEnvironment.WebRootPath;
        //                        //string fileName = @"capital_budgeting.xlsx";
        //                        //FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
        //                        ////
        //                        //using (ExcelPackage package = new ExcelPackage(file))
        //                        //{
        //                        //    ExcelWorksheet wsCapitalStructure = null;                                  
        //                        //     wsCapitalStructure = package.Workbook.Worksheets["CircularReference"];

        //                        //    ExcelCalculationOption calculateOptions = new ExcelCalculationOption();
        //                        //    calculateOptions.AllowCirculareReferences = true;

        //                        //    wsCapitalStructure = InsertInputToExcel(OutputDatasVMList, (MarginalTaxDatas!=null && MarginalTaxDatas.Value!=null? Convert.ToDecimal(MarginalTaxDatas.Value) :0), (DVRatio != null && DVRatio.Value != null ? Convert.ToDecimal(DVRatio.Value) : 0), (UCCInputDatas != null && UCCInputDatas.Value != null ? Convert.ToDecimal(UCCInputDatas.Value) : 0), (CostofDebt != null && CostofDebt.Value != null ? Convert.ToDecimal(CostofDebt.Value) : 0),Convert.ToInt32(projectVM.StartingYear),Convert.ToInt32(projectVM.NoOfYears), wsCapitalStructure);

        //                        //    wsCapitalStructure.Calculate(calculateOptions);                               
        //                        //    package.Save();
        //                        //    int initial = 3;

        //                        //    //Debt Capacity
        //                        //    OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //    OutputDatasVM.Id = 0;
        //                        //    OutputDatasVM.ProjectId = projectVM.Id;
        //                        //    OutputDatasVM.HeaderId = 0;
        //                        //    OutputDatasVM.SubHeader = "";
        //                        //    OutputDatasVM.LineItem = "Debt Capacity";
        //                        //    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        //    OutputDatasVM.HasMultiYear = true;
        //                        //    OutputDatasVM.Value = null;
        //                        //    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        //    first= 0;
        //                        //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        //   // !has(wsCapitalStructure.Cells[11, initial + first])
        //                        //    foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        //    {
        //                        //        dumyValue = new ProjectOutputValuesViewModel();
        //                        //        dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                        //        dumyValue.Year = TempValue.Year;
        //                        //        double value = 0;                                        

        //                        //        value = wsCapitalStructure.Cells[11, initial + first].Value!=null && wsCapitalStructure.Cells[11, initial + first].Value !=""? Convert.ToDouble(wsCapitalStructure.Cells[11, initial + first].Value) :0;
        //                        //        dumyValue.Value = value;                                        
        //                        //        dumyValue.BasicValue = UnitConversion.getBasicValueforCurrency(OutputDatasVM.UnitId, dumyValue.Value);
        //                        //        OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        //        first++;
        //                        //    }
        //                        //    OutputDatasVMList.Add(OutputDatasVM);

        //                        //    //Interest Paid @ rD
        //                        //    OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //    OutputDatasVM.Id = 0;
        //                        //    OutputDatasVM.ProjectId = projectVM.Id;
        //                        //    OutputDatasVM.HeaderId = 0;
        //                        //    OutputDatasVM.SubHeader = "";
        //                        //    OutputDatasVM.LineItem = "Interest Paid @ rD";
        //                        //    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        //    OutputDatasVM.HasMultiYear = true;
        //                        //    OutputDatasVM.Value = null;
        //                        //    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        //    first = 0;
        //                        //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //    foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        //    {
        //                        //        dumyValue = new ProjectOutputValuesViewModel();
        //                        //        dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                        //        dumyValue.Year = TempValue.Year;
        //                        //        double value = 0;

        //                        //       // value = wsCapitalStructure.Cells[12, initial + first].Value != null && wsCapitalStructure.Cells[12, initial + first].Value != "" ? Convert.ToDouble(wsCapitalStructure.Cells[12, initial + first].Value) : 0;
        //                        //        dumyValue.Value = value;
        //                        //        dumyValue.BasicValue = UnitConversion.getBasicValueforCurrency(OutputDatasVM.UnitId, dumyValue.Value);
        //                        //        OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        //        first++;
        //                        //    }
        //                        //    OutputDatasVMList.Add(OutputDatasVM);

        //                        //    //Interest Tax Shield @ Tc
        //                        //    OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //    OutputDatasVM.Id = 0;
        //                        //    OutputDatasVM.ProjectId = projectVM.Id;
        //                        //    OutputDatasVM.HeaderId = 0;
        //                        //    OutputDatasVM.SubHeader = "";
        //                        //    OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
        //                        //    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        //    OutputDatasVM.HasMultiYear = true;
        //                        //    OutputDatasVM.Value = null;
        //                        //    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        //    first = 0;
        //                        //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //    foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        //    {
        //                        //        dumyValue = new ProjectOutputValuesViewModel();
        //                        //        dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                        //        dumyValue.Year = TempValue.Year;
        //                        //        double value = 0;
        //                        //       // value = wsCapitalStructure.Cells[13, initial + first].Value != null && wsCapitalStructure.Cells[13, initial + first].Value!= "" ? Convert.ToDouble(wsCapitalStructure.Cells[13, initial + first].Value) : 0;
        //                        //        dumyValue.Value = value;
        //                        //        dumyValue.BasicValue = UnitConversion.getBasicValueforCurrency(OutputDatasVM.UnitId, dumyValue.Value);
        //                        //        OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        //        first++;

        //                        //    }
        //                        //    OutputDatasVMList.Add(OutputDatasVM);

        //                        //    //Tax Shield Value @ rU
        //                        //    OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //    OutputDatasVM.Id = 0;
        //                        //    OutputDatasVM.ProjectId = projectVM.Id;
        //                        //    OutputDatasVM.HeaderId = 0;
        //                        //    OutputDatasVM.SubHeader = "";
        //                        //    OutputDatasVM.LineItem = "Tax Shield Value @ rU";
        //                        //    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        //    OutputDatasVM.HasMultiYear = true;
        //                        //    OutputDatasVM.Value = null;
        //                        //    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        //    first = 0;
        //                        //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //    foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        //    {
        //                        //        dumyValue = new ProjectOutputValuesViewModel();
        //                        //        dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                        //        dumyValue.Year = TempValue.Year;
        //                        //        double value = 0;
        //                        //        //value = wsCapitalStructure.Cells[14, initial + first].Value != null && wsCapitalStructure.Cells[14, initial + first].Value !="" ? Convert.ToDouble(wsCapitalStructure.Cells[14, initial + first].Value) : 0;
        //                        //        dumyValue.Value = value;
        //                        //        dumyValue.BasicValue = UnitConversion.getBasicValueforCurrency(OutputDatasVM.UnitId, dumyValue.Value);
        //                        //        OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        //        first++;

        //                        //    }
        //                        //    OutputDatasVMList.Add(OutputDatasVM);

        //                        //    //Levered Value (VL = VU + T)
        //                        //    OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //    OutputDatasVM.Id = 0;
        //                        //    OutputDatasVM.ProjectId = projectVM.Id;
        //                        //    OutputDatasVM.HeaderId = 0;
        //                        //    OutputDatasVM.SubHeader = "";
        //                        //    OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        //    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        //    OutputDatasVM.HasMultiYear = true;
        //                        //    OutputDatasVM.Value = null;
        //                        //    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        //    first = 0;
        //                        //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //    foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        //    {
        //                        //        dumyValue = new ProjectOutputValuesViewModel();
        //                        //        dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                        //        dumyValue.Year = TempValue.Year;
        //                        //        double value = 0;
        //                        //       // value = wsCapitalStructure.Cells[16, initial + first].Value != null && wsCapitalStructure.Cells[16, initial + first].Value != ""? Convert.ToDouble(wsCapitalStructure.Cells[16, initial + first].Value) : 0;
        //                        //        dumyValue.Value = value;
        //                        //        dumyValue.BasicValue = UnitConversion.getBasicValueforCurrency(OutputDatasVM.UnitId, dumyValue.Value);
        //                        //        OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        //        first++;

        //                        //    }
        //                        //    OutputDatasVMList.Add(OutputDatasVM);


        //                        //    //Net Present Value
        //                        //    OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //    OutputDatasVM.Id = 0;
        //                        //    OutputDatasVM.ProjectId = projectVM.Id;
        //                        //    OutputDatasVM.HeaderId = 0;
        //                        //    OutputDatasVM.SubHeader = "";
        //                        //    OutputDatasVM.LineItem = "Net Present Value";
        //                        //    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        //    OutputDatasVM.HasMultiYear = false;
        //                        //    OutputDatasVM.Value = 0;

        //                        //    //OutputDatasVM.Value = wsCapitalStructure.Cells[17, initial + first].Value != null ? Convert.ToDouble(wsCapitalStructure.Cells[17, initial].Value) : 0;
        //                        //    //OutputDatasVM.BasicValue = UnitConversion.getBasicValueforCurrency(OutputDatasVM.UnitId, dumyValue.Value);
        //                        //    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        //    OutputDatasVMList.Add(OutputDatasVM);

        //                        //}


        //                        #endregion

        //                        #endregion
        //                    }

        //                    else if (projectVM.ValuationTechniqueId == 3)
        //                    {
        //                        #region Valuation3


        //                        //Levered Value (VL)
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //  ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
        //                        //ProjectOutputValuesViewModel lastRecord = dumyValueList.LastOrDefault();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));


        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            reverseList.Add(dumyValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //calculate Debt Capacity
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Debt Capacity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        ProjectOutputDatasViewModel LeveredValueDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Levered Value (VL)"));
        //                        //var DVRatio = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = LeveredValueDatas != null ? LeveredValueDatas.UnitId : null;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();

        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel LeveredValue = LeveredValueDatas != null && LeveredValueDatas.ProjectOutputValuesVM != null && LeveredValueDatas.ProjectOutputValuesVM.Count > 0 ? LeveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = ((LeveredValue != null ? UnitConversion.getBasicValueforNumbers(LeveredValueDatas.UnitId, LeveredValue.Value) : 0) * (DVRatio != null && DVRatio.Value != null ? Convert.ToDouble(DVRatio.Value) : 0) / 100);

        //                            dumyValue.BasicValue = value;

        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Interest Expense";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        //find debt capacity
        //                        ProjectOutputDatasViewModel DebtCapacityDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Debt Capacity"));
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DebtCapacityDatas != null ? DebtCapacityDatas.UnitId : null;
        //                        first = 0;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel DebtCapacityValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

        //                                value = ((DebtCapacityValue != null ? UnitConversion.getBasicValueforNumbers(DebtCapacityDatas.UnitId, DebtCapacityValue.Value) : 0) * (CostofDebt != null && CostofDebt.Value != null ? Convert.ToDouble(CostofDebt.Value) : 0) / 100);

        //                            }

        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Income
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Income";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        ProjectOutputDatasViewModel InterestExpenseDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Interest Expense"));

        //                        highunit = UnitConversion.getHigherDenominationUnit(InterestExpenseDatas.UnitId, OperatingIncomeDatas.UnitId);
        //                        //Plus: Depreciation
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            ProjectOutputValuesViewModel InterestExpenseValue = InterestExpenseDatas != null && InterestExpenseDatas.ProjectOutputValuesVM != null && InterestExpenseDatas.ProjectOutputValuesVM.Count > 0 ? InterestExpenseDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas != null && OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = (((OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) - (InterestExpenseValue != null ? UnitConversion.getBasicValueforNumbers(InterestExpenseDatas.UnitId, InterestExpenseValue.Value) : 0)) * (1 - (MarginalTaxDatas != null && MarginalTaxDatas.Value != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0) / 100));

        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Plus: Net Borrowing
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Plus: Net Borrowing";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DebtCapacityDatas != null ? DebtCapacityDatas.UnitId : null;
        //                        first = 0;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            //formula = C38+C40-C41-C42
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            ProjectOutputValuesViewModel SameyearValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel PrevYearValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

        //                            value = (SameyearValue != null ? UnitConversion.getBasicValueforNumbers(DebtCapacityDatas.UnitId, SameyearValue.Value) : 0) - (PrevYearValue != null ? UnitConversion.getBasicValueforNumbers(DebtCapacityDatas.UnitId, PrevYearValue.Value) : 0);
                    //Net Present Value
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowtoEquityDatas != null ? FreeCashFlowtoEquityDatas.UnitId : null;
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPv);

        //                        OutputDatasVMList.Add(OutputDatasVM);



        //                        #endregion
        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 5)
        //                    {
        //                        #region Valuation 5
        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //formula=Interest coverage ratio =Interest expense/FCF (C187)* marginal Tax * Unleavered value of first year


        //                        var UnleaveredShortDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Value @ rU") && x.SubHeader == "APV Based Valuation - Short Approach");

        //                        // formula==$C$15(Marginal Tax)*$C$187*C208(Unleavered Value)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "PV of Interest Tax Shield";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleaveredShortDatas.UnitId;

        //                        var SLValue = UnleaveredShortDatas != null && UnleaveredShortDatas.ProjectOutputValuesVM != null && UnleaveredShortDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredShortDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;


        //                        double ULvalue = (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, ULvalue);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Levered Value (VL = VU + T)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;


        //                        double LValue = ((InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0)) + (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, LValue);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Short Approach";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;



        //                        double NPV = ((InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0)) + (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        OutputDatasVM.ProjectOutputValuesVM = UnleaveredShortDatas != null && UnleaveredShortDatas.ProjectOutputValuesVM != null ? UnleaveredShortDatas.ProjectOutputValuesVM : null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Interest Paid @ rD";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;

        //                            double value = 0;
        //                            ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            value = (FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) * (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0);

        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Debt Capacity
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Debt Capacity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        //Interest paid
        //                        var InterestpaidLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Paid @ rD" && x.SubHeader == "APV Based Valuation - Long Approach");

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year + 1) : null;
        //                            value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0);
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Interest Tax Shield @ Tc
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
        //                        first = 0;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                dumyValue.BasicValue = dumyValue.Value = value;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0);
        //                                dumyValue.BasicValue = value;
        //                                dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            }
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Tax Shield Value @ rU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Tax Shield Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;
        //                        var InterestTaxShieldLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Tax Shield @ Tc" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestTaxShieldLongDatas.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of InterestTaxShieldLongDatas L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel InterestTaxShieldValue = InterestTaxShieldLongDatas != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestTaxShieldLongDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((InterestTaxShieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestTaxShieldLongDatas.UnitId, InterestTaxShieldValue.Value) : 0) + lastYearValue) / (1 + (UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) / 100 : 0));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        //formula = Unlevered Value @ rU+ Tax Shield Value @ rU
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;
        //                        var UleaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        var InterestShieldValueLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Tax Shield Value @ rU" && x.SubHeader == "APV Based Valuation - Long Approach");

        //                        highunit = UnitConversion.getHigherDenominationUnit((UleaveredLongDatas != null ? UleaveredLongDatas.UnitId : null), (InterestShieldValueLongDatas != null ? InterestShieldValueLongDatas.UnitId : null));
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel LUnleaveredValue = UleaveredLongDatas != null && UleaveredLongDatas.ProjectOutputValuesVM != null && UleaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? UleaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                ProjectOutputValuesViewModel InterestDhieldValue = InterestShieldValueLongDatas != null && InterestShieldValueLongDatas.ProjectOutputValuesVM != null && InterestShieldValueLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestShieldValueLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (LUnleaveredValue != null ? UnitConversion.getBasicValueforCurrency(UleaveredLongDatas.UnitId, LUnleaveredValue.Value) : 0) + (InterestDhieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestShieldValueLongDatas.UnitId, InterestDhieldValue.Value) : 0);
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "APV Based Valuation - Long Approach";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        var leaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)" && x.SubHeader == "APV Based Valuation - Long Approach");
        //                        highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredLongDatas.UnitId);
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                        var levLongValue = leaveredLongDatas != null && leaveredLongDatas.ProjectOutputValuesVM != null && leaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? leaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

        //                        NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        #endregion
        //                    }

        //                    else if (projectVM.ValuationTechniqueId == 6)
        //                    {
        //                        #region Valuation 6
        //                        //Unlevered Value @ rU ($M) - VU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Debt Capacity
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Debt Capacity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        var InputDebtCapacity = InputDatasList.Find(x => x.LineItem == "Fixed Schedule & Predetermined Debt Level, Dt");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InputDebtCapacity.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            var InputDebtCapacityValue = InputDebtCapacity != null && InputDebtCapacity.ProjectInputValuesVM != null && InputDebtCapacity.ProjectInputValuesVM.Count > 0 ? InputDebtCapacity.ProjectInputValuesVM.Find(x => x.Year == (TempValue.Year)) : null;

        //                            value = InputDebtCapacityValue != null ? UnitConversion.getBasicValueforCurrency(InputDebtCapacity.UnitId, InputDebtCapacityValue.Value) : 0;
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        // Interest Paid
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Interest Paid @ rD";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        //  OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;


        //                        //  OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

        //                        var DebtCapacityDatas = OutputDatasVMList.Find(x => x.LineItem == "Debt Capacity");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DebtCapacityDatas.UnitId;

        //                        var CostOfDebtInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));

        //                        first = 0;
        //                        double FirstYearValue = 0;

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderBy(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            //change formula
        //                            //ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            //value = (FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) * (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0);

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel DebtCapacityValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;
        //                                value = (DebtCapacityValue != null ? UnitConversion.getBasicValueforCurrency(DebtCapacityDatas.UnitId, DebtCapacityValue.Value) : 0) * ((CostOfDebtInputDatas != null && CostOfDebtInputDatas.Value != null ? Convert.ToDouble(CostOfDebtInputDatas.Value) : 0) / 100);
        //                            }

        //                            dumyValue.BasicValue = value;
        //                            FirstYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Interest Tax Shield @ Tc
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();


        //                        var InterestpaidLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Paid @ rD");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
        //                        first = 0;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                dumyValue.BasicValue = dumyValue.Value = value;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0);
        //                                dumyValue.BasicValue = value;
        //                                dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            }
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Tax Shield Value @ rU
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Tax Shield Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;

        //                        //var CostofDebt = InputDatasList.Find(x => x.LineItem == "Cost of Debt");


        //                        var InterestTaxShieldLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Tax Shield @ Tc");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestTaxShieldLongDatas.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L43+L45)/(1+$C$26); // Next year data of InterestTaxShieldLongDatas L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel InterestTaxShieldValue = InterestTaxShieldLongDatas != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestTaxShieldLongDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((InterestTaxShieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestTaxShieldLongDatas.UnitId, InterestTaxShieldValue.Value) : 0) + lastYearValue) / (1 + (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //formula = Unlevered Value @ rU+ Tax Shield Value @ rU
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        first = 0;
        //                        lastYearValue = 0;

        //                        var UleaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU");
        //                        var InterestShieldValueLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Tax Shield Value @ rU");
        //                        highunit = UnitConversion.getHigherDenominationUnit((UleaveredLongDatas != null ? UleaveredLongDatas.UnitId : null), (InterestShieldValueLongDatas != null ? InterestShieldValueLongDatas.UnitId : null));
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                ProjectOutputValuesViewModel LUnleaveredValue = UleaveredLongDatas != null && UleaveredLongDatas.ProjectOutputValuesVM != null && UleaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? UleaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                ProjectOutputValuesViewModel InterestDhieldValue = InterestShieldValueLongDatas != null && InterestShieldValueLongDatas.ProjectOutputValuesVM != null && InterestShieldValueLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestShieldValueLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                                value = (LUnleaveredValue != null ? UnitConversion.getBasicValueforCurrency(UleaveredLongDatas.UnitId, LUnleaveredValue.Value) : 0) + (InterestDhieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestShieldValueLongDatas.UnitId, InterestDhieldValue.Value) : 0);
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        // formula==$C$15(Marginal Tax)*$C$187*C208(Unleavered Value)
        //                        //OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        //OutputDatasVM.Id = 0;
        //                        //OutputDatasVM.ProjectId = projectVM.Id;
        //                        //OutputDatasVM.HeaderId = 0;
        //                        //OutputDatasVM.SubHeader = "";
        //                        //OutputDatasVM.LineItem = "PV of Interest Tax Shield";
        //                        //OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        //OutputDatasVM.HasMultiYear = false;
        //                        //OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleaveredDatasObj.UnitId;
        //                        //double ULvalue = (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (lastYearValue);
        //                        //OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, ULvalue);
        //                        //OutputDatasVM.ProjectOutputValuesVM = null;
        //                        //OutputDatasVMList.Add(OutputDatasVM);

        //                        //Net Present Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;

        //                        var leaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)");

        //                        highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredLongDatas.UnitId);
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                        var levLongValue = leaveredLongDatas != null && leaveredLongDatas.ProjectOutputValuesVM != null && leaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? leaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        double NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
        //                        OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        #endregion

        //                    }
        //                    else if (projectVM.ValuationTechniqueId == 8)
        //                    {
        //                        #region Valuation8

        //                        //Unlevered Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Unlevered Value @ rU";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        int first = 0;
        //                        double lastYearValue = 0;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;

        //                            if (first == 0)
        //                            {
        //                                value = 0;
        //                                first++;
        //                            }
        //                            else
        //                            {
        //                                //formula ==(L346+L349)/(1+$C$328); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
        //                                ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
        //                                value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
        //                            }
        //                            dumyValue.BasicValue = value;
        //                            lastYearValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        // PV of Interest Tax Shield
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "PV of Interest Tax Shield";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        // OutputDatasVM.Value = null;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;

        //                        var ICRInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Intetrest coverage ratio =Interest expense/FCF"));
        //                        var CBInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
        //                        ProjectOutputDatasViewModel UnleveredValueDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU");

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleveredValueDatas != null ? UnleveredValueDatas.UnitId : null;

        //                        // $C$15 $C$327  ((1 + C328) / (1 + C329)) * C349

        //                        ProjectOutputValuesViewModel UnLeveredValue = UnleveredValueDatas != null && UnleveredValueDatas.ProjectOutputValuesVM != null && UnleveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

        //                        OutputDatasVM.Value = ((MarginalTaxDatas != null && MarginalTaxDatas.Value != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0) / 100) * ((ICRInputDatas != null && ICRInputDatas.Value != null ? Convert.ToDouble(ICRInputDatas.Value) : 0) / 100) * ((1 + ((UCCInputDatas != null && UCCInputDatas.Value != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100)) / (1 + ((CBInputDatas != null && CBInputDatas.Value != null ? Convert.ToDouble(CBInputDatas.Value) : 0) / 100))) * (UnLeveredValue != null ? UnLeveredValue.Value : 0);
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Levered Value
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;

        //                        ProjectOutputDatasViewModel PVInterestDatas = OutputDatasVMList.Find(x => x.LineItem == "PV of Interest Tax Shield");

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleveredValueDatas != null ? UnleveredValueDatas.UnitId : null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = PVInterestDatas != null ? PVInterestDatas.UnitId : null;

        //                        // C349 + C350

        //                        // ProjectOutputValuesViewModel PVInterestValue = PVInterestDatas != null && PVInterestDatas.ProjectOutputValuesVM != null && PVInterestDatas.ProjectOutputValuesVM.Count > 0 ? PVInterestDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

        //                        // OutputDatasVM.Value = (UnLeveredValue != null ? UnLeveredValue.Value : 0) + (PVInterestValue != null ? PVInterestValue.Value : 0);
        //                        OutputDatasVM.Value = (UnLeveredValue != null ? UnLeveredValue.Value : 0) + (PVInterestDatas != null ? PVInterestDatas.Value : 0);

        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        //Net Present Value ($M)
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Net Present Value";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        OutputDatasVM.ProjectOutputValuesVM = null;

        //                        ProjectOutputDatasViewModel LeveredValueDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)");
        //                        ProjectOutputDatasViewModel FreeCashFlowDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));

        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = LeveredValueDatas != null ? LeveredValueDatas.UnitId : null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;

        //                        // C351 + C346

        //                        //  ProjectOutputValuesViewModel LeveredValue = LeveredValueDatas != null && LeveredValueDatas.ProjectOutputValuesVM != null && LeveredValueDatas.ProjectOutputValuesVM.Count > 0 ? LeveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
        //                        ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowDatas != null && FreeCashFlowDatas.ProjectOutputValuesVM != null && FreeCashFlowDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

        //                        //  OutputDatasVM.Value = (LeveredValue != null ? LeveredValue.Value : 0) + (FreeCashValue != null ? FreeCashValue.Value : 0);
        //                        OutputDatasVM.Value = (LeveredValueDatas != null ? LeveredValueDatas.Value : 0) + (FreeCashValue != null ? FreeCashValue.Value : 0);

        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        #endregion
        //                    }

        //                    if (projectVM.ValuationTechniqueId != 8)
        //                    {
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "IRR (Internal Rate of Return) of Free Cash Flows";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
        //                        OutputDatasVM.HasMultiYear = false;
        //                        try
        //                        {
        //                            OutputDatasVM.Value = Financial.Irr(FCFarray) * 100;
        //                        }
        //                        catch (Exception ss)
        //                        {

        //                            OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(3, FCFarray.Sum());
        //                        }
        //                        OutputDatasVM.ProjectOutputValuesVM = null;
        //                        OutputDatasVMList.Add(OutputDatasVM);
        //                    }

        //                }



        //            }

        //        }


        //    }

        //    catch (Exception ss)
        //    {

        //        // return OutputDatasVMList;
        //    }
        //    return OutputDatasVMList;
        //}
        //                            dumyValue.BasicValue = value;

        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Free Cash Flow
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Free Cash Flow to Equity";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        FCFarray = new List<double>();
        //                        highunit = UnitConversion.getHigherDenominationUnit(UnleaveredIncomeDatasObj.UnitId, PlusDepreciationDatasObj.UnitId);
        //                        //Plus: Depreciation
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

        //                        ProjectOutputDatasViewModel NetIncomeDatas = OutputDatasVMList.Find(x => x.LineItem == "Net Income");
        //                        ProjectOutputDatasViewModel PlusBorrowingDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Plus: Net Borrowing"));

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            //NetIncomeDatas +PlusDepreciationDatasObj -LessCapexValue
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel NetIncomeValue = NetIncomeDatas != null && NetIncomeDatas.ProjectOutputValuesVM != null && NetIncomeDatas.ProjectOutputValuesVM.Count > 0 ? NetIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel PlusDepreciationValue = PlusDepreciationDatasObj != null && PlusDepreciationDatasObj.ProjectOutputValuesVM != null && PlusDepreciationDatasObj.ProjectOutputValuesVM.Count > 0 ? PlusDepreciationDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel LessCapexValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel LessNWCValue = LessNWCDatasObj != null && LessNWCDatasObj.ProjectOutputValuesVM != null && LessNWCDatasObj.ProjectOutputValuesVM.Count > 0 ? LessNWCDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel PlusBorrowingValue = PlusBorrowingDatas != null && PlusBorrowingDatas.ProjectOutputValuesVM != null && PlusBorrowingDatas.ProjectOutputValuesVM.Count > 0 ? PlusBorrowingDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            value = ((NetIncomeValue != null ? UnitConversion.getBasicValueforNumbers(NetIncomeDatas.UnitId, NetIncomeValue.Value) : 0) + (PlusDepreciationValue != null ? UnitConversion.getBasicValueforNumbers(PlusDepreciationDatasObj.UnitId, PlusDepreciationValue.Value) : 0) - (LessCapexValue != null ? UnitConversion.getBasicValueforNumbers(LessCapexDatasObj.UnitId, LessCapexValue.Value) : 0) - (LessNWCValue != null ? UnitConversion.getBasicValueforNumbers(LessNWCDatasObj.UnitId, LessNWCValue.Value) : 0) + (PlusBorrowingValue != null ? UnitConversion.getBasicValueforNumbers(PlusBorrowingDatas.UnitId, PlusBorrowingValue.Value) : 0));

        //                            dumyValue.BasicValue = value;
        //                            FCFarray.Add(value);
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);


        //                        //Discount Factor
        //                        //formula = 
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discount Factor";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = null;
        //                        var Costofequity = InputDatasList.Find(x => x.LineItem.Contains("Cost of Equity"));

        //                        double TValue = 1 + (Costofequity != null && Costofequity.Value != null ? Convert.ToDouble(Costofequity.Value) / 100 : 0);
        //                        int i = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {
        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            value = Math.Pow(TValue, i);
        //                            i++;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = dumyValue.BasicValue;
        //                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //                        //Discounted Cash Flow
        //                        //formula = Free Cash flow/Discount Factor
        //                        OutputDatasVM = new ProjectOutputDatasViewModel();
        //                        OutputDatasVM.Id = 0;
        //                        OutputDatasVM.ProjectId = projectVM.Id;
        //                        OutputDatasVM.HeaderId = 0;
        //                        OutputDatasVM.SubHeader = "";
        //                        OutputDatasVM.LineItem = "Discounted Cash Flow";
        //                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
        //                        OutputDatasVM.HasMultiYear = true;
        //                        OutputDatasVM.Value = null;
        //                        ProjectOutputDatasViewModel FreeCashFlowtoEquityDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow to Equity"));
        //                        ProjectOutputDatasViewModel DiscountFactorDatas = OutputDatasVMList.Find(x => x.LineItem == "Discount Factor");
        //                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowtoEquityDatas != null ? FreeCashFlowtoEquityDatas.UnitId : null;
        //                        double NPv = 0;
        //                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

        //                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
        //                        {

        //                            dumyValue = new ProjectOutputValuesViewModel();
        //                            // dumyValue = TempValue;
        //                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
        //                            dumyValue.Year = TempValue.Year;
        //                            double value = 0;
        //                            ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowtoEquityDatas != null && FreeCashFlowtoEquityDatas.ProjectOutputValuesVM != null && FreeCashFlowtoEquityDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowtoEquityDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
        //                            ProjectOutputValuesViewModel DiscountFactorValues = DiscountFactorDatas != null && DiscountFactorDatas.ProjectOutputValuesVM != null && DiscountFactorDatas.ProjectOutputValuesVM.Count > 0 ? DiscountFactorDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

        //                            value = (FreeCashValue != null ? UnitConversion.getBasicValueforNumbers(FreeCashFlowtoEquityDatas.UnitId, FreeCashValue.Value) : 0) / (DiscountFactorValues != null ? UnitConversion.getBasicValueforNumbers(DiscountFactorDatas.UnitId, DiscountFactorValues.Value) : 0);


        //                            NPv = NPv + value;
        //                            dumyValue.BasicValue = value;
        //                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
        //                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
        //                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
        //                        }
        //                        OutputDatasVMList.Add(OutputDatasVM);

        //    


        private List<ProjectOutputDatasViewModel> GetProjectOutput(ProjectsViewModel projectVM)
        {

            // List<ProjectOutputDatasViewModel> SummaryList = new List<ProjectOutputDatasViewModel>()
            List<ProjectOutputDatasViewModel> OutputDatasVMList = new List<ProjectOutputDatasViewModel>();
            try
            {
                //Calculate Output by Project Inputs

                if (projectVM != null && projectVM.ProjectInputDatasVM != null && projectVM.ProjectInputDatasVM.Count > 0)
                {

                    //create dumy year Value List
                    List<ProjectOutputValuesViewModel> dumyValueList = new List<ProjectOutputValuesViewModel>();
                    if (projectVM.StartingYear != null && projectVM.StartingYear != 0 && projectVM.NoOfYears != null && projectVM.NoOfYears > 0)
                    {
                        ProjectOutputValuesViewModel dumyValue = new ProjectOutputValuesViewModel();

                        for (int i = 0; i < projectVM.NoOfYears; i++)
                        {
                            dumyValue = new ProjectOutputValuesViewModel();
                            dumyValue.Id = 0;
                            dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = Convert.ToInt32(projectVM.StartingYear) + i;
                            dumyValue.Value = null;
                            dumyValueList.Add(dumyValue);
                        }

                        ProjectOutputDatasViewModel OutputDatasVM = new ProjectOutputDatasViewModel();

                        List<ProjectInputDatasViewModel> InputDatasList = new List<ProjectInputDatasViewModel>();
                        InputDatasList = projectVM.ProjectInputDatasVM;

                        //create data //for Sales 
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Sales";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;
                        // ($O$6*C6)*($O$7*C7)/1000000
                        List<ProjectInputDatasViewModel> VolumeList = new List<ProjectInputDatasViewModel>();
                        VolumeList = InputDatasList.FindAll(x => x.LineItem == "Volume" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

                        List<ProjectInputDatasViewModel> UnitPriceList = new List<ProjectInputDatasViewModel>();
                        UnitPriceList = InputDatasList.FindAll(x => x.LineItem == "Unit Price" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();
                        int? highunit = UnitConversion.getHigherDenominationUnit(VolumeList.FirstOrDefault().UnitId, UnitPriceList.FirstOrDefault().UnitId);

                        //sales=volume*unitcost;
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {
                            dumyValue = new ProjectOutputValuesViewModel();

                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;

                            double value = 0;
                            //calculate sales for multiple years

                            if (VolumeList != null && VolumeList.Count > 0 && UnitPriceList != null && UnitPriceList.Count > 0)
                                foreach (var item in VolumeList)
                                {
                                    var UnitPriceDatas = UnitPriceList.Find(x => x.SubHeader == item.SubHeader);
                                    ProjectInputValuesViewModel UnitCostValue = UnitPriceDatas != null && UnitPriceDatas.ProjectInputValuesVM != null && UnitPriceDatas.ProjectInputValuesVM.Count > 0 ? UnitPriceDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectInputValuesViewModel VolumeValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = value + ((UnitPriceDatas != null ? UnitConversion.getBasicValueforCurrency(UnitPriceDatas.UnitId, UnitCostValue.Value) : 0) * (item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, VolumeValue.Value) : 0));
                                }




                            //get sum of all the Vlaues

                            // value = UnitConversion.getHigherDenominationUnit(UnitCostDatas.UnitId,);

                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign sum of all sales value
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);


                        //create data //for COGS 
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "COGS";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;

                        List<ProjectInputDatasViewModel> UnitCostList = new List<ProjectInputDatasViewModel>();
                        UnitCostList = InputDatasList.FindAll(x => x.LineItem == "Unit Cost" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

                        highunit = UnitConversion.getHigherDenominationUnit(VolumeList.FirstOrDefault().UnitId, UnitCostList.FirstOrDefault().UnitId);
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {
                            //calculate COGS for multiple years

                            //get sum of all the Vlaues
                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            //calculate sales for multiple years
                            if (VolumeList != null && VolumeList.Count > 0 && UnitCostList != null && UnitCostList.Count > 0)
                                foreach (var item in VolumeList)
                                {
                                    var UnitCostDatas = UnitCostList.Find(x => x.SubHeader == item.SubHeader);
                                    ProjectInputValuesViewModel UnitCostValue = UnitCostDatas != null && UnitCostDatas.ProjectInputValuesVM != null && UnitCostDatas.ProjectInputValuesVM.Count > 0 ? UnitCostDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectInputValuesViewModel VolumeValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = value + ((UnitCostDatas != null ? UnitConversion.getBasicValueforCurrency(UnitCostDatas.UnitId, UnitCostValue.Value) : 0) * (item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, VolumeValue.Value) : 0));
                                }


                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                            //value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign sum of all sales value
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        //Gross Margin
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Gross Margin";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;
                        var SalesDatas = OutputDatasVMList.Find(x => x.LineItem == "Sales");
                        var COGSDatas = OutputDatasVMList.Find(x => x.LineItem == "COGS");
                        highunit = UnitConversion.getHigherDenominationUnit(SalesDatas.UnitId, COGSDatas.UnitId);
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {

                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            ProjectOutputValuesViewModel salesValue = SalesDatas.ProjectOutputValuesVM != null && SalesDatas.ProjectOutputValuesVM.Count > 0 ? SalesDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            ProjectOutputValuesViewModel COGSValue = COGSDatas.ProjectOutputValuesVM != null && COGSDatas.ProjectOutputValuesVM.Count > 0 ? COGSDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                            value = (salesValue != null ? UnitConversion.getBasicValueforNumbers(SalesDatas.UnitId, salesValue.Value) : 0) - (COGSValue != null ? UnitConversion.getBasicValueforNumbers(COGSDatas.UnitId, COGSValue.Value) : 0);
                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);


                            // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        //all items of fixed cost with the same formula
                        //find all items of Other Fixed Cost
                        List<ProjectInputDatasViewModel> OtherFixedCostList = new List<ProjectInputDatasViewModel>();
                        OtherFixedCostList = InputDatasList.FindAll(x => x.SubHeader.Contains("Other Fixed Cost")).ToList();

                        if (OtherFixedCostList != null && OtherFixedCostList.Count > 0)
                            foreach (ProjectInputDatasViewModel fixedCostObj in OtherFixedCostList)
                            {
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = fixedCostObj.LineItem;
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = fixedCostObj.UnitId;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {

                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    ProjectInputValuesViewModel Value = fixedCostObj.ProjectInputValuesVM != null && fixedCostObj.ProjectInputValuesVM.Count > 0 ? fixedCostObj.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = value + ((fixedCostObj != null ? UnitConversion.getBasicValueforNumbers(fixedCostObj.UnitId, Value.Value) : 0));
                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));//assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                            }

                        //formula=Startyear+1 value * year Value

                        //Depreciation // Bind Depreciation Values directly from Input
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Depreciation";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;

                        // find depreciation
                        List<ProjectInputDatasViewModel> DepreciationList = new List<ProjectInputDatasViewModel>();
                        DepreciationList = InputDatasList.FindAll(x => x.LineItem == "Depreciation" && x.SubHeader.Contains("Capex & Depreciation")).ToList();
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationList != null && DepreciationList.Count > 0 ? DepreciationList.FirstOrDefault().UnitId : null;
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {
                            //Sales-COGS
                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            if (DepreciationList != null && DepreciationList.Count > 0)
                                foreach (var item in DepreciationList)
                                {
                                    ProjectInputValuesViewModel depreciationValue = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = value + ((item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, depreciationValue.Value) : 0));
                                }
                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));//assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);



                        //Operating Income 
                        //formula = Gross Margin - (sum of all items after Gross Margin)
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Operating Income";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;

                        var GrossMarginDatas = OutputDatasVMList.Find(x => x.LineItem == "Gross Margin");
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = GrossMarginDatas.UnitId;

                        var DepreciationDatas = OutputDatasVMList.Find(x => x.LineItem == "Depreciation");
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationDatas.UnitId;

                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {
                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            double sum = 0;
                            ProjectOutputValuesViewModel GrossValue = GrossMarginDatas.ProjectOutputValuesVM != null && GrossMarginDatas.ProjectOutputValuesVM.Count > 0 ? GrossMarginDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                            ProjectOutputValuesViewModel DepreciationValue = DepreciationDatas.ProjectOutputValuesVM != null && DepreciationDatas.ProjectOutputValuesVM.Count > 0 ? DepreciationDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                            foreach (ProjectInputDatasViewModel fixedCostObj in OtherFixedCostList)
                            {
                                var otherFixedValue = fixedCostObj.ProjectInputValuesVM != null && fixedCostObj.ProjectInputValuesVM.Count > 0 ? fixedCostObj.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                sum = sum + (otherFixedValue != null ? UnitConversion.getBasicValueforNumbers(fixedCostObj.UnitId, otherFixedValue.Value) : 0);
                            }

                            // value = (GrossMarginDatas != null ? UnitConversion.getBasicValueforNumbers(GrossMarginDatas.UnitId, GrossValue.Value) : 0) - sum;//-()sum of all items;

                            value = (GrossMarginDatas != null ? UnitConversion.getBasicValueforNumbers(GrossMarginDatas.UnitId, GrossValue.Value) : 0) - sum - (DepreciationDatas != null ? UnitConversion.getBasicValueforNumbers(DepreciationDatas.UnitId, DepreciationValue.Value) : 0);//-()sum of all items;

                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                            //  value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
                            //  dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);


                        //Income Tax
                        //formula = Operating Income (for diff year) * Marginal Tax
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Income Tax";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;

                        var OperatingIncomeDatas = OutputDatasVMList.Find(x => x.LineItem == "Operating Income");

                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));

                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = OperatingIncomeDatas.UnitId;
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {
                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                            value = ((OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0)) / 100;

                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                            // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);

                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        //Unlevered Net Income
                        //formula = Operating Income (for diff year) -Income Tax
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Unlevered Net Income";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;

                        var IncomeTaxDatas = OutputDatasVMList.Find(x => x.LineItem == "Income Tax");
                        highunit = UnitConversion.getHigherDenominationUnit(OperatingIncomeDatas.UnitId, IncomeTaxDatas.UnitId);
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {
                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            ProjectOutputValuesViewModel IncomeTaxValue = IncomeTaxDatas.ProjectOutputValuesVM != null && IncomeTaxDatas.ProjectOutputValuesVM.Count > 0 ? IncomeTaxDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            value = (OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) - (IncomeTaxValue != null ? UnitConversion.getBasicValueforNumbers(IncomeTaxDatas.UnitId, IncomeTaxValue.Value) : 0);

                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                            // value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, value);
                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        //NWC (Net Working Capital)
                        //formula = 
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "NWC (Net Working Capital)";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = SalesDatas.UnitId;
                        var NWCDatas = InputDatasList.Find(x => x.LineItem.Contains("NWC"));
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {

                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;

                            ProjectOutputValuesViewModel salesValue = SalesDatas.ProjectOutputValuesVM != null && SalesDatas.ProjectOutputValuesVM.Count > 0 ? SalesDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            ProjectInputValuesViewModel NWCValue = NWCDatas.ProjectInputValuesVM != null && NWCDatas.ProjectInputValuesVM.Count > 0 ? NWCDatas.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                            value = ((salesValue != null ? UnitConversion.getBasicValueforNumbers(SalesDatas.UnitId, salesValue.Value) : 0) * (NWCValue != null ? Convert.ToDouble(NWCValue.Value) : 0)) / 100;

                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        //Plus: Depreciation
                        //formula = 
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Plus: Depreciation";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;
                        var DepreciationOutputDatas = OutputDatasVMList.Find(x => x.LineItem == "Depreciation");

                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DepreciationOutputDatas != null ? DepreciationOutputDatas.UnitId : null;
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                        OutputDatasVM.ProjectOutputValuesVM = DepreciationOutputDatas != null ? DepreciationOutputDatas.ProjectOutputValuesVM : null;

                        //foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        //{
                        //    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                        //    //Sales-COGS
                        //    dumyValue = new ProjectOutputValuesViewModel();
                        //    dumyValue = TempValue;
                        //    double value = 0;


                        //    dumyValue.BasicValue = value;
                        //    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                        //   // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                        //    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        //}
                        OutputDatasVMList.Add(OutputDatasVM);

                        //Less: Capital Expenditures
                        //formula = 
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Less: Capital Expenditures";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;
                        // find Capex
                        List<ProjectInputDatasViewModel> CapexList = new List<ProjectInputDatasViewModel>();
                        CapexList = InputDatasList.FindAll(x => x.LineItem == "Capex" && x.SubHeader.Contains("Capex & Depreciation")).ToList();
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = CapexList != null && CapexList.Count > 0 ? CapexList.FirstOrDefault().UnitId : null;
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {

                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            if (CapexList != null && CapexList.Count > 0)
                                foreach (var item in CapexList)
                                {
                                    ProjectInputValuesViewModel Value = item != null && item.ProjectInputValuesVM != null && item.ProjectInputValuesVM.Count > 0 ? item.ProjectInputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = value + ((item != null ? UnitConversion.getBasicValueforNumbers(item.UnitId, Value.Value) : 0));
                                }
                            dumyValue.BasicValue = value;
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                            //dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        //Less: Increases in NWC
                        //formula = 
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Less: Increases in NWC";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;
                        ProjectOutputDatasViewModel NWCObj = OutputDatasVMList.Find(x => x.LineItem.Contains("NWC"));
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = NWCObj.UnitId;
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {

                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            ProjectOutputValuesViewModel sameValue = NWCObj != null && NWCObj.ProjectOutputValuesVM != null && NWCObj.ProjectOutputValuesVM.Count > 0 ? NWCObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            ProjectOutputValuesViewModel PrevValue = NWCObj != null && NWCObj.ProjectOutputValuesVM != null && NWCObj.ProjectOutputValuesVM.Count > 0 ? NWCObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

                            value = ((sameValue != null ? UnitConversion.getBasicValueforNumbers(NWCObj.UnitId, sameValue.Value) : 0) - (PrevValue != null ? UnitConversion.getBasicValueforNumbers(NWCObj.UnitId, PrevValue.Value) : 0));
                            dumyValue.BasicValue = value;

                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        List<double> FCFarray = new List<double>();

                        //Free Cash Flow
                        OutputDatasVM = new ProjectOutputDatasViewModel();
                        OutputDatasVM.Id = 0;
                        OutputDatasVM.ProjectId = projectVM.Id;
                        OutputDatasVM.HeaderId = 0;
                        OutputDatasVM.SubHeader = "";
                        OutputDatasVM.LineItem = "Free Cash Flow";
                        OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                        OutputDatasVM.HasMultiYear = true;
                        OutputDatasVM.Value = null;
                        ProjectOutputDatasViewModel UnleaveredIncomeDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Net Income"));
                        ProjectOutputDatasViewModel PlusDepreciationDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Plus: Depreciation"));
                        ProjectOutputDatasViewModel LessCapexDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Less: Capital Expenditures"));
                        ProjectOutputDatasViewModel LessNWCDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Less: Increases in NWC"));
                        highunit = UnitConversion.getHigherDenominationUnit(UnleaveredIncomeDatasObj.UnitId, PlusDepreciationDatasObj.UnitId);
                        //Plus: Depreciation
                        OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                        OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
                        foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                        {
                            //formula = C38+C40-C41-C42
                            dumyValue = new ProjectOutputValuesViewModel();
                            // dumyValue = TempValue;
                            dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                            dumyValue.Year = TempValue.Year;
                            double value = 0;
                            ProjectOutputValuesViewModel UnleaveredIncomeValue = UnleaveredIncomeDatasObj != null && UnleaveredIncomeDatasObj.ProjectOutputValuesVM != null && UnleaveredIncomeDatasObj.ProjectOutputValuesVM.Count > 0 ? UnleaveredIncomeDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            ProjectOutputValuesViewModel PlusDepreciationValue = PlusDepreciationDatasObj != null && PlusDepreciationDatasObj.ProjectOutputValuesVM != null && PlusDepreciationDatasObj.ProjectOutputValuesVM.Count > 0 ? PlusDepreciationDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            ProjectOutputValuesViewModel LessCapexValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            ProjectOutputValuesViewModel LessNWCValue = LessNWCDatasObj != null && LessNWCDatasObj.ProjectOutputValuesVM != null && LessNWCDatasObj.ProjectOutputValuesVM.Count > 0 ? LessNWCDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                            value = ((UnleaveredIncomeValue != null ? UnitConversion.getBasicValueforNumbers(UnleaveredIncomeDatasObj.UnitId, UnleaveredIncomeValue.Value) : 0) + (PlusDepreciationValue != null ? UnitConversion.getBasicValueforNumbers(PlusDepreciationDatasObj.UnitId, PlusDepreciationValue.Value) : 0) - (LessCapexValue != null ? UnitConversion.getBasicValueforNumbers(LessCapexDatasObj.UnitId, LessCapexValue.Value) : 0) - (LessNWCValue != null ? UnitConversion.getBasicValueforNumbers(LessNWCDatasObj.UnitId, LessNWCValue.Value) : 0));

                            dumyValue.BasicValue = value;
                            FCFarray.Add(value);
                            dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                            // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                            OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                        }
                        OutputDatasVMList.Add(OutputDatasVM);

                        if (projectVM.ValuationTechniqueId != null)

                        {
                            ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
                            var UCCInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
                            List<ProjectOutputValuesViewModel> reverseList = new List<ProjectOutputValuesViewModel>();
                            var WACC = InputDatasList.Find(x => x.LineItem.Contains("Weighted average cost of capital") || x.LineItem.Contains("WACC"));
                            var InterestCoverageDatas = InputDatasList.Find(x => x.LineItem.Contains("Interest coverage ratio =Interest expense/FCF") || x.LineItem.Contains("Intetrest coverage ratio"));
                            var TFCFValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
                            var DVRatio = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));
                            var CostofDebt = InputDatasList.Find(x => x.LineItem == "Cost of Debt");
                            if (projectVM.ValuationTechniqueId == 1 || projectVM.ValuationTechniqueId == 4)
                            {
                                #region Valuation1

                                //Levered Value (VL)
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Levered Value (VL)";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();


                                //  ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
                                //ProjectOutputValuesViewModel lastRecord = dumyValueList.LastOrDefault();
                                int first = 0;
                                double lastYearValue = 0;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
                                        ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));


                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }

                                OutputDatasVMList.Add(OutputDatasVM);

                                //Discount Factor
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Discount Factor";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = null;
                                double TValue = 1 + (WACC != null && WACC.Value != null ? Convert.ToDouble(WACC.Value) / 100 : 0);
                                int i = 0;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    value = Math.Pow(TValue, i);
                                    i++;
                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = dumyValue.BasicValue;
                                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Discounted Cash Flow
                                //formula = Free Cash flow/Discount Factor
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Discounted Cash Flow";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                ProjectOutputDatasViewModel FreeCashFlowDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
                                ProjectOutputDatasViewModel DiscountFactorDatas = OutputDatasVMList.Find(x => x.LineItem == "Discount Factor");
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;
                                double NPv = 0;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {

                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowDatas != null && FreeCashFlowDatas.ProjectOutputValuesVM != null && FreeCashFlowDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel DiscountFactorValues = DiscountFactorDatas != null && DiscountFactorDatas.ProjectOutputValuesVM != null && DiscountFactorDatas.ProjectOutputValuesVM.Count > 0 ? DiscountFactorDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = (FreeCashValue != null ? UnitConversion.getBasicValueforNumbers(FreeCashFlowDatas.UnitId, FreeCashValue.Value) : 0) / (DiscountFactorValues != null ? UnitConversion.getBasicValueforNumbers(DiscountFactorDatas.UnitId, DiscountFactorValues.Value) : 0);


                                    NPv = NPv + value;
                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Net Present Value
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Net Present Value";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;

                                OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPv);

                                OutputDatasVMList.Add(OutputDatasVM);



                                //IRR (Internal Rate of Return) of Free Cash Flows
                                //formula = 



                                #endregion
                            }
                            else if (projectVM.ValuationTechniqueId == 2 || projectVM.ValuationTechniqueId == 7)
                            {
                                int first = 0;
                                double lastYearValue = 0;
                                if (projectVM.ValuationTechniqueId == 7)
                                {
                                    #region Valuation 7

                                    //formula = Free Cash flow+ Next Year Value of Levered Value/(1+WACC)
                                    OutputDatasVM = new ProjectOutputDatasViewModel();
                                    OutputDatasVM.Id = 0;
                                    OutputDatasVM.ProjectId = projectVM.Id;
                                    OutputDatasVM.HeaderId = 0;
                                    OutputDatasVM.SubHeader = "WACC Based Valuation";
                                    OutputDatasVM.LineItem = "Levered Value @ rWACC";
                                    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                    OutputDatasVM.HasMultiYear = true;
                                    OutputDatasVM.Value = null;
                                    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                    foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                    {
                                        dumyValue = new ProjectOutputValuesViewModel();
                                        dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                        dumyValue.Year = TempValue.Year;
                                        double value = 0;
                                        if (first == 0)
                                        {
                                            value = 0;
                                            first++;
                                        }
                                        else
                                        {
                                            ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                            value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));
                                        }
                                        dumyValue.BasicValue = value;
                                        lastYearValue = value;
                                        dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                        OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                    }
                                    OutputDatasVMList.Add(OutputDatasVM);

                                    //Net Present Value
                                    OutputDatasVM = new ProjectOutputDatasViewModel();
                                    OutputDatasVM.Id = 0;
                                    OutputDatasVM.ProjectId = projectVM.Id;
                                    OutputDatasVM.HeaderId = 0;
                                    OutputDatasVM.SubHeader = "WACC Based Valuation";
                                    OutputDatasVM.LineItem = "Net Present Value";
                                    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
                                    OutputDatasVM.HasMultiYear = false;

                                    var leaveredDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value @ rWACC" && x.SubHeader == "WACC Based Valuation");

                                    highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredDatas.UnitId);
                                    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

                                    var levLongValue = leaveredDatas != null && leaveredDatas.ProjectOutputValuesVM != null && leaveredDatas.ProjectOutputValuesVM.Count > 0 ? leaveredDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
                                    double NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
                                    OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
                                    OutputDatasVM.ProjectOutputValuesVM = null;
                                    OutputDatasVMList.Add(OutputDatasVM);


                                    #endregion
                                }
                                #region Valuation2
                                //Unlevered Value @ rU ($M) - VU
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                if (projectVM.ValuationTechniqueId == 7)
                                    OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "Unlevered Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                first = 0;
                                lastYearValue = 0;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
                                        ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));


                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }

                                OutputDatasVMList.Add(OutputDatasVM);

                                #region New approach



                                List<ProjectOutputValuesViewModel> DebtCapacityValueList = new List<ProjectOutputValuesViewModel>();
                                first = 0;

                                double MarginalTaxValue = MarginalTaxDatas != null && MarginalTaxDatas.Value != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0;
                                double DVRatioValue = DVRatio != null && DVRatio.Value != null ? Convert.ToDouble
                                    (DVRatio.Value) / 100 : 0;
                                double UCCValue = UCCInputDatas != null && UCCInputDatas.Value != null ? Convert.ToDouble(UCCInputDatas.Value) / 100 : 0;
                                double CostofDebtValue = CostofDebt != null && CostofDebt.Value != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0;
                                var UnleaveredValueDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Value @ rU"));
                                double unLeaveredValue = 0;

                                //DC List
                                List<ProjectOutputValuesViewModel> DCValueList = new List<ProjectOutputValuesViewModel>();

                                // Interest paid IP value List
                                List<ProjectOutputValuesViewModel> IPValueList = new List<ProjectOutputValuesViewModel>();

                                // Interest Tax Shield ITS Value List
                                List<ProjectOutputValuesViewModel> ITSValueList = new List<ProjectOutputValuesViewModel>();

                                // Tax Shield Value TSV Value List
                                List<ProjectOutputValuesViewModel> TSVValueList = new List<ProjectOutputValuesViewModel>();

                                //Leveared Value LV List
                                List<ProjectOutputValuesViewModel> LVValueList = new List<ProjectOutputValuesViewModel>();

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    // DC
                                    var DCValue = new ProjectOutputValuesViewModel();
                                    DCValue.Id = 0; DCValue.ProjectOutputDatasId = 0; DCValue.Year = TempValue.Year;
                                    // IP
                                    var IPValue = new ProjectOutputValuesViewModel();
                                    IPValue.Id = 0; IPValue.ProjectOutputDatasId = 0; IPValue.Year = TempValue.Year;
                                    // ITS
                                    var ITSValue = new ProjectOutputValuesViewModel();
                                    ITSValue.Id = 0; ITSValue.ProjectOutputDatasId = 0; ITSValue.Year = TempValue.Year;
                                    // TSV
                                    var TSVValue = new ProjectOutputValuesViewModel();
                                    TSVValue.Id = 0; TSVValue.ProjectOutputDatasId = 0; TSVValue.Year = TempValue.Year;
                                    // LV
                                    var LVValue = new ProjectOutputValuesViewModel();
                                    LVValue.Id = 0; LVValue.ProjectOutputDatasId = 0; LVValue.Year = TempValue.Year;


                                    double DC = 0;
                                    double IP = 0;
                                    double ITS = 0;
                                    double TSV = 0;
                                    double LV = 0;

                                    if (first == 0)
                                    {
                                        //Find Last Year Value of UnLevered Value
                                        var UnleaveredValueGVM = UnleaveredValueDatas != null && UnleaveredValueDatas.ProjectOutputValuesVM != null && UnleaveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;
                                        unLeaveredValue = UnleaveredValueGVM != null ? Convert.ToDouble(UnleaveredValueGVM.Value) : 0;

                                        //for Last Year

                                        //Interest Paid
                                        IP = (DVRatioValue * CostofDebtValue * unLeaveredValue) / (1 - ((DVRatioValue * CostofDebtValue * MarginalTaxValue) / (1 + UCCValue)));

                                        //ITS= Interest Paid * Marginal Tax
                                        ITS = IP * MarginalTaxValue;
                                        first++;
                                    }
                                    else
                                    {

                                        //Find Same Year Value of UnLevered Value
                                        var UnleaveredValueGVM = UnleaveredValueDatas != null && UnleaveredValueDatas.ProjectOutputValuesVM != null && UnleaveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year)) : null;
                                        unLeaveredValue = UnleaveredValueGVM != null ? Convert.ToDouble(UnleaveredValueGVM.Value) : 0;

                                        //find Next year ITS
                                        var NextYearITS = ITSValueList != null && ITSValueList.Count > 0 ? ITSValueList.Find(x => x.Year == (TempValue.Year + 1)) : null;

                                        //find Next year TSV
                                        var NextYearTSV = TSVValueList != null && TSVValueList.Count > 0 ? TSVValueList.Find(x => x.Year == (TempValue.Year + 1)) : null;

                                        //Calculate TSV=(Next year TSV+Next year ITS)/(1+UCCValue);
                                        TSV = (((NextYearITS != null && NextYearITS.Value != null ? Convert.ToDouble(NextYearITS.Value) : 0) + (NextYearTSV != null && NextYearTSV.Value != null ? Convert.ToDouble(NextYearTSV.Value) : 0)) / (1 + UCCValue));

                                        //LV =Same year UV+ same year TSV
                                        LV = unLeaveredValue + TSV;

                                        //DC=DVRatio * LV
                                        DC = DVRatioValue * LV;

                                        //Find Last Year Value of UnLevered Value
                                        var LastyearLeaveredValue = UnleaveredValueDatas != null && UnleaveredValueDatas.ProjectOutputValuesVM != null && UnleaveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

                                        if (TempValue.Year != projectVM.StartingYear)
                                        {

                                            //IP= DVRatio * Cost of Debt*(Last year UV+((Same year ITS+same year TSV)/(1+ Unleavered Cost of capital)))
                                            // IP = DVRatioValue * CostofDebtValue * ((LastyearLeaveredValue != null ? Convert.ToDouble(LastyearLeaveredValue.Value) : 0) + ((TSV + ITS) / (1 + UCCValue)));

                                            IP = ((LastyearLeaveredValue != null ? Convert.ToDouble(LastyearLeaveredValue.Value) : 0) * (1 + UCCValue) + TSV) / (((1 + UCCValue) / (DVRatioValue * CostofDebtValue)) - MarginalTaxValue);

                                            //ITS= Interest Paid * Marginal Tax
                                            ITS = IP * MarginalTaxValue;

                                        }


                                    }

                                    //DC
                                    DCValue.Value = DC;
                                    DCValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, DCValue.Value); DCValueList.Add(DCValue);

                                    //IP
                                    IPValue.Value = IP;
                                    IPValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, IPValue.Value); IPValueList.Add(IPValue);

                                    //ITS
                                    ITSValue.Value = ITS;
                                    ITSValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, ITSValue.Value); ITSValueList.Add(ITSValue);

                                    //TSV
                                    TSVValue.Value = TSV;
                                    TSVValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, TSVValue.Value); TSVValueList.Add(TSVValue);

                                    //LV
                                    LVValue.Value = LV;
                                    LVValue.BasicValue = UnitConversion.getConvertedValueforCurrency(UnleaveredValueDatas.UnitId, LVValue.Value); LVValueList.Add(LVValue);


                                }


                                //Debt Capacity
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                if (projectVM.ValuationTechniqueId == 2)
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";
                                else
                                {
                                    OutputDatasVM.Header = "APV Based Valuation";
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";

                                }

                                OutputDatasVM.LineItem = "Debt Capacity";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.ProjectOutputValuesVM = DCValueList.Count > 0 ? DCValueList : null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Interest Paid @ rD
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                if (projectVM.ValuationTechniqueId == 2)
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";
                                else
                                {
                                    OutputDatasVM.Header = "APV Based Valuation";
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";

                                }
                                //OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "Interest Paid @ rD";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.ProjectOutputValuesVM = IPValueList.Count > 0 ? IPValueList : null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Interest Tax Shield @ Tc
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                if (projectVM.ValuationTechniqueId == 2)
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";
                                else
                                {
                                    OutputDatasVM.Header = "APV Based Valuation";
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";

                                }
                                //OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.ProjectOutputValuesVM = ITSValueList.Count > 0 ? ITSValueList : null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Tax Shield Value @ rU
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                if (projectVM.ValuationTechniqueId == 2)
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";
                                else
                                {
                                    OutputDatasVM.Header = "APV Based Valuation";
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";

                                }
                                OutputDatasVM.LineItem = "Tax Shield Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.ProjectOutputValuesVM = TSVValueList.Count > 0 ? TSVValueList : null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                if (projectVM.ValuationTechniqueId == 7)
                                {
                                    //Tax Shield Value Adjusted for Predetermined Debt ($M) - T
                                    OutputDatasVM = new ProjectOutputDatasViewModel();
                                    OutputDatasVM.Id = 0;
                                    OutputDatasVM.ProjectId = projectVM.Id;
                                    OutputDatasVM.HeaderId = 0;
                                    OutputDatasVM.Header = "APV Based Valuation";
                                    OutputDatasVM.SubHeader = "Interest Tax Shield";
                                    OutputDatasVM.LineItem = "Tax Shield Value Adjusted for Predetermined Debt";
                                    OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                    OutputDatasVM.HasMultiYear = true;
                                    OutputDatasVM.Value = null;
                                    OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                    OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                    foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                    {
                                        dumyValue = new ProjectOutputValuesViewModel();
                                        dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                        dumyValue.Year = TempValue.Year;
                                        double value = 0;
                                        var TaxShieldValue = TSVValueList != null && TSVValueList.Count > 0 ? TSVValueList.Find(x => x.Year == TempValue.Year) : null;

                                        value = ((TaxShieldValue != null ? Convert.ToDouble(TaxShieldValue.BasicValue) : 0) * (1 + UCCValue)) / (1 + CostofDebtValue);

                                        dumyValue.BasicValue = value;

                                        dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                        OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                    }
                                    OutputDatasVMList.Add(OutputDatasVM);

                                }

                                //Levered Value (VL = VU + T)
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                if (projectVM.ValuationTechniqueId == 2)
                                    OutputDatasVM.SubHeader = "Adjacent Present Value";
                                OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.ProjectOutputValuesVM = LVValueList.Count > 0 ? LVValueList : null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Net Present Value
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                if (projectVM.ValuationTechniqueId == 2)
                                    OutputDatasVM.SubHeader = "Adjacent Present Value";
                                //OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "Net Present Value";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;



                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

                                var lev2Value = LVValueList != null && LVValueList.Count > 0 ? LVValueList.Find(x => x.Year == projectVM.StartingYear) : null;
                                double NPV2 = (lev2Value != null ? Convert.ToDouble(lev2Value.Value) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.Value) : 0);
                                OutputDatasVM.Value = NPV2;
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVMList.Add(OutputDatasVM);
                                #endregion



                                #endregion
                            }

                            else if (projectVM.ValuationTechniqueId == 3)
                            {
                                #region Valuation3


                                //Levered Value (VL)
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Levered Value (VL)";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                //  ProjectOutputDatasViewModel FreeCashFlowDatasObj = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));
                                //ProjectOutputValuesViewModel lastRecord = dumyValueList.LastOrDefault();
                                int first = 0;
                                double lastYearValue = 0;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((WACC != null ? Convert.ToDouble(WACC.Value) : 0) / 100));


                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    reverseList.Add(dumyValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //calculate Debt Capacity
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Debt Capacity";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                ProjectOutputDatasViewModel LeveredValueDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Levered Value (VL)"));
                                //var DVRatio = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = LeveredValueDatas != null ? LeveredValueDatas.UnitId : null;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {

                                    dumyValue = new ProjectOutputValuesViewModel();

                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    ProjectOutputValuesViewModel LeveredValue = LeveredValueDatas != null && LeveredValueDatas.ProjectOutputValuesVM != null && LeveredValueDatas.ProjectOutputValuesVM.Count > 0 ? LeveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = ((LeveredValue != null ? UnitConversion.getBasicValueforNumbers(LeveredValueDatas.UnitId, LeveredValue.Value) : 0) * (DVRatio != null && DVRatio.Value != null ? Convert.ToDouble(DVRatio.Value) : 0) / 100);

                                    dumyValue.BasicValue = value;

                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Interest Expense";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                //find debt capacity
                                ProjectOutputDatasViewModel DebtCapacityDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Debt Capacity"));
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DebtCapacityDatas != null ? DebtCapacityDatas.UnitId : null;
                                first = 0;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    if (first == 0)
                                    {
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel DebtCapacityValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

                                        value = ((DebtCapacityValue != null ? UnitConversion.getBasicValueforNumbers(DebtCapacityDatas.UnitId, DebtCapacityValue.Value) : 0) * (CostofDebt != null && CostofDebt.Value != null ? Convert.ToDouble(CostofDebt.Value) : 0) / 100);

                                    }

                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Net Income
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Net Income";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                ProjectOutputDatasViewModel InterestExpenseDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Interest Expense"));

                                highunit = UnitConversion.getHigherDenominationUnit(InterestExpenseDatas.UnitId, OperatingIncomeDatas.UnitId);
                                //Plus: Depreciation
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {

                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    ProjectOutputValuesViewModel InterestExpenseValue = InterestExpenseDatas != null && InterestExpenseDatas.ProjectOutputValuesVM != null && InterestExpenseDatas.ProjectOutputValuesVM.Count > 0 ? InterestExpenseDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel OperatingIncomeValue = OperatingIncomeDatas != null && OperatingIncomeDatas.ProjectOutputValuesVM != null && OperatingIncomeDatas.ProjectOutputValuesVM.Count > 0 ? OperatingIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = (((OperatingIncomeValue != null ? UnitConversion.getBasicValueforNumbers(OperatingIncomeDatas.UnitId, OperatingIncomeValue.Value) : 0) - (InterestExpenseValue != null ? UnitConversion.getBasicValueforNumbers(InterestExpenseDatas.UnitId, InterestExpenseValue.Value) : 0)) * (1 - (MarginalTaxDatas != null && MarginalTaxDatas.Value != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0) / 100));

                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Plus: Net Borrowing
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Plus: Net Borrowing";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DebtCapacityDatas != null ? DebtCapacityDatas.UnitId : null;
                                first = 0;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    //formula = C38+C40-C41-C42
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    ProjectOutputValuesViewModel SameyearValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel PrevYearValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;

                                    value = (SameyearValue != null ? UnitConversion.getBasicValueforNumbers(DebtCapacityDatas.UnitId, SameyearValue.Value) : 0) - (PrevYearValue != null ? UnitConversion.getBasicValueforNumbers(DebtCapacityDatas.UnitId, PrevYearValue.Value) : 0);

                                    dumyValue.BasicValue = value;

                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Free Cash Flow
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Free Cash Flow to Equity";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                FCFarray = new List<double>();
                                highunit = UnitConversion.getHigherDenominationUnit(UnleaveredIncomeDatasObj.UnitId, PlusDepreciationDatasObj.UnitId);
                                //Plus: Depreciation
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

                                ProjectOutputDatasViewModel NetIncomeDatas = OutputDatasVMList.Find(x => x.LineItem == "Net Income");
                                ProjectOutputDatasViewModel PlusBorrowingDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Plus: Net Borrowing"));

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    //NetIncomeDatas +PlusDepreciationDatasObj -LessCapexValue
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    ProjectOutputValuesViewModel NetIncomeValue = NetIncomeDatas != null && NetIncomeDatas.ProjectOutputValuesVM != null && NetIncomeDatas.ProjectOutputValuesVM.Count > 0 ? NetIncomeDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel PlusDepreciationValue = PlusDepreciationDatasObj != null && PlusDepreciationDatasObj.ProjectOutputValuesVM != null && PlusDepreciationDatasObj.ProjectOutputValuesVM.Count > 0 ? PlusDepreciationDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel LessCapexValue = LessCapexDatasObj != null && LessCapexDatasObj.ProjectOutputValuesVM != null && LessCapexDatasObj.ProjectOutputValuesVM.Count > 0 ? LessCapexDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel LessNWCValue = LessNWCDatasObj != null && LessNWCDatasObj.ProjectOutputValuesVM != null && LessNWCDatasObj.ProjectOutputValuesVM.Count > 0 ? LessNWCDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel PlusBorrowingValue = PlusBorrowingDatas != null && PlusBorrowingDatas.ProjectOutputValuesVM != null && PlusBorrowingDatas.ProjectOutputValuesVM.Count > 0 ? PlusBorrowingDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    value = ((NetIncomeValue != null ? UnitConversion.getBasicValueforNumbers(NetIncomeDatas.UnitId, NetIncomeValue.Value) : 0) + (PlusDepreciationValue != null ? UnitConversion.getBasicValueforNumbers(PlusDepreciationDatasObj.UnitId, PlusDepreciationValue.Value) : 0) - (LessCapexValue != null ? UnitConversion.getBasicValueforNumbers(LessCapexDatasObj.UnitId, LessCapexValue.Value) : 0) - (LessNWCValue != null ? UnitConversion.getBasicValueforNumbers(LessNWCDatasObj.UnitId, LessNWCValue.Value) : 0) + (PlusBorrowingValue != null ? UnitConversion.getBasicValueforNumbers(PlusBorrowingDatas.UnitId, PlusBorrowingValue.Value) : 0));

                                    dumyValue.BasicValue = value;
                                    FCFarray.Add(value);
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00"));  //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);


                                //Discount Factor
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Discount Factor";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = null;
                                var Costofequity = InputDatasList.Find(x => x.LineItem.Contains("Cost of Equity"));

                                double TValue = 1 + (Costofequity != null && Costofequity.Value != null ? Convert.ToDouble(Costofequity.Value) / 100 : 0);
                                int i = 0;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    value = Math.Pow(TValue, i);
                                    i++;
                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = dumyValue.BasicValue;
                                    //dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Discounted Cash Flow
                                //formula = Free Cash flow/Discount Factor
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Discounted Cash Flow";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                ProjectOutputDatasViewModel FreeCashFlowtoEquityDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow to Equity"));
                                ProjectOutputDatasViewModel DiscountFactorDatas = OutputDatasVMList.Find(x => x.LineItem == "Discount Factor");
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowtoEquityDatas != null ? FreeCashFlowtoEquityDatas.UnitId : null;
                                double NPv = 0;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {

                                    dumyValue = new ProjectOutputValuesViewModel();
                                    // dumyValue = TempValue;
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowtoEquityDatas != null && FreeCashFlowtoEquityDatas.ProjectOutputValuesVM != null && FreeCashFlowtoEquityDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowtoEquityDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    ProjectOutputValuesViewModel DiscountFactorValues = DiscountFactorDatas != null && DiscountFactorDatas.ProjectOutputValuesVM != null && DiscountFactorDatas.ProjectOutputValuesVM.Count > 0 ? DiscountFactorDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;

                                    value = (FreeCashValue != null ? UnitConversion.getBasicValueforNumbers(FreeCashFlowtoEquityDatas.UnitId, FreeCashValue.Value) : 0) / (DiscountFactorValues != null ? UnitConversion.getBasicValueforNumbers(DiscountFactorDatas.UnitId, DiscountFactorValues.Value) : 0);


                                    NPv = NPv + value;
                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    // dumyValue.Value = Convert.ToDouble(value.ToString("0.00")); //assign
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Net Present Value
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Net Present Value";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowtoEquityDatas != null ? FreeCashFlowtoEquityDatas.UnitId : null;
                                OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPv);

                                OutputDatasVMList.Add(OutputDatasVM);



                                #endregion
                            }
                            else if (projectVM.ValuationTechniqueId == 5)
                            {
                                #region Valuation 5
                                //Unlevered Value @ rU ($M) - VU
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Short Approach";
                                OutputDatasVM.LineItem = "Unlevered Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                int first = 0;
                                double lastYearValue = 0;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        //formula ==(L43+L45)/(1+$C$26); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
                                        ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }

                                OutputDatasVMList.Add(OutputDatasVM);

                                //formula=Interest coverage ratio =Interest expense/FCF (C187)* marginal Tax * Unleavered value of first year


                                var UnleaveredShortDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Unlevered Value @ rU") && x.Header == "APV Based Valuation - Short Approach");

                                // formula==$C$15(Marginal Tax)*$C$187*C208(Unleavered Value)
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Short Approach";
                                OutputDatasVM.LineItem = "PV of Interest Tax Shield";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleaveredShortDatas.UnitId;

                                var SLValue = UnleaveredShortDatas != null && UnleaveredShortDatas.ProjectOutputValuesVM != null && UnleaveredShortDatas.ProjectOutputValuesVM.Count > 0 ? UnleaveredShortDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;


                                double ULvalue = (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0);
                                OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, ULvalue);
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Levered Value (VL = VU + T)
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Short Approach";
                                OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;


                                double LValue = ((InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0)) + (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0);
                                OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, LValue);
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Net Present Value
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Short Approach";
                                OutputDatasVM.LineItem = "Net Present Value";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;



                                double NPV = ((InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0)) + (SLValue != null ? Convert.ToDouble(SLValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
                                OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Unlevered Value @ rU ($M) - VU
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Long Approach";
                                OutputDatasVM.LineItem = "Unlevered Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                OutputDatasVM.ProjectOutputValuesVM = UnleaveredShortDatas != null && UnleaveredShortDatas.ProjectOutputValuesVM != null ? UnleaveredShortDatas.ProjectOutputValuesVM : null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Long Approach";
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Interest Paid @ rD";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;

                                    double value = 0;
                                    ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    value = (FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) * (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0);

                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Debt Capacity
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Long Approach";
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Debt Capacity";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                //Interest paid
                                var InterestpaidLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Paid @ rD" && x.Header == "APV Based Valuation - Long Approach");

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year + 1) : null;
                                    value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0);
                                    dumyValue.BasicValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Interest Tax Shield @ Tc
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Long Approach";
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
                                first = 0;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    if (first == 0)
                                    {
                                        dumyValue.BasicValue = dumyValue.Value = value;
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                        value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0);
                                        dumyValue.BasicValue = value;
                                        dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    }
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Tax Shield Value @ rU
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Long Approach";
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Tax Shield Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                first = 0;
                                lastYearValue = 0;
                                var InterestTaxShieldLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Tax Shield @ Tc" && x.Header == "APV Based Valuation - Long Approach");
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestTaxShieldLongDatas.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        //formula ==(L43+L45)/(1+$C$26); // Next year data of InterestTaxShieldLongDatas L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
                                        ProjectOutputValuesViewModel InterestTaxShieldValue = InterestTaxShieldLongDatas != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestTaxShieldLongDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((InterestTaxShieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestTaxShieldLongDatas.UnitId, InterestTaxShieldValue.Value) : 0) + lastYearValue) / (1 + (UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) / 100 : 0));
                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);


                                //formula = Unlevered Value @ rU+ Tax Shield Value @ rU
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Long Approach";
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                first = 0;
                                lastYearValue = 0;
                                var UleaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU" && x.Header == "APV Based Valuation - Long Approach");
                                var InterestShieldValueLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Tax Shield Value @ rU" && x.Header == "APV Based Valuation - Long Approach");

                                highunit = UnitConversion.getHigherDenominationUnit((UleaveredLongDatas != null ? UleaveredLongDatas.UnitId : null), (InterestShieldValueLongDatas != null ? InterestShieldValueLongDatas.UnitId : null));
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel LUnleaveredValue = UleaveredLongDatas != null && UleaveredLongDatas.ProjectOutputValuesVM != null && UleaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? UleaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                        ProjectOutputValuesViewModel InterestDhieldValue = InterestShieldValueLongDatas != null && InterestShieldValueLongDatas.ProjectOutputValuesVM != null && InterestShieldValueLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestShieldValueLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                        value = (LUnleaveredValue != null ? UnitConversion.getBasicValueforCurrency(UleaveredLongDatas.UnitId, LUnleaveredValue.Value) : 0) + (InterestDhieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestShieldValueLongDatas.UnitId, InterestDhieldValue.Value) : 0);
                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Net Present Value
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.Header = "APV Based Valuation - Long Approach";
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Net Present Value";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                var leaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)" && x.Header == "APV Based Valuation - Long Approach");
                                highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredLongDatas.UnitId);
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

                                var levLongValue = leaveredLongDatas != null && leaveredLongDatas.ProjectOutputValuesVM != null && leaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? leaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

                                NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
                                OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVMList.Add(OutputDatasVM);


                                #endregion
                            }

                            else if (projectVM.ValuationTechniqueId == 6)
                            {
                                #region Valuation 6
                                //Unlevered Value @ rU ($M) - VU
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "Unlevered Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                int first = 0;
                                double lastYearValue = 0;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Debt Capacity
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Debt Capacity";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                var InputDebtCapacity = InputDatasList.Find(x => x.LineItem == "Fixed Schedule & Predetermined Debt Level, Dt");
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InputDebtCapacity.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    var InputDebtCapacityValue = InputDebtCapacity != null && InputDebtCapacity.ProjectInputValuesVM != null && InputDebtCapacity.ProjectInputValuesVM.Count > 0 ? InputDebtCapacity.ProjectInputValuesVM.Find(x => x.Year == (TempValue.Year)) : null;

                                    value = InputDebtCapacityValue != null ? UnitConversion.getBasicValueforCurrency(InputDebtCapacity.UnitId, InputDebtCapacityValue.Value) : 0;
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                // Interest Paid
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Interest Paid @ rD";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                //  OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;


                                //  OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;

                                var DebtCapacityDatas = OutputDatasVMList.Find(x => x.LineItem == "Debt Capacity");
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = DebtCapacityDatas.UnitId;

                                var CostOfDebtInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));

                                first = 0;
                                double FirstYearValue = 0;

                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderBy(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    //change formula
                                    //ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                    //value = (FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) * (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0);

                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel DebtCapacityValue = DebtCapacityDatas != null && DebtCapacityDatas.ProjectOutputValuesVM != null && DebtCapacityDatas.ProjectOutputValuesVM.Count > 0 ? DebtCapacityDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year - 1)) : null;
                                        value = (DebtCapacityValue != null ? UnitConversion.getBasicValueforCurrency(DebtCapacityDatas.UnitId, DebtCapacityValue.Value) : 0) * ((CostOfDebtInputDatas != null && CostOfDebtInputDatas.Value != null ? Convert.ToDouble(CostOfDebtInputDatas.Value) : 0) / 100);
                                    }

                                    dumyValue.BasicValue = value;
                                    FirstYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Interest Tax Shield @ Tc
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Interest Tax Shield @ Tc";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();


                                var InterestpaidLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Paid @ rD");
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestpaidLongDatas.UnitId;
                                first = 0;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList)
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    if (first == 0)
                                    {
                                        dumyValue.BasicValue = dumyValue.Value = value;
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel InterestPaidValue = InterestpaidLongDatas != null && InterestpaidLongDatas.ProjectOutputValuesVM != null && InterestpaidLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestpaidLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                        value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0);
                                        dumyValue.BasicValue = value;
                                        dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    }
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Tax Shield Value @ rU
                                //formula = 
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "Interest Tax Shield";
                                OutputDatasVM.LineItem = "Tax Shield Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                first = 0;
                                lastYearValue = 0;

                                //var CostofDebt = InputDatasList.Find(x => x.LineItem == "Cost of Debt");


                                var InterestTaxShieldLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Interest Tax Shield @ Tc");
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = InterestTaxShieldLongDatas.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        //formula ==(L43+L45)/(1+$C$26); // Next year data of InterestTaxShieldLongDatas L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
                                        ProjectOutputValuesViewModel InterestTaxShieldValue = InterestTaxShieldLongDatas != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM != null && InterestTaxShieldLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestTaxShieldLongDatas.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((InterestTaxShieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestTaxShieldLongDatas.UnitId, InterestTaxShieldValue.Value) : 0) + lastYearValue) / (1 + (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0));
                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                //formula = Unlevered Value @ rU+ Tax Shield Value @ rU
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "Adjacent Present Value";
                                OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                                first = 0;
                                lastYearValue = 0;

                                var UleaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU");
                                var InterestShieldValueLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Tax Shield Value @ rU");
                                highunit = UnitConversion.getHigherDenominationUnit((UleaveredLongDatas != null ? UleaveredLongDatas.UnitId : null), (InterestShieldValueLongDatas != null ? InterestShieldValueLongDatas.UnitId : null));
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;
                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        ProjectOutputValuesViewModel LUnleaveredValue = UleaveredLongDatas != null && UleaveredLongDatas.ProjectOutputValuesVM != null && UleaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? UleaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                        ProjectOutputValuesViewModel InterestDhieldValue = InterestShieldValueLongDatas != null && InterestShieldValueLongDatas.ProjectOutputValuesVM != null && InterestShieldValueLongDatas.ProjectOutputValuesVM.Count > 0 ? InterestShieldValueLongDatas.ProjectOutputValuesVM.Find(x => x.Year == TempValue.Year) : null;
                                        value = (LUnleaveredValue != null ? UnitConversion.getBasicValueforCurrency(UleaveredLongDatas.UnitId, LUnleaveredValue.Value) : 0) + (InterestDhieldValue != null ? UnitConversion.getBasicValueforCurrency(InterestShieldValueLongDatas.UnitId, InterestDhieldValue.Value) : 0);
                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);
                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }
                                OutputDatasVMList.Add(OutputDatasVM);

                                // formula==$C$15(Marginal Tax)*$C$187*C208(Unleavered Value)
                                //OutputDatasVM = new ProjectOutputDatasViewModel();
                                //OutputDatasVM.Id = 0;
                                //OutputDatasVM.ProjectId = projectVM.Id;
                                //OutputDatasVM.HeaderId = 0;
                                //OutputDatasVM.SubHeader = "";
                                //OutputDatasVM.LineItem = "PV of Interest Tax Shield";
                                //OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
                                //OutputDatasVM.HasMultiYear = false;
                                //OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleaveredDatasObj.UnitId;
                                //double ULvalue = (InterestCoverageDatas != null ? Convert.ToDouble(InterestCoverageDatas.Value) / 100 : 0) * (MarginalTaxDatas != null ? Convert.ToDouble(MarginalTaxDatas.Value) / 100 : 0) * (lastYearValue);
                                //OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, ULvalue);
                                //OutputDatasVM.ProjectOutputValuesVM = null;
                                //OutputDatasVMList.Add(OutputDatasVM);

                                //Net Present Value
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "Adjacent Present Value";
                                OutputDatasVM.LineItem = "Net Present Value";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;

                                var leaveredLongDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)");

                                highunit = UnitConversion.getHigherDenominationUnit(FreeCashFlowDatasObj.UnitId, leaveredLongDatas.UnitId);
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = highunit;

                                var levLongValue = leaveredLongDatas != null && leaveredLongDatas.ProjectOutputValuesVM != null && leaveredLongDatas.ProjectOutputValuesVM.Count > 0 ? leaveredLongDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
                                double NPV = (levLongValue != null ? Convert.ToDouble(levLongValue.BasicValue) : 0) + (TFCFValue != null ? Convert.ToDouble(TFCFValue.BasicValue) : 0);
                                OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, NPV);
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVMList.Add(OutputDatasVM);

                                #endregion

                            }
                            else if (projectVM.ValuationTechniqueId == 8)
                            {
                                #region Valuation8

                                //Unlevered Value
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "Unlevered Value @ rU";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = true;
                                OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();

                                int first = 0;
                                double lastYearValue = 0;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatasObj.UnitId;
                                foreach (ProjectOutputValuesViewModel TempValue in dumyValueList.OrderByDescending(x => x.Year))
                                {
                                    dumyValue = new ProjectOutputValuesViewModel();
                                    dumyValue.Id = 0; dumyValue.ProjectOutputDatasId = 0;
                                    dumyValue.Year = TempValue.Year;
                                    double value = 0;

                                    if (first == 0)
                                    {
                                        value = 0;
                                        first++;
                                    }
                                    else
                                    {
                                        //formula ==(L346+L349)/(1+$C$328); // Next year data of freecashflow L43 //Next year value of same field(Levered Value (VL)) //$C$26 WACC
                                        ProjectOutputValuesViewModel FreeCashFlowValue = FreeCashFlowDatasObj != null && FreeCashFlowDatasObj.ProjectOutputValuesVM != null && FreeCashFlowDatasObj.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatasObj.ProjectOutputValuesVM.Find(x => x.Year == (TempValue.Year + 1)) : null;
                                        value = ((FreeCashFlowValue != null ? UnitConversion.getBasicValueforCurrency(FreeCashFlowDatasObj.UnitId, FreeCashFlowValue.Value) : 0) + lastYearValue) / (1 + ((UCCInputDatas != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100));
                                    }
                                    dumyValue.BasicValue = value;
                                    lastYearValue = value;
                                    dumyValue.Value = UnitConversion.getConvertedValueforCurrency(OutputDatasVM.UnitId, dumyValue.BasicValue);

                                    OutputDatasVM.ProjectOutputValuesVM.Add(dumyValue);
                                }

                                OutputDatasVMList.Add(OutputDatasVM);

                                // PV of Interest Tax Shield
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "PV of Interest Tax Shield";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                // OutputDatasVM.Value = null;
                                OutputDatasVM.ProjectOutputValuesVM = null;

                                var ICRInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Intetrest coverage ratio =Interest expense/FCF"));
                                var CBInputDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
                                ProjectOutputDatasViewModel UnleveredValueDatas = OutputDatasVMList.Find(x => x.LineItem == "Unlevered Value @ rU");

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleveredValueDatas != null ? UnleveredValueDatas.UnitId : null;

                                // $C$15 $C$327  ((1 + C328) / (1 + C329)) * C349

                                ProjectOutputValuesViewModel UnLeveredValue = UnleveredValueDatas != null && UnleveredValueDatas.ProjectOutputValuesVM != null && UnleveredValueDatas.ProjectOutputValuesVM.Count > 0 ? UnleveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

                                OutputDatasVM.Value = ((MarginalTaxDatas != null && MarginalTaxDatas.Value != null ? Convert.ToDouble(MarginalTaxDatas.Value) : 0) / 100) * ((ICRInputDatas != null && ICRInputDatas.Value != null ? Convert.ToDouble(ICRInputDatas.Value) : 0) / 100) * ((1 + ((UCCInputDatas != null && UCCInputDatas.Value != null ? Convert.ToDouble(UCCInputDatas.Value) : 0) / 100)) / (1 + ((CBInputDatas != null && CBInputDatas.Value != null ? Convert.ToDouble(CBInputDatas.Value) : 0) / 100))) * (UnLeveredValue != null ? UnLeveredValue.Value : 0);
                                OutputDatasVMList.Add(OutputDatasVM);

                                //Levered Value
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "Levered Value (VL = VU + T)";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                OutputDatasVM.ProjectOutputValuesVM = null;

                                ProjectOutputDatasViewModel PVInterestDatas = OutputDatasVMList.Find(x => x.LineItem == "PV of Interest Tax Shield");

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = UnleveredValueDatas != null ? UnleveredValueDatas.UnitId : null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = PVInterestDatas != null ? PVInterestDatas.UnitId : null;

                                // C349 + C350

                                // ProjectOutputValuesViewModel PVInterestValue = PVInterestDatas != null && PVInterestDatas.ProjectOutputValuesVM != null && PVInterestDatas.ProjectOutputValuesVM.Count > 0 ? PVInterestDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

                                // OutputDatasVM.Value = (UnLeveredValue != null ? UnLeveredValue.Value : 0) + (PVInterestValue != null ? PVInterestValue.Value : 0);
                                OutputDatasVM.Value = (UnLeveredValue != null ? UnLeveredValue.Value : 0) + (PVInterestDatas != null ? PVInterestDatas.Value : 0);

                                OutputDatasVMList.Add(OutputDatasVM);


                                //Net Present Value ($M)
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "APV Based Valuation";
                                OutputDatasVM.LineItem = "Net Present Value";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.Currency;
                                OutputDatasVM.HasMultiYear = false;
                                OutputDatasVM.ProjectOutputValuesVM = null;

                                ProjectOutputDatasViewModel LeveredValueDatas = OutputDatasVMList.Find(x => x.LineItem == "Levered Value (VL = VU + T)");
                                ProjectOutputDatasViewModel FreeCashFlowDatas = OutputDatasVMList.Find(x => x.LineItem.Contains("Free Cash Flow"));

                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = LeveredValueDatas != null ? LeveredValueDatas.UnitId : null;
                                OutputDatasVM.DefaultUnitId = OutputDatasVM.UnitId = FreeCashFlowDatas != null ? FreeCashFlowDatas.UnitId : null;

                                // C351 + C346

                                //  ProjectOutputValuesViewModel LeveredValue = LeveredValueDatas != null && LeveredValueDatas.ProjectOutputValuesVM != null && LeveredValueDatas.ProjectOutputValuesVM.Count > 0 ? LeveredValueDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;
                                ProjectOutputValuesViewModel FreeCashValue = FreeCashFlowDatas != null && FreeCashFlowDatas.ProjectOutputValuesVM != null && FreeCashFlowDatas.ProjectOutputValuesVM.Count > 0 ? FreeCashFlowDatas.ProjectOutputValuesVM.Find(x => x.Year == projectVM.StartingYear) : null;

                                //  OutputDatasVM.Value = (LeveredValue != null ? LeveredValue.Value : 0) + (FreeCashValue != null ? FreeCashValue.Value : 0);
                                OutputDatasVM.Value = (LeveredValueDatas != null ? LeveredValueDatas.Value : 0) + (FreeCashValue != null ? FreeCashValue.Value : 0);

                                OutputDatasVMList.Add(OutputDatasVM);

                                #endregion
                            }

                            if (projectVM.ValuationTechniqueId != 8)
                            {
                                OutputDatasVM = new ProjectOutputDatasViewModel();
                                OutputDatasVM.Id = 0;
                                OutputDatasVM.ProjectId = projectVM.Id;
                                OutputDatasVM.HeaderId = 0;
                                OutputDatasVM.SubHeader = "";
                                OutputDatasVM.LineItem = "IRR (Internal Rate of Return) of Free Cash Flows";
                                OutputDatasVM.ValueTypeId = (int)ValueTypeEnum.percentage;
                                OutputDatasVM.HasMultiYear = false;
                                try
                                {
                                    OutputDatasVM.Value = Financial.Irr(FCFarray) * 100;
                                }
                                catch (Exception ss)
                                {
                                    Console.WriteLine(ss.Message);

                                    OutputDatasVM.Value = UnitConversion.getConvertedValueforCurrency(3, FCFarray.Sum());
                                }
                                OutputDatasVM.ProjectOutputValuesVM = null;
                                OutputDatasVMList.Add(OutputDatasVM);
                            }

                        }



                    }

                }


            }

            catch (Exception ss)
            {
                Console.WriteLine(ss.Message);

                // return OutputDatasVMList;
            }
            return OutputDatasVMList;
        }

        [HttpGet]
        [Route("[action]")]
        public ActionResult<Object> ExportEvaluationNew(string str, long UserId, long ProjectId, int Flag)
        {
            double marginalTax = 0;
            double discountRate = 0;
            double volChangePerc = 1;
            double unitPricePerChange = 1;
            double unitCostPerChange = 1;
           // object[][][] chageFixedCost = null;
            object[][] summaryOutput = null;
           // string lastCell = null;
            string rootFolder = _hostingEnvironment.WebRootPath;
            if (string.IsNullOrEmpty(rootFolder))
            {
                return NotFound("WebRootPath is not set.");
            }
            string fileName = @"capital_budgeting.xlsx";
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            if (!file.Exists)
            {
                return NotFound($"The file {fileName} does not exist at path {rootFolder}");
            }
            var formattedCustomObject = (String)null;
            double capexDepPerChange = 1;
          
            ProjectsViewModel result = new ProjectsViewModel();

            Project project = iProject.GetSingle(x => x.Id == ProjectId);
            if (project != null)
            {
                List<ProjectInputDatasViewModel> projectInputDatasVM = new List<ProjectInputDatasViewModel>();
                List<ProjectInputValuesViewModel> projectInputValuesVM = new List<ProjectInputValuesViewModel>();
                ProjectInputDatasViewModel DatasVm = new ProjectInputDatasViewModel();

                result = mapper.Map<Project, ProjectsViewModel>(project);

                List<ProjectInputDatas> projectInputList = iProjectInputDatas.FindBy(s => s.Id != 0 && s.ProjectId == ProjectId).ToList();

                if (projectInputList != null && projectInputList.Count > 0)
                {
                    List<ProjectInputValues> projectInputValueList = iProjectInputValues.FindBy(x => x.Id != 0 && projectInputList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();

                    //get all comparable List
                    List<ProjectInputComparables> projectInputComparableList = iProjectInputComparables.FindBy(x => x.Id != 0 && projectInputList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();

                    foreach (ProjectInputDatas datas in projectInputList)
                    {
                        DatasVm = new ProjectInputDatasViewModel();

                        //get all Comparable List By DatasId
                        List<ProjectInputComparables> ComparableList = projectInputComparableList.FindAll(x => x.Id != 0 && x.ProjectInputDatasId == datas.Id).ToList();
                        ProjectInputComparablesViewModel ComparablesVM = null;

                        DatasVm = mapper.Map<ProjectInputDatas, ProjectInputDatasViewModel>(datas);

                        if (ComparableList != null && ComparableList.Count > 0)
                        {
                            DatasVm.ProjectInputComparablesVM = new List<ProjectInputComparablesViewModel>();
                            foreach (var comparable in ComparableList)
                            {
                                ComparablesVM = mapper.Map<ProjectInputComparables, ProjectInputComparablesViewModel>(comparable);
                                DatasVm.ProjectInputComparablesVM.Add(ComparablesVM);
                            }
                        }

                        if (DatasVm != null && DatasVm.ProjectInputValuesVM != null && DatasVm.ProjectInputValuesVM.Count > 0)
                        {
                            var valuesList = DatasVm.ProjectInputValuesVM;
                            DatasVm.ProjectInputValuesVM = new List<ProjectInputValuesViewModel>();
                            DatasVm.ProjectInputValuesVM = valuesList.OrderBy(x => x.Year).ToList();
                        }
                        projectInputDatasVM.Add(DatasVm);
                    }

                    if (projectInputDatasVM != null && projectInputDatasVM.Count > 0)
                    {
                        result.ProjectInputDatasVM = new List<ProjectInputDatasViewModel>();
                        result.ProjectInputDatasVM = projectInputDatasVM;
                    }
                }

            }

            Dictionary<string, object> sensiScenSummaryOutput = new Dictionary<string, object>();

            using (ExcelPackage package = new ExcelPackage(file))
            {
                var wsCapitalBudgeting = package.Workbook.Worksheets["CapitalBudgeting"];
                wsCapitalBudgeting.SelectedRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                ExcelAddress yearCountAdd = null;
                List<string> yearCoutList = new List<string>();

                for (var i = 0; i < result.NoOfYears; i++)
                {
                    var curCol = 3 + i;
                    wsCapitalBudgeting.Cells[2, curCol].Value = result.StartingYear + i;
                    wsCapitalBudgeting.Cells[3, curCol].Value = i;
                    yearCountAdd = new ExcelAddress(3, curCol, 3, curCol);
                    wsCapitalBudgeting.Cells[2, curCol].Style.Font.Size = 12;
                    wsCapitalBudgeting.Cells[3, curCol].Style.Font.Size = 12;
                    wsCapitalBudgeting.Cells[3, curCol].Style.Font.Bold = true;
                    wsCapitalBudgeting.Cells[2, curCol].Style.Font.Bold = true;

                    yearCoutList.Add(yearCountAdd.ToString());
                }

                List<ProjectInputDatasViewModel> InputDatasList = new List<ProjectInputDatasViewModel>();
                InputDatasList = result.ProjectInputDatasVM;

                int noOfYears = result.NoOfYears ?? 0; // or a default value that makes sense for your logic


                wsCapitalBudgeting = ReturnCellStyle("Total", 2, (int)(noOfYears + 3), wsCapitalBudgeting, 0);
                wsCapitalBudgeting = ReturnCellStyle("Average", 2, (int)(noOfYears + 4), wsCapitalBudgeting, 0);

                List<string> keys = new List<string>();
                List<List<double>> values = new List<List<double>>();
                int row_count = 0;
                List<List<List<string>>> listOfCellsRevenue = new List<List<List<string>>>();
                List<List<List<string>>> listOfCellsRevenue1 = new List<List<List<string>>>();
                List<List<List<string>>> listOfCellsRevenue2 = new List<List<List<string>>>();

                List<ProjectInputDatasViewModel> RevenueVariableCostTier = new List<ProjectInputDatasViewModel>();

                if (InputDatasList != null && InputDatasList.Count > 0)
                {
                    RevenueVariableCostTier = InputDatasList.FindAll(x => x.SubHeader != null && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();
                }
        


                // RevenueVariableCostTier = InputDatasList.FindAll(x => x.SubHeader.Contains("Revenue & Variable Cost")).ToList();
               
                int CountOfRevenueVariableCostTier = 0;
                CountOfRevenueVariableCostTier = (RevenueVariableCostTier.Count) / 3;

                for (int i = 1; i < CountOfRevenueVariableCostTier + 1; i++)
                {
                    List<ProjectInputDatasViewModel> RevenueVariableCostTierNew = new List<ProjectInputDatasViewModel>();
                    // RevenueVariableCostTierNew = InputDatasList.FindAll(x => x.SubHeader.Contains("Revenue & Variable Cost Tier" + i + "")).ToList();
                    RevenueVariableCostTierNew = InputDatasList.FindAll(x => x.SubHeader.Contains("Revenue & Variable Cost Tier" + i + "")).OrderByDescending(i => i.LineItem).ToList();

                    // wsCapitalBudgeting = ExcelGeneration_New(RevenueVariableCostTierNew, "Revenue & Variable Cost Tier" + i + "", wsCapitalBudgeting, 3, row_count, volChangePerc,
                    //unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsRevenue, CountOfRevenueVariableCostTier, listOfCellsRevenueNew);

                    wsCapitalBudgeting = ExcelGeneration_New(RevenueVariableCostTierNew, "Revenue & Variable Cost Tier" + i + "", wsCapitalBudgeting, 3, row_count, volChangePerc,
                   unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsRevenue);

                    if(i==1)
                    {
                        listOfCellsRevenue1 = listOfCellsRevenue;
                    }
                    else if(i==2)
                    {
                        listOfCellsRevenue2 = listOfCellsRevenue;
                    }
                    
                }

                List<List<List<string>>> listOfCellsFixed = new List<List<List<string>>>();
                List<ProjectInputDatasViewModel> OtherFixedCost = new List<ProjectInputDatasViewModel>();
                //OtherFixedCost = InputDatasList.FindAll(x => x.SubHeader.Contains("Other Fixed Cost")).ToList();
                if(InputDatasList != null && InputDatasList.Count > 0)
                {
                    OtherFixedCost = InputDatasList.FindAll(x =>x.SubHeader != null &&  x.SubHeader.Contains("Other Fixed Cost")).OrderByDescending(i => i.LineItem).ToList();
                }
                
                if (Flag == 2 && str != null && OtherFixedCost != null)
                {
                    // TODO ---- 
                    //wsCapitalBudgeting = ExcelGeneration(chageFixedCost, "Fixed Cost", wsCapitalBudgeting, 3, row_count, volChangePerc, unitPricePerChange,
                    //    unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsFixed);                  
                }
                else if (OtherFixedCost != null)
                {
                    wsCapitalBudgeting = ExcelGeneration_New(OtherFixedCost, "Other Fixed Cost", wsCapitalBudgeting, 3, row_count, volChangePerc, unitPricePerChange,
                       unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsFixed);
                }

                List<List<List<string>>> listOfCellsCapex = new List<List<List<string>>>();
                List<List<List<string>>> listOfCellsCapex1 = new List<List<List<string>>>();
                List<List<List<string>>> listOfCellsCapex2 = new List<List<List<string>>>();

                List<ProjectInputDatasViewModel> CapexDepreciation = new List<ProjectInputDatasViewModel>();
                if(InputDatasList != null && InputDatasList.Count > 0)
                {
                    CapexDepreciation = InputDatasList.FindAll(x => x.SubHeader != null && x.SubHeader.Contains("Capex & Depreciation")).ToList();
                }
                // CapexDepreciation = InputDatasList.FindAll(x => x.SubHeader.Contains("Capex & Depreciation")).ToList();

                int CountOfCapexDepreciation = 0;

                if (Flag == 2 && str != null && CapexDepreciation != null)
                {
                    //wsCapitalBudgeting = ExcelGeneration_New(CapexDepreciation, "Capex & Depreciation", wsCapitalBudgeting, 3, row_count, volChangePerc,
                    //    unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);
                }
                else if (CapexDepreciation != null)
                {
                    //int CountOfCapexDepreciation = 0;
                    CountOfCapexDepreciation = (CapexDepreciation.Count) / 2;

                    for (int i = 1; i < CountOfCapexDepreciation + 1; i++)
                    {
                        List<ProjectInputDatasViewModel> CapexDepreciationNew = new List<ProjectInputDatasViewModel>();
                        //CapexDepreciationNew = InputDatasList.FindAll(x => x.SubHeader.Contains("Capex & Depreciation" + i + "")).ToList();
                        CapexDepreciationNew = InputDatasList.FindAll(x => x.SubHeader.Contains("Capex & Depreciation" + i + "")).OrderBy(i => i.LineItem).ToList();

                        wsCapitalBudgeting = ExcelGeneration_New(CapexDepreciationNew, "Capex & Depreciation" + i + "", wsCapitalBudgeting, 3, row_count, volChangePerc,
                        unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);

                        if (i == 1)
                        {
                            listOfCellsCapex1 = listOfCellsCapex;
                        }
                        else if (i == 2)
                        {
                            listOfCellsCapex2 = listOfCellsCapex;
                        }

                    }
                }

                List<List<List<string>>> listOfCellsNWC = new List<List<List<string>>>();

                List<ProjectInputDatasViewModel> WorkingCapital = new List<ProjectInputDatasViewModel>();
                if(InputDatasList != null && InputDatasList.Count > 0)
                {
                    WorkingCapital = InputDatasList.FindAll(x => x.SubHeader != null && x.SubHeader.Contains("Working Capital")).ToList();
                }
                // WorkingCapital = InputDatasList.FindAll(x => x.SubHeader.Contains("Working Capital")).ToList();

                wsCapitalBudgeting = ExcelGeneration_New(WorkingCapital, "Working Capital", wsCapitalBudgeting, 3, row_count, volChangePerc,
                        unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsNWC);

                row_count = row_count + 4;

                List<List<List<string>>> listOfFixedSchedule = new List<List<List<string>>>(); // For Valuation Technique 6

                var marginalTaxAdd = new ExcelAddress(0,0,0,0);
                var discountRateAdd = new ExcelAddress(0,0,0,0);
                var DVRatioAdd = new ExcelAddress(0, 0, 0, 0);
                var UnleveredCostOfCapitalAdd = new ExcelAddress(0, 0, 0, 0);
                var CostOfDebtAdd = new ExcelAddress(0, 0, 0, 0);
                var WeightedAverageCostOfCapitalAdd = new ExcelAddress(0, 0, 0, 0);
                var CostOfEquityAdd = new ExcelAddress(0, 0, 0, 0);
                var IntetrestCoverageRatioAdd = new ExcelAddress(0, 0, 0, 0);
                var ComparablesUnleveredCostofCapitalAdd = new ExcelAddress(0, 0, 0, 0);
                var ProjectTargetLeverageAdd = new ExcelAddress(0, 0, 0, 0);
                var ProjectDebtCostofCapitalAdd = new ExcelAddress(0, 0, 0, 0);
                var ProjectWACCAdd = new ExcelAddress(0, 0, 0, 0);

                var marginalTaxAddColumn = "";
                var discountRateAddColumn = "";
                var DVRatioAddColumn = "";
                var UnleveredCostOfCapitalAddColumn = "";
                var CostOfDebtAddColumn = "";
                var WeightedAverageCostOfCapitalAddColumn = "";
                var CostOfEquityAddColumn = "";
                var IntetrestCoverageRatioAddColumn = "";
                var ProjectWACCAddColumn = "";
                Console.WriteLine("Flag : " + Flag);
                // Valuation Technique
                if (project != null && project.ValuationTechniqueId == 1)
                {
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    
                    if (Flag == 2 && str != null)
                    {
                        wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));                      
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var DiscountRateDatas = InputDatasList.Find(x => x.LineItem.Contains("WACC"));
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = DiscountRateDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);                           
                         discountRateAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);

                        marginalTaxAddColumn = marginalTaxAdd.Address;
                        discountRateAddColumn = discountRateAdd.Address;
                   }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Weighted Average Cost of Capital", (row_count + 1), 1, wsCapitalBudgeting, 0);

                    row_count = (row_count + 2);                  
                }
                else if (project != null && project.ValuationTechniqueId == 2)
                {                
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    if (Flag == 2 && str != null)
                    {
                        //wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        //wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var DVRatioDatas = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = DVRatioDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                        var UnleveredCostOfCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Value = UnleveredCostOfCapitalDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.Numberformat.Format = "0.00%";

                        var CostOfDebtDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Value = CostOfDebtDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.Numberformat.Format = "0.00%";

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);
                        DVRatioAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);
                        UnleveredCostOfCapitalAdd = new ExcelAddress((row_count + 2), 2, (row_count + 2), 2);
                        CostOfDebtAdd = new ExcelAddress((row_count + 3), 2, (row_count + 3), 2);

                        marginalTaxAddColumn = marginalTaxAdd.Address;
                        DVRatioAddColumn = DVRatioAdd.Address;
                        UnleveredCostOfCapitalAddColumn = UnleveredCostOfCapitalAdd.Address;
                        CostOfDebtAddColumn = CostOfDebtAdd.Address;
                    }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("D/V Ratio", (row_count + 1), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Unlevered Cost of Capital", (row_count + 2), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Cost of Debt", (row_count + 3), 1, wsCapitalBudgeting, 0);

                    row_count = (row_count + 4);
                }
                else if (project != null && project.ValuationTechniqueId == 3)
                {
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 4), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    if (Flag == 2 && str != null)
                    {
                        //wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        //wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var WeightedAverageCostOfCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("WACC"));
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = WeightedAverageCostOfCapitalDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                        var DVRatioDatas = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Value = DVRatioDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.Numberformat.Format = "0.00%";

                        var CostOfDebtDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Value = CostOfDebtDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.Numberformat.Format = "0.00%";

                        var CostOfEquityDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Equity"));
                        wsCapitalBudgeting.Cells[(row_count + 4), 2].Value = CostOfEquityDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 4), 2].Style.Numberformat.Format = "0.00%";

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);
                        WeightedAverageCostOfCapitalAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);
                        DVRatioAdd = new ExcelAddress((row_count + 2), 2, (row_count + 2), 2);
                        CostOfDebtAdd = new ExcelAddress((row_count + 3), 2, (row_count + 3), 2);
                        CostOfEquityAdd = new ExcelAddress((row_count + 4), 2, (row_count + 4), 2);

                        marginalTaxAddColumn = marginalTaxAdd.Address;
                        WeightedAverageCostOfCapitalAddColumn = WeightedAverageCostOfCapitalAdd.Address;
                        DVRatioAddColumn = DVRatioAdd.Address;
                        CostOfDebtAddColumn = CostOfDebtAdd.Address;
                        CostOfEquityAddColumn = CostOfEquityAdd.Address;
                  }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Weighted Average Cost of Capital", (row_count + 1), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("D/V Ratio", (row_count + 2), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Cost of Debt", (row_count + 3), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Cost of Equity", (row_count + 4), 1, wsCapitalBudgeting, 0);

                    row_count = (row_count + 5);
                }
                else if (project != null && project.ValuationTechniqueId == 4)
                {
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    
                    if (Flag == 2 && str != null)
                    {
                        //wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        //wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var comparableDatas = InputDatasList.Find(x => x.LineItem.Contains("comparable"));
                        // wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = comparableDatas.Value;

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);
                        marginalTaxAddColumn = marginalTaxAdd.Address;

                       // ExcelAddress comparableCountAdd = null;
                       // List<string> comparableCoutList = new List<string>();

                        for (var i = 1; i <= comparableDatas.Value; i++)
                        {
                            var curCol = 2 + i;
                            wsCapitalBudgeting.Cells[row_count + 1, curCol].Value = "Comparable" + i;
                            //wsCapitalBudgeting.Cells[3, curCol].Value = i;
                            //comparableCountAdd = new ExcelAddress(3, curCol, 3, curCol);
                            wsCapitalBudgeting.Cells[row_count + 1, curCol].Style.Font.Size = 12;
                           // wsCapitalBudgeting.Cells[3, curCol].Style.Font.Size = 12;
                           // wsCapitalBudgeting.Cells[3, curCol].Style.Font.Bold = true;
                            wsCapitalBudgeting.Cells[row_count + 1, curCol].Style.Font.Bold = true;

                            //comparableCoutList.Add(comparableCountAdd.ToString());
                        }                      
                    }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                 
                    List<ProjectInputDatasViewModel> comparable = new List<ProjectInputDatasViewModel>();
                    comparable = InputDatasList.FindAll(x => x.SubHeader.Contains("Comparable")).ToList();

                    wsCapitalBudgeting = ExcelGeneration_NewForVT4(comparable, "Example for Computing Comparable's rU", wsCapitalBudgeting, 3, row_count,out row_count);

                    row_count = (row_count + 2);

                    wsCapitalBudgeting = ReturnCellStyle("Industry or Comparables Unlevered Cost of Capital (rU-Comp)", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Project's Target Leverage (D/V) Ratio", (row_count + 1), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Project's Debt Cost of Capital (rD)", (row_count + 2), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Project's Equity Cost of Capital (rE)", (row_count + 3), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Project's rWACC (Weighted average cost of capital)", (row_count + 4), 1, wsCapitalBudgeting, 0);

                    var ComparablesUnleveredCostofCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Industry or Comparables Unlevered Cost of Capital (rU-Comp)") && x.SubHeader == "");
                    wsCapitalBudgeting.Cells[(row_count), 2].Value = ComparablesUnleveredCostofCapitalDatas.Value / 100;
                    ComparablesUnleveredCostofCapitalAdd = new ExcelAddress((row_count), 2, (row_count), 2);

                    wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.00%";

                    var ProjectTargetLeverageDatas = InputDatasList.Find(x => x.LineItem.Contains("Project's Target Leverage (D/V) Ratio") && x.SubHeader == "");
                    wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = ProjectTargetLeverageDatas.Value / 100;
                    ProjectTargetLeverageAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);

                    wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                    var ProjectDebtCostofCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Project's Debt Cost of Capital (rD)") && x.SubHeader == "");
                    wsCapitalBudgeting.Cells[(row_count + 2), 2].Value = ProjectDebtCostofCapitalDatas.Value / 100;
                    ProjectDebtCostofCapitalAdd = new ExcelAddress((row_count + 2), 2, (row_count + 2), 2);

                    wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.Numberformat.Format = "0.00%";

                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Formula = "(" + "(" + ComparablesUnleveredCostofCapitalAdd.Address + "/" + "100" + ")" + "+" + 
                    //    "(" + ProjectTargetLeverageAdd.Address + "/" + "100" + ")" + "/(1-" + "(" + ProjectTargetLeverageAdd.Address + "/" + "100" + ")" + ")" +
                    //    "*(" + "(" + ComparablesUnleveredCostofCapitalAdd.Address + "/" + "100" + ")" + "-" + 
                    //    "(" + ProjectDebtCostofCapitalAdd.Address + "/" + "100" + ")" + ")" + ")" + "*" + "100";

                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Formula = "(" + ComparablesUnleveredCostofCapitalAdd.Address + "/" + "100" + ")" + "+" +
                    //    "(" + ProjectTargetLeverageAdd.Address + "/" + "100" + ")" + "/(1-" + "(" + ProjectTargetLeverageAdd.Address + "/" + "100" + ")" + ")" +
                    //    "*(" + "(" + ComparablesUnleveredCostofCapitalAdd.Address + "/" + "100" + ")" + "-" +
                    //    "(" + ProjectDebtCostofCapitalAdd.Address + "/" + "100" + ")" + ")";

                    wsCapitalBudgeting.Cells[(row_count + 3), 2].Formula = ComparablesUnleveredCostofCapitalAdd.Address + "+" + ProjectTargetLeverageAdd.Address + 
                                                                           "/(1-" + ProjectTargetLeverageAdd.Address + ")" + 
                                                                           "*(" + ComparablesUnleveredCostofCapitalAdd.Address + "-" + ProjectDebtCostofCapitalAdd.Address + ")";

                    //var ProjectEquityCostofCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Project's Equity Cost of Capital (rE)") && x.SubHeader == "");
                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Value = ProjectEquityCostofCapitalDatas.Value / 100;

                    wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.Numberformat.Format = "0.00%";

                    //wsCapitalBudgeting.Cells[(row_count + 4), 2].Formula = "(" + "(" + ComparablesUnleveredCostofCapitalAdd.Address + "/" + "100" + ")" + "-" +
                    //    "(" + ProjectTargetLeverageAdd.Address + "/" + "100" + ")" + "*" + "(" + marginalTaxAddColumn + "/" + "100" + ")" + "*" +
                    //    "(" + ProjectDebtCostofCapitalAdd.Address + "/" + "100" + ")" + ")" + "*" + "100";

                    //wsCapitalBudgeting.Cells[(row_count + 4), 2].Formula = "(" + ComparablesUnleveredCostofCapitalAdd.Address + "/" + "100" + ")" + "-" +
                    //    "(" + ProjectTargetLeverageAdd.Address + "/" + "100" + ")" + "*" + "(" + marginalTaxAddColumn + "/" + "100" + ")" + "*" +
                    //    "(" + ProjectDebtCostofCapitalAdd.Address + "/" + "100" + ")";

                    wsCapitalBudgeting.Cells[(row_count + 4), 2].Formula = ComparablesUnleveredCostofCapitalAdd.Address + "-" + ProjectTargetLeverageAdd.Address + 
                                                                           "*" + marginalTaxAddColumn + "*" + ProjectDebtCostofCapitalAdd.Address ;

                    //var ProjectWACCDatas = InputDatasList.Find(x => x.LineItem.Contains("WACC") && x.SubHeader == "");
                    //wsCapitalBudgeting.Cells[(row_count + 4), 2].Value = ProjectWACCDatas.Value / 100;

                    wsCapitalBudgeting.Cells[(row_count + 4), 2].Style.Numberformat.Format = "0.00%";

                    ProjectWACCAdd = new ExcelAddress((row_count + 4), 2, (row_count + 4), 2);
                    ProjectWACCAddColumn = ProjectWACCAdd.Address;

                    row_count = (row_count + 5);

                }
                else if (project != null && project.ValuationTechniqueId == 5)
                {
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    if (Flag == 2 && str != null)
                    {
                        //wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        //wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var UnleveredCostOfCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = UnleveredCostOfCapitalDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                        var CostOfDebtDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Value = CostOfDebtDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.Numberformat.Format = "0.00%";

                        var IntetrestCoverageRatioDatas = InputDatasList.Find(x => x.LineItem.Contains("Intetrest coverage ratio =Interest expense/FCF"));
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Value = IntetrestCoverageRatioDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.Numberformat.Format = "0.00%";

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);
                        UnleveredCostOfCapitalAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);
                        CostOfDebtAdd = new ExcelAddress((row_count + 2), 2, (row_count + 2), 2);
                        IntetrestCoverageRatioAdd = new ExcelAddress((row_count + 3), 2, (row_count + 3), 2);

                        marginalTaxAddColumn = marginalTaxAdd.Address;
                        UnleveredCostOfCapitalAddColumn = UnleveredCostOfCapitalAdd.Address;
                        CostOfDebtAddColumn = CostOfDebtAdd.Address;
                        IntetrestCoverageRatioAddColumn = IntetrestCoverageRatioAdd.Address;
                    }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Unlevered Cost of Capital", (row_count + 1), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Cost of Debt", (row_count + 2), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Interest Coverage Ratio =Interest Expense/FCF", (row_count + 3), 1, wsCapitalBudgeting, 0);

                    row_count = (row_count + 4);
                }
                else if (project != null && project.ValuationTechniqueId == 6)
                {
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                   
                    if (Flag == 2 && str != null)
                    {
                        //wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        //wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var UnleveredCostOfCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = UnleveredCostOfCapitalDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                        var CostOfDebtDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Value = CostOfDebtDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.Numberformat.Format = "0.00%";

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);
                        UnleveredCostOfCapitalAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);
                        CostOfDebtAdd = new ExcelAddress((row_count + 2), 2, (row_count + 2), 2);

                        marginalTaxAddColumn = marginalTaxAdd.Address;
                        UnleveredCostOfCapitalAddColumn = UnleveredCostOfCapitalAdd.Address;
                        CostOfDebtAddColumn = CostOfDebtAdd.Address;
                    }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Unlevered Cost of Capital", (row_count + 1), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Cost of Debt", (row_count + 2), 1, wsCapitalBudgeting, 0);                   

                    List<ProjectInputDatasViewModel> FixedSchedule = new List<ProjectInputDatasViewModel>();
                    FixedSchedule = InputDatasList.FindAll(x => x.SubHeader.Contains("Fixed Schedule")).ToList();

                    wsCapitalBudgeting = ExcelGeneration_NewForVT6(FixedSchedule, "Fixed Schedule & Predetermined Debt Level, Dt", wsCapitalBudgeting, 3, row_count,
                                                                   out row_count, out listOfFixedSchedule);

                    row_count = (row_count + 2);
                }
                else if (project != null && project.ValuationTechniqueId == 7)
                {
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 4), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    if (Flag == 2 && str != null)
                    {
                        //wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        //wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var WeightedAverageCostOfCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("WACC"));
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = WeightedAverageCostOfCapitalDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                        var DVRatioDatas = InputDatasList.Find(x => x.LineItem.Contains("D/V Ratio"));
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Value = DVRatioDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.Numberformat.Format = "0.00%";

                        var UnleveredCostOfCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Value = UnleveredCostOfCapitalDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.Numberformat.Format = "0.00%";

                        var CostOfDebtDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
                        wsCapitalBudgeting.Cells[(row_count + 4), 2].Value = CostOfDebtDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 4), 2].Style.Numberformat.Format = "0.00%";

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);
                        WeightedAverageCostOfCapitalAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);
                        DVRatioAdd = new ExcelAddress((row_count + 2), 2, (row_count + 2), 2);
                        UnleveredCostOfCapitalAdd = new ExcelAddress((row_count + 3), 2, (row_count + 3), 2);
                        CostOfDebtAdd = new ExcelAddress((row_count + 4), 2, (row_count + 4), 2);

                        marginalTaxAddColumn = marginalTaxAdd.Address;
                        WeightedAverageCostOfCapitalAddColumn = WeightedAverageCostOfCapitalAdd.Address;
                        DVRatioAddColumn = DVRatioAdd.Address;
                        UnleveredCostOfCapitalAddColumn = UnleveredCostOfCapitalAdd.Address;
                        CostOfDebtAddColumn = CostOfDebtAdd.Address;
                      }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Weighted Average Cost of Capital", (row_count + 1), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("D/V Ratio", (row_count + 2), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Unlevered Cost of Capital", (row_count + 3), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Cost of Debt", (row_count + 4), 1, wsCapitalBudgeting, 0);

                    row_count = (row_count + 5);
                }
                else if (project != null && project.ValuationTechniqueId == 8)
                {
                    //wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    if (Flag == 2 && str != null)
                    {
                        //wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                        //wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                    }
                    else
                    {
                        var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                        wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count), 2].Style.Numberformat.Format = "0.0%";

                        var UnleveredCostOfCapitalDatas = InputDatasList.Find(x => x.LineItem.Contains("Unlevered cost of capital"));
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = UnleveredCostOfCapitalDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.Numberformat.Format = "0.00%";

                        var CostOfDebtDatas = InputDatasList.Find(x => x.LineItem.Contains("Cost of Debt"));
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Value = CostOfDebtDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 2), 2].Style.Numberformat.Format = "0.00%";

                        var IntetrestCoverageRatioDatas = InputDatasList.Find(x => x.LineItem.Contains("Intetrest coverage ratio =Interest expense/FCF"));
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Value = IntetrestCoverageRatioDatas.Value / 100;
                        wsCapitalBudgeting.Cells[(row_count + 3), 2].Style.Numberformat.Format = "0.00%";

                        marginalTaxAdd = new ExcelAddress((row_count), 2, (row_count), 2);
                        UnleveredCostOfCapitalAdd = new ExcelAddress((row_count + 1), 2, (row_count + 1), 2);
                        CostOfDebtAdd = new ExcelAddress((row_count + 2), 2, (row_count + 2), 2);
                        IntetrestCoverageRatioAdd = new ExcelAddress((row_count + 3), 2, (row_count + 3), 2);

                        marginalTaxAddColumn = marginalTaxAdd.Address;
                        UnleveredCostOfCapitalAddColumn = UnleveredCostOfCapitalAdd.Address;
                        CostOfDebtAddColumn = CostOfDebtAdd.Address;
                        IntetrestCoverageRatioAddColumn = IntetrestCoverageRatioAdd.Address;                     
                   }

                    wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate", (row_count), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Unlevered Cost of Capital", (row_count + 1), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Cost of Debt", (row_count + 2), 1, wsCapitalBudgeting, 0);
                    wsCapitalBudgeting = ReturnCellStyle("Interest Coverage Ratio =Interest Expense/FCF", (row_count + 3), 1, wsCapitalBudgeting, 0);

                    row_count = (row_count + 4);
                }

                // wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                // wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                //if (Flag == 2 && str != null)
                //{
                //    wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                //    wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;
                //}
                //else
                //{
                //    var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                //    wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value;

                //    var DiscountRateDatas = InputDatasList.Find(x => x.LineItem.Contains("WACC"));
                //    wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = DiscountRateDatas.Value;
                //}

                // wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate(%)", (row_count), 1, wsCapitalBudgeting, 0);
                // wsCapitalBudgeting = ReturnCellStyle("Weighted average cost of capital(%)", (row_count + 1), 1, wsCapitalBudgeting, 0);

                //row_count = (row_count + 2);

                if (Flag == 1 || Flag == 3)
                {                   
                    result.ProjectSummaryDatasVM = GetProjectOutput(result);
                }
                else if (Flag == 2)
                {
                    summaryOutput = DictionaryToArray(sensiScenSummaryOutput).Select(a => a.ToArray()).ToArray(); ;
                }

                if (result.ProjectSummaryDatasVM != null)
                    foreach (var outputDatasVm in result.ProjectSummaryDatasVM)
                    {
                        if (outputDatasVm != null && outputDatasVm.ProjectOutputValuesVM != null && outputDatasVm.ProjectOutputValuesVM.Count > 0)
                        {
                            var valuesList = outputDatasVm.ProjectOutputValuesVM;
                            outputDatasVm.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                            outputDatasVm.ProjectOutputValuesVM = valuesList.OrderBy(x => x.Year).ToList();
                        }
                    }

                string[,] addArray2D = new string[1, 1];

                if (result.ProjectSummaryDatasVM != null && result.ProjectSummaryDatasVM.Count > 0 &&
                    result.ProjectSummaryDatasVM[0] != null && result.ProjectSummaryDatasVM[0].ProjectOutputValuesVM != null)
                {
                    addArray2D = new string[result.ProjectSummaryDatasVM.Count, result.ProjectSummaryDatasVM[0].ProjectOutputValuesVM.Count + 1];
                }

                                

                // string[,] addArray2D = new string[result.ProjectSummaryDatasVM.Count, result.ProjectSummaryDatasVM[0].ProjectOutputValuesVM.Count + 1];

            if (result != null && result.ProjectSummaryDatasVM != null && result.ProjectSummaryDatasVM.Count > 0){

                for (int o = 0; o < result.ProjectSummaryDatasVM.Count; o++)
                {
                    if (result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM != null)
                    {
                        for (int p = 0; p < result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM.Count; p++)
                        {
                            try
                            {
                                var conditionCheck = result.ProjectSummaryDatasVM[o].LineItem;

                                if (p == 0)
                                {
                                    
                                    string UnitValue = null;

                                    if (result.ProjectSummaryDatasVM[o].UnitId == 1)
                                    {
                                        UnitValue = "$";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 2)
                                    {
                                        UnitValue = "$K";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 3)
                                    {
                                        UnitValue = "$M";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 4)
                                    {
                                        UnitValue = "$B";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 5)
                                    {
                                        UnitValue = "$T";
                                    }

                                    if (conditionCheck.Contains("Sales") || conditionCheck.Contains("COGS") ||
                                    conditionCheck.Contains("Unlevered Net Income") || conditionCheck.Contains("Free Cash Flow") ||
                                    conditionCheck.Contains("Gross Margin") || conditionCheck.Contains("Operating Income") ||
                                     conditionCheck.Contains("Net Present Value") || conditionCheck.Contains("Discounted Cash Flow"))
                                    {
                                        wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 0);
                                        addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                                    }
                                    else
                                    {
                                        if (conditionCheck == "Discount Factor")
                                        {
                                            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + "$" + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 1);
                                            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                                        }
                                        else
                                        {
                                            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 1);
                                            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                                        }

                                    }
                                }

                                if (conditionCheck == "Depreciation")
                                {
                                    wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM[p].Value), (row_count + o), (p + 3), wsCapitalBudgeting, 1);
                                    wsCapitalBudgeting.Cells[(row_count + o), p + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                }                                
                                //wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM[p].Value), (row_count + o), (p + 3), wsCapitalBudgeting, 1);
                                var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                                //addArray2D[o, p] = Convert.ToString(addr.ToString());
                                addArray2D[o, p+1] = Convert.ToString(addr.ToString());
                            }
                            catch (IndexOutOfRangeException)
                            {
                                var strr = "";
                                wsCapitalBudgeting = ReturnCellStyle(strr, (row_count + o), (p + 3), wsCapitalBudgeting, 1);
                                var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                                addArray2D[o, p] = Convert.ToString(addr.ToString());
                            }
                        }
                    }
                    else if (result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM == null)
                    {
                        int p = 0;
                        var conditionCheck = result.ProjectSummaryDatasVM[o].LineItem;

                        string UnitValue = null;

                        if (result.ProjectSummaryDatasVM[o].UnitId == 1)
                        {
                            UnitValue = "$";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 2)
                        {
                            UnitValue = "$K";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 3)
                        {
                            UnitValue = "$M";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 4)
                        {
                            UnitValue = "$B";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 5)
                        {
                            UnitValue = "$T";
                        }

                        if (conditionCheck.Contains("Net Present Value") ||
                                conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                        {

                            if (conditionCheck == "IRR (Internal Rate of Return) of Free Cash Flows")
                            {
                                wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + "%" + ")", (row_count + o), (1), wsCapitalBudgeting, 0);
                                addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                            }
                            else
                            {
                                wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (1), wsCapitalBudgeting, 0);
                                addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                            }

                        }
                        else
                        {
                            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (1), wsCapitalBudgeting, 1);
                            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                        }

                        //wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(result.ProjectSummaryDatasVM[o].Value), (row_count + o), (3), wsCapitalBudgeting, 1);
                        var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                        //addArray2D[o, p] = Convert.ToString(addr.ToString());
                        addArray2D[o, p+1] = Convert.ToString(addr.ToString());
                    }
                }
            }

                List<string[,]> addArray2DList = new List<string[,]>();
                addArray2DList.Add(addArray2D);

                List<ProjectsViewModel> resultData = new List<ProjectsViewModel>();
                resultData.Add(result);
               
                List<string> sales = new List<string>();
                List<string> cogs = new List<string>();

                List<string> grossMargin = new List<string>();

                //var grossMargin = RevenueFormula(listOfCellsRevenue, out sales, out cogs);
                if (CountOfRevenueVariableCostTier==1)
                {
                     grossMargin = RevenueFormula(listOfCellsRevenue, out sales, out cogs);
                }
                else if (CountOfRevenueVariableCostTier == 2)
                {
                     grossMargin = RevenueFormula1(listOfCellsRevenue, listOfCellsRevenue1, out sales, out cogs);
                }
                else if (CountOfRevenueVariableCostTier == 3)
                {
                     grossMargin = RevenueFormula2(listOfCellsRevenue, listOfCellsRevenue1, listOfCellsRevenue2, out sales, out cogs);
                }                    
                var operatingIncome2D = GetCellAddress(addArray2D, "COGS", "Operating Income");
                var freeCashFlow2D = GetCellAddress(addArray2D, "Income Tax", "Free Cash Flow");
                var operatingIncome = FormulaGeneration(operatingIncome2D, 0);
                var freeCashFlow = FormulaGeneration(freeCashFlow2D, 1);

         




                wsCapitalBudgeting = InsertInputToExcel(resultData, wsCapitalBudgeting, row_count,listOfCellsRevenue,sales,cogs, 
                                                        grossMargin, listOfCellsFixed, operatingIncome, freeCashFlow, addArray2DList,
                                                        marginalTaxAddColumn, discountRateAddColumn,listOfCellsNWC, 
                                                        listOfCellsCapex, listOfCellsCapex1, listOfCellsCapex2, CountOfCapexDepreciation, listOfFixedSchedule, yearCoutList,
                                                        WeightedAverageCostOfCapitalAddColumn, DVRatioAddColumn, CostOfDebtAddColumn, CostOfEquityAddColumn,
                                                        UnleveredCostOfCapitalAddColumn, IntetrestCoverageRatioAddColumn, ProjectWACCAddColumn);
                             
                wsCapitalBudgeting.Calculate();
                
               wsCapitalBudgeting.Cells.AutoFitColumns();

                #region "Commented Code"

                //for (int o = 0; o < result.ProjectSummaryDatasVM.Count; o++)
                //{
                //    if (result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM != null)
                //    {
                //        for (int p = 0; p < result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM.Count; p++)
                //        {
                //            try
                //            {
                //                if (p == 0)
                //                {
                //                    var conditionCheck = result.ProjectSummaryDatasVM[o].LineItem;

                //                    string UnitValue = null;

                //                    if (result.ProjectSummaryDatasVM[o].UnitId == 1)
                //                    {
                //                        UnitValue = "$";
                //                    }
                //                    else if (result.ProjectSummaryDatasVM[o].UnitId == 2)
                //                    {
                //                        UnitValue = "$K";
                //                    }
                //                    else if (result.ProjectSummaryDatasVM[o].UnitId == 3)
                //                    {
                //                        UnitValue = "$M";
                //                    }
                //                    else if (result.ProjectSummaryDatasVM[o].UnitId == 4)
                //                    {
                //                        UnitValue = "$B";
                //                    }
                //                    else if (result.ProjectSummaryDatasVM[o].UnitId == 5)
                //                    {
                //                        UnitValue = "$T";
                //                    }

                //                    if (conditionCheck.Contains("Sales") || conditionCheck.Contains("COGS") ||
                //                    conditionCheck.Contains("Unlevered Net Income") || conditionCheck.Contains("Free Cash Flow") ||
                //                    conditionCheck.Contains("Gross Margin") || conditionCheck.Contains("Operating Income") ||
                //                     conditionCheck.Contains("Net Present Value"))
                //                    {
                //                        wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 0);
                //                        addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                //                    }
                //                    else
                //                    {
                //                        if (conditionCheck == "Discount Factor")
                //                        {
                //                            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + "$" + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 1);
                //                            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                //                        }
                //                        else
                //                        {
                //                            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 1);
                //                            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                //                        }

                //                    }
                //                }
                //               // wsCapitalBudgeting.Cells[(row_count + o), (p + 3)].Formula = "";
                //                wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM[p].Value), (row_count + o), (p + 3), wsCapitalBudgeting, 1);                              
                //                var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                //                addArray2D[o, p] = Convert.ToString(addr.ToString());
                //            }
                //            catch (IndexOutOfRangeException)
                //            {
                //                var strr = "";
                //                wsCapitalBudgeting = ReturnCellStyle(strr, (row_count + o), (p + 3), wsCapitalBudgeting, 1);
                //                var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                //                addArray2D[o, p] = Convert.ToString(addr.ToString());
                //            }
                //        }
                //    }
                //    else if (result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM == null)
                //    {
                //        int p = 0;
                //        var conditionCheck = result.ProjectSummaryDatasVM[o].LineItem;

                //        string UnitValue = null;

                //        if (result.ProjectSummaryDatasVM[o].UnitId == 1)
                //        {
                //            UnitValue = "$";
                //        }
                //        else if (result.ProjectSummaryDatasVM[o].UnitId == 2)
                //        {
                //            UnitValue = "$K";
                //        }
                //        else if (result.ProjectSummaryDatasVM[o].UnitId == 3)
                //        {
                //            UnitValue = "$M";
                //        }
                //        else if (result.ProjectSummaryDatasVM[o].UnitId == 4)
                //        {
                //            UnitValue = "$B";
                //        }
                //        else if (result.ProjectSummaryDatasVM[o].UnitId == 5)
                //        {
                //            UnitValue = "$T";
                //        }

                //        if (conditionCheck.Contains("Net Present Value") ||
                //                conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                //        {

                //            if (conditionCheck == "IRR (Internal Rate of Return) of Free Cash Flows")
                //            {
                //                wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + "%" + ")", (row_count + o), (1), wsCapitalBudgeting, 0);
                //                addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                //            }
                //            else
                //            {
                //                wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (1), wsCapitalBudgeting, 0);
                //                addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                //            }

                //        }
                //        else
                //        {
                //            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (1), wsCapitalBudgeting, 1);
                //            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                //        }

                //        wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(result.ProjectSummaryDatasVM[o].Value), (row_count + o), (3), wsCapitalBudgeting, 1);
                //        var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                //        addArray2D[o, p] = Convert.ToString(addr.ToString());
                //    }
                //}

                #endregion

                // Circular reference

                //ExcelPackage excelPackage = new ExcelPackage();
                //excelPackage.Workbook.Worksheets.Add("CapitalBudgeting", wsCapitalBudgeting);
                //package.Save();

                //string filePath = "D:\\LMS\\Sowfin\\backend\\Sowfin.API\\wwwroot\\capital_budgeting.xlsx";

                //LoadOptions LoadOptions = new LoadOptions();
                ////var objWB = new Aspose.Cells.Workbook(wsCapitalBudgeting + "CapitalBudgeting.xls", LoadOptions);
                //var objWB = new Aspose.Cells.Workbook(filePath, LoadOptions);
                //objWB.Settings.Iteration = true;
                //CalculationOptions copts = new CalculationOptions();
                //CircularMonitor cm = new CircularMonitor();
                //copts.CalculationMonitor = cm;
                //objWB.CalculateFormula(copts);

                ExcelPackage excelPackage = new ExcelPackage();
                excelPackage.Workbook.Worksheets.Add("CapitalBudgeting", wsCapitalBudgeting);
                //package.Save();
                ExcelPackage epOut = excelPackage;
                byte[] myStream = epOut.GetAsByteArray();
                var inputAsString = Convert.ToBase64String(myStream);
                formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);
                excelPackage.Dispose();
                return Ok(formattedCustomObject);
            }
        }

        private long GetValueUnit(List<ProjectsViewModel> result , int i)
        {
            long Value = 0;

            if (result[0].ProjectSummaryDatasVM[i].UnitId == 1)
            {
                Value = 1;
            }
            else if (result[0].ProjectSummaryDatasVM[i].UnitId == 2)
            {
                Value = 1000;
            }
            else if (result[0].ProjectSummaryDatasVM[i].UnitId == 3)
            {
                Value = 1000000;
            }
            else if (result[0].ProjectSummaryDatasVM[i].UnitId == 4)
            {
                Value = 1000000000;
            }
            else if (result[0].ProjectSummaryDatasVM[i].UnitId == 5)
            {
                Value = 1000000000000;
            }

            return Value;
        }

        //public void GetValueUnit1(int OutputUnitId, out long InputUnitId1 , out long InputUnitId2, out long InputUnitId3, out long InputUnitId4,
        //                                             long InputUnitId11,  long InputUnitId22,  long InputUnitId33,  long InputUnitId44)
        //{
        //    long Value = 0;
        //    long Input1 = 1;
        //    long Input2 = 1;
        //    long Input3 = 1;
        //    long Input4 = 1;

        //    if (OutputUnitId != InputUnitId11)
        //    {
        //        if (OutputUnitId == 1)
        //        {
        //            Input1 = 1;
        //        }
        //        else if (OutputUnitId == 2)
        //        {
        //            Input1 = 1000;
        //        }
        //        else if (OutputUnitId == 3)
        //        {
        //            Input1 = 1000000;
        //        }
        //        else if (OutputUnitId == 4)
        //        {
        //            Input1 = 1000000000;
        //        }
        //        else if (OutputUnitId == 5)
        //        {
        //            Input1 = 1000000000000;
        //        }
        //    }

        //    if (OutputUnitId != InputUnitId22)
        //    {
        //        if (OutputUnitId == 1)
        //        {
        //            Input2 = 1;
        //        }
        //        else if (OutputUnitId == 2)
        //        {
        //            Input2 = 1000;
        //        }
        //        else if (OutputUnitId == 3)
        //        {
        //            Input2 = 1000000;
        //        }
        //        else if (OutputUnitId == 4)
        //        {
        //            Input2 = 1000000000;
        //        }
        //        else if (OutputUnitId == 5)
        //        {
        //            Input2 = 1000000000000;
        //        }
        //    }

        //    if (OutputUnitId != InputUnitId33)
        //    {
        //        if (OutputUnitId == 1)
        //        {
        //            Input3 = 1;
        //        }
        //        else if (OutputUnitId == 2)
        //        {
        //            Input3 = 1000;
        //        }
        //        else if (OutputUnitId == 3)
        //        {
        //            Input3 = 1000000;
        //        }
        //        else if (OutputUnitId == 4)
        //        {
        //            Input3 = 1000000000;
        //        }
        //        else if (OutputUnitId == 5)
        //        {
        //            Input3 = 1000000000000;
        //        }
        //    }

        //    if (OutputUnitId != InputUnitId44)
        //    {
        //        if (OutputUnitId == 1)
        //        {
        //            Input4 = 1;
        //        }
        //        else if (OutputUnitId == 2)
        //        {
        //            Input4 = 1000;
        //        }
        //        else if (OutputUnitId == 3)
        //        {
        //            Input4 = 1000000;
        //        }
        //        else if (OutputUnitId == 4)
        //        {
        //            Input4 = 1000000000;
        //        }
        //        else if (OutputUnitId == 5)
        //        {
        //            Input4 = 1000000000000;
        //        }
        //    }

        //    InputUnitId1 = Input1;
        //    InputUnitId2 = Input2;
        //    InputUnitId3 = Input3;
        //    InputUnitId4 = Input4;

        //   // return Value;
        //}

        private long GetConvertedUnitValue(int OutputUnitId)
        {
            long Value = 1;

                if (OutputUnitId == 1)
                {
                Value = 1;
                }
                else if (OutputUnitId == 2)
                {
                Value = 1000;
                }
                else if (OutputUnitId == 3)
                {
                Value = 1000000;
                }
                else if (OutputUnitId == 4)
                {
                Value = 1000000000;
                }
                else if (OutputUnitId == 5)
                {
                Value = 1000000000000;
                }

             return Value;
        }
        
         

        //  [HttpPost]


        private ExcelWorksheet InsertInputToExcel(List<ProjectsViewModel> projectVM, ExcelWorksheet wsCapitalBudgeting, int OutPutRowcount,
                                                  List<List<List<string>>> list, List<string> salesResult, List<string> cogsResult,
                                                   List<string> GrossMarginResult, List<List<List<string>>> OtherFixedCostlist,
                                                   List<string> OperatingIncomeResult, List<string> FreeCashFlowResult,
                                                   List<string[,]> addArray2DList, string marginalTaxAddColumn, string discountRateAddColumn,
                                                   List<List<List<string>>> listOfCellsNWC, List<List<List<string>>> listOfCellsCapex,
                                                   List<List<List<string>>> listOfCellsCapex1, List<List<List<string>>> listOfCellsCapex2, int CountOfCapexDepreciation ,
                                                   List<List<List<string>>> listOfFixedSchedule, List<string> yearCoutList,
                                                   string WeightedAverageCostOfCapitalAddColumn, string DVRatioAddColumn, string CostOfDebtAddColumn, 
                                                   string CostOfEquityAddColumn,string UnleveredCostOfCapitalAddColumn, string IntetrestCoverageRatioAddColumn,
                                                   string ProjectWACCAddColumn)
        {

          

            List<string> salesCellList = new List<string>();
            List<string> nwcCellList = new List<string>();
            List<string> DebtCapacityCellList = new List<string>();
         
            string sum = null;
            int Unleveredcount = 0; // For Valuation Technique 5            
            int NPVcount = 0; // For Valuation Technique 5
            int Unleveredcount1 = 0; // For Valuation Technique 5
            int NPVcount1 = 0; // For Valuation Technique 5

                if (projectVM == null || projectVM.Count == 0 || projectVM[0] == null || projectVM[0].ProjectSummaryDatasVM == null)
                {
                    // Return the original worksheet or null if projectVM or its properties are not properly initialized
                    return wsCapitalBudgeting; // or return null;
                }


            for (int o = 0; o < projectVM[0].ProjectSummaryDatasVM.Count; o++)
            {
                if (projectVM[0].ProjectSummaryDatasVM[o].ProjectOutputValuesVM != null)
                {
                    for (int p = 0; p < projectVM[0].ProjectSummaryDatasVM[o].ProjectOutputValuesVM.Count; p++)
                    {
                        var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;

                        if (conditionCheck.Contains("Sales"))
                        {
                            long SalesConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + salesResult[p] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + salesResult[p] + ")" + "/" + SalesConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            salesCellList.Add(addArray2DList[0][o, p + 1]);
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("COGS"))
                        {
                            long COGSConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + cogsResult[p] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + cogsResult[p] + ")" + "/" + COGSConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("Gross Margin"))
                        {
                            long GrossMarginConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + GrossMarginResult[p] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + GrossMarginResult[p] + ")" + "/" + GrossMarginConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("SG&A"))
                        {
                            long SGAConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + OtherFixedCostlist[0][0][p] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + OtherFixedCostlist[0][0][p] + ")" + "/" + SGAConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("R&D"))
                        {
                            long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + OtherFixedCostlist[0][1][p] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + OtherFixedCostlist[0][1][p] + ")" + "/" + ConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck == "Depreciation")
                        {
                            //DepreciationUnitId = (int)projectVM[0].ProjectSummaryDatasVM[o].UnitId;
                        }
                        else if (conditionCheck.Contains("Operating Income"))
                        {
                            //GetValueUnit1((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId, out GrossMarginUnitId, out SGAUnitId, out RDUnitId,
                            //               out DepreciationUnitId , GrossMarginUnitId , SGAUnitId , RDUnitId , DepreciationUnitId);

                            string GrossMarginCellValue = "";
                            string SGACellValue = "";
                            string RDCellValue = "";
                            string DepreciationCellValue = "";

                            long OperatingIncomeConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);

                            //if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o-4].UnitId)
                            //{
                                ///long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                                //GrossMarginCellValue = "(" + addArray2DList[0][o - 4, p + 1] + "/" + ConvertedValue + ")";
                                long GrossMarginConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-4].UnitId);
                                GrossMarginCellValue = "(" + addArray2DList[0][o - 4, p + 1] + "*" + GrossMarginConvertedValue + ")";
                            //}
                            //else
                            //{
                              //  GrossMarginCellValue = "(" + addArray2DList[0][o - 4, p + 1] + ")";
                            //}

                            //if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o-3].UnitId)
                            //{
                                //long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                                //SGACellValue = "(" + addArray2DList[0][o - 3, p + 1] + "/" + ConvertedValue + ")";
                                long SGAConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-3].UnitId);
                                SGACellValue = "(" + addArray2DList[0][o - 3, p + 1] + "*" + SGAConvertedValue + ")";
                            //}
                            //else
                            //{
                             //   SGACellValue = "(" + addArray2DList[0][o - 3, p + 1] + ")";
                            //}

                            //if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o-2].UnitId)
                            //{
                                //long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                                // RDCellValue = "(" + addArray2DList[0][o - 2, p + 1] + "/" + ConvertedValue + ")";
                                long RDConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-2].UnitId);
                                RDCellValue = "(" + addArray2DList[0][o - 2, p + 1] + "*" + RDConvertedValue + ")";
                            //}
                            //else
                            //{
                            //    RDCellValue = "(" + addArray2DList[0][o - 2, p + 1] + ")";
                           // }

                           // if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o-1].UnitId)
                           // {
                                //long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                                //DepreciationCellValue = "(" + addArray2DList[0][o - 1, p + 1] + "/" + ConvertedValue + ")";
                                long DepreciationConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-1].UnitId);
                                DepreciationCellValue = "(" + addArray2DList[0][o - 1, p + 1] + "*" + DepreciationConvertedValue + ")";
                           // }
                           // else
                            //{
                            //    DepreciationCellValue = "(" + addArray2DList[0][o - 1, p + 1] + ")";
                            //}

                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = GrossMarginCellValue + "-" + SGACellValue + "-" + RDCellValue + "-" + DepreciationCellValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + GrossMarginCellValue + "-" + SGACellValue + "-" +
                                                                                                    RDCellValue + "-" + DepreciationCellValue + ")" + "/" + OperatingIncomeConvertedValue;

                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + addArray2DList[0][o - 4, p + 1] + "/" + GrossMarginUnitId + ")" +
                            //                                                                  "(" + addArray2DList[0][o - 3, p + 1] + "/" + SGAUnitId + ")" +
                            //                                                                  "(" + addArray2DList[0][o - 2, p + 1] + "/" + RDUnitId + ")" +
                            //                                                                  "(" + addArray2DList[0][o - 1, p + 1] + "/" + DepreciationUnitId + ")";

                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + OperatingIncomeResult[p] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("Income Tax"))
                        {
                            long IncomeTaxConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);                           
                            long OperatingIncomeConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o - 1].UnitId);
                            
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = addArray2DList[0][o - 1, p + 1] + "*" + "(" + marginalTaxAddColumn + "/" + "100" + ")";
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = addArray2DList[0][o - 1, p + 1] + "*" + marginalTaxAddColumn;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + "(" + addArray2DList[0][o - 1, p + 1] + "*" + OperatingIncomeConvertedValue + ")" 
                                                                                               + "*" + marginalTaxAddColumn + ")" + "/" + IncomeTaxConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck == "Unlevered Net Income")
                        {
                            long UnleveredNetIncomeConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            long OperatingIncomeConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-2].UnitId);
                            long IncomeTaxConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-1].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + addArray2DList[0][o - 2, p + 1] + "-" + addArray2DList[0][o - 1, p + 1] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + "(" + addArray2DList[0][o - 2, p + 1] + "*" + OperatingIncomeConvertedValue + ")" +  "-" +
                                                                                                    "(" + addArray2DList[0][o - 1, p + 1] + "*" + IncomeTaxConvertedValue + ")" + ")" + "/" + UnleveredNetIncomeConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("NWC (Net Working Capital)"))
                        {
                            long NWCConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            long SalesConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-9].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = salesCellList[p] + "*" + (listOfCellsNWC[0][0][p] + "/" + 100);
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + "(" + salesCellList[p] + "*" + SalesConvertedValue + ")"  + "*" + "(" + (listOfCellsNWC[0][0][p] + "/" + 100) + ")" + ")" + "/" + NWCConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            nwcCellList.Add(addArray2DList[0][o, p + 1]);
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("Plus: Depreciation"))
                        {
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = addArray2DList[0][o - 5, p + 1];
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                        }
                        else if (conditionCheck.Contains("Less: Capital Expenditures"))
                        {                          
                            long CapitalExpendituresConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + listOfCellsCapex[0][0][p] + ")";
                            if (CountOfCapexDepreciation == 1)
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + listOfCellsCapex[0][0][p] + ")" + "/" + CapitalExpendituresConvertedValue;
                            }
                            else if (CountOfCapexDepreciation == 2)
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + "(" + listOfCellsCapex[0][0][p] + ")" + "+" + "(" + listOfCellsCapex1[0][0][p] + ")" + ")" + "/" + CapitalExpendituresConvertedValue;
                            }
                            else if (CountOfCapexDepreciation == 3)
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + "(" + listOfCellsCapex[0][0][p] + ")" + "+" + "(" + listOfCellsCapex1[0][0][p] + ")" + "+" + "(" + listOfCellsCapex2[0][0][p] + ")" + ")" + "/" + CapitalExpendituresConvertedValue;
                            }
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + listOfCellsCapex[0][0][p] + "/" + CapitalExpendituresConvertedValue + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if (conditionCheck.Contains("Less: Increases in NWC"))
                        {
                            long IncreasesinNWCConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            long NetWorkingCapitalConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-3].UnitId);
                            if (p == 0)
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = nwcCellList[p];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + nwcCellList[p] + "*" + NetWorkingCapitalConvertedValue + ")" + "/" + IncreasesinNWCConvertedValue;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            }
                            else
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + nwcCellList[p] + "-" + nwcCellList[p - 1] + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + "(" + nwcCellList[p] + "*" + NetWorkingCapitalConvertedValue + ")" + "-" +
                                                                                                  "(" + nwcCellList[p - 1] + "*" + NetWorkingCapitalConvertedValue + ")" + ")" + "/" + IncreasesinNWCConvertedValue;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            }                            
                        }
                        else if (conditionCheck == "Free Cash Flow")
                        {
                            string UnleveredNetIncomeCellValue = "";
                            string PlusDepreciationCellValue = "";
                            string LessCapitalExpendituresCellValue = "";
                            string LessIncreasesNWCCellValue = "";

                            #region "Commented"

                            //if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o - 5].UnitId)
                            //{
                            //    long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //    UnleveredNetIncomeCellValue = "(" + addArray2DList[0][o - 5, p + 1] + "/" + ConvertedValue + ")";
                            //}
                            //else
                            //{
                            //    UnleveredNetIncomeCellValue = "(" + addArray2DList[0][o - 5, p + 1] + ")";
                            //}

                            //if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o - 3].UnitId)
                            //{
                            //    long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //    PlusDepreciationCellValue = "(" + addArray2DList[0][o - 3, p + 1] + "/" + ConvertedValue + ")";
                            //}
                            //else
                            //{
                            //    PlusDepreciationCellValue = "(" + addArray2DList[0][o - 3, p + 1] + ")";
                            //}

                            //if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o - 2].UnitId)
                            //{
                            //    long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //    LessCapitalExpendituresCellValue = "(" + addArray2DList[0][o - 2, p + 1] + "/" + ConvertedValue + ")";
                            //}
                            //else
                            //{
                            //    LessCapitalExpendituresCellValue = "(" + addArray2DList[0][o - 2, p + 1] + ")";
                            //}

                            //if ((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId != (int)projectVM[0].ProjectSummaryDatasVM[o - 1].UnitId)
                            //{
                            //    long ConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                            //    LessIncreasesNWCCellValue = "(" + addArray2DList[0][o - 1, p + 1] + "/" + ConvertedValue + ")";
                            //}
                            //else
                            //{
                            //    LessIncreasesNWCCellValue = "(" + addArray2DList[0][o - 1, p + 1] + ")";
                            //}

                            #endregion

                                long FreeCashFlowConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);
                           
                                long UnleveredNetIncomeConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-5].UnitId);
                                UnleveredNetIncomeCellValue = "(" + addArray2DList[0][o - 5, p + 1] + "*" + UnleveredNetIncomeConvertedValue + ")";

                                long PlusDepreciationConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-3].UnitId);
                                PlusDepreciationCellValue = "(" + addArray2DList[0][o - 3, p + 1] + "*" + PlusDepreciationConvertedValue + ")";

                                long LessCapitalExpendituresConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-2].UnitId);
                                LessCapitalExpendituresCellValue = "(" + addArray2DList[0][o - 2, p + 1] + "*" + LessCapitalExpendituresConvertedValue + ")";

                                long LessIncreasesNWCConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o-1].UnitId);
                                LessIncreasesNWCCellValue = "(" + addArray2DList[0][o - 1, p + 1] + "*" + LessIncreasesNWCConvertedValue + ")";

                            // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = UnleveredNetIncomeCellValue + "+" + PlusDepreciationCellValue + "-" + LessCapitalExpendituresCellValue + "-" + LessIncreasesNWCCellValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + UnleveredNetIncomeCellValue + "+" + PlusDepreciationCellValue + "-" + LessCapitalExpendituresCellValue + "-" + LessIncreasesNWCCellValue + ")" + "/" + FreeCashFlowConvertedValue;

                            // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + FreeCashFlowResult[p] + ")";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else if ((conditionCheck == "Free Cash Flow to Equity") && (projectVM[0].ValuationTechniqueId == 3))
                        {
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = addArray2DList[0][o - 2, p + 1] + "+" +
                            //   addArray2DList[0][o - 9, p + 1] + "-" + addArray2DList[0][o - 8, p + 1] + "-" + addArray2DList[0][o - 7, p + 1]
                            //   + "+" + addArray2DList[0][o - 1, p + 1];
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;

                            string NetIncomeCellValue = "";
                            string PlusDepreciationCellValue = "";
                            string LessCapitalExpendituresCellValue = "";
                            string LessIncreasesNWCCellValue = "";
                            string NetBorrowingCellValue = "";

                            long FreeCashFlowEquityConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o].UnitId);

                            long NetIncomeConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o - 2].UnitId);
                            NetIncomeCellValue = "(" + addArray2DList[0][o - 2, p + 1] + "*" + NetIncomeConvertedValue + ")";

                            long PlusDepreciationConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o - 9].UnitId);
                            PlusDepreciationCellValue = "(" + addArray2DList[0][o - 9, p + 1] + "*" + PlusDepreciationConvertedValue + ")";

                            long LessCapitalExpendituresConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o - 8].UnitId);
                            LessCapitalExpendituresCellValue = "(" + addArray2DList[0][o - 8, p + 1] + "*" + LessCapitalExpendituresConvertedValue + ")";

                            long LessIncreasesNWCConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o - 7].UnitId);
                            LessIncreasesNWCCellValue = "(" + addArray2DList[0][o - 7, p + 1] + "*" + LessIncreasesNWCConvertedValue + ")";

                            long NetBorrowingConvertedValue = GetConvertedUnitValue((int)projectVM[0].ProjectSummaryDatasVM[o - 1].UnitId);
                            NetBorrowingCellValue = "(" + addArray2DList[0][o - 1, p + 1] + "*" + NetBorrowingConvertedValue + ")";

                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + NetIncomeCellValue + "+" + PlusDepreciationCellValue + "-" + LessCapitalExpendituresCellValue + "-" + LessIncreasesNWCCellValue + "+" + NetBorrowingCellValue + ")" + "/" + FreeCashFlowEquityConvertedValue;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }

                        if (projectVM[0].ValuationTechniqueId == 1)
                        {                           
                            if (conditionCheck.Contains("Levered Value (VL)"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + discountRateAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + discountRateAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Discount Factor"))
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + ((1 + "+" + (discountRateAddColumn + "/" + "100")) + ")" + "^" + yearCoutList[p]);
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + ((1 + "+" + (discountRateAddColumn)) + ")" + "^" + yearCoutList[p]);
                                 wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            }
                            else if (conditionCheck.Contains("Discounted Cash Flow"))
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = addArray2DList[0][o - 3, p + 1] + "/" + addArray2DList[0][o - 1, p + 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 2)
                        {
                            if (conditionCheck.Contains("Unlevered Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {                                       
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Debt Capacity"))
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + DVRatioAddColumn + ")" + "*" + addArray2DList[0][o + 4, p + 1];
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                ////DebtCapacityCellList.Add(addArray2DList[0][o, p + 1]);
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;

                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + DVRatioAddColumn + ")" + "*" + addArray2DList[0][o + 4, i];
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }

                            }
                            else if (conditionCheck == "Interest Paid @ rD")
                            {
                                //for (int i = 0; i < (int)projectVM[0].NoOfYears; i++)
                                //{
                                //    if (i == 0)
                                //    {
                                //        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = "";
                                //    }
                                //    else
                                //    {
                                //        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i] + "*" + "(" + CostOfDebtAddColumn + "/" + "100" + ")";
                                //        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i] + "*" + "(" + CostOfDebtAddColumn + ")";
                                //        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Numberformat.Format = "$#,##0.00";
                                //        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                //    }
                                //}

                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        //IP = (DVRatioValue * CostofDebtValue * unLeaveredValue) / (1 - ((DVRatioValue * CostofDebtValue * MarginalTaxValue) / (1 + UCCValue)));
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + DVRatioAddColumn + "*" + CostOfDebtAddColumn + "*" + addArray2DList[0][o - 2, i - 1] + ")" + "/" +
                                                                                                          "(" + 1 + "-" + "((" + DVRatioAddColumn + "*" + CostOfDebtAddColumn + "*" + marginalTaxAddColumn + ")" + "/" +
                                                                                                          "(" + 1 + "+" + UnleveredCostOfCapitalAddColumn + ")))";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                    else
                                    {
                                        if (i != 1)
                                        {
                                            //IP = ((LastyearLeaveredValue != null ? Convert.ToDouble(LastyearLeaveredValue.Value) : 0) *(1+UCCValue)+TSV)/(((1+UCCValue)/(DVRatioValue*CostofDebtValue))-MarginalTaxValue);
                                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "((" + addArray2DList[0][o - 2, i - 1] + "*(1+" + UnleveredCostOfCapitalAddColumn + ")" + "+" +
                                                                                                                     addArray2DList[0][o + 2, i + 1] + ")/((1+" + UnleveredCostOfCapitalAddColumn + ")/(" +
                                                                                                                     DVRatioAddColumn + "*" + CostOfDebtAddColumn + "))-" + marginalTaxAddColumn + ")";
                                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                        }
                                    }
                                }

                            }
                            else if (conditionCheck == "Interest Tax Shield @ Tc")
                            {
                                for (int i = 0; i < (int)projectVM[0].NoOfYears; i++)
                                {
                                    if (i == 0)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i + 1] + "*" + "(" + marginalTaxAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }                              
                            }
                            else if (conditionCheck.Contains("Tax Shield Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Levered Value (VL = VU + T)"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = addArray2DList[0][o - 5, i] + "+" + addArray2DList[0][o - 1, i];
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 3)
                        {
                            if (conditionCheck.Contains("Levered Value (VL)"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + WeightedAverageCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + WeightedAverageCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                           else if (conditionCheck.Contains("Debt Capacity"))
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + DVRatioAddColumn + "/" + "100" + ")" + "*" + addArray2DList[0][o - 1, p + 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + DVRatioAddColumn + ")" + "*" + addArray2DList[0][o - 1, p + 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                DebtCapacityCellList.Add(addArray2DList[0][o, p + 1]);
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            }
                           else if (conditionCheck.Contains("Interest Expense"))
                            {
                                for (int i = 0; i < (int)projectVM[0].NoOfYears; i++)
                                {
                                    if (i == 0)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i] + "*" + "(" + CostOfDebtAddColumn + "/" + "100" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i] + "*" + "(" + CostOfDebtAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck == "Net Income")
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + addArray2DList[0][o - 11, p + 1] + "-" + addArray2DList[0][o - 1, p + 1] + ")"
                                //                                                                   + "*" + "(1 -" + "(" + marginalTaxAddColumn + "/" + "100" + ")" + ")" ;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + addArray2DList[0][o - 11, p + 1] + "-" + addArray2DList[0][o - 1, p + 1] + ")"
                                                                                                   + "*" + "(1 -" + "(" + marginalTaxAddColumn + ")" + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Plus: Net Borrowing"))
                            {
                                if (p == 0)
                                {
                                    wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = DebtCapacityCellList[p];
                                    wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                }
                                else
                                {
                                    wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + DebtCapacityCellList[p] + "-" + DebtCapacityCellList[p - 1] + ")";
                                    wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                    wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                                }
                            }
                            else if (conditionCheck.Contains("Discount Factor"))
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + ((1 + "+" + (CostOfEquityAddColumn + "/" + "100")) + ")" + "^" + yearCoutList[p]);
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + ((1 + "+" + (CostOfEquityAddColumn)) + ")" + "^" + yearCoutList[p]);
                                 wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Discounted Cash Flow"))
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = addArray2DList[0][o - 2, p + 1] + "/" + addArray2DList[0][o - 1, p + 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 4)
                        {
                            if (conditionCheck.Contains("Levered Value (VL)"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + ProjectWACCAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + ProjectWACCAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Discount Factor"))
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + ((1 + "+" + (ProjectWACCAddColumn + "/" + "100")) + ")" + "^" + yearCoutList[p]);
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + ((1 + "+" + (ProjectWACCAddColumn)) + ")" + "^" + yearCoutList[p]);
                                 wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Discounted Cash Flow"))
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = addArray2DList[0][o - 3, p + 1] + "/" + addArray2DList[0][o - 1, p + 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 5)
                        {
                            if (conditionCheck.Contains("Unlevered Value @ rU") && Unleveredcount == 0)
                            {
                                Unleveredcount = 1;
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Unlevered Value @ rU") && Unleveredcount1 == 1)
                            {
                                Unleveredcount1 = 2;
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 5, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 5, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck == "Interest Paid @ rD")
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + IntetrestCoverageRatioAddColumn + "/" + "100" + ")" + "*" +
                                //                                                              addArray2DList[0][o - 6, p + 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = "(" + IntetrestCoverageRatioAddColumn + ")" + "*" +
                                                                                              addArray2DList[0][o - 6, p + 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                            }
                            else if (conditionCheck.Contains("Debt Capacity"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = addArray2DList[0][o - 1, i + 1] + "/" + "(" + CostOfDebtAddColumn + "/" + "100" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = addArray2DList[0][o - 1, i + 1] + "/" + "(" + CostOfDebtAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Interest Tax Shield @ Tc"))
                            {
                                for (int i = 0; i < (int)projectVM[0].NoOfYears; i++)
                                {
                                    if (i == 0)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 2, i + 1] + "*" + "(" + marginalTaxAddColumn + "/" + "100" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 2, i + 1] + "*" + "(" + marginalTaxAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Tax Shield Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Levered Value (VL = VU + T)"))
                            {
                                NPVcount1 = 1;
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = addArray2DList[0][o - 5, i] + "+" + addArray2DList[0][o - 1, i];
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }

                        }
                        else if (projectVM[0].ValuationTechniqueId == 6)
                        {
                            if (conditionCheck.Contains("Unlevered Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Debt Capacity"))
                            {
                               wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Formula = listOfFixedSchedule[0][0][p];
                               wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Numberformat.Format = "$#,##0.00";
                               wsCapitalBudgeting.Cells[(OutPutRowcount + o), (p + 3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck == "Interest Paid @ rD")
                            {
                                for (int i = 0; i < (int)projectVM[0].NoOfYears; i++)
                                {
                                    if (i == 0)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i] + "*" + "(" + CostOfDebtAddColumn + "/" + "100" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i] + "*" + "(" + CostOfDebtAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }
                            }                            
                            else if (conditionCheck.Contains("Interest Tax Shield @ Tc"))
                            {
                                for (int i = 0; i < (int)projectVM[0].NoOfYears; i++)
                                {
                                    if (i == 0)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i + 1] + "*" + "(" + marginalTaxAddColumn + "/" + "100" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i + 1] + "*" + "(" + marginalTaxAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Tax Shield Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + CostOfDebtAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + CostOfDebtAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Levered Value (VL = VU + T)"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = addArray2DList[0][o - 5, i] + "+" + addArray2DList[0][o - 1, i];
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }

                        }
                        else if (projectVM[0].ValuationTechniqueId == 7)
                        {
                            if (conditionCheck == ("Levered Value @ rWACC"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + WeightedAverageCostOfCapitalAddColumn + ")" + ")";
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                           else if (conditionCheck.Contains("Unlevered Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 3, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Debt Capacity"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + DVRatioAddColumn + ")" + "*" + addArray2DList[0][o + 5, i];
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }

                            }
                            else if (conditionCheck == "Interest Paid @ rD")
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        //IP = (DVRatioValue * CostofDebtValue * unLeaveredValue) / (1 - ((DVRatioValue * CostofDebtValue * MarginalTaxValue) / (1 + UCCValue)));
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + DVRatioAddColumn + "*" + CostOfDebtAddColumn + "*" + addArray2DList[0][o - 2, i - 1] + ")" + "/" +
                                                                                                          "(" + 1 + "-" + "((" + DVRatioAddColumn + "*" + CostOfDebtAddColumn + "*" + marginalTaxAddColumn + ")" + "/" +
                                                                                                          "(" + 1 + "+" + UnleveredCostOfCapitalAddColumn + ")))";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                    else
                                    {
                                        if (i != 1)
                                        {
                                            //IP = ((LastyearLeaveredValue != null ? Convert.ToDouble(LastyearLeaveredValue.Value) : 0) *(1+UCCValue)+TSV)/(((1+UCCValue)/(DVRatioValue*CostofDebtValue))-MarginalTaxValue);
                                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + addArray2DList[0][o - 2, i - 1] + "*(1+" + UnleveredCostOfCapitalAddColumn + ")" + "+" +
                                                                                                                     addArray2DList[0][o + 2, i + 1] + ")/(((1+" + UnleveredCostOfCapitalAddColumn + ")/(" +
                                                                                                                     DVRatioAddColumn + "*" + CostOfDebtAddColumn + "))-" + marginalTaxAddColumn + ")";
                                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                            wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                        }
                                    }
                                }

                            }
                            else if (conditionCheck == "Interest Tax Shield @ Tc")
                            {
                                for (int i = 0; i < (int)projectVM[0].NoOfYears; i++)
                                {
                                    if (i == 0)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Formula = addArray2DList[0][o - 1, i + 1] + "*" + "(" + marginalTaxAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 3)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Tax Shield Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+"  + UnleveredCostOfCapitalAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck.Contains("Tax Shield Value Adjusted for Predetermined Debt"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = addArray2DList[0][o - 1, i + 1] + "*(1+" + UnleveredCostOfCapitalAddColumn + ")/(1+" + CostOfDebtAddColumn + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                            else if (conditionCheck == ("Levered Value (VL = VU + T)"))
                            {
                                NPVcount1 = 1;
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = addArray2DList[0][o - 6, i] + "+" + addArray2DList[0][o - 2, i];
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 8)
                        {
                            if (conditionCheck.Contains("Unlevered Value @ rU"))
                            {
                                for (int i = (int)projectVM[0].NoOfYears; i > 0; i--)
                                {
                                    if (i == projectVM[0].NoOfYears)
                                    {
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "";
                                    }
                                    else
                                    {
                                        //wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                        //    "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Formula = "(" + (addArray2DList[0][o - 1, i + 1] + "+" + addArray2DList[0][o, i + 1]) + ")" +
                                            "/" + "(" + 1 + "+" + "(" + UnleveredCostOfCapitalAddColumn + ")" + ")";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Numberformat.Format = "$#,##0.00";
                                        wsCapitalBudgeting.Cells[(OutPutRowcount + o), (i + 2)].Style.Font.Size = 12;
                                    }
                                }
                            }

                        }
                        
                    }                   
                }
                else if (projectVM[0].ProjectSummaryDatasVM[o].ProjectOutputValuesVM == null)
                {
                    for (int p = 0; p < (int)projectVM[0].NoOfYears; p++)
                    {
                        if (projectVM[0].ValuationTechniqueId == 1)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;

                            if (conditionCheck.Contains("Net Present Value"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;

                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, p + 1];
                                }
                                else if (p == CountNoOfYears)
                                {
                                    //sum = sum + "+" + addArray2DList[0][o - 1, p + 1];
                                    sum = "SUM" + "(" + sum + ":" + addArray2DList[0][o - 1, p + 1] + ")";
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Font.Size = 12;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            }
                            else if (conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;

                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 5, p + 1];
                                }
                                //else if (p == 9)
                                else if (p == CountNoOfYears)
                                {
                                    //sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 5, p + 1] + ")" + "*" + "100";
                                    sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 5, p + 1] + ")";
                                }
                                // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "0.0%";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Font.Size = 12;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 2)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;

                            if (conditionCheck.Contains("Net Present Value") && NPVcount == 0)
                            {
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, 1] + "+" + addArray2DList[0][o - 7, 1];
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Font.Size = 12;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            }
                            else if (conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 8, p + 1];
                                }
                                //else if (p == 9)
                                else if (p == CountNoOfYears)
                                {
                                    //sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 8, p + 1] + ")" + "*" + "100";
                                    sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 8, p + 1] + ")";
                                }
                                // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "0.0%";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Font.Size = 12;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 3)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;

                            if (conditionCheck.Contains("Net Present Value"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;

                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, p + 1];
                                }
                                else if (p == CountNoOfYears)
                                {
                                    //sum = sum + "+" + addArray2DList[0][o - 1, p + 1];
                                    sum = "SUM" + "(" + sum + ":" + addArray2DList[0][o - 1, p + 1] + ")";
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = sum ;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;

                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 4, p + 1];
                                }
                                //else if (p == 9)
                                else if (p == CountNoOfYears)
                                {
                                    //sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 4, p + 1] + ")" + "*" + "100";
                                    sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 4, p + 1] + ")";
                                }
                                // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "0.0%";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 4)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;

                            if (conditionCheck.Contains("Net Present Value"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;

                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, p + 1];
                                }
                                else if (p == CountNoOfYears)
                                {
                                    // sum = sum + "+" + addArray2DList[0][o - 1, p + 1];
                                    sum = "SUM" + "(" + sum + ":" + addArray2DList[0][o - 1, p + 1] + ")";
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula =  sum ;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;

                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 5, p + 1];
                                }
                                //else if (p == 9)
                                else if (p == CountNoOfYears)
                                {
                                    //sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 5, p + 1] + ")" + "*" + "100";
                                    sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 5, p + 1] + ")";
                                }
                                // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "0.0%";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 5)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;
                            
                            if (conditionCheck.Contains("PV of Interest Tax Shield"))
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = "(" + marginalTaxAddColumn + "/" + "100" + ")" + "*" +
                                //                                                            "(" + IntetrestCoverageRatioAddColumn + "/" + "100" + ")" + "*" +
                                //                                                              addArray2DList[0][o - 1, 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = "(" + marginalTaxAddColumn + ")" + "*" +
                                                                                            "(" + IntetrestCoverageRatioAddColumn + ")" + "*" +
                                                                                              addArray2DList[0][o - 1, 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Levered Value (VL = VU + T)"))
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = addArray2DList[0][o - 2, 1] + "+" + addArray2DList[0][o - 1, 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Net Present Value") && NPVcount == 0)
                            {
                                NPVcount = 1;
                                Unleveredcount1 = 1;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, 1] + "+" + addArray2DList[0][o - 4, 1];
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Net Present Value") && NPVcount1 == 1)
                            {
                                NPVcount1 = 2;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, 1] + "+" + addArray2DList[0][o - 11, 1];
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 12, p + 1];
                                }
                                //else if (p == 9)
                                else if (p == CountNoOfYears)
                                {
                                    //sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 12, p + 1] + ")" + "*" + "100";
                                    sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 12, p + 1] + ")";
                                }
                                // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "0.0%";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 6)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;

                            if (conditionCheck.Contains("Net Present Value") && NPVcount == 0)
                            {
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, 1] + "+" + addArray2DList[0][o - 7, 1];
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }                            
                            else if (conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 8, p + 1];
                                }
                                //else if (p == 9)
                                else if (p == CountNoOfYears)
                                {
                                    //sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 8, p + 1] + ")" + "*" + "100";
                                    sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 8, p + 1] + ")";
                                }
                                // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "0.0%";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 7)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;

                            if (conditionCheck.Contains("Net Present Value") && NPVcount == 0)
                            {
                                NPVcount = 1;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, 1] + "+" + addArray2DList[0][o - 2, 1];
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Net Present Value") && NPVcount1 == 1)
                            {
                                NPVcount1 = 2;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, 1] + "+" + addArray2DList[0][o - 10, 1];
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {
                                int CountNoOfYears = (int)projectVM[0].NoOfYears - 1;
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 11, p + 1];
                                }
                                //else if (p == 9)
                                else if (p == CountNoOfYears)
                                {
                                    //sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 8, p + 1] + ")" + "*" + "100";
                                    sum = "IRR" + "(" + sum + ":" + addArray2DList[0][o - 11, p + 1] + ")";
                                }
                                // wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = sum;
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "0.0%";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                        }
                        else if (projectVM[0].ValuationTechniqueId == 8)
                        {
                            //int p = 0;
                            var conditionCheck = projectVM[0].ProjectSummaryDatasVM[o].LineItem;
                            
                            if (conditionCheck.Contains("PV of Interest Tax Shield"))
                            {
                                //wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = "(" + marginalTaxAddColumn + "/" + "100" + ")" + "*" +
                                //                                                                  "(" + IntetrestCoverageRatioAddColumn + "/" + "100" + ")" + "*" +
                                //                                                                  "((1+" + "(" + UnleveredCostOfCapitalAddColumn + "/" + "100" + ")" +
                                //                                                                  ")" + "/(1+" + "(" + CostOfDebtAddColumn + "/" + "100" + ")" +
                                //                                                                  "))*" + addArray2DList[0][o - 1, 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = "(" + marginalTaxAddColumn + ")" + "*" +
                                                                                                  "(" + IntetrestCoverageRatioAddColumn + ")" + "*" +
                                                                                                  "((1+" + "(" + UnleveredCostOfCapitalAddColumn + ")" +
                                                                                                  ")" + "/(1+" + "(" + CostOfDebtAddColumn + ")" +
                                                                                                  "))*" + addArray2DList[0][o - 1, 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Levered Value (VL = VU + T)"))
                            {
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Formula = addArray2DList[0][o - 2, 1] + "+" + addArray2DList[0][o - 1, 1];
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), 3].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                            else if (conditionCheck.Contains("Net Present Value"))
                            {
                                if (p == 0)
                                {
                                    sum = addArray2DList[0][o - 1, 1] + "+" + addArray2DList[0][o - 4, 1];
                                }
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Formula = "(" + sum + ")";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Numberformat.Format = "$#,##0.00";
                                wsCapitalBudgeting.Cells[(OutPutRowcount + o), (3)].Style.Font.Size = 12;
                            }
                        }                        
                    }                        
                }
            }

            return wsCapitalBudgeting;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <param name="UserId"></param>
        /// <param name="ProjectId"></param>
        /// <param name="Flag"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("[action]")]
        public ActionResult<Object> ExportEvaluation(string str, long UserId, long ProjectId, int Flag)
        {
            double marginalTax = 0;
            double discountRate = 0;
            double volChangePerc = 1;
            double unitPricePerChange = 1;
            double unitCostPerChange = 1;
            object[][][] chageFixedCost = null;
            object[][] summaryOutput = null;
            string lastCell = null;
            string rootFolder = _hostingEnvironment.WebRootPath;
            string fileName = @"capital_budgeting.xlsx";
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            var formattedCustomObject = (String)null;
            double capexDepPerChange = 1;
                               
            List<CapitalBudgeting> capitalBudgeting = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).ToList();

            ProjectsViewModel result = new ProjectsViewModel();

            Project project = iProject.GetSingle(x => x.Id == ProjectId);
            if (project != null)
            {             
                List<ProjectInputDatasViewModel> projectInputDatasVM = new List<ProjectInputDatasViewModel>();
                List<ProjectInputValuesViewModel> projectInputValuesVM = new List<ProjectInputValuesViewModel>();
                ProjectInputDatasViewModel DatasVm = new ProjectInputDatasViewModel();

                  result = mapper.Map<Project, ProjectsViewModel>(project);

                List<ProjectInputDatas> projectInputList = iProjectInputDatas.FindBy(s => s.Id != 0 && s.ProjectId == ProjectId).ToList();

                if (projectInputList != null && projectInputList.Count > 0)
                {

                    //get all values List
                    List<ProjectInputValues> projectInputValueList = iProjectInputValues.FindBy(x => x.Id != 0 && projectInputList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();

                    foreach (ProjectInputDatas datas in projectInputList)
                    {
                        DatasVm = new ProjectInputDatasViewModel();

                        DatasVm = mapper.Map<ProjectInputDatas, ProjectInputDatasViewModel>(datas);
                        if (DatasVm != null && DatasVm.ProjectInputValuesVM != null && DatasVm.ProjectInputValuesVM.Count > 0)
                        {
                            var valuesList = DatasVm.ProjectInputValuesVM;
                            DatasVm.ProjectInputValuesVM = new List<ProjectInputValuesViewModel>();
                            DatasVm.ProjectInputValuesVM = valuesList.OrderBy(x => x.Year).ToList();
                        }
                        projectInputDatasVM.Add(DatasVm);
                    }
                  
                    if (projectInputDatasVM != null && projectInputDatasVM.Count > 0)
                    {
                        result.ProjectInputDatasVM = new List<ProjectInputDatasViewModel>();
                        result.ProjectInputDatasVM = projectInputDatasVM;
                       // result.ProjectSummaryDatasVM = GetProjectOutput(result);
                    }                   
                }

            }    

            Dictionary<string, object> sensiScenSummaryOutput = new Dictionary<string, object>();
            // var redisKey = "Evaluation_" + UserId.ToString() + "::" + ProjectId.ToString();
            // var present = iSowfinCache.IsInCache(redisKey);
            // if (present == false)
            // {
            //     capitalBudgeting = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).ToList();
            //     iSowfinCache.Set(redisKey, capitalBudgeting[0]);
            // }
            // else
            // {
            //     capitalBudgeting.Add(iSowfinCache.Get<CapitalBudgeting>(redisKey));
            // }

                TableData tableData = new TableData(); // or return null if needed

           
                // TableData tableData = JsonConvert.DeserializeObject<TableData>(capitalBudgeting[0].RawTableData);
                if (capitalBudgeting != null && capitalBudgeting.Count > 0 && capitalBudgeting[0].RawTableData != null)
                {
                     tableData = JsonConvert.DeserializeObject<TableData>(capitalBudgeting[0].RawTableData);
                }

                
            
            //EvalSummaryOutput evalSummaryOutput = JsonConvert.DeserializeObject<EvalSummaryOutput>(capitalBudgeting[0].SummaryOutput);

            if (Flag == 2 && str != null)
            {
                // TODO ---- 
                //List<Dictionary<string, object>> body = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(str);
                //AggregateList arrgaggregateList = JsonConvert.DeserializeObject<AggregateList>(capitalBudgeting[0].TableData);
                //for (int i = 0; i < body.Count; i++)
                //{
                //    SensiScenarioNpv(body[i], capitalBudgeting, out marginalTax, out discountRate,
                //        out volChangePerc, out unitPricePerChange, out unitCostPerChange,
                //        out chageFixedCost, out capexDepPerChange, out sensiScenSummaryOutput);

                //}
            }

            using (ExcelPackage package = new ExcelPackage(file))
            {
                var wsCapitalBudgeting = package.Workbook.Worksheets["CapitalBudgeting"];
                wsCapitalBudgeting.SelectedRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                
                ExcelAddress yearCountAdd = null;
                List<string> yearCoutList = new List<string>();
         if (capitalBudgeting != null && capitalBudgeting.Count > 0 && capitalBudgeting[0] != null){

                for (var i = 0; i < capitalBudgeting[0].NoOfYears; i++)
                //for (var i = 0; i < result.NoOfYears; i++)
                    {
                    var curCol = 3 + i;
                     wsCapitalBudgeting.Cells[2, curCol].Value = capitalBudgeting[0].StartingYear + i;
                    //wsCapitalBudgeting.Cells[2, curCol].Value = result.StartingYear + i;  
                    wsCapitalBudgeting.Cells[3, curCol].Value = i;
                    yearCountAdd = new ExcelAddress(3, curCol, 3, curCol);
                    wsCapitalBudgeting.Cells[2, curCol].Style.Font.Size = 12;
                    wsCapitalBudgeting.Cells[3, curCol].Style.Font.Size = 12;
                    wsCapitalBudgeting.Cells[3, curCol].Style.Font.Bold = true;
                    wsCapitalBudgeting.Cells[2, curCol].Style.Font.Bold = true;

                    yearCoutList.Add(yearCountAdd.ToString());
                }
         }

                List<ProjectInputDatasViewModel> InputDatasList = new List<ProjectInputDatasViewModel>();
                InputDatasList = result.ProjectInputDatasVM;

                // TODO - 
                if (tableData != null && 
                    tableData.RevenueVariableCostTier != null && 
                    tableData.RevenueVariableCostTier.Count() > 0 && 
                    tableData.RevenueVariableCostTier[0] != null && 
                    tableData.RevenueVariableCostTier[0].Count() > 0 && 
                    tableData.RevenueVariableCostTier[0][0] != null &&
                    wsCapitalBudgeting != null)
                {
                wsCapitalBudgeting = ReturnCellStyle("Total", 2, (tableData.RevenueVariableCostTier[0][0].Length + 2), wsCapitalBudgeting, 0);
                wsCapitalBudgeting = ReturnCellStyle("Average", 2, (tableData.RevenueVariableCostTier[0][0].Length + 3), wsCapitalBudgeting, 0);

                wsCapitalBudgeting = ReturnCellStyle("Total", 2, (int)(result.NoOfYears + 3), wsCapitalBudgeting, 0);
                wsCapitalBudgeting = ReturnCellStyle("Average", 2, (int)(result.NoOfYears + 4), wsCapitalBudgeting, 0);

                }

                List<string> keys = new List<string>();
                List<List<double>> values = new List<List<double>>();
                int row_count = 0;
                List<List<List<string>>> listOfCellsRevenue = new List<List<List<string>>>();

                List<ProjectInputDatasViewModel> RevenueVariableCostTier = new List<ProjectInputDatasViewModel>();
                // VolumeList = InputDatasList.FindAll(x => x.LineItem == "Volume" && x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

                if (InputDatasList != null && InputDatasList.Count > 0)
                {
                    RevenueVariableCostTier = InputDatasList.FindAll(x => x.SubHeader.Contains("Revenue & Variable Cost")).ToList();
                }
                // RevenueVariableCostTier = InputDatasList.FindAll(x => x.SubHeader.Contains("Revenue & Variable Cost")).ToList();

            if (tableData != null && tableData.RevenueVariableCostTier != null && wsCapitalBudgeting != null)
            {
                wsCapitalBudgeting = ExcelGeneration(
                    tableData.RevenueVariableCostTier,
                    "Revenue",
                    wsCapitalBudgeting,
                    3,
                    row_count,
                    volChangePerc,
                    unitPricePerChange,
                    unitCostPerChange,
                    capexDepPerChange,
                    out row_count,
                    out listOfCellsRevenue
                );
            }

                // wsCapitalBudgeting = ExcelGeneration(tableData.RevenueVariableCostTier, "Revenue", wsCapitalBudgeting, 3, row_count, volChangePerc,
                //     unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsRevenue);
           
                int CountOfRevenueVariableCostTier = 0;
                CountOfRevenueVariableCostTier = (RevenueVariableCostTier.Count) / 3;
   
                for (int i = 1; i < CountOfRevenueVariableCostTier + 1; i++)
                {
                    List<ProjectInputDatasViewModel> RevenueVariableCostTierNew = new List<ProjectInputDatasViewModel>();
                    RevenueVariableCostTierNew = InputDatasList.FindAll(x => x.SubHeader.Contains("Revenue & Variable Cost Tier"+i+"")).ToList();

                   // wsCapitalBudgeting = ExcelGeneration_New(RevenueVariableCostTierNew, "Revenue & Variable Cost Tier"+i+"", wsCapitalBudgeting, 3, row_count, volChangePerc,
                   //unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsRevenue);
                }

                //wsCapitalBudgeting = ExcelGeneration_New(RevenueVariableCostTier, "Revenue & Variable Cost Tier 1", wsCapitalBudgeting, 3, row_count, volChangePerc,
                //   unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsRevenue);

                //wsCapitalBudgeting.SelectedRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                List<List<List<string>>> listOfCellsFixed = new List<List<List<string>>>();
                List<ProjectInputDatasViewModel> OtherFixedCost = new List<ProjectInputDatasViewModel>();
                if(InputDatasList != null && InputDatasList.Count > 0)
                {
                    OtherFixedCost = InputDatasList.FindAll(x => x.SubHeader.Contains("Other Fixed Cost")).ToList();
                }
                // OtherFixedCost = InputDatasList.FindAll(x => x.SubHeader.Contains("Other Fixed Cost")).ToList();

                //if (Flag == 2 && str != null && tableData.OtherFixedCost != null)
                if (Flag == 2 && str != null && OtherFixedCost != null && OtherFixedCost.Count > 0)
                {
                    // TODO ---- 
                    //wsCapitalBudgeting = ExcelGeneration(chageFixedCost, "Fixed Cost", wsCapitalBudgeting, 3, row_count, volChangePerc, unitPricePerChange,
                    //    unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsFixed);                  
                }
                // else if (tableData.OtherFixedCost != null)
                else if (OtherFixedCost != null && tableData != null && tableData.RevenueVariableCostTier != null && wsCapitalBudgeting != null )
                {
                    wsCapitalBudgeting = ExcelGeneration(tableData.OtherFixedCost, "Fixed Cost", wsCapitalBudgeting, 3, row_count, volChangePerc, unitPricePerChange,
                        unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsFixed);
                    //wsCapitalBudgeting = ExcelGeneration_New(OtherFixedCost, "Other Fixed Cost", wsCapitalBudgeting, 3, row_count, volChangePerc, unitPricePerChange,
                    //   unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsFixed);
                }

                List<List<List<string>>> listOfCellsCapex = new List<List<List<string>>>();

                List<ProjectInputDatasViewModel> CapexDepreciation = new List<ProjectInputDatasViewModel>();
                if(InputDatasList != null && InputDatasList.Count > 0)
                {
                    CapexDepreciation = InputDatasList.FindAll(x =>  x.SubHeader != null && x.SubHeader.Contains("Capex & Depreciation")).ToList();
                }

                // if (Flag == 2 && str != null && tableData.CapexDepreciation != null)
                if (Flag == 2 && str != null && CapexDepreciation != null)
                {
                    //wsCapitalBudgeting = ExcelGeneration(tableData.CapexDepreciation, "Capex&Depreciation", wsCapitalBudgeting, 3, row_count, volChangePerc,
                    //    unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);
                    //wsCapitalBudgeting = ExcelGeneration_New(CapexDepreciation, "Capex & Depreciation", wsCapitalBudgeting, 3, row_count, volChangePerc,
                    //    unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);
                }
                // else if (tableData.CapexDepreciation != null)
                else if (CapexDepreciation != null)
                {

                    int CountOfCapexDepreciation = 0;
                    CountOfCapexDepreciation = (CapexDepreciation.Count) / 2;

                    for (int i = 1; i < CountOfCapexDepreciation + 1; i++)
                    {
                        List<ProjectInputDatasViewModel> CapexDepreciationNew = new List<ProjectInputDatasViewModel>();
                        CapexDepreciationNew = InputDatasList.FindAll(x => x.SubHeader.Contains("Capex & Depreciation"+i+"")).ToList();

                        //wsCapitalBudgeting = ExcelGeneration_New(CapexDepreciationNew, "Capex & Depreciation"+i+"", wsCapitalBudgeting, 3, row_count, volChangePerc,
                        //unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);
                    }

                    if( tableData != null && tableData.RevenueVariableCostTier != null && wsCapitalBudgeting != null){
                         wsCapitalBudgeting = ExcelGeneration(tableData.CapexDepreciation, "Capex&Depreciation", wsCapitalBudgeting, 3, row_count, volChangePerc,
                        unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);
                    }

                  
                    //wsCapitalBudgeting = ExcelGeneration_New(CapexDepreciation, "Capex & Depreciation", wsCapitalBudgeting, 3, row_count, volChangePerc,
                    //    unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);
                }
                List<string> nwcAdd = new List<string>();

                List<ProjectInputDatasViewModel> WorkingCapital = new List<ProjectInputDatasViewModel>();
                if(InputDatasList != null && InputDatasList.Count > 0)
                {
                    WorkingCapital = InputDatasList.FindAll(x => x.SubHeader.Contains("Working Capital")).ToList();
                }

                // TODO ---- 
                if (tableData != null && tableData.WorkingCapital != null && wsCapitalBudgeting != null){
                 wsCapitalBudgeting = ExcelGenerationEx(tableData.WorkingCapital, "Working Capital", wsCapitalBudgeting, 3, row_count, out row_count, out nwcAdd);
            }
                //wsCapitalBudgeting = ExcelGeneration1(WorkingCapital, "Working Capital", wsCapitalBudgeting, 3, row_count, volChangePerc,
                //        unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out nwcAdd);
                //wsCapitalBudgeting = ExcelGeneration_New(WorkingCapital, "Working Capital", wsCapitalBudgeting, 3, row_count, volChangePerc,
                //        unitPricePerChange, unitCostPerChange, capexDepPerChange, out row_count, out listOfCellsCapex);
                 row_count =+ 6;

                row_count = row_count + 4;

                wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                wsCapitalBudgeting.Cells[(5 + row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                wsCapitalBudgeting.Cells[(row_count + 1), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                wsCapitalBudgeting.Cells[(row_count), 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                if (Flag == 2 && str != null)
                {
                    wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = marginalTax;
                    wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = discountRate;                    
                }
                else
                {

                    if (capitalBudgeting != null && capitalBudgeting.Count > 0)
                {
                    wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = capitalBudgeting[0].MarginalTaxRate;
                    wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = capitalBudgeting[0].DiscountRate;
                }
                

    // Check if InputDatasList is not null or empty before finding elements
            if (InputDatasList != null && InputDatasList.Count > 0)
            {
                var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                if (MarginalTaxDatas != null)
                {
                    wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value;
                }
            

                var DiscountRateDatas = InputDatasList.Find(x => x.LineItem.Contains("WACC"));
                if (DiscountRateDatas != null)
                {
                    wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = DiscountRateDatas.Value;
                }
            
            }


                    // wsCapitalBudgeting.Cells[(5 + row_count), 2].Value = capitalBudgeting[0].MarginalTaxRate;
                    // wsCapitalBudgeting.Cells[(5 + row_count + 1), 2].Value = capitalBudgeting[0].DiscountRate;

                    // var MarginalTaxDatas = InputDatasList.Find(x => x.LineItem.Contains("Marginal Tax rate"));
                    // wsCapitalBudgeting.Cells[(row_count), 2].Value = MarginalTaxDatas.Value;

                    // var DiscountRateDatas = InputDatasList.Find(x => x.LineItem.Contains("WACC"));
                    // wsCapitalBudgeting.Cells[(row_count + 1), 2].Value = DiscountRateDatas.Value;
                }

                wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate(%)", (5 + row_count), 1, wsCapitalBudgeting, 0);
                wsCapitalBudgeting = ReturnCellStyle("Marginal Tax Rate(%)", (row_count), 1, wsCapitalBudgeting, 0);

                var marginalTaxAdd = new ExcelAddress((5 + row_count), 2, (5 + row_count), 2);

                 wsCapitalBudgeting = ReturnCellStyle("Cost of Capital(%)", (5 + row_count + 1), 1, wsCapitalBudgeting, 0);
                wsCapitalBudgeting = ReturnCellStyle("Weighted average cost of capital(%)", (row_count + 1), 1, wsCapitalBudgeting, 0);

                var discountRateAdd = new ExcelAddress((5 + row_count + 1), 2, (5 + row_count + 1), 2);

                row_count = (row_count + 10);///seprate the input tables and output
                row_count = (row_count + 2);
                if (Flag == 1 || Flag == 3)
                {
                     summaryOutput = JsonConvert.DeserializeObject<object[][]>(capitalBudgeting[0].SummaryOutput);
                        result.ProjectSummaryDatasVM = GetProjectOutput(result);
                }
                else if (Flag == 2)
                {
                    summaryOutput = DictionaryToArray(sensiScenSummaryOutput).Select(a => a.ToArray()).ToArray(); ;
                }

                if (result.ProjectSummaryDatasVM != null)
                    foreach (var outputDatasVm in result.ProjectSummaryDatasVM)
                    {
                        if (outputDatasVm != null && outputDatasVm.ProjectOutputValuesVM != null && outputDatasVm.ProjectOutputValuesVM.Count > 0)
                        {
                            var valuesList = outputDatasVm.ProjectOutputValuesVM;
                            outputDatasVm.ProjectOutputValuesVM = new List<ProjectOutputValuesViewModel>();
                            outputDatasVm.ProjectOutputValuesVM = valuesList.OrderBy(x => x.Year).ToList();                           
                        }
                    }

                string[,] addArray2D = new string[1,1];
                string[,] addArray2D1 = new string[1, 1];

                if (summaryOutput != null && summaryOutput.Length > 0 && summaryOutput[0] != null)
                {
                    addArray2D = new string[summaryOutput.Length, summaryOutput[0].Length];
                }


                if (result.ProjectSummaryDatasVM != null && result.ProjectSummaryDatasVM.Count > 0 && 
                    result.ProjectSummaryDatasVM[0] != null && 
                    result.ProjectSummaryDatasVM[0].ProjectOutputValuesVM != null &&
                    result.ProjectSummaryDatasVM[0].ProjectOutputValuesVM.Count > 0)
                {
                   addArray2D1 = new string[result.ProjectSummaryDatasVM.Count, result.ProjectSummaryDatasVM[0].ProjectOutputValuesVM.Count];
                }
            if (result != null && result.ProjectSummaryDatasVM != null && result.ProjectSummaryDatasVM.Count > 0)
            {
                for (int o = 0; o < result.ProjectSummaryDatasVM.Count; o++)
                {

                    if (result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM != null)
                    {
                        for (int p = 0; p < result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM.Count; p++)
                        {
                            try
                            {
                                if (p == 0)
                                {
                                    var conditionCheck = result.ProjectSummaryDatasVM[o].LineItem;

                                    string UnitValue = null;

                                    if (result.ProjectSummaryDatasVM[o].UnitId == 1)
                                    {
                                        UnitValue = "$";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 2)
                                    {
                                        UnitValue = "$K";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 3)
                                    {
                                        UnitValue = "$M";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 4)
                                    {
                                        UnitValue = "$B";
                                    }
                                    else if (result.ProjectSummaryDatasVM[o].UnitId == 5)
                                    {
                                        UnitValue = "$T";
                                    }   

                                    if (conditionCheck.Contains("Sales") || conditionCheck.Contains("COGS") ||
                                    conditionCheck.Contains("Unlevered Net Income") || conditionCheck.Contains("Free Cash Flow") ||
                                    conditionCheck.Contains("Gross Margin") || conditionCheck.Contains("Operating Income") ||
                                     conditionCheck.Contains("Net Present Value"))
                                    {
                                        wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 0);
                                        addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                                    }
                                    else
                                    {
                                        if (conditionCheck == "Discount Factor")
                                        {
                                            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + "$" + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 1);
                                            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                                        }
                                        else
                                        {
                                            wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (p + 1), wsCapitalBudgeting, 1);
                                            addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                                        }
                                        
                                    }
                                }
                                // else
                                // {
                                wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM[p].Value), (row_count + o), (p + 3), wsCapitalBudgeting, 1);
                                var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                                addArray2D[o, p] = Convert.ToString(addr.ToString());
                                // }
                            }
                            catch (IndexOutOfRangeException)
                            {
                                var strr = "";
                                wsCapitalBudgeting = ReturnCellStyle(strr, (row_count + o), (p + 3), wsCapitalBudgeting, 1);
                                var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                                addArray2D[o, p] = Convert.ToString(addr.ToString());
                            }
                        }
                    }
                    else if (result.ProjectSummaryDatasVM[o].ProjectOutputValuesVM == null)
                    {
                            int p = 0;
                            var conditionCheck = result.ProjectSummaryDatasVM[o].LineItem;

                        string UnitValue = null;

                        if (result.ProjectSummaryDatasVM[o].UnitId == 1)
                        {
                            UnitValue = "$";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 2)
                        {
                            UnitValue = "$K";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 3)
                        {
                            UnitValue = "$M";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 4)
                        {
                            UnitValue = "$B";
                        }
                        else if (result.ProjectSummaryDatasVM[o].UnitId == 5)
                        {
                            UnitValue = "$T";
                        }

                        if (conditionCheck.Contains("Net Present Value") || 
                                conditionCheck.Contains("IRR (Internal Rate of Return) of Free Cash Flows"))
                            {

                            if (conditionCheck == "IRR (Internal Rate of Return) of Free Cash Flows")
                            {
                                wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + "%" + ")", (row_count + o), (1), wsCapitalBudgeting, 0);
                                addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                            }
                            else
                            {
                                wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (1), wsCapitalBudgeting, 0);
                                addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                            }
                            
                            }
                            else
                            {
                                wsCapitalBudgeting = ReturnCellStyle(result.ProjectSummaryDatasVM[o].LineItem + "(" + UnitValue + ")", (row_count + o), (1), wsCapitalBudgeting, 1);
                                addArray2D[o, p] = result.ProjectSummaryDatasVM[o].LineItem;
                            }

                        wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(result.ProjectSummaryDatasVM[o].Value), (row_count + o), (3), wsCapitalBudgeting, 1);
                        var addr = new ExcelAddress((row_count + o), (p + 3), (row_count + o), (p + 3));
                        addArray2D[o, p] = Convert.ToString(addr.ToString());
                    }
                }
            }

            if(summaryOutput != null && summaryOutput.Length > 0)
           {

                for (int o = 0; o < summaryOutput.Length; o++)
                {
                    for (int p = 0; p < summaryOutput[0].Length; p++)
                    {
                        try
                        {
                            if (p == 0)
                            {
                                var conditionCheck = summaryOutput[o][p].ToString();
                              
                                if (conditionCheck.Contains("Sales") || conditionCheck.Contains("COGS") ||
                                conditionCheck.Contains("Unlevered Net Income") || conditionCheck.Contains("Free Cash Flow") ||
                                conditionCheck.Contains("Gross Margin") || conditionCheck.Contains("Operating Income") ||
                                conditionCheck.Contains("Discounted Cash Flow") || conditionCheck.Contains("Net Present Value"))
                                {
                                    wsCapitalBudgeting = ReturnCellStyle(summaryOutput[o][p].ToString(), (row_count + o), (p + 1), wsCapitalBudgeting, 0);
                                    addArray2D[o, p] = summaryOutput[o][p].ToString();
                                }
                                else
                                {
                                    wsCapitalBudgeting = ReturnCellStyle(summaryOutput[o][p].ToString(), (row_count + o), (p + 1), wsCapitalBudgeting, 1);
                                    addArray2D[o, p] = summaryOutput[o][p].ToString();
                                }
                            }
                            else
                            {
                                wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(summaryOutput[o][p]), (row_count + o), (p + 2), wsCapitalBudgeting, 1);
                                var addr = new ExcelAddress((row_count + o), (p + 2), (row_count + o), (p + 2));
                                addArray2D[o, p] = Convert.ToString(addr.ToString());
                            }
                        }
                        catch (IndexOutOfRangeException)
                        {
                            var strr = "";
                            wsCapitalBudgeting = ReturnCellStyle(strr, (row_count + o), (p + 2), wsCapitalBudgeting, 1);
                            var addr = new ExcelAddress((row_count + o), (p + 2), (row_count + o), (p + 2));
                            addArray2D[o, p] = Convert.ToString(addr.ToString());
                        }
                    }
                }
            }

                List<string> sales = new List<string>();
                List<string> cogs = new List<string>();

                var crossMargin = RevenueFormula(listOfCellsRevenue, out sales, out cogs);
                var operatingIncome2D = GetCellAddress(addArray2D, "COGS", "Operating Income");
                var freeCashFlow2D = GetCellAddress(addArray2D, "Income Tax", "Free Cash Flow");
                var operatingIncome = FormulaGeneration(operatingIncome2D, 0);
                var freeCashFlow = FormulaGeneration(freeCashFlow2D, 1);
                string sum = null;

                List<string> salesCellList = new List<string>();
                List<string> nwcCellList = new List<string>();                
                for (int k = 0; k < addArray2D.GetLength(0); k++)
                {
                    for (int l = 1; l < addArray2D.GetLength(1); l++)
                    {
                        //Modified for change of 2 decimal places only during export file. package.Workbook.Worksheets
                        wsCapitalBudgeting.SelectedRange[addArray2D[k, l]].Style.Numberformat.Format = "00.00";
                        //wsCapitalBudgeting.SelectedRange[addArray2D[k, l]].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                        if (addArray2D[k, 0].Contains("Sales"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + sales[l - 1] + ")";
                            salesCellList.Add(addArray2D[k, l]);
                        }
                        else if (addArray2D[k, 0].Contains("COGS"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + cogs[l - 1] + ")";
                        }
                        else if (addArray2D[k, 0].Contains("Gross Margin"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + crossMargin[l - 1] + ")";
                        }
                        else if (addArray2D[k, 0].Contains("Operating Income"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + operatingIncome[l - 1] + ")";
                        }
                        else if (addArray2D[k, 0].Contains("Income Tax"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = addArray2D[k - 1, l] + "*" + "(" + marginalTaxAdd + "/" + "100" + ")";
                           
                        }
                        else if (addArray2D[k, 0].Contains("Unlevered Net Income"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + addArray2D[k - 2, l] + "-" + addArray2D[k - 1, l] + ")";
                           
                        }
                        else if (addArray2D[k, 0].Contains("NWC(Net Working Capital)"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = salesCellList[l - 1] + "*" + (nwcAdd[l - 1] + "/" + 100);
                            nwcCellList.Add(addArray2D[k, l]);
                           
                        }
                        else if (addArray2D[k, 0].Contains("Less: Increases in NWC"))
                        {
                            if (l == 1)
                            {
                                wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "";
                            }
                            else
                            {
                                wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + nwcCellList[l - 1] + "-" + nwcCellList[l - 2] + ")";
                                
                            }
                        }
                        else if (addArray2D[k, 0].Contains("Free Cash Flow"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + freeCashFlow[l - 1] + ")";
                            

                        }
                        else if (addArray2D[k, 0].Contains("Discount Factor"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = "(" + ((1 + "+" + (discountRateAdd + "/" + "100")) + ")" + "^" + yearCoutList[l - 1]);
                            

                        }
                        else if (addArray2D[k, 0].Contains("Discount Cash Flow"))
                        {
                            wsCapitalBudgeting.Cells[addArray2D[k, l]].Formula = addArray2D[k - 3, l] + "/" + addArray2D[k - 1, l];
                            
                        }
                        else if (addArray2D[k, 0].Contains("Net Present Value"))
                        {

                            if (l == 1)
                            {
                                sum = addArray2D[k - 1, l];
                            }
                            else
                            {
                                sum = sum + "+" + addArray2D[k - 1, l];
                            }
                            wsCapitalBudgeting.Cells[addArray2D[k, 1]].Formula = "(" + sum + ")";
                            lastCell = addArray2D[k, 1].ToString();
                        }

                        if (addArray2D[k, 0].Contains("Levered Value"))
                        {
                            for (int r = (addArray2D.GetLength(1) - 1); r > 0; r--)
                            {
                                if (r == (addArray2D.GetLength(1) - 1))
                                {
                                    wsCapitalBudgeting.Cells[addArray2D[k, r]].Formula = "";
                                }
                                else
                                {
                                    wsCapitalBudgeting.Cells[addArray2D[k, r]].Formula = "(" + (addArray2D[k - 1, r + 1] + "+" + addArray2D[k, r + 1]) + ")" +
                                        "/" + "(" + 1 + "+" + "(" + discountRateAdd + "/" + "100" + ")" + ")";
                                }
                            }
                        }
                    }
                }

                // var lastAddr = new ExcelAddress(lastCell);

            ExcelAddress lastAddr = null;

            if (!string.IsNullOrEmpty(lastCell))
            {
   
                lastAddr = new ExcelAddress(lastCell);
            }




                if (lastAddr != null && lastAddr.Start != null)
                {
                    row_count = lastAddr.Start.Row + 5;
                }
                if (Flag == 3 && str != null)
                {
                    List<Dictionary<string, object>> body = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(str);
                    List<List<object>> listOfList = new List<List<object>>();
                    var key = body[0].Keys.ToList();
                    List<Object> objectList = new List<Object>(key);
                    listOfList.Add(objectList);
                    for (int i = 0; i < body.Count; i++)
                    {
                        var bodyValues = body[i].Values.ToList();
                        listOfList.Add(bodyValues);
                    }
                    for (int i = 0; i < listOfList.Count; i++)
                    {
                        for (int j = 0; j < listOfList[i].Count; j++)
                        {
                            wsCapitalBudgeting = ReturnCellStyle(Convert.ToString(listOfList[i][j]), (row_count + i), j + 1, wsCapitalBudgeting, 0);
                        }
                    }
                }

                ExcelPackage excelPackage = new ExcelPackage();
                excelPackage.Workbook.Worksheets.Add("CapitalBudgeting", wsCapitalBudgeting);
                //package.Save();
                ExcelPackage epOut = excelPackage;
                byte[] myStream = epOut.GetAsByteArray();
                var inputAsString = Convert.ToBase64String(myStream);
                formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);
                excelPackage.Dispose();
                return Ok(formattedCustomObject);
            }
        }

        /// <summary>
        /// Deprected
        /// </summary>
        /// <param name="UserId"></param>
        /// <param name="ProjectId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("AllCapitalBudgeting_old/{UserId}/{ProjectId}")]
        public ActionResult<Object> AllCapitalBudgeting_old(long UserId, long ProjectId)
        {
            List<CapitalBudgeting> capitalBudgeting = new List<CapitalBudgeting>();

            var redisKey = "Evaluation_" + UserId.ToString() + "::" + ProjectId.ToString();
            var present = iSowfinCache.IsInCache(redisKey);
            if (present == false)
            {
                capitalBudgeting = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).ToList();
                if (capitalBudgeting.Count != 0)
                {
                    iSowfinCache.Set(redisKey, capitalBudgeting[0]);
                }
            }
            else
            {
                capitalBudgeting.Add(iSowfinCache.Get<CapitalBudgeting>(redisKey));
            }

            if (capitalBudgeting == null)
            {
                return NotFound();
            }
            else
            {
                return Ok(capitalBudgeting);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserId"></param>
        /// <param name="ProjectId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("AllCapitalBudgeting/{UserId}/{ProjectId}")]
        public ActionResult<Object> AllCapitalBudgeting(long UserId, long ProjectId)
        {
            List<CapitalBudgeting> capitalBudgeting = new List<CapitalBudgeting>();

            capitalBudgeting = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).OrderByDescending(x=>x.Id).ToList();

            if (capitalBudgeting == null)
            {
                return NotFound();
            }
            else
            {
                return Ok(capitalBudgeting);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserId"></param>
        /// <param name="ProjectId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("Sensitivity/Graph/{UserId}/{ProjectId}")]
        public ActionResult<Object> Sensitivity(long UserId, long ProjectId)
        {   
            List<CapitalBudgeting> capitalBudgeting = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).ToList();
            if( capitalBudgeting != null && capitalBudgeting.Count() > 0)
            {
             
            TableData tableData = JsonConvert.DeserializeObject<TableData>(capitalBudgeting[0].RawTableData);

            AggregateList aggregateList = JsonConvert.DeserializeObject<AggregateList>(capitalBudgeting[0].TableData);
            List<Dictionary<string, List<List<double>>>> result = new List<Dictionary<string, List<List<double>>>>();
            result.Add(ProvideMinMax(aggregateList.Volume, 0, "Volume(M)", capitalBudgeting)); //find the range of values passing zero for total
            result.Add(MinAvgMax(capitalBudgeting[0].MarginalTaxRate, "MarginalTax(%)", capitalBudgeting)); //find the range of values
            result.Add(MinAvgMax(capitalBudgeting[0].DiscountRate, "DiscountRate(%)", capitalBudgeting));
            result.Add(ProvideMinMax(aggregateList.UnitPrice, 1, "UnitPrice($M)", capitalBudgeting)); //find the range of values passing one for average 
            result.Add(ProvideMinMax(aggregateList.UnitCost, 1, "UnitCost($M)", capitalBudgeting)); //find the range of values passing one for average 
            if (tableData.CapexDepreciation.Length != 0)
            {
                result.Add(ProvideMinMax(aggregateList.Capex, 0, "Capex($M)", capitalBudgeting));
            }

            if (tableData.OtherFixedCost.Length != 0)
            {
                List<string> keys = new List<string>();
                List<List<double>> values = new List<List<double>>();
                var yourDictionary = FixedCost(tableData.OtherFixedCost, out keys, out values);   // form key value pair for fixed cost table
                var arrayOfAllKeys = yourDictionary.Keys.ToArray();
                foreach (string str in arrayOfAllKeys)
                {
                    result.Add(ProvideMinMax(yourDictionary[str], 0, str, capitalBudgeting));   //find the range of values fixedcost
                }
            }

            string dictResult = JsonConvert.SerializeObject(result);
            return Ok(result);
            }
            else
            {
                return NotFound("No data found");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="body"></param>
        /// <param name="UserId"></param>
        /// <param name="ProjectId"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("Scenario/Sensitivity/{UserId}/{ProjectId}")]
        public ActionResult<Object> ScenarioSensitivity([FromBody] List<Dictionary<string, object>> body, long UserId, long ProjectId)
        {
            //List < Dictionary<string, object>
            //var capBuds = await iCapBud.GetCapBud(UserId, ProjectId);
            //var stuff = JsonConvert.DeserializeObject(capBuds[0].DictionaryKeys);
            List<CapitalBudgeting> capBudList = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).ToList();

            // var redisKey = "Evaluation_" + UserId.ToString() + "::" + ProjectId.ToString();
            // var present = iSowfinCache.IsInCache(redisKey);
            // if (present == false)
            // {
            //     capBudList = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).ToList();
            //     iSowfinCache.Set(redisKey, capBudList[0]);
            // }
            // else
            // {
            //     capBudList.Add(iSowfinCache.Get<CapitalBudgeting>(redisKey));
            // }
            if(capBudList == null || capBudList.Count() == 0)
            {
                return NotFound("No data found");
            }
            List<double> npvList = new List<double>();
            for (int i = 0; i < body.Count; i++)
            {
                npvList.Add(SensiScenarioNpv(body[i], capBudList, out _, out _, out _, out _, out _, out _, out _, out _));
            }

            return Ok(npvList);
        }

        [HttpPost]
        [Route("AddTables")]
        public ActionResult<Object> AddTables([FromBody] CapitalBugetingTablesViewModel model)
        {
            try
            {
                if (model.Id == 0)
                {
                    CapitalBugetingTables capitalBugetingTables = new CapitalBugetingTables
                    {
                        ProjectId = model.ProjectId,
                        UserId = model.UserId,
                        StartingYear = model.StartingYear,
                        NoOfYears = model.NoOfYears,
                        DiscountRate = model.DiscountRate,
                        MarginalTax = model.MarginalTax,
                        EvalFlag = model.EvalFlag,
                        ApprovalFlag = model.ApprovalFlag,
                        Tables = model.Tables
                    };
                    iCapitalBugetingTables.Add(capitalBugetingTables);
                    iCapitalBudgeting.Commit();
                    return Ok(new { message = "Successfully Added Tables", statusCode = 200, tableId = capitalBugetingTables.Id });
                }
                else
                {
                    CapitalBugetingTables UpdateCapitalBugetingTables = new CapitalBugetingTables
                    {
                        Id = model.Id,
                        ProjectId = model.ProjectId,
                        UserId = model.UserId,
                        StartingYear = model.StartingYear,
                        NoOfYears = model.NoOfYears,
                        DiscountRate = model.DiscountRate,
                        MarginalTax = model.MarginalTax,
                        EvalFlag = model.EvalFlag,
                        ApprovalFlag = model.ApprovalFlag,
                        Tables = model.Tables
                    };
                    iCapitalBugetingTables.Update(UpdateCapitalBugetingTables);
                    iCapitalBudgeting.Commit();
                    return Ok(new { message = "Successfully Update Tables", statusCode = 200, tableId = UpdateCapitalBugetingTables.Id });
                }
            }
            catch (Exception ex)
            {

                return BadRequest(new { message = ex.Message.ToString(), statusCode = 400 });
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserId"></param>
        /// <param name="ProjectId"></param>
        /// <returns></returns>


        [HttpGet]
        [Route("GetAllTables/{UserId}/{ProjectId}")]
        public ActionResult<Object> GetAllTables(long UserId, long ProjectId)
        {
            try
            {
                var tables = iCapitalBugetingTables.FindBy(s => s.ProjectId == ProjectId && s.UserId == UserId).OrderByDescending(x=>x.Id);
                if (tables == null)
                {
                    return NotFound(new { message = "Table data not found", statusCode = 404 });
                }
                return Ok(new { result = tables, statusCode = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), statusCode = 400 });
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetTables/{Id}")]
        public ActionResult<Object> GetTables(long Id)
        {
            try
            {
                var tables = iCapitalBugetingTables.GetSingle(s => s.Id == Id);
                if (tables == null)
                {
                    return NotFound(new { message = "Table data not found", statusCode = 404 });
                }
                return Ok(new { result = tables, statusCode = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), statusCode = 400 });
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("AddCapitalBudgetingSnapshots")]
        public ActionResult<Object> AddCapitalBudgetingSnapshots([FromBody] SnapshotsViewSnapshots model)
        {
            try
            {
                Snapshots snapshot = new Snapshots
                {
                    SnapShot = model.SnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    ProjectId = model.ProjectId,
                    SnapShotType = CAPITALBUDGETINGSNAPSHOT,
                    NPV = model.NVP,
                    CNPV = model.CNVP
                };
                iSnapShots.Add(snapshot);
                iSnapShots.Commit();
                return Ok(new { result = snapshot.Id, message = "Succesfully added Snapshots", code = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ProjecId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("CapitalBudgetingSnapshots/{ProjecId}")]
        public ActionResult<Object> CapitalBudgetingSnapshots(long ProjecId)
        {
            try
            {
                var SnapShot = iSnapShots.FindBy(s => s.SnapShotType == CAPITALBUDGETINGSNAPSHOT && s.ProjectId == ProjecId);
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetCapitalBudgetingSnapshot/{Id}")]
        public ActionResult<Object> GetCapitalBudgetingSnapshot(long Id)
        {
            try
            {
                var SnapShot = iSnapShots.GetSingle(s => s.SnapShotType == CAPITALBUDGETINGSNAPSHOT && s.Id == Id);
                if (SnapShot == null)
                {
                    return NotFound("No data found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("AddSensitivitySnapshots")]
        public ActionResult<Object> AddSensitivitySnapshots([FromBody] SensiBody model)
        {
            try
            {
                Snapshots snapshot = new Snapshots
                {
                    SnapShot = model.SensitivitySnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    ProjectId = model.ProjectId,
                    SnapShotType = SENSITIVITYSNAPSHOT,
                    NPV = model.Npv,
                    CNPV = model.ChangeNpv
                };
                iSnapShots.Add(snapshot);
                iSnapShots.Commit();
                return Ok(new { result = snapshot.Id, message = "Succesfully added Snapshots", code = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ProjecId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SensitvitySnapshots/{ProjecId}")]
        public ActionResult<Object> SensitvitySnapshots(long ProjecId)
        {
            try
            {
                var SnapShot = iSnapShots.FindBy(s => s.SnapShotType == SENSITIVITYSNAPSHOT && s.ProjectId == ProjecId);
                if (SnapShot == null)
                {
                    return NotFound("No data found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetSensitvitySnapshot/{Id}")]
        public ActionResult<Object> GetSensitvitySnapshot(long Id)
        {
            try
            {
                var SnapShot = iSnapShots.FindBy(s => s.SnapShotType == SENSITIVITYSNAPSHOT && s.Id == Id);
                if (SnapShot == null)
                {
                    return NotFound("no data found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("AddScenarioSnapshots")]
        public ActionResult<Object> AddScenarioSnapshots([FromBody] Snapshots model)
        {
            try
            {
                Snapshots snapshot = new Snapshots
                {
                    SnapShot = model.SnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    ProjectId = model.ProjectId,
                    SnapShotType = SCENARIOSNAPSHOT,
                    NPV = model.NPV,
                    CNPV = model.CNPV
                };

                iSnapShots.Add(snapshot);
                iSnapShots.Commit();

                return Ok(new { result = snapshot.Id, message = "Succesfully added Snapshots", code = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ProjecId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("ScenarioSnapshots/{ProjecId}")]
        public ActionResult<Object> ScenarioSnapshots(long ProjecId)
        {
            try
            {
                var SnapShot = iSnapShots.FindBy(s => s.SnapShotType == SCENARIOSNAPSHOT && s.ProjectId == ProjecId);
                if (SnapShot == null)
                {
                    return NotFound("No data found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetScenarioSnapshot/{Id}")]
        public ActionResult<Object> GetScenarioSnapshot(long Id)
        {
            try
            {
                var SnapShot = iSnapShots.FindBy(s => s.SnapShotType == SCENARIOSNAPSHOT && s.Id == Id);
                if (SnapShot == null)
                {
                    return NotFound("No data found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        [Route("ApprovalStatus/{UserId}/{ProjectId}")]
        public ActionResult<Object> GetApprovalStatus(long UserId, long ProjectId)
        {
            try
            {
                var capBud = iCapitalBudgeting.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).FirstOrDefault();
                if (capBud == null)
                {
                    return NotFound("No data found");
                }
                else
                {
                    return Ok(new
                    {
                        result = new Approval
                        {
                            ApprovalFlag = capBud.ApprovalFlag,
                            //SummaryOutput = capBud.SummaryOutput,
                            ApprovalComment = capBud.ApprovalComment,
                           // NPV = capBud.NPV
                        },
                        message = "Success",
                        code = 200
                    });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        //[HttpPost]
        //[Route("ApprovalStatus")]
        //public ActionResult<Object> UpdateApprovalStatus([FromBody] Approval model)
        //{
        //    try
        //    {
        //        var capBud = iCapitalBudgeting.GetSingle(s => s.Id == model.Id);
        //        if (capBud != null)
        //        {
        //            capBud.ApprovalFlag = 1;
        //            capBud.ApprovalComment = model.ApprovalComment;
        //            iCapitalBudgeting.Update(capBud);
        //            iCapitalBudgeting.Commit();
        //            return Ok(new { result = capBud,comment= capBud.ApprovalComment, message = "Successfully approved", code = 200 });
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(new { message = ex.Message.ToString(), code = 400 });
        //    }
        //    return Ok(new { result = model.Id, message = "Approval comment not saved", code = 200 });
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <param name="capBudList"></param>
        /// <param name="marginalTax"></param>
        /// <param name="discountRate"></param>
        /// <param name="volChangePerc"></param>
        /// <param name="unitPricePerChange"></param>
        /// <param name="unitCostPerChange"></param>
        /// <param name="chageFixedCost"></param>
        /// <param name="capexDepPerChange"></param>
        /// <param name="sensiSceSummaryOutput"></param>
        /// <returns></returns>
        public static double SensiScenarioNpv(Dictionary<string, object> model, List<CapitalBudgeting> capBudList,
             out double marginalTax, out double discountRate,
            out double volChangePerc, out double unitPricePerChange,
            out double unitCostPerChange, out object[][][] chageFixedCost, out double capexDepPerChange, out Dictionary<string, object> sensiSceSummaryOutput)
        {
            List<double> npvList = new List<double>();
            var objectDictionary = new Dictionary<string, object>();
            double volPercentChange = 1;
            double unitPricePercentChange = 1;
            double unitCostPercentChange = 1;
            object[][][] fixedCostObj = null;
            double outputMarginalTax = 0;
            double outputDiscountRate = 0;
            double capexDepChange = 1;

            // if (capBudList == null || capBudList.Count == 0)
            // {
            //     return 0;
            // }

            AggregateList aggregateList = JsonConvert.DeserializeObject<AggregateList>(capBudList[0].TableData);

            TableData data = JsonConvert.DeserializeObject<TableData>(capBudList[0].RawTableData);


            if (model.Keys.Contains("volume(M)"))
            {
                objectDictionary["Volume"] = FindForRevenueCapex(model["volume(M)"], aggregateList.Volume, out volPercentChange, 0); // passing flag zero to use total 
            }
            else
            {
                objectDictionary.Add("Volume", aggregateList.Volume);
            }
            if (model.Keys.Contains("unitPrice($M)"))
            {
                objectDictionary["UnitPrice"] = FindForRevenueCapex(model["unitPrice($M)"], aggregateList.UnitPrice, out unitPricePercentChange, 1); // passing flag one to use average
            }
            else
            {
                objectDictionary.Add("UnitPrice", aggregateList.UnitPrice);
            }
            if (model.Keys.Contains("unitCost($M)"))
            {
                objectDictionary["UnitCost"] = FindForRevenueCapex(model["unitCost($M)"], aggregateList.UnitCost, out unitCostPercentChange, 1); // passing flag one to use average 
            }
            else
            {
                objectDictionary.Add("UnitCost", aggregateList.UnitCost);
            }

            if (model.Keys.Contains("discountRate(%)"))
            {
                outputDiscountRate = ParseDouble(model["discountRate(%)"]);
                objectDictionary.Add("DiscountRate", model["discountRate(%)"]);
            }
            else
            {
                outputDiscountRate = capBudList[0].DiscountRate;
                objectDictionary.Add("DiscountRate", capBudList[0].DiscountRate);
            }
            if (model.Keys.Contains("marginalTax(%)"))
            {
                outputMarginalTax = ParseDouble(model["marginalTax(%)"]);
                objectDictionary.Add("MarginalTax", model["marginalTax(%)"]);
            }
            else
            {
                outputMarginalTax = capBudList[0].MarginalTaxRate;
                objectDictionary.Add("MarginalTax", capBudList[0].MarginalTaxRate);
            }
            if (model.Keys.Contains("capex($M)"))
            {
                objectDictionary["Capex"] = FindForRevenueCapex(model["capex($M)"], aggregateList.Capex, out capexDepChange, 0); // passing flag zero to use total
                objectDictionary["TotalDepreciation"] = FindForRevenueCapex(model["capex($M)"], aggregateList.TotalDepreciation, out _, 0); // passing flag zero to use total
            }
            else
            {
                objectDictionary.Add("Capex", aggregateList.Capex);
                objectDictionary.Add("TotalDepreciation", aggregateList.TotalDepreciation);
            }

            objectDictionary["Fixed"] = (FindFixedCost(model, capBudList, aggregateList.Fixed, out fixedCostObj));
            objectDictionary.Add("NWC", aggregateList.NWC);

            var summaryOutput = MathCapitalBudgeting.SummaryOutput(objectDictionary);
            List<string> keys = new List<string>();
            List<List<double>> values = new List<List<double>>();
            var yourDictionary = FixedCost(fixedCostObj, out keys, out values);
            List<string> depKey = new List<string>();
            List<List<double>> depValue = new List<List<double>>();
            List<string> capKeys = new List<string>();
            List<List<double>> capVaule = new List<List<double>>();
            Dictionary<string, object> result = new Dictionary<string, object>();

            result.Add("Sales ($M)", summaryOutput["sales"]);
            result.Add("COGS ($M)", summaryOutput["cogs"]);
            result.Add("Gross Margin ($M)", summaryOutput["grossMargins"]);

            if (data.OtherFixedCost.Length != 0)
            {
                for (int i = 0; i < keys.Count; i++)
                {
                    result.Add(keys[i], yourDictionary[keys[i]].Select(x => x = x / 1000000));
                }
            }
            if (data.CapexDepreciation.Length != 0)
            {
                depKey = GetDep(data.CapexDepreciation, out depValue, out capKeys, out capVaule);
                for (int i = 0; i < depKey.Count; i++)
                {
                    result.Add(depKey[i] + Convert.ToString(i + 1), depValue[i].Select(x => x = (x * capexDepChange) / 1000000));
                }
            }
            result.Add("Operating Income ($M)", summaryOutput["operatingIncomes"]);
            result.Add("Income Tax ($M)", summaryOutput["incomeTaxs"]);
            result.Add("Unlevered Net Income ($M)", summaryOutput["unleveredNetIcomes"]);
            result.Add("NWC(Net Working Capital)($M)", summaryOutput["nwcs"]);
            if (data.CapexDepreciation != null)
            {
                for (int i = 0; i < depKey.Count; i++)
                {
                    result.Add("Plus:" + " " + depKey[i] + Convert.ToString(i + 1), depValue[i].Select(x => x = (x * capexDepChange) / 1000000));
                }
                for (int i = 0; i < capKeys.Count; i++)
                {
                    result.Add("Less:" + " " + capKeys[i] + Convert.ToString(i + 1), capVaule[i].Select(x => x = (x * capexDepChange) / 1000000));
                }
            }
            result.Add("Less: Increases in NWC", summaryOutput["increaseNwcs"]);
            result.Add("Free Cash Flow", summaryOutput["freeCashFlows"]);
            result.Add("Levered Value (VL) ($M)", summaryOutput["leveredValues"]);
            result.Add("Discount Factor", summaryOutput["discountFactors"]);
            result.Add("Discount Cash Flow ($M)", summaryOutput["discountCashFlow"]);
            result.Add("Net Present Value ($M)", summaryOutput["npv"]);

            List<double> npv = (List<double>)summaryOutput["npv"];
            chageFixedCost = fixedCostObj;
            marginalTax = outputMarginalTax;
            discountRate = outputDiscountRate;
            volChangePerc = volPercentChange;
            unitPricePerChange = unitPricePercentChange;
            unitCostPerChange = unitCostPercentChange;
            capexDepPerChange = capexDepChange;
            sensiSceSummaryOutput = result;
            return npv[0];
        }

        //this function is to find cal the change in percent and  return the percent change for revenue table vol,unit price,unit cost        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <param name="aggregateList"></param>
        /// <param name="percentChange"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public static object FindForRevenueCapex(object value, List<double> aggregateList, out double percentChange, int flag)
        {
            List<double> finalAggList = new List<double>();
            double total = aggregateList.Sum();
            double average = total / aggregateList.Count;
            double changeInValue = 0;
            if (flag == 0) //Volume use sum
            {
                changeInValue = Convert.ToDouble(value) / total;
                foreach (double aggregateValues in aggregateList)
                {
                    finalAggList.Add(aggregateValues * changeInValue);
                }
            }
            else if (flag == 1)  /// Average unit price and unit cost   use average
            {
                changeInValue = Convert.ToDouble(value) / average;
                foreach (double aggregateValues in aggregateList)
                {
                    finalAggList.Add(aggregateValues * changeInValue);
                }
            }
            percentChange = changeInValue;

            return finalAggList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <param name="capBud"></param>
        /// <param name="aggregateListFixed"></param>
        /// <param name="fixedCostResult"></param>
        /// <returns></returns>
        public static object FindFixedCost(Dictionary<string, object> model, List<CapitalBudgeting> capBud,
            List<double> aggregateListFixed, out object[][][] fixedCostResult)
        {
            List<object> fixedCoList = new List<object>();
            TableData tableData = JsonConvert.DeserializeObject<TableData>(capBud[0].RawTableData);
            var keys = model.Keys.ToList();
            foreach (string key in keys)
            {
                foreach (object[] values in tableData.OtherFixedCost[0])
                {
                    if ((string)values[0] == key)
                    {
                        double sum = 0;
                        double changeVal = 0;
                        for (int i = 1; i < values.Length; i++)
                        {
                            sum += ParseDouble(values[i]);
                        }
                        var ku = model[key];
                        changeVal = Convert.ToDouble(model[key]) / sum;
                        List<object> obj = new List<object>();
                        obj.Add(key);
                        for (int j = 1; j < values.Length; j++)
                        {
                            obj.Add(Convert.ToDouble(ParseDoubleEx(values[j])) * changeVal);
                        }
                        for (int y = 0; y < tableData.OtherFixedCost[0].Length; y++)
                        {
                            if ((string)(tableData.OtherFixedCost[0][y][0]) == key)
                            {
                                obj.CopyTo(tableData.OtherFixedCost[0][y]);
                            }
                        }
                        obj.Clear();
                    }
                }
            }
            List<List<double>> fixedcostMultiList = CombineLists(tableData.OtherFixedCost, capBud[0].NoOfYears);
            List<double> fixedcostSum = ProvideAggregate(fixedcostMultiList, 0);///flag because the new fixed cost table is formed

            fixedCostResult = tableData.OtherFixedCost;

            return fixedcostSum;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Array"></param>
        /// <returns></returns>
        public static List<double> ExtractValues(object[] Array)
        {
            List<double> doubleArray = new List<double>();
            foreach (object[] value in Array)
            {
                for (int m = 1; m < value.Length; m++)
                {
                    doubleArray.Add(ParseDouble(value[m]));
                }
            }
            return doubleArray;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <param name="innerValue"></param>
        /// <param name="noOfYears"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        // public static List<List<double>> ProvideListOfList(int index, object[][][] innerValue, int noOfYears, int flag)
        // {
        //     List<List<double>> listOfList = new List<List<double>>();
        //     if (innerValue.Length != 0)
        //     {
        //         foreach (object[][] value in innerValue)
        //         {
        //             List<double> list = new List<double>();
        //             if ((value[index].Length - 1) == noOfYears || (value[index].Length) == noOfYears)
        //             {
        //                 if (flag == 0)
        //                 {
        //                     for (int i = 1; i < value[0].Length; i++)
        //                     {
        //                         list.Add(ParseDouble(value[index][i]));
        //                     }

        //                     listOfList.Add(list);
        //                 }
        //                 else if (flag == 1)
        //                 {
        //                     for (int i = 0; i < value[0].Length; i++)
        //                     {
        //                         list.Add(ParseDouble(value[index][i]));
        //                     }

        //                     listOfList.Add(list);
        //                 }
        //             }
        //             else
        //             {
        //                 return null;
        //             }
        //         }
        //     }
        //     else
        //     {
        //         List<double> list = new List<double>();
        //         for (int j = 0; j < noOfYears; j++)
        //         {
        //             list.Add(0);
        //         }
        //         listOfList.Add(list);
        //     }
        //     return listOfList;
        // }

public static List<List<double>> ProvideListOfList(int index, object[][][] innerValue, int noOfYears, int flag)
{
    List<List<double>> listOfList = new List<List<double>>();
    if (innerValue.Length != 0)
    {
        foreach (object[][] value in innerValue)
        {
            if (value.Length > index && value[index] != null && ((value[index].Length - 1) == noOfYears || value[index].Length == noOfYears))
            {
                List<double> list = new List<double>();
                if (flag == 0 && value[index].Length > 0)
                {
                    // Start from 1 to skip the first index if required
                    for (int i = 1; i < value[index].Length; i++)
                    {
                        list.Add(ParseDouble(value[index][i]));
                    }
                }
                else if (flag == 1 && value[index].Length > 0)
                {
                    // Include all entries in the list
                    for (int i = 0; i < value[index].Length; i++)
                    {
                        list.Add(ParseDouble(value[index][i]));
                    }
                }
                listOfList.Add(list);
            }
            else
            {
                // Return null or handle the case where array lengths are not matching the expected noOfYears
                // For instance, you might want to add an empty list or log a warning
                return null;
            }
        }
    }
    else
    {
        List<double> list = new List<double>();
        for (int j = 0; j < noOfYears; j++)
        {
            list.Add(0);
        }
        listOfList.Add(list);
    }
    return listOfList;
}


        /// <summary>
        /// 
        /// </summary>
        /// <param name="table"></param>
        /// <param name="noOfYears"></param>
        /// <returns></returns>
        public static List<List<double>> CombineLists(object[][][] table, int noOfYears)
        {
            List<List<double>> listOfList = new List<List<double>>();
            if (table.Length != 0)
            {
                foreach (object[][] outerValue in table)
                {
                    List<double> list = new List<double>();
                    foreach (object[] innerValue in outerValue)
                    {
                        if ((innerValue.Length - 1) == noOfYears)
                        {
                            List<double> lists = new List<double>();
                            for (int u = 1; u < innerValue.Length; u++)
                            {
                                lists.Add(ParseDouble(innerValue[u])); ;
                            }
                            listOfList.Add(lists);
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
            else
            {
                List<double> list = new List<double>();
                for (int j = 0; j < noOfYears; j++)
                {
                    list.Add(0);
                }
                listOfList.Add(list);
            }

            return listOfList;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="listOfList"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public static List<double> ProvideAggregate(List<List<double>> listOfList, int flag)
        {
                if (listOfList == null || listOfList.Count == 0)
                {
                    return new List<double> { 0 };
                }


            List<double> sum = new List<double>();

            for (int w = 0; w < listOfList[0].Count; w++)
            {
                double colSum = 0;
                for (int e = 0; e < listOfList.Count; e++)
                {
                    if (flag == 0)
                    {
                        colSum += (listOfList[e][w] / 1000000);
                    }
                    else
                    {
                        colSum += listOfList[e][w];
                    }

                }
                sum.Add(colSum);
            }
            return sum;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="values"></param>
        /// <param name="useSumRAverage"></param>
        /// <param name="str"></param>
        /// <param name="capitalBudgeting"></param>
        /// <returns></returns>
        public static Dictionary<string, List<List<double>>> ProvideMinMax(List<double> values, int useSumRAverage, string str, List<CapitalBudgeting> capitalBudgeting)
        {
            double sum = 0;
            double average = 0;
            double sumRAverage = 0;
            foreach (double result in values)
            {
                sum += result;
            }
            average = (sum / values.Count);
            if (useSumRAverage == 0)
            {
                sumRAverage = sum;
            }
            else
            {
                sumRAverage = average;
            }
            List<double> resultList = new List<double>();
            List<double> npvResultList = new List<double>();
            List<List<double>> completeResult = new List<List<double>>(); //This combine both range and npvlist
            resultList = Calculate(sumRAverage);
            npvResultList = NpvList(resultList, capitalBudgeting, str);

            if (str != "Volume(M)" && str != "MarginalTax(%)" && str != "DiscountRate(%)" && str != "UnitPrice($M)" && str != "UnitCost($M)" && str != "Capex($M)")
            {
                completeResult.Add(resultList.Select(x => x = x / 1000000).ToList());
            }
            else
            {
                completeResult.Add(resultList);
            }

            completeResult.Add(npvResultList);
            Dictionary<string, List<List<double>>> resultDict = new Dictionary<string, List<List<double>>>();
            resultDict.Add((str), completeResult);
            return resultDict;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <param name="str"></param>
        /// <param name="capitalBudgeting"></param>
        /// <returns></returns>
        public static Dictionary<string, List<List<double>>> MinAvgMax(double value, string str,
            List<CapitalBudgeting> capitalBudgeting)
        {
            List<double> resultList = new List<double>();
            List<double> npvResultList = new List<double>();
            List<List<double>> completeResult = new List<List<double>>(); //This combine both range and npvlist
            resultList = Calculate(value);
            npvResultList = NpvList(resultList, capitalBudgeting, str);
            completeResult.Add(resultList);
            completeResult.Add(npvResultList);
            Dictionary<string, List<List<double>>> resultDict = new Dictionary<string, List<List<double>>>();
            resultDict.Add((str), completeResult);
            return resultDict;
        }

        public static Dictionary<string, List<double>> FixedCost(object[][][] fixedCosts, out List<string>
            fixedKeys, out List<List<double>> fixedValues)
        {
            List<string> keys = new List<string>();
            Dictionary<string, List<double>> resultDict = new Dictionary<string, List<double>>();
            List<List<double>> valueList = new List<List<double>>();
            foreach (object[][] outerValue in fixedCosts)
            {
                foreach (object[] innerValue in outerValue)
                {
                    List<double> lists = new List<double>();
                    for (int u = 1; u < innerValue.Length; u++)
                    {
                        lists.Add(ParseDouble(innerValue[u]));
                    }
                    keys.Add(innerValue[0].ToString());
                    resultDict.Add(innerValue[0].ToString(), lists);
                    valueList.Add(lists);
                }
            }
            fixedKeys = keys;
            fixedValues = valueList;
            return resultDict;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="depList"></param>
        /// <param name="depValues"></param>
        /// <param name="capKeys"></param>
        /// <param name="CapValues"></param>
        /// <returns></returns>
       public static List<string> GetDep(object[][][] depList, out List<List<double>> depValues,
            out List<string> capKeys, out List<List<double>> CapValues)
        {
            List<string> Depkeys = new List<string>();
            List<List<double>> DepList = new List<List<double>>();
            List<string> Capkeys = new List<string>();
            List<List<double>> CapList = new List<List<double>>();

            for (int i = 0; i < depList.Length; i++)
            {
                // Ensure depList[i] is not null and has the required length
                if (depList[i] != null && depList[i].Length > 1 && depList[i][1] != null && depList[i][0] != null)
                {
                    List<double> listDep = new List<double>();
                    for (int j = 1; j < depList[i][1].Length; j++)
                    {
                        // Ensure that the current element is not null before parsing
                        if (depList[i][1][j] != null)
                        {
                            listDep.Add(ParseDouble(depList[i][1][j]));
                        }
                    }

                    List<double> listCap = new List<double>();
                    for (int j = 1; j < depList[i][0].Length; j++)
                    {
                        if (depList[i][0][j] != null)
                        {
                            listCap.Add(ParseDouble(depList[i][0][j]));
                        }
                    }

                    // Check that the first element exists before adding to the keys
                    if (depList[i][0].Length > 0 && depList[i][0][0] != null)
                    {
                        Capkeys.Add(depList[i][0][0].ToString());  // for capex
                    }

                    if (depList[i][1].Length > 0 && depList[i][1][0] != null)
                    {
                        Depkeys.Add(depList[i][1][0].ToString());  // for dep
                    }

                    CapList.Add(listCap);
                    DepList.Add(listDep);
                }
                else
                {
                    Console.WriteLine($"Warning: depList[{i}] is null or does not have the required structure.");
                }
            }

            capKeys = Capkeys;
            CapValues = CapList;
            depValues = DepList;
            return Depkeys;
        }

        public static List<double> Calculate(double value)
        {
            List<double> resultList = new List<double>();
            resultList.Add((value * 0.5));
            resultList.Add((value * 0.6));
            resultList.Add((value * 0.7));
            resultList.Add((value * 0.8));
            resultList.Add((value * 0.9));
            resultList.Add(value);
            resultList.Add((value * 1.1));
            resultList.Add((value * 1.2));
            resultList.Add((value * 1.3));
            resultList.Add((value * 1.4));
            resultList.Add(value * 1.5);
            return resultList;
        }

        // listOfRange can be volume range , marginaltax range , discount and fixed cost 

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listOfRange"></param>
        /// <param name="capitalBudgeting"></param>
        /// <param name="str"></param>
        /// <returns></returns>
        public static List<double> NpvList(List<double> listOfRange, List<CapitalBudgeting> capitalBudgeting, string str)
        {
            AggregateList aggregateList = JsonConvert.DeserializeObject<AggregateList>(capitalBudgeting[0].TableData);
            TableData tableData = JsonConvert.DeserializeObject<TableData>(capitalBudgeting[0].RawTableData);
            List<double> MarginalTaxList = new List<double>();
            MarginalTaxList.Add(capitalBudgeting[0].MarginalTaxRate);
            List<double> DiscountRateList = new List<double>();
            DiscountRateList.Add(capitalBudgeting[0].DiscountRate);
            List<double> npvList = new List<double>();

            if (str == "Volume(M)")
            {
                npvList = RevenueCapexNpv(listOfRange, aggregateList, MarginalTaxList, DiscountRateList, str);
            }
            else if (str == "UnitPrice($M)")
            {
                npvList = RevenueCapexNpv(listOfRange, aggregateList, MarginalTaxList, DiscountRateList, str);
            }
            else if (str == "UnitCost($M)")
            {
                npvList = RevenueCapexNpv(listOfRange, aggregateList, MarginalTaxList, DiscountRateList, str);
            }
            else if (str == "MarginalTax(%)")
            {
                npvList = MarginalTaxNpv(listOfRange, aggregateList, MarginalTaxList, DiscountRateList);
            }
            else if (str == "DiscountRate(%)")
            {
                npvList = DiscountNpv(listOfRange, aggregateList, MarginalTaxList, DiscountRateList, capitalBudgeting);
            }
            else if (str == "Capex($M)" && tableData.CapexDepreciation.Length != 0)
            {
                npvList = RevenueCapexNpv(listOfRange, aggregateList, MarginalTaxList, DiscountRateList, str);
            }
            else if (tableData.OtherFixedCost.Length != 0)
            {
                npvList = FixedCostNpv(listOfRange, aggregateList, MarginalTaxList, DiscountRateList, capitalBudgeting, str);
            }

            return npvList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listVolume"></param>
        /// <param name="aggregateList"></param>
        /// <param name="MarginalTaxList"></param>
        /// <param name="DiscountRateList"></param>
        /// <param name="str"></param>
        /// <returns></returns>
        public static List<double> RevenueCapexNpv(List<double> listVolume, AggregateList aggregateList,
            List<double> MarginalTaxList, List<double> DiscountRateList, string str)
        {
            List<double> npvList = new List<double>();

            double total = aggregateList.Volume.Sum();
            double averageUnitPrice = (aggregateList.UnitPrice.Sum()) / (aggregateList.UnitPrice.Count());
            double averageUnitCost = (aggregateList.UnitCost.Sum()) / (aggregateList.UnitCost.Count());
            double averageCapex = (aggregateList.Capex.Sum()) / (aggregateList.Capex.Count());
            foreach (double value in listVolume)
            {
                var objectDictionary = new Dictionary<string, object>();
                if (str == "Volume(M)")
                {
                    objectDictionary.Add("Volume", GenerateAggregateListList(aggregateList.Volume, total, value));
                }
                else
                {
                    objectDictionary.Add("Volume", aggregateList.Volume);
                }

                if (str == "UnitPrice($M)")
                {
                    objectDictionary.Add("UnitPrice", GenerateAggregateListList(aggregateList.UnitPrice, averageUnitPrice, value));
                }
                else
                {
                    objectDictionary.Add("UnitPrice", aggregateList.UnitPrice);
                }

                if (str == "UnitCost($M)")
                {
                    objectDictionary.Add("UnitCost", GenerateAggregateListList(aggregateList.UnitCost, averageUnitCost, value));
                }
                else
                {
                    objectDictionary.Add("UnitCost", aggregateList.UnitCost);
                }
                if (str == "Capex($M)")
                {
                    objectDictionary.Add("Capex", GenerateAggregateListList(aggregateList.Capex, averageCapex, value));
                    objectDictionary.Add("TotalDepreciation", GenerateAggregateListList(aggregateList.TotalDepreciation, averageCapex, value));

                }
                else
                {
                    objectDictionary.Add("Capex", aggregateList.Capex);
                    objectDictionary.Add("TotalDepreciation", aggregateList.TotalDepreciation);
                }

                objectDictionary.Add("Fixed", aggregateList.Fixed);
                objectDictionary.Add("NWC", aggregateList.NWC);
                objectDictionary.Add("MarginalTax", MarginalTaxList[0]);
                objectDictionary.Add("DiscountRate", DiscountRateList[0]);
                var summaryOutput = MathCapitalBudgeting.SummaryOutput(objectDictionary);
                List<double> npv = (List<double>)summaryOutput["npv"];
                npvList.Add(npv[0]);
            }
            return npvList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="aggrList"></param>
        /// <param name="totalRAverage"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static List<double> GenerateAggregateListList(List<double> aggrList, double totalRAverage, double value)
        {
            List<double> finalAggList = new List<double>();
            var volumeChange = value / totalRAverage;
            foreach (double aggregateValues in aggrList)
            {
                finalAggList.Add(aggregateValues * volumeChange);
            }
            return finalAggList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listDiscount"></param>
        /// <param name="aggregateList"></param>
        /// <param name="MarginalTaxList"></param>
        /// <param name="DiscountRateList"></param>
        /// <param name="capitalBudgeting"></param>
        /// <returns></returns>
        public static List<double> DiscountNpv(List<double> listDiscount, AggregateList aggregateList,
            List<double> MarginalTaxList, List<double> DiscountRateList, List<CapitalBudgeting> capitalBudgeting)
        {

            List<double> npvList = new List<double>();
            foreach (double value in listDiscount)
            {
                var objectDictionary = new Dictionary<string, object>();
                objectDictionary.Add("Volume", aggregateList.Volume);
                objectDictionary.Add("UnitPrice", aggregateList.UnitPrice);
                objectDictionary.Add("UnitCost", aggregateList.UnitCost);
                objectDictionary.Add("Fixed", aggregateList.Fixed);
                objectDictionary.Add("NWC", aggregateList.NWC);
                objectDictionary.Add("Capex", aggregateList.Capex);
                objectDictionary.Add("MarginalTax", MarginalTaxList[0]);
                objectDictionary.Add("DiscountRate", value);
                objectDictionary.Add("TotalDepreciation", aggregateList.TotalDepreciation);
                var summaryOutput = MathCapitalBudgeting.SummaryOutput(objectDictionary);
                List<double> npv = (List<double>)summaryOutput["npv"];
                npvList.Add(npv[0]);
            }
            return npvList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listMarignalTax"></param>
        /// <param name="aggregateList"></param>
        /// <param name="MarginalTaxList"></param>
        /// <param name="DiscountRateList"></param>
        /// <returns></returns>
        public static List<double> MarginalTaxNpv(List<double> listMarignalTax, AggregateList aggregateList,
            List<double> MarginalTaxList, List<double> DiscountRateList)
        {
            List<double> npvList = new List<double>();
            foreach (double value in listMarignalTax)
            {
                var objectDictionary = new Dictionary<string, object>();
                objectDictionary.Add("Volume", aggregateList.Volume);
                objectDictionary.Add("UnitPrice", aggregateList.UnitPrice);
                objectDictionary.Add("UnitCost", aggregateList.UnitCost);
                objectDictionary.Add("Fixed", aggregateList.Fixed);
                objectDictionary.Add("NWC", aggregateList.NWC);
                objectDictionary.Add("Capex", aggregateList.Capex);
                objectDictionary.Add("MarginalTax", value);
                objectDictionary.Add("DiscountRate", DiscountRateList[0]);
                objectDictionary.Add("TotalDepreciation", aggregateList.TotalDepreciation);
                var summaryOutput = MathCapitalBudgeting.SummaryOutput(objectDictionary);
                List<double> npv = (List<double>)summaryOutput["npv"];
                npvList.Add(npv[0]);
            }
            return npvList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listFixedCost"></param>
        /// <param name="aggregateList"></param>
        /// <param name="MarginalTaxList"></param>
        /// <param name="DiscountRateList"></param>
        /// <param name="capitalBudgeting"></param>
        /// <param name="str"></param>
        /// <returns></returns>
        public static List<double> FixedCostNpv(List<double> listFixedCost, AggregateList aggregateList,
            List<double> MarginalTaxList, List<double> DiscountRateList, List<CapitalBudgeting> capitalBudgeting, string str)
        {
            List<double> npvList = new List<double>();
            foreach (double divideBy in listFixedCost)
            {
                var rawTbData = JsonConvert.DeserializeObject<TableData>(capitalBudgeting[0].RawTableData);
                foreach (object[] values in rawTbData.OtherFixedCost[0])
                {
                    if ((string)values[0] == str)
                    {
                        double sum = 0;
                        double changeVal = 0;
                        for (int i = 1; i < values.Length; i++)
                        {
                            sum += ParseDouble(values[i]);
                        }
                        changeVal = divideBy / sum;
                        changeVal = ParseNaNInfinity(changeVal);
                        List<object> obj = new List<object>();
                        obj.Add(str);
                        for (int j = 1; j < values.Length; j++)
                        {
                            obj.Add(ParseDouble(values[j]) * changeVal);
                        }
                        for (int y = 0; y < rawTbData.OtherFixedCost[0].Length; y++)
                        {
                            if ((string)(rawTbData.OtherFixedCost[0][y][0]) == str)
                            {
                                obj.CopyTo(rawTbData.OtherFixedCost[0][y]);
                            }
                        }
                        obj.Clear();
                    }
                }
                List<List<double>> fixedcostMultiList = CombineLists(rawTbData.OtherFixedCost, capitalBudgeting[0].NoOfYears);
                List<double> fixedcostSum = ProvideAggregate(fixedcostMultiList, 0);
                var objectDictionary = new Dictionary<string, object>();

                objectDictionary.Add("Volume", aggregateList.Volume);
                objectDictionary.Add("UnitPrice", aggregateList.UnitPrice);
                objectDictionary.Add("UnitCost", aggregateList.UnitCost);
                objectDictionary.Add("Fixed", fixedcostSum);
                objectDictionary.Add("NWC", aggregateList.NWC);
                objectDictionary.Add("Capex", aggregateList.Capex);
                objectDictionary.Add("MarginalTax", MarginalTaxList[0]);
                objectDictionary.Add("DiscountRate", DiscountRateList[0]);
                objectDictionary.Add("TotalDepreciation", aggregateList.TotalDepreciation);
                var summaryOutput = MathCapitalBudgeting.SummaryOutput(objectDictionary);
                List<double> npv = (List<double>)summaryOutput["npv"];
                npvList.Add(npv[0]);
            }
            return npvList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="wsSheetName"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public static ExcelWorksheet ReturnCellStyle(string cellName, int row, int col, ExcelWorksheet wsSheetName, int flag)
        {
            wsSheetName.Cells[row, col].Value = cellName;
            wsSheetName.Cells[row, col].Style.Font.Size = 12;
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableArray"></param>
        /// <param name="str"></param>
        /// <param name="wsSheetName"></param>
        /// <param name="col_count"></param>
        /// <param name="vertical_count"></param>
        /// <param name="volPerCentChange"></param>
        /// <param name="unitPricePerChange"></param>
        /// <param name="unitCostPerChange"></param>
        /// <param name="capexPerChange"></param>
        /// <param name="row_count"></param>
        /// <param name="listOfFormula"></param>
        /// <returns></returns>
        public static ExcelWorksheet ExcelGeneration(Object[][][] tableArray, string str, ExcelWorksheet wsSheetName, int col_count, int vertical_count,
            double volPerCentChange, double unitPricePerChange, double unitCostPerChange, double capexPerChange, out int row_count, out List<List<List<string>>> listOfFormula)
        {   /// the volume needs to multipled with the percentchange for the change to value
            /// pass one when change ////else pass the change to be multipled
            int table_no = 1;
            List<List<List<string>>> outer = new List<List<List<string>>>();
            foreach (object[][] revTableValue in tableArray)
            {
                List<List<string>> middle = new List<List<string>>();
                for (int p = 0; p < revTableValue.Length; p++)
                {
                    ExcelAddress addr = null;
                    List<string> inner = new List<string>();
                    string combinedString = null;

                    for (int year = 0; year < revTableValue[p].Length; year++)
                    {

                        double sum = 0;

                        wsSheetName = ReturnCellStyle(str + " " + "Table" + "" + table_no.ToString(), (4 + vertical_count), 1, wsSheetName, 0);
                        wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (revTableValue[p][year]);
                        wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        if (year != 0 && str == "Revenue" && p == 0)  // volume 
                        {
                            wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (ParseDouble(revTableValue[p][year]) * volPerCentChange); //// this for change in percent in volumne

                        }
                        else if (year != 0 && str == "Revenue" && p == 1) // unit price
                        {
                            wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (ParseDouble(revTableValue[p][year]) * unitPricePerChange); //// this is for change in percent in unit price

                        }
                        else if (year != 0 && str == "Revenue" && p == 2) ///unit cost
                        {
                            wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (ParseDouble(revTableValue[p][year]) * unitCostPerChange);  /// this is for change in percent in unit cost

                        }
                        else if (year != 0 && str == "Capex&Depreciation")
                        {
                            wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (ParseDouble(revTableValue[p][year]) * capexPerChange);  /// this is for change in percent in capex

                        }
                        else
                        {
                            wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = ParseDoubleEx(revTableValue[p][year]);

                        }
                        addr = new ExcelAddress((5 + p + vertical_count), (col_count + year), (5 + p + vertical_count), (col_count + year));
                        if (year != (revTableValue[p].Length - 1) && str != "Capex&Depreciation")
                        {
                            sum += ParseDouble(revTableValue[p][year + 1]);
                            if (combinedString == null)
                            {
                                combinedString = addr.ToString();
                            }
                            else
                            {
                                combinedString = combinedString + "+" + addr;
                            }

                            wsSheetName.Cells[(5 + p + vertical_count), (revTableValue[p].Length + 2)].Value = sum;
                            wsSheetName.Cells[(5 + p + vertical_count), (revTableValue[p].Length + 2)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            wsSheetName.Cells[(5 + p + vertical_count), (revTableValue[p].Length + 3)].Value = (sum / (revTableValue[p].Length - 1));
                            wsSheetName.Cells[(5 + p + vertical_count), (revTableValue[p].Length + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        }

                        inner.Add(addr.ToString());
                    }
                    wsSheetName.Cells[addr.ToString()].Formula = "(" + combinedString + ")";

                    middle.Add(inner.Take(inner.Count() - 1).ToList());
                }

                vertical_count = vertical_count + revTableValue.Length + 1;
                table_no++;
                outer.Add(middle);
            }

            row_count = vertical_count;
            listOfFormula = outer;
            return wsSheetName;
        }
        public static ExcelWorksheet ExcelGeneration_New(List<ProjectInputDatasViewModel> projectVM, string str, ExcelWorksheet wsSheetName, int col_count, int vertical_count,
            double volPerCentChange, double unitPricePerChange, double unitCostPerChange, double capexPerChange, out int row_count, out List<List<List<string>>> listOfFormula)
        {   /// the volume needs to multipled with the percentchange for the change to value
            /// pass one when change ////else pass the change to be multipled
            /// 

            int table_no = 1;
            int p = 0;
            List<List<List<string>>> outer = new List<List<List<string>>>();
            List<List<string>> middle = new List<List<string>>();

            List<string> VolumeList = new List<string>();
            List<string> UnitPriceList = new List<string>();
            List<string> UnitCostList = new List<string>();
            List<string> SGAList = new List<string>();
            List<string> RDList = new List<string>();
            List<string> NWCList = new List<string>();
            List<string> CapexList = new List<string>();

            foreach (ProjectInputDatasViewModel TempValue in projectVM)
            {
                //List<List<string>> middle = new List<List<string>>();

                // for(int p=0;p<TempValue.ProjectInputValuesVM.Count; p++)
                // for (int p = 0; p < projectVM.Count; p++)
                // {
                ExcelAddress addr = null;

                List<string> inner = new List<string>();
                string combinedString = null;

                int ExecuteOnceForLineItem = 0;
                double sum = 0;
                long BasicValue = 0;
                // for(int year = 0;year<TempValue.ProjectInputValuesVM[p].Year;year++) 
                for (int year = 0; year < TempValue.ProjectInputValuesVM.Count; year++)
                {
                    //double sum = 0;
                    // wsSheetName = ReturnCellStyle(str + " " + "Table" + "" + table_no.ToString(), (4 + vertical_count), 1, wsSheetName, 0);
                    wsSheetName = ReturnCellStyle(str , (4 + vertical_count), 1, wsSheetName, 0);
                    //wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (TempValue.ProjectInputValuesVM[0][year]);

                    //wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem);
                    //wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                   
                    if (ExecuteOnceForLineItem == 0)
                    {
                        string UnitValue = null;

                        if (TempValue.UnitId == 1)
                        {
                            UnitValue = "$";
                            BasicValue = 1;
                        }
                        else if (TempValue.UnitId == 2)
                        {
                            UnitValue = "$K";
                            BasicValue = 1000;
                        }
                        else if (TempValue.UnitId == 3)
                        {
                            UnitValue = "$M";
                            BasicValue = 1000000;
                        }
                        else if (TempValue.UnitId == 4)
                        {
                            UnitValue = "$B";
                            BasicValue = 1000000000;
                        }
                        else if (TempValue.UnitId == 5)
                        {
                            UnitValue = "$T";
                            BasicValue = 1000000000000;
                        }

                        if (TempValue.LineItem == "Volume")
                        {                                                   
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;                            
                        }
                        else if (TempValue.LineItem == "Unit Price")
                        {
                            wsSheetName.Cells[(5 + 1 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(5 + 1 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;                            
                        }
                        else if (TempValue.LineItem == "Unit Cost")
                        {
                            wsSheetName.Cells[(5 + 2 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(5 + 2 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "R&D")
                        {
                            wsSheetName.Cells[(5 + 1 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(5 + 1 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "SG&A")
                        {
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "Capex")
                        {
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "Depreciation")
                        {
                            wsSheetName.Cells[(5 + 1 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(5 + 1 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "NWC")
                        {
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) + "(" + "%" + ")";
                            wsSheetName.Cells[(5 + 0 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                    }

                    //    if (OneTimeExecute == 1)
                    //{
                    //    // wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1)].Value = (TempValue.LineItem);
                    //    wsSheetName.Cells[(5 + p + vertical_count), (col_count + year)].Value = (TempValue.ProjectInputValuesVM[year].Value);
                    //    wsSheetName.Cells[(5 + p + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    //   // OneTimeExecute = OneTimeExecute + 1;
                    //}

                    //if (year != 0 && str == "Revenue" && p == 0)  // volume
                    //if (str == "Revenue" && p == 0)  // volume
                    if (TempValue.LineItem == "Volume")  // volume
                    {
                        // wsSheetName.Cells[(5 + p + vertical_count), (col_count  + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value) * volPerCentChange); //// this for change in percent in volumne
                        wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value) * volPerCentChange); //// this for change in percent in volumne
                        //wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        addr = new ExcelAddress((5 + 0 + vertical_count), (col_count + year), (5 + 0 + vertical_count), (col_count + year));
                        VolumeList.Add(addr.ToString());
                    }
                    //else if (str == "Revenue" && p == 1) // unit price
                    else if (TempValue.LineItem == "Unit Price") // unit price
                    {
                        wsSheetName.Cells[(5 + 1 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value) * unitPricePerChange); //// this is for change in percent in unit price
                        //wsSheetName.Cells[(5 + 1 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        addr = new ExcelAddress((5 + 1 + vertical_count), (col_count + year), (5 + 1 + vertical_count), (col_count + year));
                        UnitPriceList.Add(addr.ToString());
                    }
                    // else if (str == "Revenue" && p == 2) ///unit cost
                    else if (TempValue.LineItem == "Unit Cost") ///unit cost
                    {
                        wsSheetName.Cells[(5 + 2 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value) * unitCostPerChange);  /// this is for change in percent in unit cost
                        //wsSheetName.Cells[(5 + 2 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        addr = new ExcelAddress((5 + 2 + vertical_count), (col_count + year), (5 + 2 + vertical_count), (col_count + year));
                        UnitCostList.Add(addr.ToString());
                    }
                    else if (str == "Other Fixed Cost" && TempValue.LineItem == "R&D")
                    {
                        wsSheetName.Cells[(5 + 1 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value));
                        //wsSheetName.Cells[(5 + 1 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        addr = new ExcelAddress((5 + 1 + vertical_count), (col_count + year), (5 + 1 + vertical_count), (col_count + year));
                        RDList.Add(addr.ToString());
                    }
                    else if (str == "Other Fixed Cost" && TempValue.LineItem == "SG&A")
                    {
                        wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value));
                        // wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        addr = new ExcelAddress((5 + 0 + vertical_count), (col_count + year), (5 + 0 + vertical_count), (col_count + year));
                        SGAList.Add(addr.ToString());
                    }
                    else if (TempValue.LineItem == "Capex")
                    {
                        wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value) * capexPerChange);  /// this is for change in percent in capex
                        //wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        //wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Style.Locked = true;
                        addr = new ExcelAddress((5 + 0 + vertical_count), (col_count + year), (5 + 0 + vertical_count), (col_count + year));
                        CapexList.Add(addr.ToString());
                    }
                    else if (TempValue.LineItem == "Depreciation")
                    {
                        wsSheetName.Cells[(5 + 1 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value) * capexPerChange);  /// this is for change in percent in capex
                        //wsSheetName.Cells[(5 + 1 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        addr = new ExcelAddress((5 + 1 + vertical_count), (col_count + year), (5 + 1 + vertical_count), (col_count + year));
                    }
                    else if (str == "Working Capital" && TempValue.LineItem == "NWC")
                    {
                        wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value));
                        // wsSheetName.Cells[(5 + 0 + vertical_count), (col_count + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        addr = new ExcelAddress((5 + 0 + vertical_count), (col_count + year), (5 + 0 + vertical_count), (col_count + year));
                        NWCList.Add(addr.ToString());
                    }
                   
                    //addr = new ExcelAddress((5 + p + vertical_count), (col_count + year), (5 + p + vertical_count), (col_count + year));
                    // if (year != (TempValue.ProjectInputValuesVM.Count - 1) && str != "Capex&Depreciation")
                    if (year != (TempValue.ProjectInputValuesVM.Count) && str != "Capex & Depreciation")
                    {
                        // sum += ParseDouble(TempValue.ProjectInputValuesVM[year+1].Value);
                        sum += ParseDouble(TempValue.ProjectInputValuesVM[year].Value);
                        if (combinedString == null)
                        {
                            combinedString = addr.ToString();
                        }
                        else
                        {
                            combinedString = combinedString + "+" + addr;
                        }

                        if (TempValue.LineItem == "Volume")
                        {
                            if (year == TempValue.ProjectInputValuesVM.Count - 1)
                            {
                                wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Formula = "SUM(" + VolumeList[0] + ":" + VolumeList[year] + ")" ;
                                wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Formula ="AVERAGE(" + VolumeList[0] + ":" + VolumeList[year] + ")";
                            }
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = sum;
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputValuesVM.Count));
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                        else if (TempValue.LineItem == "Unit Price")
                        {
                            if (year == TempValue.ProjectInputValuesVM.Count - 1)
                            {
                                wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = "NA";
                                wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Formula = "SUMPRODUCT(" + VolumeList[0] + ":" + VolumeList[year] + "," +  UnitPriceList[0] + ":" + UnitPriceList[year] + ")/SUM(" + VolumeList[0] + ":" + VolumeList[year] + ")";
                            }
                            //wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = sum;
                            //wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            //wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputValuesVM.Count));
                            //wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                        else if (TempValue.LineItem == "Unit Cost")
                        {
                            if (year == TempValue.ProjectInputValuesVM.Count - 1)
                            {
                                wsSheetName.Cells[(5 + 2 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = "NA";
                                wsSheetName.Cells[(5 + 2 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Formula = "SUMPRODUCT(" + VolumeList[0] + ":" + VolumeList[year] + "," + UnitCostList[0] + ":" + UnitCostList[year] + ")/SUM(" + VolumeList[0] + ":" + VolumeList[year] + ")";
                            }
                            //wsSheetName.Cells[(5 + 2 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = sum;
                            //wsSheetName.Cells[(5 + 2 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                           // wsSheetName.Cells[(5 + 2 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputValuesVM.Count));
                            //wsSheetName.Cells[(5 + 2 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                        else if (TempValue.LineItem == "SG&A")
                        {
                            if (year == TempValue.ProjectInputValuesVM.Count - 1)
                            {
                                wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Formula = "SUM(" + SGAList[0] + ":" + SGAList[year] + ")";
                                wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Formula = "AVERAGE(" + SGAList[0] + ":" + SGAList[year] + ")";
                            }
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = sum;
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputValuesVM.Count));
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                        else if (TempValue.LineItem == "R&D")
                        {
                            if (year == TempValue.ProjectInputValuesVM.Count - 1)
                            {
                                wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Formula = "SUM(" + RDList[0] + ":" + RDList[year] + ")";
                                wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Formula = "AVERAGE(" + RDList[0] + ":" + RDList[year] + ")";
                            }
                            //wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = sum;
                            //wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            // wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputValuesVM.Count));
                            //wsSheetName.Cells[(5 + 1 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                        else if (TempValue.LineItem == "NWC")
                        {
                            if (year == TempValue.ProjectInputValuesVM.Count - 1)
                            {
                                wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = "NA";
                                wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Formula = "AVERAGE(" + NWCList[0] + ":" + NWCList[year] + ")";
                            }
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = sum;                            
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputValuesVM.Count));
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                        else if (TempValue.LineItem == "Capex")
                        {
                            if (year == TempValue.ProjectInputValuesVM.Count - 1)
                            {
                                wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Formula = "SUM(" + CapexList[0] + ":" + CapexList[year] + ")";
                            }                           
                        }

                        //wsSheetName.Cells[(5 + p + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Value = sum;
                        //               wsSheetName.Cells[(5 + p + vertical_count), (TempValue.ProjectInputValuesVM.Count + 3)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        //               wsSheetName.Cells[(5 + p + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputValuesVM.Count));
                        //               wsSheetName.Cells[(5 + p + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    }

                    //inner.Add(addr.ToString());
                    inner.Add(addr.ToString() + "*" + BasicValue);
                    //if (listOfFormulaNew.Count == 0)
                    //{
                    //    inner.Add("(" + addr.ToString() + "*" + BasicValue + ")");
                    //}
                    //else
                    //{
                    //    inner.Add("(" + addr.ToString() + "*" + BasicValue + ")" + "+" + listOfFormulaNew[0][p][year]);
                    //}
                }

                // TODO
                //wsSheetName.Cells[addr.ToString()].Formula = "(" + combinedString + ")";
                //middle.Add(inner.Take(inner.Count() - 1).ToList());
                middle.Add(inner.Take(inner.Count()).ToList());

                p = p + 1;
                //  }

                //TODO
                //    vertical_count = vertical_count + projectVM.Count + 1;
                //     table_no++;
                //    outer.Add(middle);
            }

            vertical_count = vertical_count + projectVM.Count + 1;
            table_no++;
            outer.Add(middle);

            row_count = vertical_count;
             listOfFormula = outer;
            //listOfFormula = middle;
            return wsSheetName;          
        }

        public static ExcelWorksheet ExcelGeneration_NewForVT4(List<ProjectInputDatasViewModel> projectVM, string str, ExcelWorksheet wsSheetName, int col_count, int vertical_count,
                                                               out int row_count)
        {   
            int table_no = 1;
            int p = 0;
            List<List<List<string>>> outer = new List<List<List<string>>>();
            List<List<string>> middle = new List<List<string>>();
            foreach (ProjectInputDatasViewModel TempValue in projectVM)
            {
                ExcelAddress addr = null;
                List<string> inner = new List<string>();
                string combinedString = null;

                List<string> IndustryList = new List<string>();

                int ExecuteOnceForLineItem = 0;
                double sum = 0;
                for (int year = 0; year < TempValue.ProjectInputComparablesVM.Count; year++)
                {                
                    wsSheetName = ReturnCellStyle(str, (1 + vertical_count), 1, wsSheetName, 0);
                  
                    if (ExecuteOnceForLineItem == 0)
                    {                       
                        if (TempValue.LineItem == "Project's Equity Cost of Capital (rE)")
                        {
                            wsSheetName.Cells[(2 + 0 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) ;
                            wsSheetName.Cells[(2 + 0 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "Project's Debt Cost of Capital (rD)")
                        {
                            wsSheetName.Cells[(2 + 1 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) ;
                            wsSheetName.Cells[(2 + 1 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "Project's Target Leverage (D/V) Ratio")
                        {
                            wsSheetName.Cells[(2 + 2 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem) ;
                            wsSheetName.Cells[(2 + 2 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                        else if (TempValue.LineItem == "Industry or Comparables Unlevered Cost of Capital (rU-Comp)")
                        {
                            wsSheetName.Cells[(2 + 3 + vertical_count), (col_count - 1 + year)].Value = (TempValue.LineItem);
                            wsSheetName.Cells[(2 + 3 + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }                        
                    }

                    if (TempValue.LineItem == "Project's Equity Cost of Capital (rE)")
                    {                    
                        wsSheetName.Cells[(2 + 0 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputComparablesVM[year].Value)) / 100;
                        wsSheetName.Cells[(2 + 0 + vertical_count), (col_count + year)].Style.Numberformat.Format = "0.00%";
                        addr = new ExcelAddress((2 + 0 + vertical_count), (col_count + year), (2 + 0 + vertical_count), (col_count + year));
                    }
                    else if (TempValue.LineItem == "Project's Debt Cost of Capital (rD)")
                    {
                        wsSheetName.Cells[(2 + 1 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputComparablesVM[year].Value)) / 100;
                        wsSheetName.Cells[(2 + 1 + vertical_count), (col_count + year)].Style.Numberformat.Format = "0.00%";
                        addr = new ExcelAddress((2 + 1 + vertical_count), (col_count + year), (2 + 1 + vertical_count), (col_count + year));
                    }
                    else if (TempValue.LineItem == "Project's Target Leverage (D/V) Ratio")
                    {
                        wsSheetName.Cells[(2 + 2 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputComparablesVM[year].Value)) / 100;
                        wsSheetName.Cells[(2 + 2 + vertical_count), (col_count + year)].Style.Numberformat.Format = "0.00%";
                        addr = new ExcelAddress((2 + 2 + vertical_count), (col_count + year), (2 + 2 + vertical_count), (col_count + year));
                    }
                    else if (TempValue.LineItem == "Industry or Comparables Unlevered Cost of Capital (rU-Comp)")
                    {
                        wsSheetName.Cells[(2 + 3 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputComparablesVM[year].Value)) / 100;
                        wsSheetName.Cells[(2 + 3 + vertical_count), (col_count + year)].Style.Numberformat.Format = "0.00%";
                        addr = new ExcelAddress((2 + 3 + vertical_count), (col_count + year), (2 + 3 + vertical_count), (col_count + year));
                        IndustryList.Add(addr.ToString());
                    }
                    
                    //addr = new ExcelAddress((2 + p + vertical_count), (col_count + year), (2 + p + vertical_count), (col_count + year));
                    // if (year != (TempValue.ProjectInputValuesVM.Count - 1) && str != "Capex&Depreciation")
                    if (year != (TempValue.ProjectInputComparablesVM.Count))
                    {
                        // sum += ParseDouble(TempValue.ProjectInputValuesVM[year+1].Value);
                        sum += ParseDouble(TempValue.ProjectInputComparablesVM[year].Value);
                        if (combinedString == null)
                        {
                            combinedString = addr.ToString();
                        }
                        else
                        {
                            combinedString = combinedString + "+" + addr;
                        }

                        if (TempValue.LineItem == "Industry or Comparables Unlevered Cost of Capital (rU-Comp)")
                        {
                            if (year == TempValue.ProjectInputComparablesVM.Count - 1)
                            {
                                wsSheetName.Cells[(2 + 4 + vertical_count), 2].Value = "Average (Industry or Comparables Unlevered Cost of Capital (rU-Comp))";
                                wsSheetName.Cells[(2 + 4 + vertical_count), 3].Formula = "AVERAGE(" + IndustryList[0] + ":" + IndustryList[year] + ")";
                                wsSheetName.Cells[(2 + 4 + vertical_count), 3].Style.Numberformat.Format = "0.00%";
                            }                  
                            //wsSheetName.Cells[(2 + 3 + vertical_count), (TempValue.ProjectInputComparablesVM.Count + 4)].Value = (sum / (TempValue.ProjectInputComparablesVM.Count));
                            //wsSheetName.Cells[(5 + 0 + vertical_count), (TempValue.ProjectInputValuesVM.Count + 4)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }                       
                    }
                    inner.Add(addr.ToString());
                }

                // TODO
                //wsSheetName.Cells[addr.ToString()].Formula = "(" + combinedString + ")";
                //middle.Add(inner.Take(inner.Count() - 1).ToList());
                middle.Add(inner.Take(inner.Count()).ToList());

                p = p + 1;
                //  }

                //TODO
                //    vertical_count = vertical_count + projectVM.Count + 1;
                //     table_no++;
                //    outer.Add(middle);
            }

            vertical_count = vertical_count + projectVM.Count + 1;
            table_no++;
            outer.Add(middle);

            row_count = vertical_count;
            //listOfFormula = outer;
            //listOfFormula = middle;
            return wsSheetName;
        }

        public static ExcelWorksheet ExcelGeneration_NewForVT6(List<ProjectInputDatasViewModel> projectVM, string str, ExcelWorksheet wsSheetName, int col_count, int vertical_count,
                                                               out int row_count, out List<List<List<string>>> listOfFormula)
        {   

            int table_no = 1;
            int p = 0;
            List<List<List<string>>> outer = new List<List<List<string>>>();
            List<List<string>> middle = new List<List<string>>();
            foreach (ProjectInputDatasViewModel TempValue in projectVM)
            {
                //List<List<string>> middle = new List<List<string>>();

                ExcelAddress addr = null;
                List<string> inner = new List<string>();
                string combinedString = null;

                int ExecuteOnceForLineItem = 0;

                for (int year = 0; year < TempValue.ProjectInputValuesVM.Count; year++)
                {                
                    //wsSheetName = ReturnCellStyle(str, (3 + vertical_count), 1, wsSheetName, 0);
                   
                    if (ExecuteOnceForLineItem == 0)
                    {
                        string UnitValue = null;

                        if (TempValue.UnitId == 1)
                        {
                            UnitValue = "$";
                        }
                        else if (TempValue.UnitId == 2)
                        {
                            UnitValue = "$K";
                        }
                        else if (TempValue.UnitId == 3)
                        {
                            UnitValue = "$M";
                        }
                        else if (TempValue.UnitId == 4)
                        {
                            UnitValue = "$B";
                        }
                        else if (TempValue.UnitId == 5)
                        {
                            UnitValue = "$T";
                        }

                        if (TempValue.LineItem == "Fixed Schedule & Predetermined Debt Level, Dt")
                        {
                            wsSheetName.Cells[(3 + 0 + vertical_count), (col_count - 2 + year)].Value = (TempValue.LineItem) + "(" + UnitValue + ")";
                            wsSheetName.Cells[(3 + 0 + vertical_count), (col_count - 2 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            ExecuteOnceForLineItem = ExecuteOnceForLineItem + 1;
                        }
                    }

                    if (TempValue.LineItem == "Fixed Schedule & Predetermined Debt Level, Dt") 
                    {                        
                        wsSheetName.Cells[(3 + 0 + vertical_count), (col_count + year)].Value = (ParseDouble(TempValue.ProjectInputValuesVM[year].Value));
                        wsSheetName.Cells[(3 + 0 + vertical_count), (col_count + year)].Style.Numberformat.Format = "$#,##0.00";
                    }
                    
                    addr = new ExcelAddress((3 + p + vertical_count), (col_count + year), (3 + p + vertical_count), (col_count + year));
                    if (year != (TempValue.ProjectInputValuesVM.Count))
                    {
                        if (combinedString == null)
                        {
                            combinedString = addr.ToString();
                        }
                        else
                        {
                            combinedString = combinedString + "+" + addr;
                        }                       
                    }
                    inner.Add(addr.ToString());
                }
                middle.Add(inner.Take(inner.Count()).ToList());
                p = p + 1;
            }

            vertical_count = vertical_count + projectVM.Count + 1;
            table_no++;
            outer.Add(middle);

            row_count = vertical_count;
            listOfFormula = outer;
            //listOfFormula = middle;
            return wsSheetName;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableArray"></param>
        /// <param name="str"></param>
        /// <param name="wsSheetName"></param>
        /// <param name="col_count"></param>
        /// <param name="vertical_count"></param>
        /// <param name="row_count"></param>
        /// <param name="resultAddList"></param>
        /// <returns></returns>
        public static ExcelWorksheet ExcelGenerationEx(Object[][] tableArray, string str, ExcelWorksheet wsSheetName, int col_count, int vertical_count, out int row_count,
            out List<string> resultAddList)
        {
            int table_no = 1;
            List<string> nwcAddList = new List<string>();
            for (int p = 0; p < tableArray.Length; p++)
            {
                for (int year = 0; year < tableArray[p].Length; year++)
                {
                    wsSheetName = ReturnCellStyle(str + " " + table_no.ToString(), (4 + vertical_count), 1, wsSheetName, 0);

                    wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left ;

                    if (year == 0)
                    {
                        wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = (tableArray[p][year]);
                    }
                    else
                    {
                        wsSheetName.Cells[(5 + p + vertical_count), (col_count - 1 + year)].Value = ParseDouble(tableArray[p][year]);
                    }

                    var nwcAdd = new ExcelAddress((5 + p + vertical_count), (col_count - 1 + year), (5 + p + vertical_count), (col_count - 1 + year));
                    if (year != 0 && year != tableArray[p].Length)
                    {
                        nwcAddList.Add(nwcAdd.ToString());
                    }

                }
                row_count = vertical_count + tableArray.Length + 1;
                table_no++;
            }
            row_count = vertical_count;
            resultAddList = nwcAddList;

            return wsSheetName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="salesResult"></param>
        /// <param name="cogsResult"></param>
        /// <returns></returns>
        private static List<string> RevenueFormula(List<List<List<string>>> list, out List<string> salesResult, out List<string> cogsResult)
        {
            List<string> salesList = new List<string>();
            List<string> cogsList = new List<string>();
            List<string> grossList = new List<string>();

                if (list == null || list.Count == 0 || list[0] == null || list[0].Count == 0 || list[0][0] == null)
                {
                    // Return empty results if the input list is not valid
                    salesResult = new List<string> { "N/A" };
                    cogsResult = new List<string> { "N/A" };
                    return new List<string> { "N/A" };
                }
            for (int w = 0; w < list[0][0].Count; w++)
            {
                string grossMargin = null;
                string sales = null;
                string cogs = null;
                int x = 0;
                for (int e = 0; e < list[0].Count; e++)
                {
                    string aggregate = null;
                    for (int q = 0; q < list.Count; q++)
                    {
                        if (aggregate == null)
                        {
                            aggregate = list[q][e][w];
                        }
                        else
                        {
                            aggregate = aggregate + "+" + list[q][e][w];
                        }

                    }
                    if (sales == null)
                    {
                        sales = aggregate;
                        cogs = aggregate;
                    }
                    else
                    {
                        //if (!sales.Contains('*'))
                        if (x == 0)
                        {
                            //sales = "(" + sales + "/" + "1000000" + ")" + "*" + "(" + aggregate + "/" + "1000000" + ")";  
                            //sales = "(" + sales + "*" + aggregate + ")" + "/" + "1000000";
                            //sales = "(" + sales + "*" + aggregate + ")";
                            sales = "(" + sales + ")" + "*" + "(" + aggregate + ")";
                            x = 1;
                        }
                        else
                        {
                            //cogs = "(" + cogs + "/" + "1000000" + ")" + "*" + "(" + aggregate + "/" + "1000000" + ")";
                            //cogs = "(" + cogs + "*" + aggregate + ")" + "/" + "1000000";
                            //cogs = "(" + cogs + "*" + aggregate + ")";
                            cogs = "(" + cogs + ")" + "*" + "(" + aggregate + ")";
                        }
                    }
                    grossMargin = sales + "-" + cogs;
                }
                salesList.Add(sales);
                cogsList.Add(cogs);
                grossList.Add(grossMargin);
            }

            salesResult = salesList;
            cogsResult = cogsList;

            return grossList;
        }

        public static List<string> RevenueFormula1(List<List<List<string>>> list, List<List<List<string>>> list1, out List<string> salesResult, out List<string> cogsResult)
        {
            List<string> salesList = new List<string>();
            List<string> cogsList = new List<string>();
            List<string> grossList = new List<string>();

            for (int w = 0; w < list[0][0].Count; w++)
            {
                string grossMargin = null;
                string sales = null;
                string cogs = null;

                string sales1 = null;
                string cogs1 = null;

                int x = 0;
                for (int e = 0; e < list[0].Count; e++)
                {
                    string aggregate = null;
                    string aggregate1 = null;

                    for (int q = 0; q < list.Count; q++)
                    {
                        if (aggregate == null)
                        {
                            aggregate = list[q][e][w];
                            aggregate1 = list1[q][e][w];
                        }
                        else
                        {
                            aggregate = aggregate + "+" + list[q][e][w];
                            aggregate1 = aggregate1 + "+" + list1[q][e][w];
                        }

                    }
                    if (sales == null)
                    {
                        sales = aggregate;
                        cogs = aggregate;

                        sales1 = aggregate1;
                        cogs1 = aggregate1;

                    }
                    else
                    {
                        if (x == 0)
                        {
                            sales = "(" + sales + ")" + "*" + "(" + aggregate + ")" + "+" + "(" + sales1 + ")" + "*" + "(" + aggregate1 + ")";
                            x = 1;
                        }
                        else
                        {
                            cogs = "(" + cogs + ")" + "*" + "(" + aggregate + ")" + "+" + "(" + cogs1 + ")" + "*" + "(" + aggregate1 + ")";
                        }
                    }
                    grossMargin = "(" + sales + ")" + "-" + "(" + cogs + ")";
                }
                salesList.Add(sales);
                cogsList.Add(cogs);
                grossList.Add(grossMargin);
            }

            salesResult = salesList;
            cogsResult = cogsList;

            return grossList;
        }

        public static List<string> RevenueFormula2(List<List<List<string>>> list, List<List<List<string>>> list1, List<List<List<string>>> list2, 
                                                   out List<string> salesResult, out List<string> cogsResult)
        {
            List<string> salesList = new List<string>();
            List<string> cogsList = new List<string>();
            List<string> grossList = new List<string>();

            for (int w = 0; w < list[0][0].Count; w++)
            {
                string grossMargin = null;
                string sales = null;
                string cogs = null;

                string sales1 = null;
                string cogs1 = null;

                string sales2 = null;
                string cogs2 = null;

                int x = 0;
                for (int e = 0; e < list[0].Count; e++)
                {
                    string aggregate = null;
                    string aggregate1 = null;
                    string aggregate2 = null;

                    for (int q = 0; q < list.Count; q++)
                    {
                        if (aggregate == null)
                        {
                            aggregate = list[q][e][w];
                            aggregate1 = list1[q][e][w];
                            aggregate2 = list1[q][e][w];
                        }
                        else
                        {
                            aggregate = aggregate + "+" + list[q][e][w];
                            aggregate1 = aggregate1 + "+" + list1[q][e][w];
                            aggregate2 = aggregate2 + "+" + list2[q][e][w];
                        }

                    }
                    if (sales == null)
                    {
                        sales = aggregate;
                        cogs = aggregate;

                        sales1 = aggregate1;
                        cogs1 = aggregate1;

                        sales2 = aggregate2;
                        cogs2 = aggregate2;
                    }
                    else
                    {
                        if (x == 0)
                        {
                            sales = "(" + sales + ")" + "*" + "(" + aggregate + ")" + "+" + "(" + sales1 + ")" + "*" + "(" + aggregate1 + ")" + "+" + "(" + sales2 + ")" + "*" + "(" + aggregate2 + ")";
                            x = 1;
                        }
                        else
                        {
                            cogs = "(" + cogs + ")" + "*" + "(" + aggregate + ")" + "+" + "(" + cogs1 + ")" + "*" + "(" + aggregate1 + ")" + "+" + "(" + cogs2 + ")" + "*" + "(" + aggregate2 + ")";
                        }
                    }
                    grossMargin = "(" + sales + ")" + "-" + "(" + cogs + ")";
                }
                salesList.Add(sales);
                cogsList.Add(cogs);
                grossList.Add(grossMargin);
            }

            salesResult = salesList;
            cogsResult = cogsList;

            return grossList;
        }

        public static List<string> OtherFixedCostFormula(List<List<List<string>>> list, out List<string> sgaResult, out List<string> rdResult)
        {
            List<string> sgaList = new List<string>();
            List<string> rdList = new List<string>();
            for (int w = 0; w < list[0][0].Count; w++)
            {
                string sga = null;
                string rd = null;

                for (int e = 0; e < list[0].Count; e++)
                {
                    string aggregate = null;
                    for (int q = 0; q < list.Count; q++)
                    {
                        if (aggregate == null)
                        {
                            aggregate = list[q][e][w];
                        }
                        else
                        {
                            aggregate = aggregate + "+" + list[q][e][w];
                        }

                    }
                    if (sga == null)
                    {
                        sga = aggregate;
                        rd = aggregate;
                    }
                    else
                    {
                        if (!sga.Contains('*'))
                        {
                          //  sga = "(" + sga + "/" + "1000000" + ")" + "*" + "(" + aggregate + "/" + "1000000" + ")";
                        }
                        else
                        {
                           // rd = "(" + rd + "/" + "1000000" + ")" + "*" + "(" + aggregate + "/" + "1000000" + ")";
                        }
                    }
                }
                sgaList.Add(sga);
                rdList.Add(rd);
            }

            sgaResult = sgaList;
            rdResult = rdList;

            return null;
        }

        //this function is to generate formula for operating income and capex&dep

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listOfList"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public static List<string> FormulaGeneration(List<List<string>> listOfList, int flag)
        {
            List<string> result = new List<string>();
            if (listOfList == null || listOfList.Count == 0 || listOfList[0] == null || listOfList[0].Count == 0)
            {
                return result; 
            }
            for (int i = 0; i < listOfList[0].Count(); i++)
            {
                string sum = null;
                for (int y = 0; y < listOfList.Count(); y++)
                {
                    if (sum == null)
                    {
                        sum = listOfList[y][i];
                    }
                    else
                    {
                        if (listOfList[y][i] != null && listOfList[y][i] != "")
                        {
                            if (flag == 0) /// formula generation for operating income
                            {
                                sum = sum + "-" + listOfList[y][i];
                            }
                            else if (flag == 1)
                            {
                                sum = sum + listOfList[y][i];
                            }
                        }
                    }
                }
                result.Add("(" + sum + ")");
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="array"></param>
        /// <returns></returns>
        public static object[][][] ConvertTo3DArray(List<List<List<object>>> array)
        {
            object[][][] result = new object[array.Count][][];
            for (int i = 0; i < array.Count; i++)
            {
                result[i] = new object[array[0].Count][];
                for (int j = 0; j < array[0].Count; j++)
                {
                    result[i][j] = new object[array[0][0].Count];
                    for (int m = 0; m < array[0][0].Count; m++)
                    {
                        result[i][j][m] = array[i][j][m];
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static Double ParseDouble(Object obj)
        {
            if ((obj == null) || (obj.ToString() == ""))
            {
                return 0;
            }

            return Convert.ToDouble(obj);
        }

        //This functon to add string if present
        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static Object ParseDoubleEx(Object obj)
        {
            if ((obj == null) || (obj.ToString() == ""))
            {
                return 0;
            }
            else if (obj.ToString() != "")
            {
                return Convert.ToString(obj);
            }
            else
            {
                return Convert.ToDouble(obj);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="changeVal"></param>
        /// <returns></returns>
        public static double ParseNaNInfinity(double changeVal)
        {
            if (Double.IsNaN(Convert.ToDouble(changeVal)) || Double.IsInfinity(Convert.ToDouble(changeVal)))
            {
                changeVal = 0;
            }
            return changeVal;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dictionary"></param>
        /// <returns></returns>
        public static List<List<object>> DictionaryToArray(Dictionary<string, object> dictionary)
        {
            List<List<object>> listOfList = new List<List<object>>();
            string[] keys = dictionary.Keys.ToArray();
            for (int i = 0; i < keys.Length; i++)
            {
                List<object> list = new List<object>();
                list.Add(keys[i]);
                var strings = ((IEnumerable)dictionary[keys[i]]).Cast<object>()
                                   .Select(x => x == null ? x : x.ToString())
                                   .ToArray();
                for (int j = 0; j < strings.Length; j++)
                {
                    list.Add(strings[j]);
                }
                listOfList.Add(list);
            }
            return listOfList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="addArray2D"></param>
        /// <param name="cellOne"></param>
        /// <param name="cellTwo"></param>
        /// <returns></returns>
        public static List<List<string>> GetCellAddress(string[,] addArray2D, string cellOne, string cellTwo)
        {
            List<List<string>> cellAddress = new List<List<string>>();
            if (addArray2D == null)
            {
                return cellAddress;
            }
            for (int k = 0; k < addArray2D.GetLength(0); k++)
            {
                if (addArray2D[k, 0] != null && addArray2D[k, 0].Contains(cellOne))
                {
                    for (int m = k + 1; m < addArray2D.GetLength(0); m++)
                    {
                        List<string> collectAddOperatingCal = new List<string>();
                        if (addArray2D[m, 0].Contains(cellTwo))
                            break;

                        for (int l = 1; l < addArray2D.GetLength(1); l++)
                        {
                            if (!addArray2D[m, 0].Contains("NWC (Net Working Capital)"))
                            {
                                if (addArray2D[m, 0].Contains("Plus"))
                                {
                                    collectAddOperatingCal.Add("+" + addArray2D[m, l]);
                                }
                                else if (addArray2D[m, 0].Contains("Less"))
                                {
                                    collectAddOperatingCal.Add("-" + addArray2D[m, l]);
                                }
                                else
                                {
                                    collectAddOperatingCal.Add(addArray2D[m, l]);
                                }
                            }
                            else
                            {
                                collectAddOperatingCal.Add("");
                            }
                        }

                        cellAddress.Add(collectAddOperatingCal);
                    }
                }
            }

            return cellAddress;
        }


        // V Changes 12-sep-2019
        [HttpGet]
        [Route("ProjectEvaluation/{UserId}/{ProjectId}")]
        public ActionResult<Object> GetProjectEvaluation(long UserId, long ProjectId)
        {
            try
            {

                //var capBud = iCapitalBudgeting.GetLatestSingle(s => s.UserId == UserId && s.ProjectId == ProjectId);
                var capBud = iCapitalBudgeting.GetSingle(s => s.UserId == UserId && s.ProjectId == ProjectId);
                if (capBud == null)
                {
                    return NotFound("No data found");
                }
                else
                {

                    Dictionary<string, object> objectDictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(capBud.TableData);
                    double MarginalTax = (double)objectDictionary["MarginalTax"];
                    double DiscountRate = (double)objectDictionary["DiscountRate"];
                    objectDictionary.Remove("MarginalTax");
                    objectDictionary.Remove("DiscountRate");


                    Dictionary<string, List<double>> objectDictionary2 = new Dictionary<string, List<double>>();
                    List<double> obj1;
                    foreach (var x in objectDictionary)
                    {
                        obj1 = new List<double>();
                        obj1 = JsonConvert.DeserializeObject<List<double>>(Convert.ToString(x.Value));
                        objectDictionary2.Add(x.Key, obj1);
                    }

                    obj1 = new List<double>();
                    obj1.Add(MarginalTax);
                    objectDictionary2.Add("MarginalTax", obj1);
                    obj1 = new List<double>();
                    obj1.Add(DiscountRate);
                    objectDictionary2.Add("DiscountRate", obj1);

                    Evaluation data = new Evaluation();
                    data.tableData = new TableData();
                    string rawTableData = capBud.RawTableData;
                    List<string> keys = new List<string>();
                    List<List<double>> values = new List<List<double>>();

                    data.tableData = JsonConvert.DeserializeObject<TableData>(capBud.RawTableData);

                    var summaryOutput = MathCapitalBudgeting.SummaryOutput_Double(objectDictionary2);
                    var yourDictionary = FixedCost(data.tableData.OtherFixedCost, out keys, out values);
                    List<string> depKey = new List<string>();
                    List<List<double>> depValue = new List<List<double>>();
                    List<string> capKeys = new List<string>();
                    List<List<double>> capVaule = new List<List<double>>();
                    Dictionary<string, object> result = new Dictionary<string, object>();

                    result.Add("Sales ($M)", summaryOutput["sales"]);
                    result.Add("COGS ($M)", summaryOutput["cogs"]);
                    result.Add("Gross Margin ($M)", summaryOutput["grossMargins"]);

                    if (data.tableData.OtherFixedCost != null)
                    {
                        for (int i = 0; i < keys.Count; i++)
                        {
                            result.Add(keys[i], yourDictionary[keys[i]].Select(x => x = x / toMill));
                        }
                    }

                    if (data.tableData.CapexDepreciation != null)
                    {
                        depKey = GetDep(data.tableData.CapexDepreciation, out depValue, out capKeys, out capVaule);
                        for (int i = 0; i < depKey.Count; i++)
                        {
                            result.Add(depKey[i] + Convert.ToString(i + 1), depValue[i].Select(x => x = x / toMill));
                        }
                    }

                    result.Add("Operating Income ($M)", summaryOutput["operatingIncomes"]);
                    result.Add("Income Tax ($M)", summaryOutput["incomeTaxs"]);
                    result.Add("Unlevered Net Income ($M)", summaryOutput["unleveredNetIcomes"]);
                    result.Add("NWC(Net Working Capital)($M)", summaryOutput["nwcs"]);

                    if (data.tableData.CapexDepreciation != null)
                    {
                        for (int i = 0; i < depKey.Count; i++)
                        {
                            result.Add("Plus:" + " " + depKey[i] + Convert.ToString(i + 1), depValue[i].Select(x => x = x / toMill));
                        }
                        for (int i = 0; i < capKeys.Count; i++)
                        {
                            result.Add("Less:" + " " + capKeys[i] + Convert.ToString(i + 1), capVaule[i].Select(x => x = x / toMill));
                        }
                    }

                    result.Add("Less: Increases in NWC", summaryOutput["increaseNwcs"]);
                    result.Add("Free Cash Flow", summaryOutput["freeCashFlows"]);
                    result.Add("Levered Value (VL) ($M)", summaryOutput["leveredValues"]);
                    result.Add("Discount Factor", summaryOutput["discountFactors"]);
                    result.Add("Discount Cash Flow ($M)", summaryOutput["discountCashFlow"]);
                    result.Add("Net Present Value ($M)", summaryOutput["npv"]);
                    //result.Add("npv", summaryOutput["npv"]);

                    var summaryString = JsonConvert.SerializeObject(DictionaryToArray(result));
                    List<object> res = new List<object>();
                    res.Add(summaryString);


                    var tableArray = iCapitalBugetingTables.FindBy(s => s.UserId == UserId && s.ProjectId == ProjectId).OrderByDescending(x=>x.Id).ToArray();
                    // var tableID = iCapitalBugetingTables.GetSingle(s => s.UserId == UserId && s.ProjectId == ProjectId).Id;
                    var tableID = tableArray.Length!=0 ? tableArray[0] :  null;
                    //return Ok(new
                    //{
                    //    result = new Approval
                    //    {
                    //        Evaluation = summaryString
                    //    },
                    //    message = "Success",
                    //    code = 200
                    //}

                    return Ok(new
                    {
                        //*remove capitalBudgeting Values when React Build Uploaded
                        result = res,
                        message = "Success",
                        id = capBud.Id,
                        tableID = tableID!=null ? tableID.Id :0,
                        code = 200,
                        capitalBudgeting = capBud,
                        startingYear= capBud.StartingYear,
                        approvalComment= capBud.ApprovalComment
                    }

                     );
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

    }
}
