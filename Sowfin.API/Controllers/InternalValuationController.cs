using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Sowfin.API.ViewModels.InternalValuation;
using Sowfin.Data.Abstract;
using Sowfin.Data.Common.Enum;
using Sowfin.Model.Entities;
using AutoMapper;
using Sowfin.API.ViewModels;
using Sowfin.Data.Common.Helper;
//using Sowfin.API.Controllers;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class InternalValuationController : ControllerBase
    {
        //private readonly FindataController findataController;
        private readonly ICostOfCapital_IValuation iCostOfCapital_IValuation = null;
        private readonly ITaxRates_IValuation iTaxRates_IValuation = null;
        private readonly IInitialSetup_IValuation iInitialSetup_IValuation = null;
        private readonly IPayoutPolicy_IValuation iPayoutPolicy_IValuation = null;
        private readonly IInterest_IValuation iInterest_IValuation = null;
        private readonly ICostOfCapital icostOfCapitalMaster = null;
        private readonly ICapitalStructure iCapitalStructure = null;
        private readonly IHistoryElementMapping iHistoryElementMapping = null;
        private readonly IIntegratedElementMaster iIntegratedElementMaster = null;
        private readonly IIntegratedFinancialStmt iIntegratedFinancialStmt = null;
        private readonly ILineItemInfoRepository iLineItemInfoRepository = null;
        private readonly IRawHistoricalValues iRawHistoricalValues = null;
        private readonly IHistoryAnalysisAndForcastRatio iHistoryAnalysisAndForcastRatio = null;
        private readonly IForcastRatioElementMaster iForcastRatioElementMaster = null;
        private readonly IExplicitPeriod_HistoryForcastRatio iExplicitPeriod_HistoryForcastRatio = null;
        private readonly IROICDatas iROICDatas = null;
        private readonly IROICValues iROICValues = null;
        private readonly IROIC_ExplicitValues iROIC_ExplicitValues = null;
        private readonly IIntegratedDatas iIntegratedDatas = null;
        private readonly IForcastRatioDatas iForcastRatioDatas = null;
        private readonly IIntegratedValues iIntegratedValues = null;
        private readonly IForcastRatioValues iForcastRatioValues = null;
        private readonly IIntegrated_ExplicitValues iIntegrated_ExplicitValues = null;
        private readonly IForcastRatio_ExplicitValues iForcastRatio_ExplicitValues = null;
        IFilings iFilings;
        IDatas iDatas;
        IValues iValues;
        IMapper mapper;
        ICIKStatus iCIKStatus;
        IHistory iHistory;
        public InternalValuationController(IInitialSetup_IValuation _iInitialSetup_IValuation, ICostOfCapital _iCostOfCapitalMaster, IMapper mapper, ICapitalStructure _iCapitalStructure,
            IPayoutPolicy_IValuation _iPayoutPolicy_IValuation, IInterest_IValuation _iInterest_IValuation, ICostOfCapital_IValuation _iCostOfCapital_IValuation,
            ITaxRates_IValuation _iTaxRates_IValuation, IHistoryElementMapping _iHistoryElementMapping, IIntegratedElementMaster _iIntegratedElementMaster,
            IIntegratedFinancialStmt _iIntegratedFinancialStmt, ILineItemInfoRepository _iLineItemInfoRepository, IRawHistoricalValues _iRawHistoricalValues,
            IHistoryAnalysisAndForcastRatio _iHistoryAnalysisAndForcastRatio, IForcastRatioElementMaster _iForcastRatioElementMaster, IFilings _iFilings, IDatas _iDatas,
            IValues _iValues, IROICDatas _iROICDatas, IROICValues _iROICValues, IROIC_ExplicitValues _iROIC_ExplicitValues,
            IExplicitPeriod_HistoryForcastRatio _iExplicitPeriod_HistoryForcastRatio, IIntegratedDatas _iIntegratedDatas, IIntegratedValues _iIntegratedValues,
            IIntegrated_ExplicitValues _iIntegrated_ExplicitValues, IForcastRatioDatas _iForcastRatioDatas, IForcastRatioValues _iForcastRatioValues,
            IForcastRatio_ExplicitValues _iForcastRatio_ExplicitValues, ICIKStatus _iCIKStatus, IHistory _iHistory/*, FindataController _findataController*/)
        {

            //findataController = _findataController;
            iInitialSetup_IValuation = _iInitialSetup_IValuation;
            iPayoutPolicy_IValuation = _iPayoutPolicy_IValuation;
            iInterest_IValuation = _iInterest_IValuation;
            iCostOfCapital_IValuation = _iCostOfCapital_IValuation;
            iTaxRates_IValuation = _iTaxRates_IValuation;
            icostOfCapitalMaster = _iCostOfCapitalMaster;
            iCapitalStructure = _iCapitalStructure;
            iHistoryElementMapping = _iHistoryElementMapping;
            iIntegratedElementMaster = _iIntegratedElementMaster;
            iIntegratedFinancialStmt = _iIntegratedFinancialStmt;
            iRawHistoricalValues = _iRawHistoricalValues;
            iLineItemInfoRepository = _iLineItemInfoRepository;
            iForcastRatioElementMaster = _iForcastRatioElementMaster;
            iHistoryAnalysisAndForcastRatio = _iHistoryAnalysisAndForcastRatio;
            iExplicitPeriod_HistoryForcastRatio = _iExplicitPeriod_HistoryForcastRatio;
            iFilings = _iFilings;
            iCIKStatus = _iCIKStatus;
            iDatas = _iDatas;
            iValues = _iValues;
            iROICDatas = _iROICDatas;
            iROICValues = _iROICValues;
            iIntegratedDatas = _iIntegratedDatas;
            iIntegratedValues = _iIntegratedValues;
            iIntegrated_ExplicitValues = _iIntegrated_ExplicitValues;
            iROIC_ExplicitValues = _iROIC_ExplicitValues;
            iForcastRatioDatas = _iForcastRatioDatas;
            iForcastRatioValues = _iForcastRatioValues;
            iForcastRatio_ExplicitValues = _iForcastRatio_ExplicitValues;
            this.mapper = mapper;
            iHistory = _iHistory;
        }

        #region  InitialSetup_IValuation 

        /// <summary>
        ///   Get InitialSetup By Id
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetInitialSetup/{UserId}")]
        public ActionResult<Object> GetInitialSetup(long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                var tblInitialSetUpObj = iInitialSetup_IValuation.FindBy(s => s.UserId == UserId && s.IsActive == true).OrderByDescending(x => x.Id).FirstOrDefault();
                if (tblInitialSetUpObj == null)
                {
                    return NotFound("No Valuation records found for the given UserId.");
                }

                Initialsetup_IValuationViewModel InitialSetUpObj = new Initialsetup_IValuationViewModel(); ;
                if (tblInitialSetUpObj == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    // map table to vm
                    InitialSetUpObj = mapper.Map<InitialSetup_IValuation, Initialsetup_IValuationViewModel>(tblInitialSetUpObj);

                    if (!string.IsNullOrEmpty(InitialSetUpObj.CIKNumber))
                    {
                        FilingsTable filingsTable = iFilings.GetSingle(x => x.CIK == InitialSetUpObj.CIKNumber);
                        if (filingsTable != null)
                            InitialSetUpObj.ParentCompany = filingsTable.CompanyName;
                    }
                }
                return Ok(InitialSetUpObj);
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }


        /// <summary>
        ///   Get InitialSetup By InitialSetupId
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetInitialSetupById/{Id}/{UserId}")]
        public ActionResult<Object> GetInitialSetupById(long Id, long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                SetAllInitialSetUpFalse(UserId);
                var InitialSetUptblObj = iInitialSetup_IValuation.FindBy(s => s.Id == Id).OrderByDescending(x => x.Id).First();

                if(InitialSetUptblObj == null)
                {
                    return NotFound("No Valuation records found for the given Id.");
                }
                Initialsetup_IValuationViewModel InitialSetUpObj = new Initialsetup_IValuationViewModel();
                if (InitialSetUptblObj == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    InitialSetUptblObj.IsActive = true;
                    iInitialSetup_IValuation.Update(InitialSetUptblObj);
                    iInitialSetup_IValuation.Commit();

                    // map table to vm
                    InitialSetUpObj = mapper.Map<InitialSetup_IValuation, Initialsetup_IValuationViewModel>(InitialSetUptblObj);

                    if (!string.IsNullOrEmpty(InitialSetUpObj.CIKNumber))
                    {
                        FilingsTable filingsTable = iFilings.GetSingle(x => x.CIK == InitialSetUpObj.CIKNumber);
                        if (filingsTable != null)
                            InitialSetUpObj.ParentCompany = filingsTable.CompanyName;
                    }


                }
                return Ok(InitialSetUpObj);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        ///   Get InitialSetup List By Id
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetInitialSetupListByUserId/{UserId}")]
        public ActionResult<Object> GetInitialSetupListByUserId(long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                List<Initialsetup_IValuationViewModel> InitialSetUpObj = new List<Initialsetup_IValuationViewModel>();

                var tblInitialSetUpObj = iInitialSetup_IValuation.FindBy(s => s.UserId == UserId) != null ? iInitialSetup_IValuation.FindBy(s => s.UserId == UserId).OrderByDescending(x => x.Id).ToList() : null;
                var tblFilingsListObj = iFilings.FindBy(s => s.Id != 0).ToList();
                if (tblInitialSetUpObj == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    if (tblFilingsListObj != null && tblFilingsListObj.Count > 0)
                        foreach (var obj in tblInitialSetUpObj)
                        {
                            Initialsetup_IValuationViewModel tempInitialSetUpObj = mapper.Map<InitialSetup_IValuation, Initialsetup_IValuationViewModel>(obj);
                            tempInitialSetUpObj.ParentCompany = tblFilingsListObj.Find(x => x.CIK == obj.CIKNumber) != null ? tblFilingsListObj.Find(x => x.CIK == obj.CIKNumber).CompanyName : "";
                            InitialSetUpObj.Add(tempInitialSetUpObj);
                        }
                }
                return Ok(InitialSetUpObj);
            }
            catch (Exception ss)
            {
                return BadRequest(new { StatusCode = 0, message = Convert.ToString(ss.Message) });
            }
        }

        [HttpGet]
        [Route("GetInitialSetupListByUserId_Integrated/{UserId}")]
        public ActionResult<Object> GetInitialSetupListByUserId_Integrated(long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                List<Initialsetup_IValuationViewModel> InitialSetUpObj = new List<Initialsetup_IValuationViewModel>();

                var tblIntegratedDatasObj = iIntegratedDatas.FindBy(s => s.Id != 0).ToList();
                var tblInitialSetUpObjAll = iInitialSetup_IValuation.FindBy(s => s.UserId == UserId).OrderByDescending(x => x.Id).ToList();

                //var tblInitialSetUpObj = iInitialSetup_IValuation.FindBy(s => s.UserId == UserId).OrderByDescending(x => x.Id).ToList();
                var tblFilingsListObj = iFilings.FindBy(s => s.Id != 0).ToList();
                if (tblInitialSetUpObjAll == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    List<InitialSetup_IValuation> tblInitialSetUpObj = new List<InitialSetup_IValuation>();

                    if (tblIntegratedDatasObj != null && tblIntegratedDatasObj.Count > 0)
                    {
                        foreach (InitialSetup_IValuation objTemp in tblInitialSetUpObjAll)
                        {
                            var intregatedData = tblIntegratedDatasObj.Find(x => x.InitialSetupId == objTemp.Id);
                            if (intregatedData != null)
                            {
                                tblInitialSetUpObj.Add(objTemp);
                            }
                        }
                    }


                    if (tblFilingsListObj != null && tblFilingsListObj.Count > 0)
                        foreach (var obj in tblInitialSetUpObj)
                        {
                            Initialsetup_IValuationViewModel tempInitialSetUpObj = mapper.Map<InitialSetup_IValuation, Initialsetup_IValuationViewModel>(obj);
                            tempInitialSetUpObj.ParentCompany = tblFilingsListObj.Find(x => x.CIK == obj.CIKNumber) != null ? tblFilingsListObj.Find(x => x.CIK == obj.CIKNumber).CompanyName : "";
                            InitialSetUpObj.Add(tempInitialSetUpObj);
                        }
                }
                return Ok(InitialSetUpObj);
            }
            catch (Exception ss)
            {
                return BadRequest(new { StatusCode = 0, message = Convert.ToString(ss.Message) });
            }
        }

        private bool SetAllInitialSetUpFalse(long UserId)
        {
            bool flag = false;
            //set false to all initial setup
            //update
            try
            {
                List<InitialSetup_IValuation> initialSetupListObj = iInitialSetup_IValuation.FindBy(x => x.UserId == UserId).ToList();

                if (initialSetupListObj != null && initialSetupListObj.Count > 0)
                {
                    foreach (InitialSetup_IValuation item in initialSetupListObj)
                    {
                        item.IsActive = false;
                    }
                    iInitialSetup_IValuation.UpdatedMany(initialSetupListObj);
                    iInitialSetup_IValuation.Commit();
                    flag = true;
                }
            }
            catch (Exception ss)
            {

            }
            return flag;
        }



        /// <summary>
        ///  Save internal Valuation Tax Rates
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("EvaluateInitialSetup")]
        public ActionResult<Object> EvaluateInitialSetup([FromBody] Initialsetup_IValuationViewModel model)
        {
            try
            {
                /////Check if same CIK year and description already exist in company
                var CHK_InitialSetUptblObj = iInitialSetup_IValuation.GetSingle(x => x.CIKNumber == model.CIKNumber && x.YearFrom == model.YearFrom && x.YearTo == model.YearTo && x.Company == model.Company);
                if (CHK_InitialSetUptblObj != null)
                {
                    return Ok(new
                    {
                        StatusCode = 3,
                        Message = "Data already exists for these inputs."
                    });
                }
                ///////////////////////////////////


                /////Data saved in CIKStatus table
                bool isDataSave = InsertCIKStatus(model.CIKNumber);
                ///////////////////////////////////

                //bool isCurrentyear = false;
                if (model.YearTo == Convert.ToInt32(DateTime.Now.Year.ToString()))
                {
                    //isCurrentyear = true;
                    var FilingList = iFilings.FindBy(x => x.CIK == model.CIKNumber).ToList();
                    var DataList = FilingList != null && FilingList.Count > 0 ? iDatas.FindBy(t => FilingList.Any(m => m.Id == t.FilingId)).ToList() : null;
                    var ValueList = DataList != null && DataList.Count > 0 ? iValues.FindBy(t => DataList.Any(m => m.Id == t.DataId && Convert.ToInt32(Convert.ToDateTime(t.FilingDate).Year) == model.YearTo)).ToList() : null;
                    if (ValueList == null || ValueList.Count == 0)
                    {
                        return Ok(new
                        {
                            StatusCode = 2,
                            Message = "No data available for year " + model.YearTo.ToString() + " please decrease Ending/Current Year value."
                        });
                    }
                }
                InitialSetup_IValuation InitialSetUptblObj = new InitialSetup_IValuation();
                SetAllInitialSetUpFalse(Convert.ToInt64(model.UserId));
                Initialsetup_IValuationViewModel InitialSetup_IValuationObj = new Initialsetup_IValuationViewModel();
                //add
                InitialSetUptblObj = new InitialSetup_IValuation
                {
                    CIKNumber = model.CIKNumber,
                    Company = model.Company,
                    YearFrom = model.YearFrom,
                    YearTo = model.YearTo,
                    SourceId = model.SourceId,
                    UserId = model.UserId,
                    CreatedDate = System.DateTime.UtcNow,
                    ModifiedDate = System.DateTime.UtcNow,
                    ExplicitYearCount = model.ExplicitYearCount,
                    IsActive = true
                };
                iInitialSetup_IValuation.Add(InitialSetUptblObj);
                iInitialSetup_IValuation.Commit();
                SetInitialToFilings(InitialSetUptblObj.Id, InitialSetUptblObj.CIKNumber);
                //// 
                //GetInitialSetupById
                // map table to vm
                InitialSetup_IValuationObj = mapper.Map<InitialSetup_IValuation, Initialsetup_IValuationViewModel>(InitialSetUptblObj);

                if (!string.IsNullOrEmpty(InitialSetup_IValuationObj.CIKNumber))
                {
                    FilingsTable filingsTable = iFilings.GetSingle(x => x.CIK == InitialSetup_IValuationObj.CIKNumber);
                    if (filingsTable != null)
                        InitialSetup_IValuationObj.ParentCompany = filingsTable.CompanyName;
                }
                return Ok(new
                {
                    StatusCode = 1,
                    Message = "Saved successfully.",
                    InitialSetup_IValuationObj = InitialSetup_IValuationObj
                });
                //return Ok(InitialSetup_IValuationObj);

            }
            catch (Exception ss)
            {
                return BadRequest(Convert.ToString(ss.Message));
            }
        }

        private bool InsertCIKStatus(string cik)
        {
            try
            {
                CIKStatus tabledat = iCIKStatus.GetSingle(x => x.CIK == cik);
                if (tabledat == null)
                {
                    CIKStatus CIKStatusObj = new CIKStatus();
                    CIKStatusObj.CIK = cik;
                    CIKStatusObj.Status = 0;
                    CIKStatusObj.Remark = "";
                    iCIKStatus.Add(CIKStatusObj);
                    iCIKStatus.Commit();
                }
                return true;
            }
            catch
            {
                return false;
            }

        }

        private bool SetInitialToFilings(long initialSetupId, string cik)
        {
            bool flag = false;
            List<FilingsTable> FilingsListInitialSetup = new List<FilingsTable>();
            FilingsListInitialSetup = iFilings.FindBy(x => x.CIK == cik).ToList();
            if (FilingsListInitialSetup != null && FilingsListInitialSetup.Count > 0)
            {
                foreach (FilingsTable Ft in FilingsListInitialSetup)
                {
                    Ft.InitialSetupId = initialSetupId;
                }
                iFilings.UpdatedMany(FilingsListInitialSetup);
                iFilings.Commit();
                flag = true;
            }
            return flag;
        }


        private InitialSetup_IValuation GetInitialSetupId(long UserId)
        {
            var getinitialSetupvar = iInitialSetup_IValuation.FindBy(s => s.UserId == UserId && s.IsActive == true) != null ? iInitialSetup_IValuation.FindBy(s => s.UserId == UserId && s.IsActive == true).OrderByDescending(x => x.Id).FirstOrDefault() : null;
            //var getinitialSetupvar = iInitialSetup_IValuation.FindBy(s => s.UserId == UserId && s.IsActive == true) != null ? iInitialSetup_IValuation.FindBy(s => s.UserId == UserId && s.IsActive == true).OrderByDescending(x => x.Id).First() : null;
            InitialSetup_IValuation getinitialSetup = getinitialSetupvar;
            return getinitialSetup;
        }


        #endregion

        #region  TaxRates_IValuation 

        /// <summary>
        ///   Get tax Rates By Id
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetTaxRates/{UserId}")]
        public ActionResult<Object> GetTaxRates(long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                var getinitialSetup = GetInitialSetupId(UserId);
                long InitialSetupId = getinitialSetup != null ? getinitialSetup.Id : 0;
                var TaxRatesObj = InitialSetupId != 0 ? iTaxRates_IValuation.GetSingle(s => s.InitialSetupId == InitialSetupId) : null;
                resultObject.id = 0;
                return Ok(new
                {
                    InitialSetup = getinitialSetup,
                    TaxRates = TaxRatesObj
                });
            }
            catch (Exception ss)
            {
                return BadRequest();
            }
        }

        /// <summary>
        ///  Save internal Valuation Tax Rates
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("EvaluateTaxRates")]
        public ActionResult<Object> EvaluateTaxRates([FromBody] InternalTaxRatesViewModel model)
        {
            try
            {
                if (model.Id == 0)
                {
                    //add 
                    TaxRates_IValuation TaxRates_IValuationObj = new TaxRates_IValuation
                    {
                        Statutory_Federal = model.Statutory_Federal,
                        Marginal = model.Marginal,
                        Operating = model.Operating,
                        CompanyId = model.CompanyId,
                        UserId = model.UserId,
                        CreatedDate = System.DateTime.UtcNow,
                        ModifiedDate = System.DateTime.UtcNow,
                        InitialSetupId = model.InitialSetupId
                    };

                    iTaxRates_IValuation.Add(TaxRates_IValuationObj);
                    iTaxRates_IValuation.Commit();

                    return Ok(TaxRates_IValuationObj);
                }
                else
                {
                    //update
                    TaxRates_IValuation TaxRates_IValuationObj = new TaxRates_IValuation
                    {
                        Id = model.Id,
                        Statutory_Federal = model.Statutory_Federal,
                        Marginal = model.Marginal,
                        Operating = model.Operating,
                        CompanyId = model.CompanyId,
                        UserId = model.UserId,
                        ModifiedDate = System.DateTime.UtcNow,
                        InitialSetupId = model.InitialSetupId
                    };
                    iTaxRates_IValuation.Update(TaxRates_IValuationObj);
                    iTaxRates_IValuation.Commit();

                    return Ok(TaxRates_IValuationObj);
                }
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }

        #endregion

        #region Cost Of Capital

        /// <summary>
        ///   Get GetInternalCostOfCapital By Id
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetInternalCostOfCapital/{UserId}")]
        public ActionResult<Object> GetInternalCostOfCapital(long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                double? defaultRiskFreeRate = null;
                double? defaultCostOfDebt = null;
                var getinitialSetup = GetInitialSetupId(UserId);
                long InitialSetupId = getinitialSetup != null ? getinitialSetup.Id : 0;
                var costofCapitalObj = InitialSetupId != 0 ? iCostOfCapital_IValuation.GetSingle(s => s.InitialSetupId == InitialSetupId) : null;
                //if (costofCapitalObj == null)
                //{
                var costOfcapital = icostOfCapitalMaster.FindBy(x => x.UserId == UserId).OrderByDescending(x => x.Id).ToArray();
                if (costOfcapital.Length != 0)
                {
                    defaultRiskFreeRate = costOfcapital[0].RiskFreeRate;
                }

                var CapitalStructure = iCapitalStructure.FindBy(s => s.UserId == UserId).OrderByDescending(x => x.Id).ToArray();
                if (CapitalStructure.Length != 0)
                {
                    defaultCostOfDebt = CapitalStructure[0].debt.CostOfDebt;
                }
                //}
                return Ok(new
                {
                    costOfCapital = costofCapitalObj,
                    InitialSetup = getinitialSetup,
                    defaultRiskFreeRate = defaultRiskFreeRate,
                    defaultCostOfDebt = defaultCostOfDebt
                });
            }
            catch (Exception ss)
            {
                return BadRequest(ss.Message.ToString());
            }
        }


        /// <summary>
        ///  Save internal Valuation Cost of Capital
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("EvaluateInternalCostOfCapital")]
        public ActionResult<Object> EvaluateInternalCostOfCapital([FromBody] InternalCostOfCapitalViewModel model)
        {
            try
            {
                if (model.Id == 0)
                {
                    //add 
                    CostOfCapital_IValuation CostOfCapital_IValuationObj = new CostOfCapital_IValuation
                    {
                        RiskFreeRate = model.RiskFreeRate,
                        CostOfDebt = model.CostOfDebt,
                        WeightedAverage = model.WeightedAverage,
                        CompanyId = model.CompanyId,
                        UserId = model.UserId,
                        CreatedDate = System.DateTime.UtcNow,
                        ModifiedDate = System.DateTime.UtcNow,
                        InitialSetupId = model.InitialSetupId
                    };

                    iCostOfCapital_IValuation.Add(CostOfCapital_IValuationObj);
                    iCostOfCapital_IValuation.Commit();

                    return Ok(CostOfCapital_IValuationObj);
                }
                else
                {
                    //get cost of capital and check if Discount Rate or WACC changed or not
                    CostOfCapital_IValuation CostOfCapital_IValuationObj = iCostOfCapital_IValuation.GetSingle(x => x.Id == model.Id);
                    if (CostOfCapital_IValuationObj != null)
                    {
                        decimal? prevValue = CostOfCapital_IValuationObj.WeightedAverage;
                        CostOfCapital_IValuationObj.RiskFreeRate = model.RiskFreeRate;
                        CostOfCapital_IValuationObj.CostOfDebt = model.CostOfDebt;
                        CostOfCapital_IValuationObj.WeightedAverage = model.WeightedAverage;
                        CostOfCapital_IValuationObj.CompanyId = model.CompanyId;
                        CostOfCapital_IValuationObj.ModifiedDate = System.DateTime.UtcNow;

                        iCostOfCapital_IValuation.Update(CostOfCapital_IValuationObj);
                        iCostOfCapital_IValuation.Commit();

                        if (prevValue != model.WeightedAverage)
                        {
                            //change Discount Rate in ROIC and Reflexes 

                            List<ROICDatas> ROICDatasList = new List<ROICDatas>();
                            List<ROICValues> allROICValuesList = new List<ROICValues>();
                            List<ROIC_ExplicitValues> allROIC_ExplicitValuesList = new List<ROIC_ExplicitValues>();

                            ROICDatasList = iROICDatas.FindBy(x => x.InitialSetupId == model.InitialSetupId).ToList();

                            if (ROICDatasList != null && ROICDatasList.Count > 0)
                            {
                                //get all Hisstorical Values of ROIC
                                //allROICValuesList = iROICValues.FindBy(t => ROICDatasList.Any(m => m.Id == t.ROICDatasId && m.StatementTypeId==(int)StatementTypeEnum.DCF2)).ToList();
                                //if (allROICValuesList != null && allROICValuesList.Count > 0)
                                //{
                                //    iROICValues.DeleteMany(allROICValuesList);
                                //    iROICValues.Commit();
                                //}

                                ////get all Explicit Values of ROIC
                                //allROIC_ExplicitValuesList = iROIC_ExplicitValues.FindBy(t => ROICDatasList.Any(m => m.Id == t.ROICDatasId)).ToList();
                                //if (allROIC_ExplicitValuesList != null && allROIC_ExplicitValuesList.Count > 0)
                                //{
                                //    iROIC_ExplicitValues.DeleteMany(allROIC_ExplicitValuesList);
                                //    iROIC_ExplicitValues.Commit();
                                //}

                                //List<Integrated_ExplicitValues> dumyexplicitValuesList = new List<Integrated_ExplicitValues>();
                                //List<IntegratedValues> dumyValuesList = new List<IntegratedValues>();
                                //Integrated_ExplicitValues explicitValue;
                                //InitialSetup_IValuation InitialSetup_IValuationObj = iInitialSetup_IValuation.GetSingle(x=>x.Id==model.InitialSetupId);

                                //if (InitialSetup_IValuationObj != null)
                                //{
                                //    IntegratedValues integratedvalue;
                                //    int startyear = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                                //    for (int i = startyear; i <= InitialSetup_IValuationObj.YearTo; i++)
                                //    {
                                //        integratedvalue = new IntegratedValues();
                                //        integratedvalue.Year = Convert.ToString(startyear);
                                //        integratedvalue.Value = "";
                                //        dumyValuesList.Add(integratedvalue);
                                //        startyear = startyear + 1;
                                //    }
                                //    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearTo);
                                //    for (int i = 1; i <= InitialSetup_IValuationObj.ExplicitYearCount + 1; i++)
                                //    {
                                //        explicitValue = new Integrated_ExplicitValues();
                                //        year = year + 1;
                                //        explicitValue.Year = Convert.ToString(year);
                                //        explicitValue.Value = "";
                                //        dumyexplicitValuesList.Add(explicitValue);
                                //    }



                                string discountRate = "";
                                string growthduring_Terminal = "";
                                string RONIC = "";
                                string PVof_FCF = "";
                                string Baseof_Terinal = "";
                                string PVof_Terinal = "";
                                int totalYearCount = 0;
                                List<ROICDatas> CalculatedROICDatasList = ROICDatasList.FindAll(x => x.StatementTypeId != (int)StatementTypeEnum.ROIC && x.StatementTypeId != (int)StatementTypeEnum.FCF).ToList();

                                //get all Explicit Values of ROIC
                                allROIC_ExplicitValuesList = iROIC_ExplicitValues.FindBy(t => CalculatedROICDatasList.Any(m => m.Id == t.ROICDatasId)).ToList();
                                if (allROIC_ExplicitValuesList != null && allROIC_ExplicitValuesList.Count > 0)
                                {
                                    iROIC_ExplicitValues.DeleteMany(allROIC_ExplicitValuesList);
                                    iROIC_ExplicitValues.Commit();
                                }

                                foreach (ROICDatas ROICDatasObj in CalculatedROICDatasList)
                                {
                                    if (ROICDatasObj.LineItem == "Discount Rate")
                                    {
                                        //update value
                                        ROICDatasObj.DtValue = Convert.ToString(model.WeightedAverage);
                                        iROICDatas.Update(ROICDatasObj);
                                        iROICDatas.Commit();
                                        //get discount Rate from Cost of Capital
                                        discountRate = ROICDatasObj.DtValue;
                                    }
                                    else if (ROICDatasObj.LineItem == "Growth During Terminal Period")
                                    {
                                        growthduring_Terminal = ROICDatasObj.DtValue;
                                    }
                                    else if (ROICDatasObj.LineItem == "Return on New Invested Capital (RONIC)")
                                    {
                                        RONIC = ROICDatasObj.DtValue;
                                    }
                                    else if (ROICDatasObj.LineItem == "DCF")
                                    {
                                        int i = 1;
                                        List<ROIC_ExplicitValues> roic_ExplicitValuesList = allROIC_ExplicitValuesList.FindAll(x => x.ROICDatasId == ROICDatasObj.Id);
                                        ROICDatas FCFDatas = ROICDatasList.Find(x => x.LineItem == "Free Cash Flow (FCF)");
                                        //saave ROICValues

                                        foreach (ROIC_ExplicitValues ROICitem in roic_ExplicitValuesList)
                                        {
                                            double value = 0;
                                            ROIC_ExplicitValues FCFExplicitValue = FCFDatas != null & FCFDatas.ROIC_ExplicitValues != null && FCFDatas.ROIC_ExplicitValues.Count > 0 ? FCFDatas.ROIC_ExplicitValues.Find(x => x.Year == ROICitem.Year) : null;

                                            //Math.Pow(100.00, 3.00)
                                            value = FCFExplicitValue != null && !string.IsNullOrEmpty(FCFExplicitValue.Value) && FCFExplicitValue.Value != "0" ? Convert.ToDouble(FCFExplicitValue.Value) / Math.Pow(1 + (!string.IsNullOrEmpty(discountRate) ? Convert.ToDouble(discountRate) : 0), i) : 0;


                                            ROICitem.Value = value.ToString("0.");
                                            i++;
                                        }

                                        iROIC_ExplicitValues.UpdatedMany(roic_ExplicitValuesList);
                                        iROIC_ExplicitValues.Commit();

                                    }
                                    else if (ROICDatasObj.LineItem == "PV of FCF During Explicit Forecast Period")
                                    {
                                        ROICDatas FCFDatas = ROICDatasList.Find(x => x.LineItem == "DCF");
                                        //get all the Explicit Values of DCF
                                        List<ROIC_ExplicitValues> dcfExplicitList = FCFDatas != null && FCFDatas.ROIC_ExplicitValues != null && FCFDatas.ROIC_ExplicitValues.Count > 0 ? FCFDatas.ROIC_ExplicitValues : null;
                                        if (dcfExplicitList != null && dcfExplicitList.Count > 0)
                                        {
                                            double value = 0;
                                            foreach (var dcfValue in dcfExplicitList)
                                            {
                                                value = value + (dcfValue != null && !string.IsNullOrEmpty(dcfValue.Value) ? Convert.ToDouble(dcfValue.Value) : 0);
                                            }
                                            PVof_FCF = Convert.ToString(value.ToString("0.##"));
                                            ROICDatasObj.DtValue = PVof_FCF;

                                            iROICDatas.Update(ROICDatasObj);
                                            iROICDatas.Commit();
                                        }
                                    }
                                    //Base for Terminal Value "Base for Terminal Value"
                                    else if (ROICDatasObj.LineItem == "Base for Terminal Value")
                                    {
                                        //Baseof_Terinal
                                        ROICDatas noPlatDatas = ROICDatasList.Find(x => x.LineItem == "NOPLAT");
                                        //get all the Explicit Values of DCF
                                        List<ROIC_ExplicitValues> dcfExplicitList = noPlatDatas != null && noPlatDatas.ROIC_ExplicitValues != null && noPlatDatas.ROIC_ExplicitValues.Count > 0 ? noPlatDatas.ROIC_ExplicitValues : null;
                                        if (dcfExplicitList != null && dcfExplicitList.Count > 0)
                                        {
                                            double NoplatTerminalvalue = 0;
                                            double value = 0;
                                            foreach (var dcfValue in dcfExplicitList)
                                            {
                                                NoplatTerminalvalue = (dcfValue != null && !string.IsNullOrEmpty(dcfValue.Value) ? Convert.ToDouble(dcfValue.Value) : 0);
                                            }
                                            value = ((NoplatTerminalvalue) * (1 - ((!string.IsNullOrEmpty(growthduring_Terminal) ? Convert.ToDouble(growthduring_Terminal) : 0) / (!string.IsNullOrEmpty(RONIC) ? Convert.ToDouble(RONIC) : 0)))) / (((!string.IsNullOrEmpty(discountRate) ? Convert.ToDouble(discountRate) : 0) - (!string.IsNullOrEmpty(growthduring_Terminal) ? Convert.ToDouble(growthduring_Terminal) : 0)) / 100);

                                            //PVof_FCF = Convert.ToString(value.ToString("0.##"));

                                            Baseof_Terinal = Convert.ToString(value.ToString("0.##"));
                                            ROICDatasObj.DtValue = Baseof_Terinal;

                                            iROICDatas.Update(ROICDatasObj);
                                            iROICDatas.Commit();
                                            totalYearCount = dcfExplicitList.Count;
                                        }
                                    }
                                    else if (ROICDatasObj.LineItem == "PV of Terminal Value")
                                    {

                                        //Baseof_Terinal
                                        double value = 0;
                                        value = (!string.IsNullOrEmpty(Baseof_Terinal) ? Convert.ToDouble(Baseof_Terinal) : 0) / Math.Pow((1 + ((!string.IsNullOrEmpty(discountRate) ? Convert.ToDouble(discountRate) : 0) / 100)), (totalYearCount));


                                        PVof_Terinal = Convert.ToString(value.ToString("0.##"));
                                        ROICDatasObj.DtValue = PVof_Terinal;

                                        iROICDatas.Update(ROICDatasObj);
                                        iROICDatas.Commit();

                                    }
                                    else if (ROICDatasObj.LineItem == "Total Value of Operations")
                                    {


                                        //Baseof_Terinal
                                        double value = 0;
                                        value = (!string.IsNullOrEmpty(PVof_Terinal) ? Convert.ToDouble(PVof_Terinal) : 0) + (!string.IsNullOrEmpty(PVof_FCF) ? Convert.ToDouble(PVof_FCF) : 0);


                                        ROICDatasObj.DtValue = Convert.ToString(value.ToString("0.##"));
                                        //ROICDatasObj.DtValue = PVof_Terinal;

                                        iROICDatas.Update(ROICDatasObj);
                                        iROICDatas.Commit();

                                    }

                                }

                            }

                        }

                    }

                    return Ok(CostOfCapital_IValuationObj);
                }

                ////update
                //CostOfCapital_IValuation CostOfCapital_IValuationObj = new CostOfCapital_IValuation
                //{
                //    Id = model.Id,
                //    RiskFreeRate = model.RiskFreeRate,
                //    CostOfDebt = model.CostOfDebt,
                //    WeightedAverage = model.WeightedAverage,
                //    CompanyId = model.CompanyId,
                //    UserId = model.UserId,
                //    ModifiedDate = System.DateTime.Now,
                //    InitialSetupId = model.InitialSetupId
                //};




            }
            catch (Exception ss)
            {
                return BadRequest(Convert.ToString(ss.Message));
            }
        }

        #endregion

        #region Interest

        /// <summary>
        ///   Get GetInterest_Ivaluation By Id
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetInterest_IValuation/{UserId}")]
        public ActionResult<Object> GetInterest_IValuation(long UserId)
        {

            try
            {
                ResultObject resultObject = new ResultObject();
                InitialSetup_IValuation InitialSetup_IValuationObj = GetInitialSetupId(UserId);
                long InitialSetupId = InitialSetup_IValuationObj != null ? InitialSetup_IValuationObj.Id : 0;
                var interest_IValuationListObj = InitialSetupId != 0 ? iInterest_IValuation.FindBy(s => s.InitialSetupId == InitialSetupId && s.UserId == UserId).ToArray() : null;

                //Date:23-July-2020 | Added By : anonymous | Enh. : Single Units Enum
                List<ValueTextWrapper> currencyValueList = EnumHelper.GetEnumDescriptions<CurrencyValueEnum>();


                // single data for every row
                DataTable dtInterestIncome = new DataTable();
                dtInterestIncome.Columns.Add("year");
                dtInterestIncome.Columns.Add("yearValue");
                int yearcount = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) - Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtInterestIncome.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtInterestIncome.Rows.Add(newRow);
                    year = year + 1;
                }
                DataTable dtInterestexpense = new DataTable();
                dtInterestexpense.Columns.Add("year");
                dtInterestexpense.Columns.Add("yearValue");
                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtInterestexpense.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtInterestexpense.Rows.Add(newRow);
                    year = year + 1;
                }
                //for (int i = 0; i < interest_IValuationListObj.Length; i++)
                //{
                //    dtInterestIncome.Rows[i][1] = interest_IValuationListObj[i].Interest_Income;
                //    dtInterestexpense.Rows[i][1] = interest_IValuationListObj[i].Interest_Expense;
                //}
                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                for (int i = 0; i <= yearcount; i++)
                {
                    var ValueObj = interest_IValuationListObj.FirstOrDefault(x => x.Year == year);
                    if (ValueObj != null)
                    {
                        dtInterestIncome.Rows[i][1] = ValueObj.Interest_Income;
                        dtInterestexpense.Rows[i][1] = ValueObj.Interest_Expense;
                        year = year + 1;
                    }
                }
                return Ok(new
                {
                    InitialSetup = InitialSetup_IValuationObj,
                    Interest_Income = ConvertDataTableToList(dtInterestIncome),
                    Interest_Expense = ConvertDataTableToList(dtInterestexpense),
                    CurrencyValueList = currencyValueList
                });


            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message.ToString());
            }

        }

        private List<Dictionary<string, object>> ConvertDataTableToList(DataTable table)
        {
            var list = new List<Dictionary<string, object>>();
            foreach (DataRow row in table.Rows)
            {
                var dict = new Dictionary<string, object>();
                foreach (DataColumn col in table.Columns)
                {
                    dict[col.ColumnName] = row[col];
                }
                list.Add(dict);
            }
            return list;
        }


        /// <summary>
        ///  Save internal Valuation Tax Rates
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("EvaluateInternal_Interest")]
        public ActionResult<Object> EvaluateInternal_Interest([FromBody] Interest_IValuationViewModel model)
        {

            try
            {
                List<Interest_IValuation> PayoutPolicy_IValuationListObj = new List<Interest_IValuation>();
                if (model != null)
                {
                    //first delete prev data 
                    iInterest_IValuation.DeleteWhere(x => x.InitialSetupId == model.InitialSetupId);
                    iInterest_IValuation.Commit();

                    //add new data for new years in payout policy
                    Console.WriteLine("above deserialize");
                    var dtInterest_IncomeT = model.Interest_Income != null ? JsonConvert.DeserializeObject<DataTable>(model.Interest_Income) : null;
                    var dtInterest_ExpenseT = model.Interest_Expense != null ? JsonConvert.DeserializeObject<DataTable>(model.Interest_Expense) : null;

                   
                    Console.WriteLine("below deserialize");
                    for (int i = 0; i < model.Yearcount; i++)
                    {
                        Interest_IValuation PayoutPolicy_IValuationObj = new Interest_IValuation();
                        PayoutPolicy_IValuationObj = new Interest_IValuation
                        {
                            TableData = "",
                            UserId = model.UserId,
                            InitialSetupId = model.InitialSetupId,
                            CreatedDate = System.DateTime.UtcNow,
                            ModifiedDate = System.DateTime.UtcNow,
                            Year = dtInterest_IncomeT.Columns.Count != 0 ? Convert.ToInt32(dtInterest_IncomeT.Rows[i][0]) : dtInterest_ExpenseT.Columns.Count != 0 ? Convert.ToInt32(dtInterest_ExpenseT.Rows[i][0]) : 2000,

                            Interest_Income = dtInterest_IncomeT.Rows.Count != 0 ? Convert.ToString(dtInterest_IncomeT.Rows[i][1]) : "",
                            Interest_Expense = dtInterest_ExpenseT.Rows.Count != 0 ? Convert.ToString(dtInterest_ExpenseT.Rows[i][1]) : "",

                        };
                        iInterest_IValuation.Add(PayoutPolicy_IValuationObj);
                        iInterest_IValuation.Commit();
                        PayoutPolicy_IValuationListObj.Add(PayoutPolicy_IValuationObj);
                    }
                }
                 Console.WriteLine("out of if block");

                // single data for every row
                DataTable dtInterest_Income = new DataTable();
                dtInterest_Income.Columns.Add("year");
                dtInterest_Income.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtInterest_Income.NewRow();
                    newRow[0] = obj.Year;
                    dtInterest_Income.Rows.Add(newRow);
                }

                DataTable dtInterest_Expense = new DataTable();
                dtInterest_Expense.Columns.Add("year");
                dtInterest_Expense.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtInterest_Expense.NewRow();
                    newRow[0] = obj.Year;
                    dtInterest_Expense.Rows.Add(newRow);
                }

                for (int i = 0; i < PayoutPolicy_IValuationListObj.Count; i++)
                {
                    dtInterest_Income.Rows[i][1] = PayoutPolicy_IValuationListObj[i].Interest_Income;
                    dtInterest_Expense.Rows[i][1] = PayoutPolicy_IValuationListObj[i].Interest_Expense;

                }
                Console.WriteLine("above return");

                return Ok(new
                {
                    // PayoutPolicy = iPayoutPolicy_IValuationObj,
                    InitialSetupId = model.InitialSetupId,
                   
                    Interest_Income = ConvertDataTableToList(dtInterest_Income),
                    Interest_Expense = ConvertDataTableToList(dtInterest_Expense)
                });
            }
            catch (Exception ex)
            {
                 Console.WriteLine("inside the exception");
                return BadRequest(ex.Message);
            }
        }

        #endregion

        #region PayOut Policy

        /// <summary>
        ///   Get GetPayoutPolicy_IValuation By Id
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetPayoutPolicy_IValuation/{UserId}")]
        public ActionResult<Object> GetPayoutPolicy_IValuation(long UserId)
        {
            try
            {
                int timelineStatus = 0;
                string message = "";
                ResultObject resultObject = new ResultObject();
                // InitialSetup_IValuation InitialSetup_IValuationObj = GetInitialSetupId(UserId);
                InitialSetup_IValuation InitialSetup_IValuationObj = GetInitialSetupId(UserId);
                if (InitialSetup_IValuationObj == null)
                {
                    return NotFound("Initial setup data not found for the given UserId.");
                }

                long InitialSetupId = InitialSetup_IValuationObj != null ? InitialSetup_IValuationObj.Id : 0;
                // var iPayoutPolicy_IValuationObj = InitialSetupId != 0 ? iPayoutPolicy_IValuation.FindBy(s => s.InitialSetupId == InitialSetupId && s.UserId == UserId).ToArray() : null;
                var iPayoutPolicy_IValuationObj = InitialSetupId != 0 ? iPayoutPolicy_IValuation.FindBy(s => s.InitialSetupId == InitialSetupId && s.UserId == UserId).ToArray() : null;
                if (iPayoutPolicy_IValuationObj == null)
                {
                    return NotFound("Payout policy data not found.");
                }

                if (iPayoutPolicy_IValuationObj.Length <= 0)
                {
                    return NotFound("no elements inside");
                }


                string Unit_TotalOngoingDividend = iPayoutPolicy_IValuationObj.Length != 0 ? iPayoutPolicy_IValuationObj[0].Unit_TotalOngoingDividend : "";
                string Unit_OneTimeDividend = iPayoutPolicy_IValuationObj.Length != 0 ? iPayoutPolicy_IValuationObj[0].Unit_OneTimeDividend : "";
                string Unit_StockBuyback = iPayoutPolicy_IValuationObj.Length != 0 ? iPayoutPolicy_IValuationObj[0].Unit_StockBuyback : "";
                string unit_Basic = iPayoutPolicy_IValuationObj.Length != 0 ? iPayoutPolicy_IValuationObj[0].Unit_WASOBasic : "";
                string unit_diluted = iPayoutPolicy_IValuationObj.Length != 0 ? iPayoutPolicy_IValuationObj[0].Unit_WASODiluted : "";
                string unit_shareRepurchased = iPayoutPolicy_IValuationObj.Length != 0 ? iPayoutPolicy_IValuationObj[0].Unit_ShareRepurchased : "";
                bool IsSaved = false;

                //Date:23-July-2020 | Added By : anonymous | Enh. : Single Units Enum
                List<ValueTextWrapper> currencyValueList = EnumHelper.GetEnumDescriptions<CurrencyValueEnum>();
                List<ValueTextWrapper> numberCountList = EnumHelper.GetEnumDescriptions<NumberCountEnum>();
                // single data for every row

                DataTable dtWeightedAvgShares = new DataTable();
                dtWeightedAvgShares.Columns.Add("year");
                dtWeightedAvgShares.Columns.Add("yearValue");

                int yearcount = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) - Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtWeightedAvgShares.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtWeightedAvgShares.Rows.Add(newRow);
                    year = year + 1;
                }

                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                DataTable dtWeightedAvgShares_Diluted = new DataTable();
                dtWeightedAvgShares_Diluted.Columns.Add("year");
                dtWeightedAvgShares_Diluted.Columns.Add("yearValue");
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtWeightedAvgShares_Diluted.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtWeightedAvgShares_Diluted.Rows.Add(newRow);
                    year = year + 1;
                }

                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                DataTable dtAnnual_DPS = new DataTable();
                dtAnnual_DPS.Columns.Add("year");
                dtAnnual_DPS.Columns.Add("yearValue");
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtAnnual_DPS.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtAnnual_DPS.Rows.Add(newRow);
                    year = year + 1;
                }

                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                DataTable dtTotalAnnualDividendPayout = new DataTable();
                dtTotalAnnualDividendPayout.Columns.Add("year");
                dtTotalAnnualDividendPayout.Columns.Add("yearValue");
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtTotalAnnualDividendPayout.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtTotalAnnualDividendPayout.Rows.Add(newRow);
                    year = year + 1;
                }

                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                DataTable dtOneTimeDividendPayout = new DataTable();
                dtOneTimeDividendPayout.Columns.Add("year");
                dtOneTimeDividendPayout.Columns.Add("yearValue");
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtOneTimeDividendPayout.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtOneTimeDividendPayout.Rows.Add(newRow);
                    year = year + 1;
                }

                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                DataTable dtStockPayBackAmount = new DataTable();
                dtStockPayBackAmount.Columns.Add("year");
                dtStockPayBackAmount.Columns.Add("yearValue");
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtStockPayBackAmount.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtStockPayBackAmount.Rows.Add(newRow);
                    year = year + 1;
                }

                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                DataTable dtSharePurchased = new DataTable();
                dtSharePurchased.Columns.Add("year");
                dtSharePurchased.Columns.Add("yearValue");
                for (int i = 0; i <= yearcount; i++)
                {
                    DataRow newRow = dtSharePurchased.NewRow();
                    newRow[0] = year;
                    newRow[1] = "";
                    dtSharePurchased.Rows.Add(newRow);
                    year = year + 1;
                }
                //for (int i = 0; i < dt.Rows.Count; i++)
                //{
                //    DataRow dr = dt.Rows[i];
                //    var yeardt = dr["year"];
                //    var ForcastRationValueobj = iPayoutPolicy_IValuationObj.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt));
                //    if (ForcastRationValueobj != null)
                //    {
                //        dr["yearValue"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
                //        dr["yearValue"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
                //        dr["yearValue"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
                //        dr["yearValue"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
                //        dr["yearValue"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
                //        dr["yearValue"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
                //        dr["yearValue"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
                //    }
                //}
                //test4564


                year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);

                ///Saved Data
                if (iPayoutPolicy_IValuationObj.Length != 0)
                {
                    IsSaved = true;
                    for (int i = 0; i <= yearcount; i++)
                    {
                        var ValueObj = iPayoutPolicy_IValuationObj.FirstOrDefault(x => x.Year == year);
                        if (ValueObj != null)
                        {
                            dtWeightedAvgShares.Rows[i][1] = ValueObj.WeightedAvgShares_Basic;
                            dtWeightedAvgShares_Diluted.Rows[i][1] = ValueObj.WeightedAvgShares_Diluted;
                            dtAnnual_DPS.Rows[i][1] = ValueObj.Annual_DPS;
                            dtTotalAnnualDividendPayout.Rows[i][1] = ValueObj.TotalAnnualDividendPayout;
                            dtOneTimeDividendPayout.Rows[i][1] = ValueObj.OneTimeDividendPayout;
                            dtStockPayBackAmount.Rows[i][1] = ValueObj.StockPayBackAmount;
                            dtSharePurchased.Rows[i][1] = ValueObj.SharePurchased;

                        }
                        year = year + 1;
                    }



                }
                ///Data fetch from Payout Policy History
                else
                {
                    IsSaved = false;
                    History history = iHistory.GetSingle(x => x.UserId == UserId);
                    if(history == null || history.PayoutTable == null){
                        return NotFound("History by This UserId not found");
                    }


                      
                            var PayoutTableObj = ConvterStringToObject(history.PayoutTable);
                            var ShareTable = ConvterStringToObject(history.ShareTable);
                            DataTable HistortDT = new DataTable();
                            HistortDT.Columns.Add("Name");
                            HistortDT.Columns.Add("Unit");
                            HistortDT.Columns.Add("year");
                            HistortDT.Columns.Add("yearValue");
                            int count = Convert.ToInt32(history.EndYear) - Convert.ToInt32(history.StartYear);

                            #region Show Message for default Values

                            // check if start year(History)<=startyear(Initial setup) 
                            // end (History)>=Endyear(InitialSetup)
                            if (Convert.ToInt32(InitialSetup_IValuationObj.YearFrom) >= Convert.ToInt32(history.StartYear) && Convert.ToInt32(InitialSetup_IValuationObj.YearTo) <= Convert.ToInt32(history.EndYear))
                            {
                                timelineStatus = 1;
                                message = "Dafault values are coming from History of Payout Policy module";

                            }
                            else if ((Convert.ToInt32(InitialSetup_IValuationObj.YearFrom) < Convert.ToInt32(history.StartYear) && Convert.ToInt32(InitialSetup_IValuationObj.YearTo) > Convert.ToInt32(history.StartYear)) || (Convert.ToInt32(InitialSetup_IValuationObj.YearFrom) < Convert.ToInt32(history.EndYear) && Convert.ToInt32(InitialSetup_IValuationObj.YearTo) > Convert.ToInt32(history.EndYear)))
                            {
                                timelineStatus = 2;
                                message = "Available year values are coming from History of Payout Policy module";
                            }
                            else
                            {
                                timelineStatus = 0;
                                message = "";
                            }

                            #endregion

                            if (PayoutTableObj != null && PayoutTableObj.Length > 0)
                            {
                                for (int j = 0; j <= PayoutTableObj.Length - 1; j++)
                                {
                                    int yr = Convert.ToInt32(history.StartYear);
                                    for (int i = 0; i <= count + 2; i++)
                                    {
                                        DataRow newRow = HistortDT.NewRow();
                                        newRow[0] = PayoutTableObj[j][0];
                                        newRow[1] = PayoutTableObj[j][1].ToString().Substring(Convert.ToInt32(PayoutTableObj[j][1].ToString().Length) - 1).ToUpper();
                                        if (i > 1)
                                        {
                                            newRow[2] = yr;
                                            newRow[3] = ConvertedValue(PayoutTableObj[j][1].ToString(), Convert.ToDouble(PayoutTableObj[j][i]));
                                            yr = yr + 1;
                                        }
                                        HistortDT.Rows.Add(newRow);
                                    }
                                }
                            }
                            if (ShareTable != null && ShareTable.Length > 0)
                            {
                                for (int k = 0; k <= ShareTable.Length - 1; k++)
                                {
                                    int yr = Convert.ToInt32(history.StartYear);
                                    for (int i = 0; i <= count + 2; i++)
                                    {
                                        DataRow newRow = HistortDT.NewRow();
                                        newRow[0] = ShareTable[k][0];
                                        newRow[1] = ShareTable[k][1].ToString().Substring(Convert.ToInt32(ShareTable[k][1].ToString().Length) - 1).ToUpper();
                                        if (i > 1)
                                        {
                                            newRow[2] = yr;
                                            newRow[3] = ConvertedValue(ShareTable[k][1].ToString(), Convert.ToDouble(ShareTable[k][i]));
                                            yr = yr + 1;
                                        }
                                        HistortDT.Rows.Add(newRow);
                                    }

                                }
                            }

                            List<TempHistory> TempHistoryList = new List<TempHistory>();
                            for (int l = 0; l < HistortDT.Rows.Count; l++)
                            {
                                TempHistory tempHistory = new TempHistory();
                                tempHistory.Name = HistortDT.Rows[l]["Name"].ToString();
                                tempHistory.Unit = HistortDT.Rows[l]["Unit"].ToString();
                                tempHistory.year = HistortDT.Rows[l]["year"].ToString();
                                tempHistory.yearValue = HistortDT.Rows[l]["yearValue"].ToString();
                                TempHistoryList.Add(tempHistory);
                            }
                            year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
                            for (int i = 0; i <= yearcount; i++)
                            {
                                dtWeightedAvgShares.Rows[i][1] = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Basic" && x.year == year.ToString()) != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Basic" && x.year == year.ToString()).yearValue) ? TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Basic" && x.year == year.ToString()).yearValue : "";
                                dtWeightedAvgShares_Diluted.Rows[i][1] = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Diluted" && x.year == year.ToString()) != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Diluted" && x.year == year.ToString()).yearValue) ? TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Diluted" && x.year == year.ToString()).yearValue : "";

                                dtTotalAnnualDividendPayout.Rows[i][1] = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.Find(x => x.Name == "Total Ongoing DIvident Payout-Annual" && x.year == year.ToString()) != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Total Ongoing DIvident Payout-Annual" && x.year == year.ToString()).yearValue) ? TempHistoryList.Find(x => x.Name == "Total Ongoing DIvident Payout-Annual" && x.year == year.ToString()).yearValue : "";
                                dtOneTimeDividendPayout.Rows[i][1] = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.Find(x => x.Name == "One Time Divident Payout" && x.year == year.ToString()) != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "One Time Divident Payout" && x.year == year.ToString()).yearValue) ? TempHistoryList.Find(x => x.Name == "One Time Divident Payout" && x.year == year.ToString()).yearValue : "";
                                dtStockPayBackAmount.Rows[i][1] = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.Find(x => x.Name == "Stock Buyback Amount" && x.year == year.ToString()) != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Stock Buyback Amount" && x.year == year.ToString()).yearValue) ? TempHistoryList.Find(x => x.Name == "Stock Buyback Amount" && x.year == year.ToString()).yearValue : "";
                                dtSharePurchased.Rows[i][1] = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.Find(x => x.Name == "Shares Repurchased" && x.year == year.ToString()) != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Shares Repurchased" && x.year == year.ToString()).yearValue) ? TempHistoryList.Find(x => x.Name == "Shares Repurchased" && x.year == year.ToString()).yearValue : "";
                                dtAnnual_DPS.Rows[i][1] = dtWeightedAvgShares.Rows[i][1] != null && dtWeightedAvgShares.Rows[i][1] != "" && Convert.ToInt32(dtWeightedAvgShares.Rows[i][1]) != 0 ? Convert.ToDouble((dtTotalAnnualDividendPayout.Rows[i][1] != null && dtTotalAnnualDividendPayout.Rows[i][1] != "" ? Convert.ToDouble(dtTotalAnnualDividendPayout.Rows[i][1]) : 0) * 4 / (Convert.ToDouble(dtWeightedAvgShares.Rows[i][1]))).ToString("0.##") : "";
                                year = year + 1;
                            }
                            unit_Basic = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.FirstOrDefault(x => x.Name == "Weighed Avg Shares Outstanding - Basic" && x.year == "") != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Basic" && x.year == "").Unit) ? TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Basic" && x.year == "").Unit : "";
                            unit_diluted = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.FirstOrDefault(x => x.Name == "Weighed Avg Shares Outstanding - Diluted" && x.year == "") != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Diluted" && x.year == "").Unit) ? TempHistoryList.Find(x => x.Name == "Weighed Avg Shares Outstanding - Diluted" && x.year == "").Unit : "";
                            Unit_TotalOngoingDividend = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.FirstOrDefault(x => x.Name == "Total Ongoing DIvident Payout-Annual" && x.year == "") != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Total Ongoing DIvident Payout-Annual" && x.year == "").Unit) ? TempHistoryList.Find(x => x.Name == "Total Ongoing DIvident Payout-Annual" && x.year == "").Unit : "";
                            Unit_OneTimeDividend = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.FirstOrDefault(x => x.Name == "One Time Divident Payout" && x.year == "") != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "One Time Divident Payout" && x.year == "").Unit) ? TempHistoryList.Find(x => x.Name == "One Time Divident Payout" && x.year == "").Unit : "";
                            Unit_StockBuyback = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.FirstOrDefault(x => x.Name == "Stock Buyback Amount" && x.year == "") != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Stock Buyback Amount" && x.year == "").Unit) ? TempHistoryList.Find(x => x.Name == "Stock Buyback Amount" && x.year == "").Unit : "";
                            unit_shareRepurchased = TempHistoryList != null && TempHistoryList.Count > 0 && TempHistoryList.FirstOrDefault(x => x.Name == "Shares Repurchased" && x.year == "") != null && !string.IsNullOrEmpty(TempHistoryList.Find(x => x.Name == "Shares Repurchased" && x.year == "").Unit) ? TempHistoryList.Find(x => x.Name == "Shares Repurchased" && x.year == "").Unit : "";
                        
                    

                }


                return Ok(new
                {
                    InitialSetup = InitialSetup_IValuationObj,
                    WeightedAvgShares_Basic = ConvertDataTableToList(dtWeightedAvgShares),
                    WeightedAvgShares_Diluted = ConvertDataTableToList(dtWeightedAvgShares_Diluted),
                    Annual_DPS = ConvertDataTableToList(dtAnnual_DPS),
                    TotalAnnualDividendPayout = ConvertDataTableToList(dtTotalAnnualDividendPayout),
                    OneTimeDividendPayout = ConvertDataTableToList(dtOneTimeDividendPayout),
                    StockPayBackAmount = ConvertDataTableToList(dtStockPayBackAmount),
                    SharePurchased = ConvertDataTableToList(dtSharePurchased),
                    Unit_TotalOngoingDividend = Unit_TotalOngoingDividend,
                    Unit_OneTimeDividend = Unit_OneTimeDividend,
                    Unit_StockBuyback = Unit_StockBuyback,
                    Unit_WASOBasic = unit_Basic,
                    Unit_WASODiluted = unit_diluted,
                    Unit_ShareRepurchased = unit_shareRepurchased,
                    CurrencyValueList = currencyValueList,
                    NumberCountList = numberCountList,
                    IsSaved = IsSaved,
                    TimelineStatusId = timelineStatus,
                    Message = message
                });
            }
            catch (Exception ss)
            {
                return BadRequest(ss.Message);
            }
        }

        

        double ConvertedValue(string unit, double value)
        {
            double convertedValue = value;
            if (unit.ToUpper() == "M" || unit.ToUpper() == "$M")
            {
                convertedValue = value / 1000000;
            }
            else if (unit.ToUpper() == "K" || unit.ToUpper() == "$K")
            {
                convertedValue = value / 1000;
            }
            else if (unit.ToUpper() == "B" || unit.ToUpper() == "$B")
            {
                convertedValue = value / 1000000000;
            }
            else if (unit.ToUpper() == "T" || unit.ToUpper() == "$T")
            {
                convertedValue = value / 1000000000000;
            }
            else if (unit.ToUpper() == "P" || unit.ToUpper() == "$P")
            {
                convertedValue = value / 1000000000000000;
            }
            return convertedValue;
        }

        private static object[][] ConvterStringToObject(string str)
        {
            var obj = JsonConvert.DeserializeObject<Object[][]>(str);
            return obj;
        }

        /// <summary>
        ///  Save internal Valuation Tax Rates
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("EvaluateInternalPayoutPolicy")]
        public ActionResult<Object> EvaluateInternalPayoutPolicy([FromBody] PayoutPolicy_IValuationViewModel model)
        {
            try
            {
                List<PayoutPolicy_IValuation> PayoutPolicy_IValuationListObj = new List<PayoutPolicy_IValuation>();
                if (model != null)
                {
                    Console.WriteLine("inside if block 1");
                    //first delete prev data 
                    iPayoutPolicy_IValuation.DeleteWhere(x => x.InitialSetupId == model.InitialSetupId);
                    iPayoutPolicy_IValuation.Commit();
                    Console.WriteLine("inside if block below commit");
                    //add new data for new years in payout policy
                    var dtWeightedAvgSharesT = model.WeightedAvgShares_Basic != null ? JsonConvert.DeserializeObject<DataTable>(model.WeightedAvgShares_Basic) : null;
                    var dtweightedAvgShares_DilutedT = model.WeightedAvgShares_Diluted != null ? JsonConvert.DeserializeObject<DataTable>(model.WeightedAvgShares_Diluted) : null;
                    var dtAnnual_DPST = model.Annual_DPS != null ? JsonConvert.DeserializeObject<DataTable>(model.Annual_DPS) : null;
                    var dtTotalAnnualDividendPayoutT = model.TotalAnnualDividendPayout != null ? JsonConvert.DeserializeObject<DataTable>(model.TotalAnnualDividendPayout) : null;
                    var dtOneTimeDividendPayoutT = model.OneTimeDividendPayout != null ? JsonConvert.DeserializeObject<DataTable>(model.OneTimeDividendPayout) : null;
                    var dtStockPayBackAmountT = model.StockPayBackAmount != null ? JsonConvert.DeserializeObject<DataTable>(model.StockPayBackAmount) : null;
                    var dtSharePurchasedT = model.SharePurchased != null ? JsonConvert.DeserializeObject<DataTable>(model.SharePurchased) : null;
                  
                    Console.WriteLine("inside if block below deserailze");

                    if (dtWeightedAvgSharesT == null || dtweightedAvgShares_DilutedT == null || dtAnnual_DPST == null || 
                    dtTotalAnnualDividendPayoutT == null || dtOneTimeDividendPayoutT == null || 
                    dtStockPayBackAmountT == null || dtSharePurchasedT == null)
                    {
                        return BadRequest("One or more DataTable deserializations failed.");
                    }
                    Console.WriteLine("inside if block first badrequest");

                    for (int i = 0; i < model.Yearcount; i++)
                    {
                        Console.WriteLine("inside if block inside for block");
                        PayoutPolicy_IValuation PayoutPolicy_IValuationObj = new PayoutPolicy_IValuation();
                        PayoutPolicy_IValuationObj = new PayoutPolicy_IValuation
                        {
                            TableData = "",
                            UserId = model.UserId,
                            InitialSetupId = model.InitialSetupId,
                            CreatedDate = System.DateTime.UtcNow,
                            ModifiedDate = System.DateTime.UtcNow,
                            Unit_TotalOngoingDividend = model.Unit_TotalOngoingDividend,
                            Unit_OneTimeDividend = model.Unit_OneTimeDividend,
                            Unit_StockBuyback = model.Unit_StockBuyback,
                            Unit_WASOBasic = model.Unit_WASOBasic,
                            Unit_WASODiluted = model.Unit_WASODiluted,
                            Unit_ShareRepurchased = model.Unit_ShareRepurchased,

                            Year = dtWeightedAvgSharesT.Columns.Count > 0 ? Convert.ToInt32(dtWeightedAvgSharesT.Rows[i][0]) : 
                                dtweightedAvgShares_DilutedT.Columns.Count > 0 ? Convert.ToInt32(dtweightedAvgShares_DilutedT.Rows[i][0]) : 
                                dtAnnual_DPST.Columns.Count > 0 ? Convert.ToInt32(dtAnnual_DPST.Rows[i][0]) : 
                                dtTotalAnnualDividendPayoutT.Columns.Count > 0 ? Convert.ToInt32(dtTotalAnnualDividendPayoutT.Rows[i][0]) : 
                                dtOneTimeDividendPayoutT.Columns.Count > 0 ? Convert.ToInt32(dtOneTimeDividendPayoutT.Rows[i][0]) : 
                                dtStockPayBackAmountT.Columns.Count > 0 ? Convert.ToInt32(dtStockPayBackAmountT.Rows[i][0]) : 
                                dtSharePurchasedT.Columns.Count > 0 ? Convert.ToInt32(dtSharePurchasedT.Rows[i][0]) : 2000,

                            WeightedAvgShares_Basic = dtWeightedAvgSharesT.Columns.Count > 1 ? Convert.ToString(dtWeightedAvgSharesT.Rows[i][1]) : "",
                            WeightedAvgShares_Diluted = dtweightedAvgShares_DilutedT.Columns.Count > 1 ? Convert.ToString(dtweightedAvgShares_DilutedT.Rows[i][1]) : "",
                            Annual_DPS = dtAnnual_DPST.Columns.Count > 1 ? Convert.ToString(dtAnnual_DPST.Rows[i][1]) : "",
                            TotalAnnualDividendPayout = dtTotalAnnualDividendPayoutT.Columns.Count > 1 ? Convert.ToString(dtTotalAnnualDividendPayoutT.Rows[i][1]) : "",
                            OneTimeDividendPayout = dtOneTimeDividendPayoutT.Columns.Count > 1 ? Convert.ToString(dtOneTimeDividendPayoutT.Rows[i][1]) : "",
                            StockPayBackAmount = dtStockPayBackAmountT.Columns.Count > 1 ? Convert.ToString(dtStockPayBackAmountT.Rows[i][1]) : "",
                            SharePurchased = dtSharePurchasedT.Columns.Count > 1 ? Convert.ToString(dtSharePurchasedT.Rows[i][1]) : "",
                        };
                        Console.WriteLine("inside if block above addiing to repo ");
                        iPayoutPolicy_IValuation.Add(PayoutPolicy_IValuationObj);
                        Console.WriteLine("inside if block below addiing to repo befor commit ");
                        iPayoutPolicy_IValuation.Commit();
                        Console.WriteLine("inside if block below addiing to repo below commit ");

                        PayoutPolicy_IValuationListObj.Add(PayoutPolicy_IValuationObj);
                    }



                    // for (int i = 0; i < model.Yearcount; i++)
                    // {
                    //     Console.WriteLine("inside if block inside for block");
                    //     PayoutPolicy_IValuation PayoutPolicy_IValuationObj = new PayoutPolicy_IValuation();
                    //     PayoutPolicy_IValuationObj = new PayoutPolicy_IValuation
                    //     {
                    //         TableData = "",
                    //         UserId = model.UserId,
                    //         InitialSetupId = model.InitialSetupId,
                    //         CreatedDate = System.DateTime.Now,
                    //         ModifiedDate = System.DateTime.Now,
                    //         Unit_TotalOngoingDividend = model.Unit_TotalOngoingDividend,
                    //         Unit_OneTimeDividend = model.Unit_OneTimeDividend,
                    //         Unit_StockBuyback = model.Unit_StockBuyback,
                    //         Unit_WASOBasic = model.Unit_WASOBasic,
                    //         Unit_WASODiluted = model.Unit_WASODiluted,
                    //         Unit_ShareRepurchased = model.Unit_ShareRepurchased,

                    //         Year = dtWeightedAvgSharesT.Columns.Count != 0 ? Convert.ToInt32(dtWeightedAvgSharesT.Rows[i][0]) : dtweightedAvgShares_DilutedT.Columns.Count != 0 ? Convert.ToInt32(dtweightedAvgShares_DilutedT.Rows[i][0]) : dtAnnual_DPST.Columns.Count != 0 ? Convert.ToInt32(dtAnnual_DPST.Rows[i][0]) : dtTotalAnnualDividendPayoutT.Columns.Count != 0 ? Convert.ToInt32(dtTotalAnnualDividendPayoutT.Rows[i][0]) : dtOneTimeDividendPayoutT.Columns.Count != 0 ? Convert.ToInt32(dtOneTimeDividendPayoutT.Rows[i][0]) : dtStockPayBackAmountT.Columns.Count != 0 ? Convert.ToInt32(dtStockPayBackAmountT.Rows[i][0]) : dtSharePurchasedT.Columns.Count != 0 ? Convert.ToInt32(dtSharePurchasedT.Rows[i][0]) : 2000,

                    //         WeightedAvgShares_Basic = dtWeightedAvgSharesT.Rows.Count != 0 ? Convert.ToString(dtWeightedAvgSharesT.Rows[i][1]) : "",
                    //         WeightedAvgShares_Diluted = dtweightedAvgShares_DilutedT.Rows.Count != 0 ? Convert.ToString(dtweightedAvgShares_DilutedT.Rows[i][1]) : "",
                    //         Annual_DPS = dtAnnual_DPST.Rows.Count != 0 ? Convert.ToString(dtAnnual_DPST.Rows[i][1]) : "",
                    //         TotalAnnualDividendPayout = dtTotalAnnualDividendPayoutT.Rows.Count != 0 ? Convert.ToString(dtTotalAnnualDividendPayoutT.Rows[i][1]) : "",
                    //         OneTimeDividendPayout = dtOneTimeDividendPayoutT.Rows.Count != 0 ? Convert.ToString(dtOneTimeDividendPayoutT.Rows[i][1]) : "",
                    //         StockPayBackAmount = dtStockPayBackAmountT.Rows.Count != 0 ? Convert.ToString(dtStockPayBackAmountT.Rows[i][1]) : "",
                    //         SharePurchased = dtSharePurchasedT.Rows.Count != 0 ? Convert.ToString(dtSharePurchasedT.Rows[i][1]) : "",
                    //     };
                    //     Console.WriteLine("inside if block above addiing to repo ");
                    //     iPayoutPolicy_IValuation.Add(PayoutPolicy_IValuationObj);
                    //     iPayoutPolicy_IValuation.Commit();
                    //     PayoutPolicy_IValuationListObj.Add(PayoutPolicy_IValuationObj);
                    // }
                }

                //return Ok(PayoutPolicy_IValuationListObj);
                Console.WriteLine(" when model is null");
                // single data for every row
                DataTable dtWeightedAvgShares = new DataTable();
                dtWeightedAvgShares.Columns.Add("year");
                dtWeightedAvgShares.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtWeightedAvgShares.NewRow();
                    newRow[0] = obj.Year;
                    dtWeightedAvgShares.Rows.Add(newRow);
                }
                Console.WriteLine(" when model is null below for loop");

                DataTable dtWeightedAvgShares_Diluted = new DataTable();
                dtWeightedAvgShares_Diluted.Columns.Add("year");
                dtWeightedAvgShares_Diluted.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtWeightedAvgShares_Diluted.NewRow();
                    newRow[0] = obj.Year;
                    dtWeightedAvgShares_Diluted.Rows.Add(newRow);
                }
                Console.WriteLine(" when model is null below diluted for loop");

                DataTable dtAnnual_DPS = new DataTable();
                dtAnnual_DPS.Columns.Add("year");
                dtAnnual_DPS.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtAnnual_DPS.NewRow();
                    newRow[0] = obj.Year;
                    dtAnnual_DPS.Rows.Add(newRow);
                }
                Console.WriteLine(" when model is null below dtAnnual_DPS for loop"); 

                DataTable dtTotalAnnualDividendPayout = new DataTable();
                dtTotalAnnualDividendPayout.Columns.Add("year");
                dtTotalAnnualDividendPayout.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtTotalAnnualDividendPayout.NewRow();
                    newRow[0] = obj.Year;
                    dtTotalAnnualDividendPayout.Rows.Add(newRow);
                }

                Console.WriteLine(" when model is null below dtTotalAnnualDividendPayout for loop");
                

                DataTable dtOneTimeDividendPayout = new DataTable();
                dtOneTimeDividendPayout.Columns.Add("year");
                dtOneTimeDividendPayout.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtOneTimeDividendPayout.NewRow();
                    newRow[0] = obj.Year;
                    dtOneTimeDividendPayout.Rows.Add(newRow);
                }
                Console.WriteLine(" when model is null below dtOneTimeDividendPayout for loop");

                DataTable dtStockPayBackAmount = new DataTable();
                dtStockPayBackAmount.Columns.Add("year");
                dtStockPayBackAmount.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtStockPayBackAmount.NewRow();
                    newRow[0] = obj.Year;
                    dtStockPayBackAmount.Rows.Add(newRow);
                }

                Console.WriteLine(" when model is null below dtStockPayBackAmount for loop");

                DataTable dtSharePurchased = new DataTable();
                dtSharePurchased.Columns.Add("year");
                dtSharePurchased.Columns.Add("yearValue");
                foreach (var obj in PayoutPolicy_IValuationListObj)
                {
                    DataRow newRow = dtSharePurchased.NewRow();
                    newRow[0] = obj.Year;
                    dtSharePurchased.Rows.Add(newRow);
                }

                Console.WriteLine(" when model is null below dtSharePurchased for loop");

                for (int i = 0; i < PayoutPolicy_IValuationListObj.Count; i++)
                {
                    if (dtWeightedAvgShares.Columns.Count > 1)
                        dtWeightedAvgShares.Rows[i][1] = PayoutPolicy_IValuationListObj[i].WeightedAvgShares_Basic;
                    else
                        dtWeightedAvgShares.Rows[i][1] = ""; // Or handle appropriately

                    if (dtWeightedAvgShares_Diluted.Columns.Count > 1)
                        dtWeightedAvgShares_Diluted.Rows[i][1] = PayoutPolicy_IValuationListObj[i].WeightedAvgShares_Diluted;
                    else
                        dtWeightedAvgShares_Diluted.Rows[i][1] = "";

                    if (dtAnnual_DPS.Columns.Count > 1)
                        dtAnnual_DPS.Rows[i][1] = PayoutPolicy_IValuationListObj[i].Annual_DPS;
                    else
                        dtAnnual_DPS.Rows[i][1] = "";

                    if (dtTotalAnnualDividendPayout.Columns.Count > 1)
                        dtTotalAnnualDividendPayout.Rows[i][1] = PayoutPolicy_IValuationListObj[i].TotalAnnualDividendPayout;
                    else
                        dtTotalAnnualDividendPayout.Rows[i][1] = "";

                    if (dtOneTimeDividendPayout.Columns.Count > 1)
                        dtOneTimeDividendPayout.Rows[i][1] = PayoutPolicy_IValuationListObj[i].OneTimeDividendPayout;
                    else
                        dtOneTimeDividendPayout.Rows[i][1] = "";

                    if (dtStockPayBackAmount.Columns.Count > 1)
                        dtStockPayBackAmount.Rows[i][1] = PayoutPolicy_IValuationListObj[i].StockPayBackAmount;
                    else
                        dtStockPayBackAmount.Rows[i][1] = "";

                    if (dtSharePurchased.Columns.Count > 1)
                        dtSharePurchased.Rows[i][1] = PayoutPolicy_IValuationListObj[i].SharePurchased;
                    else
                        dtSharePurchased.Rows[i][1] = "";
                }


               Console.WriteLine(" when model is null below dtSharePurchased for loop");

                // for (int i = 0; i < PayoutPolicy_IValuationListObj.Count; i++)
                // {
                //     dtWeightedAvgShares.Rows[i][1] = PayoutPolicy_IValuationListObj[i].WeightedAvgShares_Basic;
                //     dtWeightedAvgShares_Diluted.Rows[i][1] = PayoutPolicy_IValuationListObj[i].WeightedAvgShares_Diluted;
                //     dtAnnual_DPS.Rows[i][1] = PayoutPolicy_IValuationListObj[i].Annual_DPS;
                //     dtTotalAnnualDividendPayout.Rows[i][1] = PayoutPolicy_IValuationListObj[i].TotalAnnualDividendPayout;
                //     dtOneTimeDividendPayout.Rows[i][1] = PayoutPolicy_IValuationListObj[i].OneTimeDividendPayout;
                //     dtStockPayBackAmount.Rows[i][1] = PayoutPolicy_IValuationListObj[i].StockPayBackAmount;
                //     dtSharePurchased.Rows[i][1] = PayoutPolicy_IValuationListObj[i].SharePurchased;
                // }


                Console.WriteLine(" when model is null  above return");
                return Ok(new
                {
                    // PayoutPolicy = iPayoutPolicy_IValuationObj,
                    InitialSetupId = model.InitialSetupId,
                    WeightedAvgShares_Basic = ConvertDataTableToList(dtWeightedAvgShares),
                    WeightedAvgShares_Diluted = ConvertDataTableToList(dtWeightedAvgShares_Diluted),
                    Annual_DPS = ConvertDataTableToList(dtAnnual_DPS),
                    TotalAnnualDividendPayout = ConvertDataTableToList(dtTotalAnnualDividendPayout),
                    OneTimeDividendPayout = ConvertDataTableToList(dtOneTimeDividendPayout),
                    StockPayBackAmount = ConvertDataTableToList(dtStockPayBackAmount),
                    SharePurchased = ConvertDataTableToList(dtSharePurchased),
                    Unit_TotalOngoingDividend = model.Unit_TotalOngoingDividend,
                    Unit_OneTimeDividend = model.Unit_OneTimeDividend,
                    Unit_StockBuyback = model.Unit_StockBuyback,
                    Unit_WASOBasic = model.Unit_WASOBasic,
                    Unit_WASODiluted = model.Unit_WASODiluted,
                    Unit_ShareRepurchased = model.Unit_ShareRepurchased
                });
            }
            catch (Exception ss)
            {
                Console.WriteLine(" when model is inside exception");
                return BadRequest(ss.Message);
            }
        }


        private bool SaveReflexesofpoayouyPolicy(long initialSetupId, List<PayoutPolicy_IValuation> PayoutPolicy_IValuationList)
        {
            bool flag = false;
            // #region Bind Payout Policy
            try
            {

                List<ForcastRatio_ExplicitValuesViewModel> explicitValuesList = new List<ForcastRatio_ExplicitValuesViewModel>();
                //ForcastRatio_ExplicitValuesViewModel explicitValue;
                InitialSetup_IValuation InitialSetup_IValuationObj = iInitialSetup_IValuation.GetSingle(x => x.Id == initialSetupId);

                // get Payout Policy Statement  from Forcast Ratio
                List<ForcastRatioDatas> forcastRatioDatasList = iForcastRatioDatas.FindBy(x => x.InitialSetupId == initialSetupId).ToList();
                if (forcastRatioDatasList != null && forcastRatioDatasList.Count > 0) // check if Forcast Ratio Saved or not
                {
                    List<ForcastRatioDatas> tblPayoutForCastRatioListObj = new List<ForcastRatioDatas>();
                    tblPayoutForCastRatioListObj = forcastRatioDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.PayoutPolicyForcast).ToList();
                    if (tblPayoutForCastRatioListObj != null && tblPayoutForCastRatioListObj.Count > 0)
                    {
                        //delete Values or update values
                        // get all explicit values
                        List<ForcastRatio_ExplicitValues> ForcastRatio_ExplicitValuesListObj = iForcastRatio_ExplicitValues.FindBy(x => tblPayoutForCastRatioListObj.Any(m => m.Id == x.ForcastRatioDatasId)).ToList();

                        //get all Historical Values
                        List<ForcastRatioValues> ForcastRatioValuesListObj = iForcastRatioValues.FindBy(x => tblPayoutForCastRatioListObj.Any(m => m.Id == x.ForcastRatioDatasId)).ToList();

                        //get income after extraOrdinary items
                        IntegratedDatas incomeafter = iIntegratedDatas.GetSingle(x => x.InitialSetupId == initialSetupId && x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "NET INCOME after extraordinary items");
                        List<IntegratedValues> incomelist = new List<IntegratedValues>();
                        incomelist = incomeafter != null ? iIntegratedValues.FindBy(x => x.IntegratedDatasId == incomeafter.Id).ToList() : null;

                        string basicValue = "";
                        //change Values in Saved Datas
                        foreach (ForcastRatioDatas fortcastPayoutObj in tblPayoutForCastRatioListObj)
                        {
                            if (fortcastPayoutObj.isParent != true && fortcastPayoutObj.Sequence != 8 && fortcastPayoutObj.Sequence != 20)
                            {

                                List<ForcastRatioValues> payoutvaluesByDatas = ForcastRatioValuesListObj.FindAll(x => x.ForcastRatioDatasId == fortcastPayoutObj.Id).ToList();
                                if (payoutvaluesByDatas != null && payoutvaluesByDatas.Count > 0)
                                {
                                    if (fortcastPayoutObj.Sequence == 2)
                                    {
                                        string value = "";
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            obj.Value = payout_IValuation != null ? payout_IValuation.WeightedAvgShares_Basic : "";
                                            value = obj.Value;
                                        }
                                        basicValue = value;
                                        //explicit 
                                        List<ForcastRatio_ExplicitValues> payout_ExplicitvaluesByDatas = ForcastRatio_ExplicitValuesListObj.FindAll(x => x.ForcastRatioDatasId == fortcastPayoutObj.Id).ToList();
                                        if (payout_ExplicitvaluesByDatas != null && payout_ExplicitvaluesByDatas.Count > 0)
                                        {
                                            foreach (ForcastRatio_ExplicitValues obj in payout_ExplicitvaluesByDatas)
                                            {
                                                obj.Value = value;
                                            }
                                            //save Explicit
                                            iForcastRatio_ExplicitValues.UpdatedMany(payout_ExplicitvaluesByDatas);
                                            iForcastRatio_ExplicitValues.Commit();
                                        }

                                    }
                                    else if (fortcastPayoutObj.Sequence == 3)
                                    {
                                        string value = "";
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            obj.Value = payout_IValuation != null ? payout_IValuation.WeightedAvgShares_Diluted : "";
                                            value = obj.Value;
                                        }
                                        //explicit 
                                        List<ForcastRatio_ExplicitValues> payout_ExplicitvaluesByDatas = ForcastRatio_ExplicitValuesListObj.FindAll(x => x.ForcastRatioDatasId == fortcastPayoutObj.Id).ToList();
                                        if (payout_ExplicitvaluesByDatas != null && payout_ExplicitvaluesByDatas.Count > 0)
                                        {
                                            foreach (ForcastRatio_ExplicitValues obj in payout_ExplicitvaluesByDatas)
                                            {
                                                obj.Value = value;
                                            }
                                            //save Explicit
                                            iForcastRatio_ExplicitValues.UpdatedMany(payout_ExplicitvaluesByDatas);
                                            iForcastRatio_ExplicitValues.Commit();
                                        }
                                    }
                                    else if (fortcastPayoutObj.Sequence == 4 || fortcastPayoutObj.Sequence == 12)
                                    {
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            obj.Value = payout_IValuation != null ? payout_IValuation.Annual_DPS : "";
                                        }
                                        if (fortcastPayoutObj.Sequence == 12)
                                        {
                                            //explicit 
                                            List<ForcastRatio_ExplicitValues> payout_ExplicitvaluesByDatas = ForcastRatio_ExplicitValuesListObj.FindAll(x => x.ForcastRatioDatasId == fortcastPayoutObj.Id).ToList();
                                            if (payout_ExplicitvaluesByDatas != null && payout_ExplicitvaluesByDatas.Count > 0)
                                            {
                                                //find forcast Value
                                                ForcastRatioDatas totalDaatas = tblPayoutForCastRatioListObj.Find(x => x.Sequence == 13);
                                                foreach (ForcastRatio_ExplicitValues obj in payout_ExplicitvaluesByDatas)
                                                {
                                                    double value = 0;
                                                    ForcastRatio_ExplicitValues ForcastRatio_ExplicitValues = totalDaatas != null ? ForcastRatio_ExplicitValuesListObj.Find(x => x.ForcastRatioDatasId == totalDaatas.Id && x.Year == obj.Year) : null;
                                                    value = !string.IsNullOrEmpty(basicValue) && basicValue != "0" ? (ForcastRatio_ExplicitValues != null && !string.IsNullOrEmpty(ForcastRatio_ExplicitValues.Value) ? Convert.ToDouble(ForcastRatio_ExplicitValues.Value) : 0) / Convert.ToDouble(basicValue) : 0;
                                                    obj.Value = value.ToString("0.##");
                                                }
                                                //save Explicit
                                                iForcastRatio_ExplicitValues.UpdatedMany(payout_ExplicitvaluesByDatas);
                                                iForcastRatio_ExplicitValues.Commit();
                                            }
                                        }

                                    }
                                    else if (fortcastPayoutObj.Sequence == 5 || fortcastPayoutObj.Sequence == 13)
                                    {
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            obj.Value = payout_IValuation != null ? payout_IValuation.TotalAnnualDividendPayout : "";
                                        }
                                    }
                                    else if (fortcastPayoutObj.Sequence == 6 || fortcastPayoutObj.Sequence == 17)
                                    {
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            obj.Value = payout_IValuation != null ? payout_IValuation.OneTimeDividendPayout : "";
                                        }
                                    }
                                    else if (fortcastPayoutObj.Sequence == 7 || fortcastPayoutObj.Sequence == 19)
                                    {
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            obj.Value = payout_IValuation != null ? payout_IValuation.StockPayBackAmount : "";
                                        }
                                    }
                                    else if (fortcastPayoutObj.Sequence == 11)
                                    {
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            IntegratedValues income = incomelist != null && incomelist.Count > 0 ? incomelist.Find(x => x.Year == obj.Year) : null;
                                            double value = 0;
                                            value = payout_IValuation != null && income != null && !string.IsNullOrEmpty(income.Value) && income.Value != "0" ? Convert.ToDouble(payout_IValuation.TotalAnnualDividendPayout) / Convert.ToDouble(income.Value) : 0;
                                            obj.Value = value.ToString("0.#");
                                        }
                                    }
                                    else if (fortcastPayoutObj.Sequence == 15)
                                    {
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            IntegratedValues income = incomelist != null && incomelist.Count > 0 ? incomelist.Find(x => x.Year == obj.Year) : null;
                                            double value = 0;
                                            value = payout_IValuation != null && income != null && !string.IsNullOrEmpty(income.Value) && income.Value != "0" ? Convert.ToDouble(payout_IValuation.OneTimeDividendPayout) / Convert.ToDouble(income.Value) : 0;
                                            obj.Value = value.ToString("0.#");
                                        }
                                    }
                                    else if (fortcastPayoutObj.Sequence == 15)
                                    {
                                        foreach (ForcastRatioValues obj in payoutvaluesByDatas)
                                        {
                                            PayoutPolicy_IValuation payout_IValuation = PayoutPolicy_IValuationList.Find(x => x.Year == Convert.ToInt32(obj.Year));
                                            double value = 0;
                                            value = payout_IValuation != null && payout_IValuation.WeightedAvgShares_Basic != null && !string.IsNullOrEmpty(payout_IValuation.WeightedAvgShares_Basic) && payout_IValuation.WeightedAvgShares_Basic != "0" ? Convert.ToDouble(payout_IValuation.OneTimeDividendPayout) / Convert.ToDouble(payout_IValuation.WeightedAvgShares_Basic) : 0;
                                            obj.Value = value.ToString("0.#");
                                        }
                                    }

                                    //update and commit
                                    iForcastRatioValues.UpdatedMany(payoutvaluesByDatas);
                                    iForcastRatioValues.Commit();
                                }
                            }

                        }


                    }
                    else
                    {
                        //add years for Explicit values
                        // pending for now

                        //add new records for Payout 



                    }
                }

            }
            catch (Exception ss)
            {

            }
            //#endregion

            return flag;
        }

        #endregion

        #region Integrated Financial statement

        ///// <summary>
        /////   Get Integrated Financial statement By Id
        ///// </summary>
        ///// <param name="Id"></param>
        ///// <returns></returns>
        //[HttpGet]
        //[Route("GetIntegratedFinancialStatement/{UserId}")]
        //public ActionResult<Object> GetIntegratedFinancialStatement(long UserId)
        //{
        //    DataSet ds = new DataSet();
        //    try
        //    {
        //        ResultObject resultObject = new ResultObject();

        //        InitialSetup_IValuation InitialSetup_IValuationObj = GetInitialSetupId(UserId);
        //        if (InitialSetup_IValuationObj != null)
        //        {
        //            long InitialSetupId = InitialSetup_IValuationObj != null ? InitialSetup_IValuationObj.Id : 0;

        //            //check if initial setup exist in IntegratedFinancialStatement
        //            var IntegratedFinancialStatement = InitialSetupId != 0 ? iIntegratedFinancialStmt.FindBy(s => s.InitialSetupId == InitialSetupId).ToList() : null;
        //            if (IntegratedFinancialStatement != null && IntegratedFinancialStatement.Count != 0)
        //            {
        //                //edit Case
        //                //get no of years from initial setup                     
        //                int yearcount = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) - Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                var integratedElementMaster = iIntegratedElementMaster.AllIncluding().OrderBy(x => x.Id).ToList();
        //                if (integratedElementMaster != null && integratedElementMaster.Count > 0)
        //                {
        //                    foreach (IntegratedElementMaster obj in integratedElementMaster)
        //                    {
        //                        //create datatable and add years with historical binding
        //                        DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                        dt.Columns.Add("year");
        //                        dt.Columns.Add("yearValue");
        //                        int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                        for (int i = 0; i <= yearcount; i++)
        //                        {
        //                            DataRow newRow = dt.NewRow();
        //                            newRow[0] = year;
        //                            dt.Rows.Add(newRow);
        //                            year = year + 1;
        //                        }
        //                        for (int i = 0; i < dt.Rows.Count; i++)
        //                        {
        //                            DataRow dr = dt.Rows[i];
        //                            var yeardt = dr["year"];
        //                            var IntegratedValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == obj.Id);
        //                            if (IntegratedValueobj != null)
        //                            {
        //                                dr["yearValue"] = IntegratedValueobj.Value != null ? IntegratedValueobj.Value : null;
        //                            }
        //                        }

        //                        ds.Tables.Add(dt);
        //                        dt.Dispose();

        //                    }
        //                    //Remove special Characters from element object name
        //                    foreach (DataTable table in ds.Tables)
        //                    {
        //                        table.TableName = table.TableName.Replace(" ", "");
        //                        table.TableName = table.TableName.Replace("/", "");
        //                        table.TableName = table.TableName.Replace(",", "");
        //                        table.TableName = table.TableName.Replace("_", "");
        //                        table.TableName = table.TableName.Replace("'", "");
        //                    }
        //                }
        //                //return result
        //                return Ok(new
        //                {
        //                    dataset = ds,
        //                    InitialSetup = InitialSetup_IValuationObj
        //                });

        //            }
        //            else
        //            {
        //                //get no of years from initial setup                     
        //                int yearcount = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) - Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                var integratedElementMaster = iIntegratedElementMaster.AllIncluding().OrderBy(x => x.Id).ToList();
        //                if (integratedElementMaster != null && integratedElementMaster.Count > 0)
        //                {
        //                    foreach (IntegratedElementMaster obj in integratedElementMaster)
        //                    {




        //                        // Formula Not Required
        //                        if (obj.IsFormulaReq == false)
        //                        {
        //                            // for null or empty data
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.Empty)
        //                            {
        //                                //string tableName = obj.ElementName + "dt";
        //                                DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                dt.Columns.Add("year");
        //                                dt.Columns.Add("yearValue");
        //                                int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                for (int i = 0; i <= yearcount; i++)
        //                                {
        //                                    DataRow newRow = dt.NewRow();
        //                                    newRow[0] = year;
        //                                    dt.Rows.Add(newRow);
        //                                    year = year + 1;
        //                                }
        //                                ds.Tables.Add(dt);
        //                                dt.Dispose();
        //                            }
        //                            else if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.HistoricalData)//for direct binding from Historical Data
        //                            {

        //                                //create datatable and add years with historical binding
        //                                DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                dt.Columns.Add("year");
        //                                dt.Columns.Add("yearValue");
        //                                int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                for (int i = 0; i <= yearcount; i++)
        //                                {
        //                                    DataRow newRow = dt.NewRow();
        //                                    newRow[0] = year;
        //                                    dt.Rows.Add(newRow);
        //                                    year = year + 1;
        //                                }

        //                                var HistoricaldatamappingObj = iHistoryElementMapping.GetSingle(x => x.Id == obj.HistoryElementMappingId);
        //                                if (HistoricaldatamappingObj != null)
        //                                {
        //                                    List<LineItemInfo> Datas = iLineItemInfoRepository.FindBy(x => x.InitialSetupId == InitialSetupId && x.ElementName == HistoricaldatamappingObj.CommonElementName).ToList();
        //                                    if (Datas != null && Datas.Count > 0)
        //                                    {
        //                                        //get all values including above data ids
        //                                        List<RawHistoricalValues> ValuesListObj = iRawHistoricalValues.AllIncluding().Where(x => Datas.Any(p2 => p2.Id == x.DataId)).ToList();

        //                                        if (ValuesListObj != null && ValuesListObj.Count > 0)
        //                                        {
        //                                            for (int i = 0; i < dt.Rows.Count; i++)
        //                                            {
        //                                                DataRow dr = dt.Rows[i];
        //                                                var yeardt = dr["year"];
        //                                                var Valueobj = ValuesListObj.FirstOrDefault(x => x.FilingDate.Value.Year == Convert.ToInt32(yeardt));
        //                                                //  var Valueobj = ValuesListObj.FirstOrDefault(x=>x.FilingDate.ToString().Contains(yeardt.ToString()));                                                                                                               
        //                                                if (Valueobj != null)
        //                                                {
        //                                                    // for negative Values
        //                                                    if ((obj.ElementName == "R&D Expenses") || (obj.ElementName == "Selling, General and Admin Expenses") || (obj.ElementName == "RDepreciation") || (obj.ElementName == "Amortization") || (obj.ElementName == "Non-Operating Item: Restructuring Charges") || (obj.ElementName == "Provision for Taxes"))
        //                                                        dr["yearValue"] = Valueobj.Value != null ? -(Valueobj.Value) : null;
        //                                                    else
        //                                                        dr["yearValue"] = Valueobj.Value != null ? Valueobj.Value : null;
        //                                                }

        //                                            }

        //                                        }


        //                                    }
        //                                    else
        //                                    {
        //                                        // multiple element name case
        //                                    }
        //                                }
        //                                ds.Tables.Add(dt);
        //                                dt.Dispose();



        //                            }
        //                            else if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.Interest) //for direct binding from interest
        //                            {
        //                                List<Interest_IValuation> interest_IValuationListobj = iInterest_IValuation.FindBy(x => x.InitialSetupId == InitialSetupId).ToList();

        //                                DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                dt.Columns.Add("year");
        //                                dt.Columns.Add("yearValue");
        //                                int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                for (int i = 0; i <= yearcount; i++)
        //                                {
        //                                    DataRow newRow = dt.NewRow();
        //                                    newRow[0] = year;
        //                                    dt.Rows.Add(newRow);
        //                                    year = year + 1;
        //                                }


        //                                if (interest_IValuationListobj != null && interest_IValuationListobj.Count > 0)
        //                                {

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        //dr["yearValue"] = "New Value";
        //                                        var interestobj = interest_IValuationListobj.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt));
        //                                        if (interestobj != null)
        //                                        {
        //                                            dr["yearValue"] = obj.ReferenceName == "Interest_Income" ? Convert.ToInt32(interestobj.Interest_Income) : obj.ReferenceName == "Interest_Expense" ? -(Convert.ToInt32(interestobj.Interest_Expense)) : 0;
        //                                        }
        //                                    }

        //                                }
        //                                ds.Tables.Add(dt);
        //                                dt.Dispose();
        //                            }
        //                            else if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.HistoricalForcastRatio) // data from History Forcast Ratio
        //                            {
        //                                DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                dt.Columns.Add("year");
        //                                dt.Columns.Add("yearValue");
        //                                int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                for (int i = 0; i <= yearcount; i++)
        //                                {
        //                                    DataRow newRow = dt.NewRow();
        //                                    newRow[0] = year;
        //                                    dt.Rows.Add(newRow);
        //                                    year = year + 1;
        //                                }
        //                                // no data inserted from forcast for time being
        //                                ds.Tables.Add(dt);
        //                                dt.Dispose();
        //                            }
        //                        }
        //                        else
        //                        {
        //                            //for formula required true
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.HistoricalData)//for direct binding from interest
        //                            {
        //                                DataSet dsHistory = new DataSet();
        //                                // Cost of Goods Sold(integrated)=Cost of sales -Depreciation
        //                                if (obj.ElementName == "Cost of Goods Sold")
        //                                {
        //                                    //create datatable and add years with historical binding
        //                                    DataTable Cosdt = new DataTable("Cost of sales");
        //                                    Cosdt.Columns.Add("year");
        //                                    Cosdt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = Cosdt.NewRow();
        //                                        newRow[0] = year;
        //                                        Cosdt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    //Cost of sales
        //                                    #region Get Cost Of Sales
        //                                    var HistoricaldatamappingObj = iHistoryElementMapping.GetSingle(x => x.ItemName == "Cost of sales");
        //                                    if (HistoricaldatamappingObj != null)
        //                                    {
        //                                        List<LineItemInfo> Datas = iLineItemInfoRepository.FindBy(x => x.InitialSetupId == InitialSetupId && x.ElementName == HistoricaldatamappingObj.CommonElementName).ToList();
        //                                        if (Datas != null && Datas.Count > 0)
        //                                        {
        //                                            //get all values including above data ids
        //                                            List<RawHistoricalValues> ValuesListObj = iRawHistoricalValues.AllIncluding().Where(x => Datas.Any(p2 => p2.Id == x.DataId)).ToList();

        //                                            if (ValuesListObj != null && ValuesListObj.Count > 0)
        //                                            {
        //                                                for (int i = 0; i < Cosdt.Rows.Count; i++)
        //                                                {
        //                                                    DataRow dr = Cosdt.Rows[i];
        //                                                    var yeardt = dr["year"];
        //                                                    var Valueobj = ValuesListObj.FirstOrDefault(x => x.FilingDate.Value.Year == Convert.ToInt32(yeardt));
        //                                                    if (Valueobj != null)
        //                                                    {
        //                                                        dr["yearValue"] = Valueobj.Value != null ? Valueobj.Value : null;
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                        else
        //                                        {
        //                                            // multiple element name case
        //                                        }
        //                                    }
        //                                    dsHistory.Tables.Add(Cosdt);
        //                                    Cosdt.Dispose();
        //                                    #endregion

        //                                    //Depreciation
        //                                    #region Depreciation

        //                                    DataTable Deprdt = new DataTable("Depreciation");
        //                                    Deprdt.Columns.Add("year");
        //                                    Deprdt.Columns.Add("yearValue");
        //                                    year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = Deprdt.NewRow();
        //                                        newRow[0] = year;
        //                                        Deprdt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    //DataTable Deprdt = new DataTable("Depreciation");
        //                                    //Deprdt = Cosdt;

        //                                    var HistoricalDepreciationamappingObj = iHistoryElementMapping.GetSingle(x => x.ItemName == "Depreciation");
        //                                    if (HistoricalDepreciationamappingObj != null)
        //                                    {
        //                                        List<LineItemInfo> Datas = iLineItemInfoRepository.FindBy(x => x.InitialSetupId == InitialSetupId && x.ElementName == HistoricalDepreciationamappingObj.CommonElementName).ToList();
        //                                        if (Datas != null && Datas.Count > 0)
        //                                        {
        //                                            //get all values including above data ids
        //                                            List<RawHistoricalValues> ValuesListObj = iRawHistoricalValues.AllIncluding().Where(x => Datas.Any(p2 => p2.Id == x.DataId)).ToList();

        //                                            if (ValuesListObj != null && ValuesListObj.Count > 0)
        //                                            {
        //                                                for (int i = 0; i < Deprdt.Rows.Count; i++)
        //                                                {
        //                                                    DataRow dr = Deprdt.Rows[i];
        //                                                    var yeardt = dr["year"];
        //                                                    var Valueobj = ValuesListObj.FirstOrDefault(x => x.FilingDate.Value.Year == Convert.ToInt32(yeardt));
        //                                                    //  var Valueobj = ValuesListObj.FirstOrDefault(x=>x.FilingDate.ToString().Contains(yeardt.ToString()));
        //                                                    if (Valueobj != null)
        //                                                    {
        //                                                        dr["yearValue"] = Valueobj.Value != null ? Valueobj.Value : null;
        //                                                    }

        //                                                }

        //                                            }


        //                                        }
        //                                        else
        //                                        {
        //                                            // multiple element name case
        //                                        }
        //                                    }
        //                                    dsHistory.Tables.Add(Deprdt);
        //                                    Deprdt.Dispose();
        //                                    #endregion

        //                                    //get substracted value of "cost of capital" and "depreciation"
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var costOfCapital = dsHistory.Tables["Cost of sales"].Rows[i]["yearValue"];
        //                                        var depreciation = dsHistory.Tables["Depreciation"].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = -((costOfCapital != DBNull.Value ? Convert.ToInt32(costOfCapital) : 0) - (depreciation != DBNull.Value ? Convert.ToInt32(depreciation) : 0));
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();


        //                                }
        //                                // Accrued Expenses (Integrated) =History(Accrued compensation and benefits)+History(Accrued advertising)
        //                                if (obj.ElementName == "Accrued Expenses")
        //                                {
        //                                    //create datatable and add years with historical binding
        //                                    DataTable benefitsdt = new DataTable("Accrued compensation and benefits");
        //                                    benefitsdt.Columns.Add("year");
        //                                    benefitsdt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = benefitsdt.NewRow();
        //                                        newRow[0] = year;
        //                                        benefitsdt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }
        //                                    //Accrued compensation and benefits
        //                                    #region Get Accrued compensation and benefits
        //                                    var HistoricaldatamappingObj = iHistoryElementMapping.GetSingle(x => x.ItemName == "Accrued compensation and benefits");
        //                                    if (HistoricaldatamappingObj != null)
        //                                    {
        //                                        List<LineItemInfo> Datas = iLineItemInfoRepository.FindBy(x => x.InitialSetupId == InitialSetupId && x.ElementName == HistoricaldatamappingObj.CommonElementName).ToList();
        //                                        if (Datas != null && Datas.Count > 0)
        //                                        {
        //                                            //get all values including above data ids
        //                                            List<RawHistoricalValues> ValuesListObj = iRawHistoricalValues.AllIncluding().Where(x => Datas.Any(p2 => p2.Id == x.DataId)).ToList();

        //                                            if (ValuesListObj != null && ValuesListObj.Count > 0)
        //                                            {
        //                                                for (int i = 0; i < benefitsdt.Rows.Count; i++)
        //                                                {
        //                                                    DataRow dr = benefitsdt.Rows[i];
        //                                                    var yeardt = dr["year"];
        //                                                    var Valueobj = ValuesListObj.FirstOrDefault(x => x.FilingDate.Value.Year == Convert.ToInt32(yeardt));
        //                                                    if (Valueobj != null)
        //                                                    {
        //                                                        dr["yearValue"] = Valueobj.Value != null ? Valueobj.Value : null;
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                        else
        //                                        {
        //                                            // multiple element name case
        //                                        }
        //                                    }
        //                                    dsHistory.Tables.Add(benefitsdt);
        //                                    benefitsdt.Dispose();
        //                                    #endregion

        //                                    //Accrued advertising
        //                                    #region Accrued advertising

        //                                    DataTable advertisingdt = new DataTable("Accrued advertising");
        //                                    advertisingdt.Columns.Add("year");
        //                                    advertisingdt.Columns.Add("yearValue");
        //                                    year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = advertisingdt.NewRow();
        //                                        newRow[0] = year;
        //                                        advertisingdt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }
        //                                    var HistoricaladvertisingmappingObj = iHistoryElementMapping.GetSingle(x => x.ItemName == "Accrued advertising");
        //                                    if (HistoricaladvertisingmappingObj != null)
        //                                    {
        //                                        List<LineItemInfo> Datas = iLineItemInfoRepository.FindBy(x => x.InitialSetupId == InitialSetupId && x.ElementName == HistoricaladvertisingmappingObj.CommonElementName).ToList();
        //                                        if (Datas != null && Datas.Count > 0)
        //                                        {
        //                                            //get all values including above data ids
        //                                            List<RawHistoricalValues> ValuesListObj = iRawHistoricalValues.AllIncluding().Where(x => Datas.Any(p2 => p2.Id == x.DataId)).ToList();

        //                                            if (ValuesListObj != null && ValuesListObj.Count > 0)
        //                                            {
        //                                                for (int i = 0; i < advertisingdt.Rows.Count; i++)
        //                                                {
        //                                                    DataRow dr = advertisingdt.Rows[i];
        //                                                    var yeardt = dr["year"];
        //                                                    var Valueobj = ValuesListObj.FirstOrDefault(x => x.FilingDate.Value.Year == Convert.ToInt32(yeardt));
        //                                                    if (Valueobj != null)
        //                                                    {
        //                                                        dr["yearValue"] = Valueobj.Value != null ? Valueobj.Value : null;
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                        else
        //                                        {
        //                                            // multiple element name case
        //                                        }
        //                                    }
        //                                    dsHistory.Tables.Add(advertisingdt);
        //                                    advertisingdt.Dispose();
        //                                    #endregion

        //                                    //get sum of "Accrued compensation and benefits" and "Accrued advertising"
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var compensationNBenifits = dsHistory.Tables["Accrued compensation and benefits"].Rows[i]["yearValue"];
        //                                        var advertising = dsHistory.Tables["Accrued advertising"].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (compensationNBenifits != DBNull.Value ? Convert.ToInt32(compensationNBenifits) : 0) + (advertising != DBNull.Value ? Convert.ToInt32(advertising) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();


        //                                }

        //                            }//for same sheet calculations
        //                            else if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.IntegratedFinancialStmt)
        //                            {
        //                                // formula appplied in integrated values
        //                                //gross profit=NET SALES + Cost of Goods Sold
        //                                if (obj.ElementName == "GROSS PROFIT")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var NetSales = ds.Tables["NET SALES" + "_dt_Seq" + 1].Rows[i]["yearValue"];
        //                                        var COGS = ds.Tables["Cost of Goods Sold" + "_dt_Seq" + 2].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (NetSales != DBNull.Value ? Convert.ToInt32(NetSales) : 0) + (COGS != DBNull.Value ? Convert.ToInt32(COGS) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                //EBITDA = GrossProfit + R&D Expenses + Selling, General and Admin Expenses
        //                                else if (obj.ElementName == "EBITDA")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var GrossProfit = ds.Tables["GROSS PROFIT" + "_dt_Seq" + 3].Rows[i]["yearValue"];
        //                                        var RNDExpense = ds.Tables["R&D Expenses" + "_dt_Seq" + 4].Rows[i]["yearValue"];
        //                                        var SGA = ds.Tables["Selling, General and Admin Expenses" + "_dt_Seq" + 5].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (GrossProfit != DBNull.Value ? Convert.ToInt32(GrossProfit) : 0) + (RNDExpense != DBNull.Value ? Convert.ToInt32(RNDExpense) : 0) + (SGA != DBNull.Value ? Convert.ToInt32(SGA) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                // EBITA = EBITDA + Depreciation
        //                                else if (obj.ElementName == "EBITA")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var EBITDA = ds.Tables["EBITDA" + "_dt_Seq" + 6].Rows[i]["yearValue"];
        //                                        var Depreciation = ds.Tables["Depreciation" + "_dt_Seq" + 7].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (EBITDA != DBNull.Value ? Convert.ToInt32(EBITDA) : 0) + (Depreciation != DBNull.Value ? Convert.ToInt32(Depreciation) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                // EBIT = EBITA + Amortization
        //                                else if (obj.ElementName == "EBIT")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var EBITA = ds.Tables["EBITA" + "_dt_Seq" + 8].Rows[i]["yearValue"];
        //                                        var Amortization = ds.Tables["Amortization" + "_dt_Seq" + 9].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (EBITA != DBNull.Value ? Convert.ToInt32(EBITA) : 0) + (Amortization != DBNull.Value ? Convert.ToInt32(Amortization) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                else
        //                                //ebit=EBIT + (Interest Expense + Non-Operating Item: Interest Income + Non-Operating Item: Equity Investment Gains (Losses) + Non-Operating Item: Restructuring Charges)
        //                                if (obj.ElementName == "EBT")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var EBIT = ds.Tables["EBIT" + "_dt_Seq" + 10].Rows[i]["yearValue"];
        //                                        var InterestExpense = ds.Tables["Interest Expense" + "_dt_Seq" + 11].Rows[i]["yearValue"];
        //                                        var InterestIncome = ds.Tables["Non-Operating Item: Interest Income" + "_dt_Seq" + 12].Rows[i]["yearValue"];
        //                                        var EIGL = ds.Tables["Non-Operating Item: Equity Investment Gains (Losses)" + "_dt_Seq" + 13].Rows[i]["yearValue"];
        //                                        var RestructuringCharge = ds.Tables["Non-Operating Item: Restructuring Charges" + "_dt_Seq" + 14].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (EBIT != DBNull.Value ? Convert.ToInt32(EBIT) : 0) + ((InterestExpense != DBNull.Value ? Convert.ToInt32(InterestExpense) : 0) +
        //                                            (InterestIncome != DBNull.Value ? Convert.ToInt32(InterestIncome) : 0) +
        //                                            (EIGL != DBNull.Value ? Convert.ToInt32(EIGL) : 0) +
        //                                            (RestructuringCharge != DBNull.Value ? Convert.ToInt32(RestructuringCharge) : 0));
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                // NET INCOME before extraordinary items = EBT + Provision for Taxes
        //                                else if (obj.ElementName == "NET INCOME before extraordinary items")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var EBT = ds.Tables["EBT" + "_dt_Seq" + 15].Rows[i]["yearValue"];
        //                                        var Provision_Taxes = ds.Tables["Provision for Taxes" + "_dt_Seq" + 16].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (EBT != DBNull.Value ? Convert.ToInt32(EBT) : 0) + (Provision_Taxes != DBNull.Value ? Convert.ToInt32(Provision_Taxes) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                //NET INCOME after extraordinary items= NET INCOME before extraordinary items  + Provision for Taxes
        //                                else if (obj.ElementName == "NET INCOME after extraordinary items")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var Netincomebefore = ds.Tables["NET INCOME before extraordinary items" + "_dt_Seq" + 17].Rows[i]["yearValue"];
        //                                        var Extraordinary_Items = ds.Tables["Extraordinary Items" + "_dt_Seq" + 18].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (Netincomebefore != DBNull.Value ? Convert.ToInt32(Netincomebefore) : 0) + (Extraordinary_Items != DBNull.Value ? Convert.ToInt32(Extraordinary_Items) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                //Excess Cash= Cash and Cash Equivalents- Operating Cash
        //                                else if (obj.ElementName == "Excess Cash")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var CashEquivalent = ds.Tables["Cash and Cash Equivalents" + "_dt_Seq" + 20].Rows[i]["yearValue"];
        //                                        var OperatingCash = ds.Tables["Operating Cash" + "_dt_Seq" + 21].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (CashEquivalent != DBNull.Value ? Convert.ToInt32(CashEquivalent) : 0) - (OperatingCash != DBNull.Value ? Convert.ToInt32(OperatingCash) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                // TOTAL CURRENT ASSETS=Operating Cash+ Excess Cash + Net Receivables +Inventories +Other Current Assets +Non-Operating Item: Short-term investments +Non-Operating Item: Trading assets +Non-Operating Item: Assets held for sale
        //                                else if (obj.ElementName == "TOTAL CURRENT ASSETS")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var OperatingCash = ds.Tables["Operating Cash" + "_dt_Seq" + 21].Rows[i]["yearValue"];
        //                                        var ExcessCash = ds.Tables["Excess Cash" + "_dt_Seq" + 22].Rows[i]["yearValue"];
        //                                        var NetReceivables = ds.Tables["Net Receivables" + "_dt_Seq" + 23].Rows[i]["yearValue"];
        //                                        var Inventories = ds.Tables["Inventories" + "_dt_Seq" + 24].Rows[i]["yearValue"];
        //                                        var OtherCurrentAssets = ds.Tables["Other Current Assets" + "_dt_Seq" + 25].Rows[i]["yearValue"];
        //                                        var ShortTermInvestment = ds.Tables["Non-Operating Item: Short-term investments" + "_dt_Seq" + 26].Rows[i]["yearValue"];
        //                                        var TradingAssets = ds.Tables["Non-Operating Item: Trading assets" + "_dt_Seq" + 27].Rows[i]["yearValue"];
        //                                        var Asstesheld_Sale = ds.Tables["Non-Operating Item: Assets held for sale" + "_dt_Seq" + 28].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (OperatingCash != DBNull.Value ? Convert.ToInt32(OperatingCash) : 0) +
        //                                            (ExcessCash != DBNull.Value ? Convert.ToInt32(ExcessCash) : 0) +
        //                                            (NetReceivables != DBNull.Value ? Convert.ToInt32(NetReceivables) : 0) +
        //                                            (Inventories != DBNull.Value ? Convert.ToInt32(Inventories) : 0) +
        //                                            (OtherCurrentAssets != DBNull.Value ? Convert.ToInt32(OtherCurrentAssets) : 0) +
        //                                            (ShortTermInvestment != DBNull.Value ? Convert.ToInt32(ShortTermInvestment) : 0) +
        //                                            (TradingAssets != DBNull.Value ? Convert.ToInt32(TradingAssets) : 0) +
        //                                            (Asstesheld_Sale != DBNull.Value ? Convert.ToInt32(Asstesheld_Sale) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                // TOTAL ASSETS= TOTAL CURRENT ASSETS + Net Property, Plant & Equipment+ Good Will +Net Intangible Assets + Non-Operating Item: Marketable equity securities +Non-Operating Item: Other long-term investments + Non-Operating Item: Other long-term assets
        //                                else if (obj.ElementName == "TOTAL ASSETS")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var CurrentAssets = ds.Tables["TOTAL CURRENT ASSETS" + "_dt_Seq" + 29].Rows[i]["yearValue"];
        //                                        var PlanNEquipment = ds.Tables["Net Property, Plant & Equipment" + "_dt_Seq" + 30].Rows[i]["yearValue"];
        //                                        var GoodWill = ds.Tables["Good Will" + "_dt_Seq" + 31].Rows[i]["yearValue"];
        //                                        var IntangibleAssets = ds.Tables["Net Intangible Assets" + "_dt_Seq" + 32].Rows[i]["yearValue"];
        //                                        var MarketableEquitySecurities = ds.Tables["Non-Operating Item: Marketable equity securities" + "_dt_Seq" + 33].Rows[i]["yearValue"];
        //                                        var OtherInvestment = ds.Tables["Non-Operating Item: Other long-term investments" + "_dt_Seq" + 34].Rows[i]["yearValue"];
        //                                        var longtermAssets = ds.Tables["Non-Operating Item: Other long-term assets" + "_dt_Seq" + 35].Rows[i]["yearValue"];

        //                                        dr["yearValue"] = (CurrentAssets != DBNull.Value ? Convert.ToInt32(CurrentAssets) : 0) +
        //                                            (PlanNEquipment != DBNull.Value ? Convert.ToInt32(PlanNEquipment) : 0) +
        //                                            (GoodWill != DBNull.Value ? Convert.ToInt32(GoodWill) : 0) +
        //                                            (IntangibleAssets != DBNull.Value ? Convert.ToInt32(IntangibleAssets) : 0) +
        //                                            (MarketableEquitySecurities != DBNull.Value ? Convert.ToInt32(MarketableEquitySecurities) : 0) +
        //                                            (OtherInvestment != DBNull.Value ? Convert.ToInt32(OtherInvestment) : 0) +
        //                                            (longtermAssets != DBNull.Value ? Convert.ToInt32(longtermAssets) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                // CURRENT LIABILITIES= Short-Term Debt+ Accounts Payable + Deferred Revenue / Income + Accrued Expenses + Accrued Expenses  + Other Accrued Expenses/Liabilities +Income Taxes Payable + Non-Operating Item: Liabilities held for sale
        //                                else if (obj.ElementName == "CURRENT LIABILITIES")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];

        //                                        var Short_term_debt = ds.Tables["Short-Term Debt" + "_dt_Seq" + 38].Rows[i]["yearValue"];
        //                                        var Accounts_Payable = ds.Tables["Accounts Payable" + "_dt_Seq" + 39].Rows[i]["yearValue"];
        //                                        var DeferredRevenue = ds.Tables["Deferred Revenue/ Income" + "_dt_Seq" + 40].Rows[i]["yearValue"];
        //                                        var AccruedExpenses = ds.Tables["Accrued Expenses" + "_dt_Seq" + 41].Rows[i]["yearValue"];
        //                                        var OtherLiabilities = ds.Tables["Other Accrued Expenses/Liabilities" + "_dt_Seq" + 42].Rows[i]["yearValue"];
        //                                        var IncomeTaxes = ds.Tables["Income Taxes Payable" + "_dt_Seq" + 43].Rows[i]["yearValue"];
        //                                        var LiabilitiesHeldForsale = ds.Tables["Non-Operating Item: Liabilities held for sale" + "_dt_Seq" + 44].Rows[i]["yearValue"];


        //                                        dr["yearValue"] = (Short_term_debt != DBNull.Value ? Convert.ToInt32(Short_term_debt) : 0) +
        //                                            (Accounts_Payable != DBNull.Value ? Convert.ToInt32(Accounts_Payable) : 0) +
        //                                            (DeferredRevenue != DBNull.Value ? Convert.ToInt32(DeferredRevenue) : 0) +
        //                                            (AccruedExpenses != DBNull.Value ? Convert.ToInt32(AccruedExpenses) : 0) +
        //                                            (OtherLiabilities != DBNull.Value ? Convert.ToInt32(OtherLiabilities) : 0) +
        //                                            (IncomeTaxes != DBNull.Value ? Convert.ToInt32(IncomeTaxes) : 0) +
        //                                            (LiabilitiesHeldForsale != DBNull.Value ? Convert.ToInt32(LiabilitiesHeldForsale) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                //  TOTAL LIABILITIES = CURRENT LIABILITIES +  Long-Term Debt + Newly Issued Debt + Deferred Taxes +Non-Operating Item: Other Long-Term Liabilities
        //                                else if (obj.ElementName == "TOTAL LIABILITIES")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];

        //                                        var currentLIABILITIES = ds.Tables["CURRENT LIABILITIES" + "_dt_Seq" + 45].Rows[i]["yearValue"];
        //                                        var Long_term_debt = ds.Tables["Long-Term Debt" + "_dt_Seq" + 46].Rows[i]["yearValue"];
        //                                        var NewlyIssuedDebt = ds.Tables["Newly Issued Debt" + "_dt_Seq" + 47].Rows[i]["yearValue"];
        //                                        var DeferredRevenue = ds.Tables["Deferred Taxes" + "_dt_Seq" + 48].Rows[i]["yearValue"];
        //                                        var OtherLiabilities = ds.Tables["Non-Operating Item: Other Long-Term Liabilities" + "_dt_Seq" + 49].Rows[i]["yearValue"];


        //                                        dr["yearValue"] = (currentLIABILITIES != DBNull.Value ? Convert.ToInt32(currentLIABILITIES) : 0) +
        //                                            (Long_term_debt != DBNull.Value ? Convert.ToInt32(Long_term_debt) : 0) +
        //                                            (NewlyIssuedDebt != DBNull.Value ? Convert.ToInt32(NewlyIssuedDebt) : 0) +
        //                                            (DeferredRevenue != DBNull.Value ? Convert.ToInt32(DeferredRevenue) : 0) +
        //                                            (OtherLiabilities != DBNull.Value ? Convert.ToInt32(OtherLiabilities) : 0);

        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                //TOTAL STOCKHOLDERS' EQUITY= Common Stock and Paid-In Capital +Retained Earnings +Accumulated Other Comprehensive Income
        //                                // NET INCOME before extraordinary items = EBT + Provision for Taxes
        //                                else if (obj.ElementName == "TOTAL STOCKHOLDERS' EQUITY")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var commonStock = ds.Tables["Common Stock and Paid-In Capital" + "_dt_Seq" + 52].Rows[i]["yearValue"];
        //                                        var RetainedEarnings = ds.Tables["Retained Earnings" + "_dt_Seq" + 53].Rows[i]["yearValue"];
        //                                        var AccumulatedIncome = ds.Tables["Accumulated Other Comprehensive Income" + "_dt_Seq" + 54].Rows[i]["yearValue"];
        //                                        dr["yearValue"] = (commonStock != DBNull.Value ? Convert.ToInt32(commonStock) : 0) + (RetainedEarnings != DBNull.Value ? Convert.ToInt32(RetainedEarnings) : 0) +
        //                                            (AccumulatedIncome != DBNull.Value ? Convert.ToInt32(AccumulatedIncome) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }
        //                                //RETAINED EARNINGS (BEGINNING YEAR) = TOTAL LIABILITIES + Temporary Equity + Common Stock and Paid-In Capital +Retained Earnings + Accumulated Other Comprehensive Income
        //                                else if (obj.ElementName == "TOTAL LIABILITIES AND EQUITY")
        //                                {
        //                                    DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                                    dt.Columns.Add("year");
        //                                    dt.Columns.Add("yearValue");
        //                                    int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                                    for (int i = 0; i <= yearcount; i++)
        //                                    {
        //                                        DataRow newRow = dt.NewRow();
        //                                        newRow[0] = year;
        //                                        dt.Rows.Add(newRow);
        //                                        year = year + 1;
        //                                    }

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];

        //                                        var TOTAL_LIABILITIES = ds.Tables["TOTAL LIABILITIES" + "_dt_Seq" + 50].Rows[i]["yearValue"];
        //                                        var TemporaryEquity = ds.Tables["Temporary Equity" + "_dt_Seq" + 51].Rows[i]["yearValue"];
        //                                        var CommonStock = ds.Tables["Common Stock and Paid-In Capital" + "_dt_Seq" + 52].Rows[i]["yearValue"];
        //                                        var RetainedEarnings = ds.Tables["Retained Earnings" + "_dt_Seq" + 53].Rows[i]["yearValue"];
        //                                        var AccumulatedIncome = ds.Tables["Accumulated Other Comprehensive Income" + "_dt_Seq" + 54].Rows[i]["yearValue"];

        //                                        dr["yearValue"] = (TOTAL_LIABILITIES != DBNull.Value ? Convert.ToInt32(TOTAL_LIABILITIES) : 0) +
        //                                            (TemporaryEquity != DBNull.Value ? Convert.ToInt32(TemporaryEquity) : 0) +
        //                                            (CommonStock != DBNull.Value ? Convert.ToInt32(CommonStock) : 0) +
        //                                            (RetainedEarnings != DBNull.Value ? Convert.ToInt32(RetainedEarnings) : 0) +
        //                                            (AccumulatedIncome != DBNull.Value ? Convert.ToInt32(AccumulatedIncome) : 0);
        //                                    }
        //                                    ds.Tables.Add(dt);
        //                                    dt.Dispose();
        //                                }



        //                            }


        //                        }

        //                        //save data for every master and for every year
        //                        for (int i = 0; i < ds.Tables[obj.ElementName + "_dt_Seq" + obj.Sequence].Rows.Count; i++)
        //                        {
        //                            DataRow dr = ds.Tables[obj.ElementName + "_dt_Seq" + obj.Sequence].Rows[i];
        //                            var year = dr["year"];
        //                            var Yearvalue = dr["yearValue"];

        //                            IntegratedFinancialStmt integratedFinancialStmtsObj = new IntegratedFinancialStmt();
        //                            integratedFinancialStmtsObj.IntegratedElementMasterId = Convert.ToInt32(obj.Id);
        //                            integratedFinancialStmtsObj.InitialSetupId = InitialSetupId;
        //                            integratedFinancialStmtsObj.Year = year != DBNull.Value ? Convert.ToInt32(year) : 0;
        //                            integratedFinancialStmtsObj.Value = year != DBNull.Value ? Convert.ToString(Yearvalue) : null;
        //                            iIntegratedFinancialStmt.Add(integratedFinancialStmtsObj);
        //                            iIntegratedFinancialStmt.Commit();
        //                        }



        //                    }

        //                    //Remove special Characters from element object name
        //                    foreach (DataTable table in ds.Tables)
        //                    {
        //                        table.TableName = table.TableName.Replace(" ", "");
        //                        table.TableName = table.TableName.Replace("/", "");
        //                        table.TableName = table.TableName.Replace(",", "");
        //                        table.TableName = table.TableName.Replace("_", "");
        //                        table.TableName = table.TableName.Replace("'", "");
        //                    }
        //                }
        //                //return result
        //                return Ok(new
        //                {
        //                    dataset = ds,
        //                    InitialSetup = InitialSetup_IValuationObj
        //                });

        //            }
        //        }
        //        return Ok(new
        //        {
        //            //InterestIValuation = interestIvaluation,
        //            InitialSetup = InitialSetup_IValuationObj
        //        });
        //    }
        //    catch (Exception ss)
        //    {
        //        return BadRequest();
        //    }

        //}

        #endregion

        #region History Analysis and Forcast Ratio
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="UserId"></param>
        ///// <returns></returns>
        //[HttpGet]
        //[Route("GetHistoryandForcastRatio/{UserId}")]
        //public ActionResult<Object> GetHistoryandForcastRatio(long UserId)
        //{
        //    DataSet ds = new DataSet();
        //    DataSet ExplicitDataSet = new DataSet();
        //    try
        //    {
        //        ResultObject resultObject = new ResultObject();

        //        InitialSetup_IValuation InitialSetup_IValuationObj = GetInitialSetupId(UserId);
        //        if (InitialSetup_IValuationObj != null)
        //        {
        //            long InitialSetupId = InitialSetup_IValuationObj != null ? InitialSetup_IValuationObj.Id : 0;
        //            var HistoryAnalysisandForcastRationObj = InitialSetupId != 0 ? iHistoryAnalysisAndForcastRatio.FindBy(s => s.InitialSetupId == InitialSetupId).ToList() : null;
        //            if (HistoryAnalysisandForcastRationObj != null && HistoryAnalysisandForcastRationObj.Count != 0)
        //            {
        //                //edit Case
        //                //get no of years from initial setup                     
        //                int yearcount = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) - Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                var forcastRationElementMasterObj = iForcastRatioElementMaster.AllIncluding().OrderBy(x => x.Id).ToList();
        //                if (forcastRationElementMasterObj != null && forcastRationElementMasterObj.Count > 0)
        //                {

        //                    foreach (ForcastRatioElementMaster obj in forcastRationElementMasterObj)
        //                    {
        //                        //create datatable and add years with historical binding
        //                        DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                        dt.Columns.Add("year");
        //                        dt.Columns.Add("value");
        //                        dt.Columns.Add("IsEd");
        //                        dt.Columns["IsEd"].DataType = typeof(bool);
        //                        int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                        for (int i = 0; i <= yearcount; i++)
        //                        {
        //                            DataRow newRow = dt.NewRow();
        //                            newRow[0] = year;
        //                            newRow[2] = false;
        //                            dt.Rows.Add(newRow);
        //                            year = year + 1;
        //                        }
        //                        for (int i = 0; i < dt.Rows.Count; i++)
        //                        {
        //                            DataRow dr = dt.Rows[i];
        //                            var yeardt = dr["year"];
        //                            var ForcastRationValueobj = HistoryAnalysisandForcastRationObj.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.ForcastRatioElementMasterId == obj.Id);
        //                            if (ForcastRationValueobj != null)
        //                            {
        //                                dr["value"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
        //                            }
        //                        }

        //                        ds.Tables.Add(dt);
        //                        dt.Dispose();

        //                    }
        //                    //Remove special Characters from element object name
        //                    foreach (DataTable table in ds.Tables)
        //                    {
        //                        #region Explicit Forcast period

        //                        var ExplicitForcastListObj = iExplicitPeriod_HistoryForcastRatio.AllIncluding().OrderBy(x => x.Id).ToList();


        //                        bool isEditable = table.TableName == "SALES GROWTH" + "_dt_Seq" + 1
        //                            || table.TableName == "COGS % of Sales" + "_dt_Seq" + 2
        //                             || table.TableName == "R&D % of Sales" + "_dt_Seq" + 4
        //                              || table.TableName == "SG&A % of Sales" + "_dt_Seq" + 5
        //                               || table.TableName == "Depreciation % of Net PP&E (Beginning Year)" + "_dt_Seq" + 7
        //                                || table.TableName == "Amortization % of Net Intangible Assets (Beginning Year)" + "_dt_Seq" + 9
        //                                 || table.TableName == "Equity Investment Gains (Losses) YoY Change %" + "_dt_Seq" + 11
        //                                  || table.TableName == "Extraordinary Items % of Sales" + "_dt_Seq" + 14
        //                                   || table.TableName == "Cash Needed for Operations (Working Capital)" + "_dt_Seq" + 17
        //                                    || table.TableName == "Net Receivables % of Sales" + "_dt_Seq" + 19
        //                                     || table.TableName == "Inventories % of COGS" + "_dt_Seq" + 20
        //                                      || table.TableName == "Other Current Assets % of Sales" + "_dt_Seq" + 21
        //                                       || table.TableName == "Net PP&E % of Sales" + "_dt_Seq" + 25
        //                                        || table.TableName == "Accounts Payable % of COGS" + "_dt_Seq" + 30
        //                                         || table.TableName == "Deferred Revenue/ Income % of Sales" + "_dt_Seq" + 31
        //                                          || table.TableName == "Accrued Expenses % of Sales" + "_dt_Seq" + 32
        //                                           || table.TableName == "Other Accrued Expenses/Liabilities % of Sales" + "_dt_Seq" + 33
        //                                            || table.TableName == "Dividend Payout Ratio %(Ongoing Dividends)" + "_dt_Seq" + 53
        //                                             || table.TableName == "Dividend Payout Ratio %(One Time Dividend)" + "_dt_Seq" + 56
        //                                              || table.TableName == "One Time Dividend Payout(One Time Dividend)" + "_dt_Seq" + 58
        //                                               || table.TableName == "Stock Buyback Amount(Stock Buybacks)" + "_dt_Seq" + 59
        //                                                || table.TableName == "Shares Repurchased(Stock Buybacks)" + "_dt_Seq" + 60 ? true : false;

        //                        int ExplicitYearCount = Convert.ToInt32(InitialSetup_IValuationObj.ExplicitYearCount);
        //                        int startyear = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) + 1;
        //                        for (int i = 0; i < ExplicitYearCount; i++)
        //                        {
        //                            DataRow newRow = table.NewRow();
        //                            newRow[0] = startyear;
        //                            newRow[1] = null;
        //                            newRow[2] = isEditable;
        //                            table.Rows.Add(newRow);
        //                            startyear = startyear + 1;
        //                        }
        //                        //for (int i = 0; i < table.Rows.Count; i++)
        //                        //{
        //                        //    DataRow dr = table.Rows[i];
        //                        //    var yeardt = dr["year"];
        //                        //    var ForcastRationValueobj = ExplicitForcastListObj.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.ForcastRatioElementMasterId == obj.Id);
        //                        //    if (ForcastRationValueobj != null)
        //                        //    {
        //                        //        dr["value"] = ForcastRationValueobj.Value != null ? ForcastRationValueobj.Value : null;
        //                        //    }
        //                        //}
        //                        #endregion

        //                        table.TableName = table.TableName.Replace(" ", "");
        //                        table.TableName = table.TableName.Replace("/", "");
        //                        table.TableName = table.TableName.Replace(",", "");
        //                        table.TableName = table.TableName.Replace("_", "");
        //                        table.TableName = table.TableName.Replace("'", "");
        //                    }
        //                }
        //                //return result
        //                return Ok(new
        //                {
        //                    dataset = ds,
        //                    InitialSetup = InitialSetup_IValuationObj
        //                });
        //            }
        //            else
        //            {
        //                //get no of years from initial setup                     
        //                int yearcount = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) - Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                var forcastRationElementMasterObj = iForcastRatioElementMaster.AllIncluding().OrderBy(x => x.Id).ToList();
        //                if (forcastRationElementMasterObj != null && forcastRationElementMasterObj.Count > 0)
        //                {
        //                    //get integrated FinancialElement Values
        //                    var IntegratedFinancialStatement = InitialSetupId != 0 ? iIntegratedFinancialStmt.FindBy(s => s.InitialSetupId == InitialSetupId).ToList() : null;


        //                    // get all the element
        //                    foreach (ForcastRatioElementMaster obj in forcastRationElementMasterObj)
        //                    {
        //                        // for Historical Elements
        //                        DataTable dt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                        dt.Columns.Add("year");
        //                        dt.Columns.Add("value");
        //                        dt.Columns.Add("IsEd");
        //                        dt.Columns["IsEd"].DataType = typeof(bool);
        //                        int year = Convert.ToInt32(InitialSetup_IValuationObj.YearFrom);
        //                        for (int i = 0; i <= yearcount; i++)
        //                        {
        //                            DataRow newRow = dt.NewRow();
        //                            newRow[0] = year;
        //                            newRow[2] = false;
        //                            dt.Rows.Add(newRow);
        //                            year = year + 1;
        //                        }

        //                        int j = 0; int Neg = 1;
        //                        if (obj.HasInitialvalueZero == true)
        //                            j = 1;
        //                        if (obj.IsNegative == true)
        //                            Neg = -1;

        //                        // if Is Formula Required is false
        //                        if (obj.IsFormulaReq == false)
        //                        {
        //                            // direct Binding
        //                            // from Historical data
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.HistoricalData)
        //                            {
        //                                var HistoricaldatamappingObj = iHistoryElementMapping.GetSingle(x => x.Id == obj.HistoryElementMappingId);
        //                                if (HistoricaldatamappingObj != null)
        //                                {
        //                                    List<LineItemInfo> Datas = iLineItemInfoRepository.FindBy(x => x.InitialSetupId == InitialSetupId && x.ElementName == HistoricaldatamappingObj.CommonElementName).ToList();
        //                                    if (Datas != null && Datas.Count > 0)
        //                                    {
        //                                        //get all values including above data ids
        //                                        List<RawHistoricalValues> ValuesListObj = iRawHistoricalValues.AllIncluding().Where(x => Datas.Any(p2 => p2.Id == x.DataId)).ToList();

        //                                        if (ValuesListObj != null && ValuesListObj.Count > 0)
        //                                        {
        //                                            for (int i = 0; i < dt.Rows.Count; i++)
        //                                            {
        //                                                DataRow dr = dt.Rows[i];
        //                                                var yeardt = dr["year"];
        //                                                var Valueobj = ValuesListObj.FirstOrDefault(x => x.FilingDate.Value.Year == Convert.ToInt32(yeardt));
        //                                                if (Valueobj != null)
        //                                                {
        //                                                    double value = Valueobj.Value != null ? Convert.ToDouble(Valueobj.Value) : 0;
        //                                                    dr["value"] = Convert.ToDouble((value * Neg).ToString("0.##"));
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        // multiple element name case
        //                                    }
        //                                }
        //                            }
        //                            else
        //                            // from Integrated Financial Statement 
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.IntegratedFinancialStmt)
        //                            {
        //                                for (int i = j; i < dt.Rows.Count; i++)
        //                                {
        //                                    DataRow dr = dt.Rows[i];
        //                                    var yeardt = dr["year"];
        //                                    var currentValueobj = obj.IntegratedElementMasterId != null ? IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == obj.IntegratedElementMasterId) : null;

        //                                    double Value = (currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0);
        //                                    dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                }
        //                            }
        //                            else //from Payout Policy
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.PayoutPolicy)
        //                            {
        //                                List<PayoutPolicy_IValuation> payOutPolicyListobj = iPayoutPolicy_IValuation.FindBy(x => x.InitialSetupId == InitialSetupId).ToList();

        //                                if (payOutPolicyListobj != null && payOutPolicyListobj.Count > 0)
        //                                {

        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        //dr["value"] = "New Value";
        //                                        var Payoutobj = payOutPolicyListobj.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt));
        //                                        if (Payoutobj != null)
        //                                        {
        //                                            double Value = obj.ReferenceName == "WeightedAvgShares_Basic" ? (!string.IsNullOrEmpty(Payoutobj.WeightedAvgShares_Basic) ? Convert.ToDouble(Payoutobj.WeightedAvgShares_Basic) : 0)
        //                                                : obj.ReferenceName == "WeightedAvgShares_Diluted" ? (!string.IsNullOrEmpty(Payoutobj.WeightedAvgShares_Diluted) ? Convert.ToDouble(Payoutobj.WeightedAvgShares_Diluted) : 0)
        //                                                : obj.ReferenceName == "Annual_DPS" ? (!string.IsNullOrEmpty(Payoutobj.Annual_DPS) ? Convert.ToDouble(Payoutobj.Annual_DPS) : 0)
        //                                                : obj.ReferenceName == "TotalAnnualDividendPayout" ? (!string.IsNullOrEmpty(Payoutobj.TotalAnnualDividendPayout) ? Convert.ToDouble(Payoutobj.TotalAnnualDividendPayout) : 0)
        //                                                : obj.ReferenceName == "OneTimeDividendPayout" ? (!string.IsNullOrEmpty(Payoutobj.OneTimeDividendPayout) ? Convert.ToDouble(Payoutobj.OneTimeDividendPayout) : 0)
        //                                                : obj.ReferenceName == "StockPayBackAmount" ? (!string.IsNullOrEmpty(Payoutobj.StockPayBackAmount) ? Convert.ToDouble(Payoutobj.StockPayBackAmount) : 0)
        //                                                : obj.ReferenceName == "SharePurchased" ? (!string.IsNullOrEmpty(Payoutobj.SharePurchased) ? Convert.ToDouble(Payoutobj.SharePurchased) : 0) : 0;
        //                                            dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                        }
        //                                    }

        //                                }
        //                            }
        //                            else //for case 6
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.Empty)
        //                            {
        //                                // no need for any code
        //                            }



        //                        }
        //                        else
        //                        {
        //                            // when isFormula Required is  true   //all calculations are manual
        //                            // when Calculations are from Integrated Elements 
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.IntegratedFinancialStmt && (IntegratedFinancialStatement != null && IntegratedFinancialStatement.Count > 0))
        //                            {

        //                                //Sales Growth  // (current Net Sales - Prev Net sales)/ Prev Net sales *100(for percentage)
        //                                if (obj.ElementName == "SALES GROWTH")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 1);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                } //COGS % of Sales
        //                                // (CostOfGoodsSold/Net Sales) *100
        //                                else if (obj.ElementName == "COGS % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var CostOfGoodsSoldObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 2);

        //                                        double Value = ((CostOfGoodsSoldObj != null && !string.IsNullOrEmpty(CostOfGoodsSoldObj.Value) ? Convert.ToDouble(CostOfGoodsSoldObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // (R&D % of Sales/Net Sales) *100
        //                                else if (obj.ElementName == "R&D % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var RNDEXpenseObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 4);

        //                                        double Value = ((RNDEXpenseObj != null && !string.IsNullOrEmpty(RNDEXpenseObj.Value) ? Convert.ToDouble(RNDEXpenseObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                //SG&A % of Sales= (Selling, General and Admin Expenses/Net Sales) * 100
        //                                else if (obj.ElementName == "SG&A % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var SGandAObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 5);

        //                                        double Value = ((SGandAObj != null && !string.IsNullOrEmpty(SGandAObj.Value) ? Convert.ToDouble(SGandAObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }// EBITDA MARGIN %= (Integrated(EBITDA(6))*100)/integrated(Net Sales(1))
        //                                else if (obj.ElementName == "EBITDA MARGIN %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var EBITDAObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 6);

        //                                        double Value = ((EBITDAObj != null && !string.IsNullOrEmpty(EBITDAObj.Value) ? Convert.ToDouble(EBITDAObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Depreciation % of Net PP&E (Beginning Year)= -(Integrated(Depreciation(7))/Integrated(Net Property, Plant & Equipment(30)))
        //                                else if (obj.ElementName == "Depreciation % of Net PP&E (Beginning Year)")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var DepreciationObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 7);
        //                                        var netPropertyObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 30);

        //                                        double Value = ((DepreciationObj != null && !string.IsNullOrEmpty(DepreciationObj.Value) ? Convert.ToDouble(DepreciationObj.Value) : 0) / (netPropertyObj != null && !string.IsNullOrEmpty(netPropertyObj.Value) ? Convert.ToDouble(netPropertyObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // EBITA MARGIN %= (Integrated(EBITA(8))*100/Integrated(Net sales(1)))
        //                                else if (obj.ElementName == "EBITA MARGIN %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var EBITAObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 8);

        //                                        double Value = ((EBITAObj != null && !string.IsNullOrEmpty(EBITAObj.Value) ? Convert.ToDouble(EBITAObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Amortization % of Net Intangible Assets (Beginning Year)= -((Integrated(Amortization(9))*100)/Integrated(Net Intangible Assets(32)) of prev year)
        //                                else if (obj.ElementName == "Amortization % of Net Intangible Assets (Beginning Year)")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var AmortizationObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 9);
        //                                        var NetIntangibleAssetsObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 32);

        //                                        double Value = ((AmortizationObj != null && !string.IsNullOrEmpty(AmortizationObj.Value) ? Convert.ToDouble(AmortizationObj.Value) : 0) / (NetIntangibleAssetsObj != null && !string.IsNullOrEmpty(NetIntangibleAssetsObj.Value) ? Convert.ToDouble(NetIntangibleAssetsObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // EBIT MARGIN % =Integrated(EBIT(10))*100/Integrated(Net Sales(1))
        //                                else if (obj.ElementName == "EBIT MARGIN %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var EBITObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 10);

        //                                        double Value = ((EBITObj != null && !string.IsNullOrEmpty(EBITObj.Value) ? Convert.ToDouble(EBITObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Equity Investment Gains (Losses) YoY Change % = ((Integrater(Non-Operating Item: Equity Investment Gains (Losses) (13) of  current )-Integrated(Non-Operating Item: Equity Investment Gains (Losses)) (13) of prev)*100)/Integrated(Non-Operating Item: Equity Investment Gains (Losses) (13) of prev)
        //                                else if (obj.ElementName == "Equity Investment Gains (Losses) YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 13);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 13);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                }
        //                                // EBT MARGIN %= Integrated(EBT(15))*100/Integrated(Net Sales(1))
        //                                else if (obj.ElementName == "EBT MARGIN %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var EBTObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 15);

        //                                        double Value = ((EBTObj != null && !string.IsNullOrEmpty(EBTObj.Value) ? Convert.ToDouble(EBTObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // NET INCOME MARGIN % (before extraordinary items) =Integrated(NET INCOME before extraordinary items(17))*100/Integrated(Net Sales(1))
        //                                else if (obj.ElementName == "NET INCOME MARGIN % (before extraordinary items)")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var extraordinaryObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 17);

        //                                        double Value = ((extraordinaryObj != null && !string.IsNullOrEmpty(extraordinaryObj.Value) ? Convert.ToDouble(extraordinaryObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Extraordinary Items % of Sales  = Integrated(Extraordinary Items(18))*100/Integrated(Net Sales(1))
        //                                else if (obj.ElementName == "Extraordinary Items % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var extraordinaryObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 18);

        //                                        double Value = ((extraordinaryObj != null && !string.IsNullOrEmpty(extraordinaryObj.Value) ? Convert.ToDouble(extraordinaryObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // NET INCOME MARGIN % (after extraordinary items) =Integrated(NET INCOME after extraordinary items(19))*100/Integrated(Net Sales(1))
        //                                else if (obj.ElementName == "NET INCOME MARGIN % (after extraordinary items)")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var NetIncomeObjObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 19);

        //                                        double Value = ((NetIncomeObjObj != null && !string.IsNullOrEmpty(NetIncomeObjObj.Value) ? Convert.ToDouble(NetIncomeObjObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Net Receivables % of Sales = Integrated(Net Receivables(23))*100/Integrated(Net Sales(1))*100
        //                                else if (obj.ElementName == "Net Receivables % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var NetReceivablesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 23);

        //                                        double Value = ((NetReceivablesObj != null && !string.IsNullOrEmpty(NetReceivablesObj.Value) ? Convert.ToDouble(NetReceivablesObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Other Current Assets % of Sales= Integrated(Other Current Assets(25)/Net Sales(1))*100
        //                                else if (obj.ElementName == "Other Current Assets % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var CurrentAssets = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 25);

        //                                        double Value = ((CurrentAssets != null && !string.IsNullOrEmpty(CurrentAssets.Value) ? Convert.ToDouble(CurrentAssets.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Net PP&E % of Sales = Integrated(Net Property, Plant & Equipment(30))*100/Integrated(Net Sales(1))*100
        //                                else if (obj.ElementName == "Net PP&E % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var CurrentAssets = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 30);

        //                                        double Value = ((CurrentAssets != null && !string.IsNullOrEmpty(CurrentAssets.Value) ? Convert.ToDouble(CurrentAssets.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Deferred Revenue/ Income % of Sales = Integrated(Deferred Revenue/ Income(40))*100/Integrated(Net Sales(1))*100
        //                                else if (obj.ElementName == "Deferred Revenue/ Income % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var DeferredRevenue = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 40);

        //                                        double Value = ((DeferredRevenue != null && !string.IsNullOrEmpty(DeferredRevenue.Value) ? Convert.ToDouble(DeferredRevenue.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                //Accrued Expenses % of Sales = Integrated(Accrued Expenses(41))*100/Integrated(Net Sales(1))*100
        //                                else if (obj.ElementName == "Accrued Expenses % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var AccruedExpenses = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 41);

        //                                        double Value = ((AccruedExpenses != null && !string.IsNullOrEmpty(AccruedExpenses.Value) ? Convert.ToDouble(AccruedExpenses.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Other Accrued Expenses/Liabilities % of Sales =Integrated(Other Accrued Expenses/Liabilities(42))*100/Integrated(Net Sales(1))*100
        //                                else if (obj.ElementName == "Other Accrued Expenses/Liabilities % of Sales")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var AccruedExpenses = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 42);

        //                                        double Value = ((AccruedExpenses != null && !string.IsNullOrEmpty(AccruedExpenses.Value) ? Convert.ToDouble(AccruedExpenses.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Non-Operating Item: Short-term investments YoY Change %=  Integrated(Non-Operating Item: Short-term investments (current)- Non-Operating Item: Short-term investments(prev)/Non-Operating Item: Short-term investments(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Short-term investments YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 26);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 26);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                }
        //                                // Non-Operating Item: Trading assets YoY Change % = Integrated(Non-Operating Item: Trading assets(current)- Non-Operating Item: Trading assets(prev)/Non-Operating Item: Trading assets(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Trading assets YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 27);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 27);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                }
        //                                // Non-Operating Item: Assets held for sale YoY Change % = Integrated(Non-Operating Item: Assets held for sale(current)- Non-Operating Item: Assets held for sale(prev)/Non-Operating Item: Assets held for sale(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Assets held for sale YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 28);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 28);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                }
        //                                // Non-Operating Item: Marketable equity securities YoY Change % = Integrated(Non-Operating Item: Marketable equity securities(current)- Non-Operating Item: Marketable equity securities(prev)/Non-Operating Item: Marketable equity securities(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Marketable equity securities YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 33);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 33);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                }
        //                                //  Non-Operating Item: Other long-term investments YoY Change % = Integrated(Non-Operating Item: Other long-term investments(current)- Non-Operating Item: Other long-term investments(prev)/Non-Operating Item: Other long-term investments(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Other long-term investments YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 34);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 34);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                }
        //                                // Non-Operating Item: Other long-term assets YoY Change % = Integrated(Non-Operating Item: Other long-term assets(current)- Non-Operating Item: Other long-term assets(prev)/Non-Operating Item: Other long-term assets(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Other long-term assets YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 35);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 35);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));

        //                                    }
        //                                }
        //                                // Non-Operating Item: Liabilities held for sale YoY Change % = Integrated(Non-Operating Item: Liabilities held for sale(current)- Non-Operating Item: Liabilities held for sale(prev)/Non-Operating Item: Liabilities held for sale(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Liabilities held for sale YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 44);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 44);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Non-Operating Item: Other Long-Term Liabilities YoY Change % = Integrated(Non-Operating Item: Other Long-Term Liabilities(current)- Non-Operating Item: Other Long-Term Liabilities(prev)/Non-Operating Item: Other Long-Term Liabilities(prev))*100
        //                                else if (obj.ElementName == "Non-Operating Item: Other Long-Term Liabilities YoY Change %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var currentValueobj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 49);
        //                                        var prevValueObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 49);

        //                                        double Value = (((currentValueobj != null && !string.IsNullOrEmpty(currentValueobj.Value) ? Convert.ToDouble(currentValueobj.Value) : 0) - (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 0)) * 100) / (prevValueObj != null && !string.IsNullOrEmpty(prevValueObj.Value) ? Convert.ToDouble(prevValueObj.Value) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Inventories % of COGS = "('Integrated(Inventories (24)/'Integrated(   Cost of Goods Sold (2)))*100"
        //                                else if (obj.ElementName == "Inventories % of COGS")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var COGSObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 2);
        //                                        var Inventories = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 24);

        //                                        double Value = ((Inventories != null && !string.IsNullOrEmpty(Inventories.Value) ? Convert.ToDouble(Inventories.Value) : 0) / (COGSObj != null && !string.IsNullOrEmpty(COGSObj.Value) ? Convert.ToDouble(COGSObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Accounts Payable % of COGS = Integrated(Accounts Payable(39)/Cost of Goods Sold(2))*100
        //                                else if (obj.ElementName == "Accounts Payable % of COGS")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var COGSObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 2);
        //                                        var AccountsPayable = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 39);

        //                                        double Value = ((AccountsPayable != null && !string.IsNullOrEmpty(AccountsPayable.Value) ? Convert.ToDouble(AccountsPayable.Value) : 0) / (COGSObj != null && !string.IsNullOrEmpty(COGSObj.Value) && COGSObj.Value != "0" ? Convert.ToDouble(COGSObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }



        //                            }
        //                            else
        //                            // data  from multipole sheets
        //                            if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.Mixed)
        //                            {
        //                                // Dividend Payout Ratio % = Forcast(Total Ongoing Dividend Payout -Annual)/Integrated(NET INCOME after extraordinary items)*100
        //                                if (obj.ElementName == "Dividend Payout Ratio %(Ongoing Dividends)")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var TOTALDividend = ds.Tables["Total Ongoing Dividend Payout -Annual" + "_dt_Seq" + 49].Rows[i]["value"];
        //                                        var NetIncomeObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 19);
        //                                        double Value = ((TOTALDividend != DBNull.Value ? Convert.ToDouble(TOTALDividend) : 0)
        //                                            / (NetIncomeObj != null && !string.IsNullOrEmpty(NetIncomeObj.Value) ? Convert.ToDouble(NetIncomeObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                else
        //                                // Dividend Payout Ratio %(One Time Dividend) = Forcast(One Time Dividend Payout)/Integrated(NET INCOME after extraordinary items)*100
        //                                 if (obj.ElementName == "Dividend Payout Ratio %(One Time Dividend)")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var TOTALDividend = ds.Tables["One Time Dividend Payout" + "_dt_Seq" + 50].Rows[i]["value"];
        //                                        var NetIncomeObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 19);
        //                                        double Value = ((TOTALDividend != DBNull.Value ? Convert.ToDouble(TOTALDividend) : 0)
        //                                            / (NetIncomeObj != null && !string.IsNullOrEmpty(NetIncomeObj.Value) ? Convert.ToDouble(NetIncomeObj.Value) : 1)) * 100;
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }

        //                            }
        //                            else
        //                                // get data from Forcast Ratio
        //                                if (obj.IntegratedReferenceId == (int)IntegratedReferenceEnum.HistoricalForcastRatio)
        //                            {
        //                                //GROSS PROFIT MARGIN %=  (100-ForcastRatio(COGS % of Sales(2)))
        //                                if (obj.ElementName == "GROSS PROFIT MARGIN %")
        //                                {
        //                                    for (int i = j; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var netSalesObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == Convert.ToInt32(yeardt) && x.IntegratedElementMasterId == 1);
        //                                        var CostOfGoodsSoldObj = IntegratedFinancialStatement.FirstOrDefault(x => x.Year == (Convert.ToInt32(yeardt) - 1) && x.IntegratedElementMasterId == 2);

        //                                        double Value = ((CostOfGoodsSoldObj != null && !string.IsNullOrEmpty(CostOfGoodsSoldObj.Value) ? Convert.ToDouble(CostOfGoodsSoldObj.Value) : 0) / (netSalesObj != null && !string.IsNullOrEmpty(netSalesObj.Value) ? Convert.ToDouble(netSalesObj.Value) : 1)) * 100;
        //                                        dr["value"] = (100 - Convert.ToDouble((Value * Neg).ToString("0.##")));
        //                                    }
        //                                }// Excess Cash =  Focast(Cash and Cash Equivalents)-Forcast(Cash Needed for Operations (Working Capital))
        //                                else if (obj.ElementName == "Excess Cash")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var NetSales = ds.Tables["Cash and Cash Equivalents" + "_dt_Seq" + 16].Rows[i]["value"];
        //                                        var COGS = ds.Tables["Cash Needed for Operations (Working Capital)" + "_dt_Seq" + 17].Rows[i]["value"];
        //                                        double Value = (NetSales != DBNull.Value ? Convert.ToDouble(NetSales) : 0) - (COGS != DBNull.Value ? Convert.ToDouble(COGS) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }// CURRENT LIABILITIES =  focast Ratio(sUM(29+30+31+32+333+34)
        //                                else if (obj.ElementName == "CURRENT LIABILITIES")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var ShortTermdebt = ds.Tables["Short-Term Debt" + "_dt_Seq" + 29].Rows[i]["value"];
        //                                        var accountPayable = ds.Tables["Accounts Payable % of COGS" + "_dt_Seq" + 30].Rows[i]["value"];
        //                                        var Deferred = ds.Tables["Deferred Revenue/ Income % of Sales" + "_dt_Seq" + 31].Rows[i]["value"];
        //                                        var AccruedExpenses = ds.Tables["Accrued Expenses % of Sales" + "_dt_Seq" + 32].Rows[i]["value"];
        //                                        var otherLibilities = ds.Tables["Other Accrued Expenses/Liabilities % of Sales" + "_dt_Seq" + 33].Rows[i]["value"];
        //                                        var LiabilitiesHFS = ds.Tables["Non-Operating Item: Liabilities held for sale YoY Change %" + "_dt_Seq" + 34].Rows[i]["value"];
        //                                        double Value = (ShortTermdebt != DBNull.Value ? Convert.ToDouble(ShortTermdebt) : 0)
        //                                            + (accountPayable != DBNull.Value ? Convert.ToDouble(accountPayable) : 0)
        //                                            + (Deferred != DBNull.Value ? Convert.ToDouble(Deferred) : 0)
        //                                            + (AccruedExpenses != DBNull.Value ? Convert.ToDouble(AccruedExpenses) : 0)
        //                                            + (otherLibilities != DBNull.Value ? Convert.ToDouble(otherLibilities) : 0)
        //                                            + (LiabilitiesHFS != DBNull.Value ? Convert.ToDouble(LiabilitiesHFS) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }// TOTAL LIABILITIES = focast Ratio(sUM(35+36+37+38)
        //                                else if (obj.ElementName == "TOTAL LIABILITIES")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var CURRENT_LIABILITIES = ds.Tables["CURRENT LIABILITIES" + "_dt_Seq" + 35].Rows[i]["value"];
        //                                        var LongTermDebt = ds.Tables["Long-Term Debt" + "_dt_Seq" + 36].Rows[i]["value"];
        //                                        var Deferred_Taxes = ds.Tables["Deferred Taxes" + "_dt_Seq" + 37].Rows[i]["value"];
        //                                        var LongTerm_Liabilities = ds.Tables["Non-Operating Item: Other Long-Term Liabilities YoY Change %" + "_dt_Seq" + 38].Rows[i]["value"];
        //                                        double Value = (CURRENT_LIABILITIES != DBNull.Value ? Convert.ToDouble(CURRENT_LIABILITIES) : 0)
        //                                            + (LongTermDebt != DBNull.Value ? Convert.ToDouble(LongTermDebt) : 0)
        //                                            + (Deferred_Taxes != DBNull.Value ? Convert.ToDouble(Deferred_Taxes) : 0)
        //                                            + (LongTerm_Liabilities != DBNull.Value ? Convert.ToDouble(LongTerm_Liabilities) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }// TOTAL STOCKHOLDERS' EQUITY = focast Ratio(sUM(41+42+43)
        //                                else if (obj.ElementName == "TOTAL STOCKHOLDERS' EQUITY")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var commonStock = ds.Tables["Common Stock and Paid-In Capital" + "_dt_Seq" + 41].Rows[i]["value"];
        //                                        var RetainedEarnings = ds.Tables["Retained Earnings" + "_dt_Seq" + 42].Rows[i]["value"];
        //                                        var Accumulated = ds.Tables["Accumulated Other Comprehensive Income (Loss)" + "_dt_Seq" + 43].Rows[i]["value"];
        //                                        double Value = (commonStock != DBNull.Value ? Convert.ToDouble(commonStock) : 0)
        //                                            + (RetainedEarnings != DBNull.Value ? Convert.ToDouble(RetainedEarnings) : 0)
        //                                            + (Accumulated != DBNull.Value ? Convert.ToDouble(Accumulated) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // TOTAL LIABILITIES AND EQUITY =  focast Ratio(sUM(39+40 +41+42+43)
        //                                else if (obj.ElementName == "TOTAL LIABILITIES AND EQUITY")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var TOTALLIABILITIES = ds.Tables["TOTAL LIABILITIES" + "_dt_Seq" + 39].Rows[i]["value"];
        //                                        var TemporaryEquity = ds.Tables["Temporary Equity" + "_dt_Seq" + 40].Rows[i]["value"];
        //                                        var commonStock = ds.Tables["Common Stock and Paid-In Capital" + "_dt_Seq" + 41].Rows[i]["value"];
        //                                        var RetainedEarnings = ds.Tables["Retained Earnings" + "_dt_Seq" + 42].Rows[i]["value"];
        //                                        var Accumulated = ds.Tables["Accumulated Other Comprehensive Income (Loss)" + "_dt_Seq" + 43].Rows[i]["value"];
        //                                        double Value = (TOTALLIABILITIES != DBNull.Value ? Convert.ToDouble(TOTALLIABILITIES) : 0)
        //                                            + (TemporaryEquity != DBNull.Value ? Convert.ToDouble(TemporaryEquity) : 0)
        //                                            + (commonStock != DBNull.Value ? Convert.ToDouble(commonStock) : 0)
        //                                            + (RetainedEarnings != DBNull.Value ? Convert.ToDouble(RetainedEarnings) : 0)
        //                                            + (Accumulated != DBNull.Value ? Convert.ToDouble(Accumulated) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // DPS (Dividends per Share -Basic) ($)(Ongoing Dividends) =Forcast(Annual DPS $)
        //                                else if (obj.ElementName == "DPS (Dividends per Share -Basic) ($)(Ongoing Dividends)")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var AnnualDPS = ds.Tables["Annual DPS $" + "_dt_Seq" + 48].Rows[i]["value"];
        //                                        double Value = (AnnualDPS != DBNull.Value ? Convert.ToDouble(AnnualDPS) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Total Ongoing Dividend Payout -Annual(Ongoing Dividends) = Forcast(Total Ongoing Dividend Payout -Annual)
        //                                else if (obj.ElementName == "Total Ongoing Dividend Payout -Annual(Ongoing Dividends)")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var ongoingdividend = ds.Tables["Total Ongoing Dividend Payout -Annual" + "_dt_Seq" + 49].Rows[i]["value"];
        //                                        double Value = (ongoingdividend != DBNull.Value ? Convert.ToDouble(ongoingdividend) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // DPS (Dividends per Share -Basic) ($)(One Time Dividend) = Forcast(One Time Dividend Payout/Weighted Average Shares Outstanding - Basic (Millions))
        //                                else if (obj.ElementName == "DPS (Dividends per Share -Basic) ($)(One Time Dividend)")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var onTimedividend = ds.Tables["One Time Dividend Payout" + "_dt_Seq" + 50].Rows[i]["value"];
        //                                        var weightedAvg = ds.Tables["Weighted Average Shares Outstanding - Basic (Millions)" + "_dt_Seq" + 46].Rows[i]["value"];
        //                                        double Value = (onTimedividend != DBNull.Value && Convert.ToString(onTimedividend) != "0" ? Convert.ToDouble(onTimedividend) : 0)
        //                                            / (weightedAvg != DBNull.Value && Convert.ToString(weightedAvg) != "0" ? Convert.ToDouble(weightedAvg) : 1);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // One Time Dividend Payout(One Time Dividend) = Forcast(One Time Dividend Payout)
        //                                else if (obj.ElementName == "One Time Dividend Payout(One Time Dividend)")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var onTimedividend = ds.Tables["One Time Dividend Payout" + "_dt_Seq" + 50].Rows[i]["value"];
        //                                        double Value = (onTimedividend != DBNull.Value ? Convert.ToDouble(onTimedividend) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }
        //                                // Stock Buyback Amount(Stock Buybacks) = Forcast(Stock Buyback Amount)
        //                                else if (obj.ElementName == "Stock Buyback Amount(Stock Buybacks)")
        //                                {
        //                                    for (int i = 0; i < dt.Rows.Count; i++)
        //                                    {
        //                                        DataRow dr = dt.Rows[i];
        //                                        var yeardt = dr["year"];
        //                                        var StockBuyBackAmount = ds.Tables["Stock Buyback Amount" + "_dt_Seq" + 51].Rows[i]["value"];
        //                                        double Value = (StockBuyBackAmount != DBNull.Value ? Convert.ToDouble(StockBuyBackAmount) : 0);
        //                                        dr["value"] = Convert.ToDouble((Value * Neg).ToString("0.##"));
        //                                    }
        //                                }



        //                            }



        //                        }
        //                        // formula need true ended
        //                        // add datatable to dataset for every element of foreach
        //                        ds.Tables.Add(dt);
        //                        dt.Dispose();


        //                        #region Explicit dataSet

        //                        //// explicit Forcast
        //                        //DataTable ExplicitDt = new DataTable(obj.ElementName + "_dt_Seq" + obj.Sequence);
        //                        //ExplicitDt.Columns.Add("year");
        //                        //ExplicitDt.Columns.Add("value");
        //                        //int ExplicitYearCount = Convert.ToInt32(InitialSetup_IValuationObj.ExplicitYearCount);
        //                        //int startyear = Convert.ToInt32(InitialSetup_IValuationObj.YearTo)+1;
        //                        //for (int i = 0; i <= ExplicitYearCount; i++)
        //                        //{
        //                        //    DataRow newRow = ExplicitDt.NewRow();
        //                        //    newRow[0] = startyear;
        //                        //    ExplicitDt.Rows.Add(newRow);
        //                        //    startyear = startyear + 1;
        //                        //}
        //                        //// add datatable to dataset for every element of foreach
        //                        //ExplicitDataSet.Tables.Add(ExplicitDt);
        //                        //ExplicitDataSet.Dispose();

        //                        #endregion

        //                    }

        //                    //Remove special Characters from element object name
        //                    foreach (DataTable table in ds.Tables)
        //                    {
        //                        #region Explicit Forcast period

        //                        bool isEditable = table.TableName == "SALES GROWTH" + "_dt_Seq" + 1
        //                            || table.TableName == "COGS % of Sales" + "_dt_Seq" + 2
        //                             || table.TableName == "R&D % of Sales" + "_dt_Seq" + 4
        //                              || table.TableName == "SG&A % of Sales" + "_dt_Seq" + 5
        //                               || table.TableName == "Depreciation % of Net PP&E (Beginning Year)" + "_dt_Seq" + 7
        //                                || table.TableName == "Amortization % of Net Intangible Assets (Beginning Year)" + "_dt_Seq" + 9
        //                                 || table.TableName == "Equity Investment Gains (Losses) YoY Change %" + "_dt_Seq" + 11
        //                                  || table.TableName == "Extraordinary Items % of Sales" + "_dt_Seq" + 14
        //                                   || table.TableName == "Cash Needed for Operations (Working Capital)" + "_dt_Seq" + 17
        //                                    || table.TableName == "Net Receivables % of Sales" + "_dt_Seq" + 19
        //                                     || table.TableName == "Inventories % of COGS" + "_dt_Seq" + 20
        //                                      || table.TableName == "Other Current Assets % of Sales" + "_dt_Seq" + 21
        //                                       || table.TableName == "Net PP&E % of Sales" + "_dt_Seq" + 25
        //                                        || table.TableName == "Accounts Payable % of COGS" + "_dt_Seq" + 30
        //                                         || table.TableName == "Deferred Revenue/ Income % of Sales" + "_dt_Seq" + 31
        //                                          || table.TableName == "Accrued Expenses % of Sales" + "_dt_Seq" + 32
        //                                           || table.TableName == "Other Accrued Expenses/Liabilities % of Sales" + "_dt_Seq" + 33
        //                                            || table.TableName == "Dividend Payout Ratio %(Ongoing Dividends)" + "_dt_Seq" + 53
        //                                             || table.TableName == "Dividend Payout Ratio %(One Time Dividend)" + "_dt_Seq" + 56
        //                                              || table.TableName == "One Time Dividend Payout(One Time Dividend)" + "_dt_Seq" + 58
        //                                               || table.TableName == "Stock Buyback Amount(Stock Buybacks)" + "_dt_Seq" + 59
        //                                                || table.TableName == "Shares Repurchased(Stock Buybacks)" + "_dt_Seq" + 60 ? true : false;

        //                        int ExplicitYearCount = Convert.ToInt32(InitialSetup_IValuationObj.ExplicitYearCount);
        //                        int startyear = Convert.ToInt32(InitialSetup_IValuationObj.YearTo) + 1;
        //                        for (int i = 0; i < ExplicitYearCount; i++)
        //                        {
        //                            DataRow newRow = table.NewRow();
        //                            newRow[0] = startyear;
        //                            newRow[1] = null;
        //                            newRow[2] = isEditable;
        //                            table.Rows.Add(newRow);
        //                            startyear = startyear + 1;
        //                        }

        //                        #endregion

        //                        table.TableName = table.TableName.Replace(" ", "");
        //                        table.TableName = table.TableName.Replace("/", "");
        //                        table.TableName = table.TableName.Replace(",", "");
        //                        table.TableName = table.TableName.Replace("_", "");
        //                        table.TableName = table.TableName.Replace("'", "");
        //                    }
        //                    return Ok(new
        //                    {
        //                        dataset = ds,
        //                        InitialSetup = InitialSetup_IValuationObj

        //                    });

        //                }
        //                //return result
        //                return Ok(new
        //                {
        //                    dataset = ds,
        //                    InitialSetup = InitialSetup_IValuationObj,
        //                    message = "no Element Exist for Forcast Ratio"
        //                });

        //            }



        //        }
        //        return Ok(new
        //        {
        //            //InterestIValuation = interestIvaluation,
        //            InitialSetup = InitialSetup_IValuationObj
        //        });
        //    }
        //    catch (Exception ss)
        //    {
        //        return BadRequest(Convert.ToString(ss.Message));
        //    }
        //}






        #endregion



    }
}