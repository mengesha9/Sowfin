using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Sowfin.API.ViewModels.FAnalysis;
using Sowfin.Data.Abstract;
using Sowfin.Data.Common.Enum;
using Sowfin.Data.Common.Helper;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using Sowfin.API.ViewModels;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FinancialAnalysisController : ControllerBase
    {
        IInitialSetup_FAnalysis iInitialSetup_FAnalysis = null;
        IFilings iFilings;
        IMapper mapper;
        ICIKStatus iCIKStatus;
        IDatas iDatas;
        IValues iValues;
        private readonly IEdgarDataRepository edgarDataRepository;
        IMarketDatas iMarketDatas;
        IMarketValues iMarketValues;
        IMixedSubDatas_FAnalysis iMixedSubDatas_FAnalysis;
        IMixedSubValues_FAnalysis iMixedSubValues_FAnalysis;
        IFAnalysis_CategoryByInitialSetup iFAnalysis_CategoryByInitialSetup;
        IIntegratedDatasFAnalysis iIntegratedDatasFAnalysis;
        IIntegratedValuesFAnalysis iIntegratedValuesFAnalysis;
        IFinancialStatementAnalysisDatas iFinancialStatementAnalysisDatas;
        IFinancialStatementAnalysisValues iFinancialStatementAnalysisValues;
        IIntegratedDatas iIntegratedDatas;
        IIntegratedValues iIntegratedValues;
        ICategoryByInitialSetup iCategoryByInitialSetup;
        IMixedSubDatas iMixedSubDatas;
        IMixedSubValues iMixedSubValues;
        IROICDatas iROICDatas;
        IROICValues iROICValues;
        public FinancialAnalysisController(IInitialSetup_FAnalysis _iInitialSetup_FAnalysis, IFilings _iFilings, IMapper mapper
            , ICIKStatus _iCIKStatus, IDatas _iDatas, IValues _iValues, IEdgarDataRepository edgarDataRepository, IMarketDatas _iMarketDatas
            , IMarketValues _iMarketValues, IFAnalysis_CategoryByInitialSetup _iFAnalysis_CategoryByInitialSetup
            , IMixedSubValues_FAnalysis _iMixedSubValues_FAnalysis, IMixedSubDatas_FAnalysis _iIMixedSubDatas_FAnalysis
            , IIntegratedDatasFAnalysis _iIntegratedDatasFAnalysis, IIntegratedValuesFAnalysis _iIntegratedValuesFAnalysis
            , IFinancialStatementAnalysisDatas _iFinancialStatementAnalysisDatas, IFinancialStatementAnalysisValues _iFinancialStatementAnalysisValues
            , IIntegratedDatas _iIntegratedDatas, IIntegratedValues _iIntegratedValues, ICategoryByInitialSetup _iCategoryByInitialSetup, IMixedSubDatas _iMixedSubDatas
            , IMixedSubValues _iMixedSubValues, IROICDatas _iROICDatas, IROICValues _iROICValues)
        {
            iInitialSetup_FAnalysis = _iInitialSetup_FAnalysis;
            iFilings = _iFilings;
            this.mapper = mapper;
            iCIKStatus = _iCIKStatus;
            iDatas = _iDatas;
            iValues = _iValues;
            this.edgarDataRepository = edgarDataRepository;
            iMarketDatas = _iMarketDatas;
            iMarketValues = _iMarketValues;
            iFAnalysis_CategoryByInitialSetup = _iFAnalysis_CategoryByInitialSetup;
            iMixedSubDatas_FAnalysis = _iIMixedSubDatas_FAnalysis;
            iMixedSubValues_FAnalysis = _iMixedSubValues_FAnalysis;
            iIntegratedDatasFAnalysis = _iIntegratedDatasFAnalysis;
            iIntegratedValuesFAnalysis = _iIntegratedValuesFAnalysis;
            iFinancialStatementAnalysisDatas = _iFinancialStatementAnalysisDatas;
            iFinancialStatementAnalysisValues = _iFinancialStatementAnalysisValues;
            iIntegratedDatas = _iIntegratedDatas;
            iIntegratedValues = _iIntegratedValues;
            iCategoryByInitialSetup = _iCategoryByInitialSetup;
            iMixedSubDatas = _iMixedSubDatas;
            iMixedSubValues = _iMixedSubValues;
            iROICDatas = _iROICDatas;
            iROICValues = _iROICValues;
        }

        #region  InitialSetup_FAnalysis 
        /// <summary>
        ///   Get InitialSetup List By Id
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetFInitialSetupListByUserId/{UserId}")]
        public ActionResult<Object> GetFInitialSetupListByUserId(long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                List<InitialSetup_FAnalysisViewModel> InitialSetUpObj = new List<InitialSetup_FAnalysisViewModel>();

                var tblInitialSetUpObj = iInitialSetup_FAnalysis.FindBy(s => s.UserId == UserId).OrderByDescending(x => x.Id).ToList();
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
                            InitialSetup_FAnalysisViewModel tempInitialSetUpObj = mapper.Map<InitialSetup_FAnalysis, InitialSetup_FAnalysisViewModel>(obj);
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


        private bool SetAllAnalysisInitialSetUpFalse(long UserId)
        {
            bool flag = false;
            //set false to all initial setup
            //update
            try
            {
                List<InitialSetup_FAnalysis> initialSetupListObj = iInitialSetup_FAnalysis.FindBy(x => x.UserId == UserId).ToList();
                if (initialSetupListObj != null && initialSetupListObj.Count > 0)
                {
                    foreach (InitialSetup_FAnalysis item in initialSetupListObj)
                    {
                        item.IsActive = false;
                    }
                    iInitialSetup_FAnalysis.UpdatedMany(initialSetupListObj);
                    iInitialSetup_FAnalysis.Commit();
                    flag = true;
                }
            }
            catch (Exception ss)
            {
                Console.WriteLine(ss.Message);

            }
            return flag;
        }


        /// <summary>
        ///   Get InitialSetup By InitialSetupId
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetFAnalysisInitialSetupById/{Id}/{UserId}")]
        public ActionResult<Object> GetFAnalysisInitialSetupById(long Id, long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                SetAllAnalysisInitialSetUpFalse(UserId);
                var InitialSetUptblObj = iInitialSetup_FAnalysis.FindBy(s => s.Id == Id).OrderByDescending(x => x.Id).First();
                InitialSetup_FAnalysisViewModel InitialSetUpObj = new InitialSetup_FAnalysisViewModel();
                if (InitialSetUptblObj == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    InitialSetUptblObj.IsActive = true;
                    iInitialSetup_FAnalysis.Update(InitialSetUptblObj);
                    iInitialSetup_FAnalysis.Commit();

                    // map table to vm
                    InitialSetUpObj = mapper.Map<InitialSetup_FAnalysis, InitialSetup_FAnalysisViewModel>(InitialSetUptblObj);

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

        [HttpPost]
        [Route("EvaluateFAnalysis_InitialSetup")]
        public ActionResult<Object> EvaluateFAnalysis_InitialSetup([FromBody] InitialSetup_FAnalysisViewModel model)
        {
            try
            {
                /////Check if same CIK year and description already exist in company
                var CHK_InitialSetUptblObj = iInitialSetup_FAnalysis.GetSingle(x => x.CIKNumber == model.CIKNumber && x.YearFrom == model.YearFrom && x.YearTo == model.YearTo && x.Company == model.Company);
                if (CHK_InitialSetUptblObj != null)
                {
                    return Ok(new
                    {
                        StatusCode = 3,
                        Message = "Data already exists for these inputs."
                    });
                }

                /////Data saved in CIKStatus table
                bool isDataSave = InsertCIKStatus(model.CIKNumber);

                //bool isCurrentyear = false;
                if (model.YearTo == Convert.ToInt32(DateTime.UtcNow.Year.ToString()))
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
                InitialSetup_FAnalysis InitialSetUptblObj = new InitialSetup_FAnalysis();
                SetAllAnalysisInitialSetUpFalse(Convert.ToInt64(model.UserId));
                model.IsActive = true;
                model.CreatedDate = model.ModifiedDate = System.DateTime.UtcNow;
                InitialSetUptblObj = mapper.Map<InitialSetup_FAnalysisViewModel, InitialSetup_FAnalysis>(model);
                iInitialSetup_FAnalysis.Add(InitialSetUptblObj);
                iInitialSetup_FAnalysis.Commit();

                // map table to vm
                InitialSetup_FAnalysisViewModel InitialSetup_IValuationObj = new InitialSetup_FAnalysisViewModel();

                InitialSetup_IValuationObj = mapper.Map<InitialSetup_FAnalysis, InitialSetup_FAnalysisViewModel>(InitialSetUptblObj);

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
            }
            catch (Exception ss)
            {
                return BadRequest(new { Message = Convert.ToString(ss.Message), StatusCode = 0 });
            }
        }


      

        [HttpGet]
        [Route("GetInitialSetup_FAnalysis/{UserId}")]
        public ActionResult<Object> GetInitialSetup_FAnalysis(long UserId)
        {
            try
            {
                ResultObject resultObject = new ResultObject();
                InitialSetup_FAnalysis tblInitialSetUpObj = new InitialSetup_FAnalysis();
                List<InitialSetup_FAnalysis> tblInitialSetUpObjList = iInitialSetup_FAnalysis.FindBy(s => s.UserId == UserId && s.IsActive == true).ToList();
                //.OrderByDescending(x => x.Id).First();
                if (tblInitialSetUpObjList != null && tblInitialSetUpObjList.Count > 0)
                {
                    tblInitialSetUpObj = tblInitialSetUpObjList.OrderByDescending(x => x.Id).First();
                }
                else
                {
                    tblInitialSetUpObj = null;
                }
                InitialSetup_FAnalysisViewModel InitialSetUpObj = new InitialSetup_FAnalysisViewModel(); ;
                if (tblInitialSetUpObj == null)
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return Ok(resultObject);
                }
                else
                {
                    // map table to vm
                    InitialSetUpObj = mapper.Map<InitialSetup_FAnalysis, InitialSetup_FAnalysisViewModel>(tblInitialSetUpObj);

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
                return BadRequest();
            }
        }

        [HttpGet]
        [Route("DeleteDataByInitialSetup_FAnalysisId/{InitialSetup_FAnalysisId}")]
        public ActionResult DeleteDataByInitialSetup_FAnalysisId(long InitialSetup_FAnalysisId)
        {
            try
            {
                //delete IntegratedDatasFAnalysis
                List<IntegratedDatasFAnalysis> IntegratedDatasList = new List<IntegratedDatasFAnalysis>();
                IntegratedDatasList = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetup_FAnalysisId).ToList();
                if (IntegratedDatasList != null && IntegratedDatasList.Count > 0)
                {
                    List<IntegratedValuesFAnalysis> IntegratedValuesFAnalysisList = iIntegratedValuesFAnalysis.FindBy(t => IntegratedDatasList.Any(m => m.Id == t.IntegratedDatasFAnalysisId)).ToList();
                    if (IntegratedValuesFAnalysisList != null && IntegratedValuesFAnalysisList.Count > 0)
                    {
                        iIntegratedValuesFAnalysis.DeleteMany(IntegratedValuesFAnalysisList);
                        iIntegratedValuesFAnalysis.Commit();
                    }
                    iIntegratedDatasFAnalysis.DeleteMany(IntegratedDatasList);
                    iIntegratedDatasFAnalysis.Commit();
                }
                //////////////////

                //delete MarketData
                List<MarketDatas> MarketDatasList = new List<MarketDatas>();
                MarketDatasList = iMarketDatas.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetup_FAnalysisId).ToList();
                if (MarketDatasList != null && MarketDatasList.Count > 0)
                {
                    List<MarketValues> MarketValuesList = iMarketValues.FindBy(t => MarketDatasList.Any(m => m.Id == t.MarketDatasId)).ToList();
                    if (MarketValuesList != null && MarketValuesList.Count > 0)
                    {
                        iMarketValues.DeleteMany(MarketValuesList);
                        iMarketValues.Commit();
                    }
                    iMarketDatas.DeleteMany(MarketDatasList);
                    iMarketDatas.Commit();
                }
                //////////////////

                //delete MixedSubDatas_FAnalysis
                List<MixedSubDatas_FAnalysis> MixedSubDatas_FAnalysisList = new List<MixedSubDatas_FAnalysis>();
                MixedSubDatas_FAnalysisList = iMixedSubDatas_FAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetup_FAnalysisId).ToList();
                if (MixedSubDatas_FAnalysisList != null && MixedSubDatas_FAnalysisList.Count > 0)
                {
                    List<MixedSubValues_FAnalysis> MixedSubValues_FAnalysisList = iMixedSubValues_FAnalysis.FindBy(t => MixedSubDatas_FAnalysisList.Any(m => m.Id == t.MixedSubDatas_FAnalysisId)).ToList();
                    if (MixedSubValues_FAnalysisList != null && MixedSubValues_FAnalysisList.Count > 0)
                    {
                        iMixedSubValues_FAnalysis.DeleteMany(MixedSubValues_FAnalysisList);
                        iMixedSubValues_FAnalysis.Commit();
                    }
                    iMixedSubDatas_FAnalysis.DeleteMany(MixedSubDatas_FAnalysisList);
                    iMixedSubDatas_FAnalysis.Commit();
                }
                //////////////////////////////

                //delete FAnalysisCategory
                List<FAnalysis_CategoryByInitialSetup> FAnalysis_CategoryByInitialSetupList = new List<FAnalysis_CategoryByInitialSetup>();
                FAnalysis_CategoryByInitialSetupList = iFAnalysis_CategoryByInitialSetup.FindBy(x => x.FAnalysis_InitialSetupId == InitialSetup_FAnalysisId).ToList();
                if (FAnalysis_CategoryByInitialSetupList != null && FAnalysis_CategoryByInitialSetupList.Count > 0)
                {
                    iFAnalysis_CategoryByInitialSetup.DeleteMany(FAnalysis_CategoryByInitialSetupList);
                    iFAnalysis_CategoryByInitialSetup.Commit();
                }
                //////////////////////////

                //delete from in initialsetup
                iInitialSetup_FAnalysis.DeleteWhere(x => x.Id == InitialSetup_FAnalysisId);
                iInitialSetup_FAnalysis.Commit();
                //////////////////////////
                return Ok(new { message = "Deleted", result = true, statusCode = 200 });
            }
            catch (Exception ss)
            {
                return BadRequest(new { message = "Error", result = false, statusCode = 400 });
            }
        }
        #endregion


        #region Raw Historical 

        // Raw Historical get api
        [HttpGet]
        [Route("GetRawHistoricalData/{cik}/{startYear?}/{endYear?}/{UserId}")]
        public ActionResult GetRawHistoricalData(string cik, int? startYear = null, int? endYear = null, long? UserId = null)
        {
            RenderResult ed = new RenderResult();

            List<FilingsArray> lst = new List<FilingsArray>();

            try
            {
                var edgarData = edgarDataRepository.GetEdgar(cik, startYear, endYear).FirstOrDefault().EdgarView;

                if (edgarData == null)
                {
                    ed.StatusCode = 0;
                    ed.Result = lst;
                }
                else
                {
                    List<List<Filings>> jsonArr = JsonConvert.DeserializeObject<List<List<Filings>>>(edgarData);

                    foreach (var json in jsonArr)
                    {
                        try
                        {


                            var fa = new FilingsArray
                            {
                                CompanyName = json[0].CompanyName,
                                StatementType = json[0].StatementType,
                                Filings = json[0]
                            };

                            // test
                            // var test = json[0].Datas();
                            int MaxLength = 0;
                            List<Values> arrlist = new List<Values>();
                            List<Values> MaxValue = new List<Values>();
                            Values Val = new Values();
                            var Removedataslist = json[0].Datas.FindAll(x => (x.Values == null || x.Values.Count == 0) && x.IsParentItem != true);
                            if (Removedataslist != null)
                            {
                                foreach (var dt in Removedataslist)
                                {
                                    json[0].Datas.Remove(dt);
                                }
                            }
                            foreach (var obj in json[0].Datas)
                            {
                                if (obj.Values != null && MaxLength < obj.Values.Count)
                                {
                                    MaxLength = obj.Values.Count;
                                    MaxValue = obj.Values;
                                }

                            }

                            if (json[0].StatementType == "INCOME")
                            {
                                string maxDate = MaxValue[MaxLength - 1].FilingDate;
                                if (maxDate != null && maxDate != "" && Convert.ToInt32(Convert.ToDateTime(maxDate).Year) < endYear)
                                {
                                    InitialSetup_FAnalysis initialSetupdata = iInitialSetup_FAnalysis.GetSingle(x => x.UserId == UserId && x.IsActive == true);
                                    if (initialSetupdata != null)
                                    {
                                        initialSetupdata.YearTo = Convert.ToInt32(Convert.ToDateTime(maxDate).Year);
                                        iInitialSetup_FAnalysis.Update(initialSetupdata);
                                        iInitialSetup_FAnalysis.Commit();
                                        ed.ShowMsg = 1;
                                    }
                                }
                            }

                            int length = 0;
                            foreach (var obj1 in fa.Filings.Datas)
                            {

                                if (obj1.Values != null)
                                {
                                    length = obj1.Values.Count;
                                    if (length != MaxLength)
                                    {
                                        int k = 0;
                                        foreach (var item in MaxValue)
                                        {
                                            Val = new Values();
                                            Val.CElementName = null;
                                            Val.CLineItem = null;
                                            Val.Value = null;
                                            var Valueobj = obj1.Values.FirstOrDefault(x => Convert.ToDateTime(x.FilingDate) == Convert.ToDateTime(item.FilingDate));
                                            if (Valueobj == null)
                                            {
                                                Val.FilingDate = item.FilingDate;
                                                obj1.Values.Insert(k, Val);
                                            }
                                            ++k;
                                        }
                                    }
                                }
                                else
                                {
                                    arrlist = new List<Values>();
                                    obj1.Values = arrlist;
                                    foreach (var item1 in MaxValue)
                                    {
                                        Val = new Values();
                                        Val.CElementName = null;
                                        Val.CLineItem = null;
                                        Val.Value = null;
                                        Val.FilingDate = item1.FilingDate;
                                        obj1.Values.Add(Val);
                                    }
                                }

                            }
                            ///////////////////////////////////////
                            lst.Add(fa);
                        }
                        catch (Exception ss)
                        {

                        }
                    }

                    // 


                    ed.StatusCode = 1;
                    ed.Result = lst;
                }

                return Ok(ed);
            }
            catch (Exception ss)
            {
                return BadRequest(Convert.ToString(ss.Message));
            }
        }

        #endregion



        #region Data Processing

        [HttpGet]
        [Route("FAnalaysisDataProcessing/{UserId}/{cik}/{startYear?}/{endYear?}")]
        public ActionResult FAnalaysiaDataProcessing(long UserId, string cik, int? startYear = null, int? endYear = null)
        {
            DataProcessingResult renderResult = new DataProcessingResult();
            List<FilingsArray> filingsArrayList = new List<FilingsArray>();
            try
            {
                ////Update InitialSetupID into Filings///
                long? InitialSetupID = null;
                List<FAnalysis_CategoryByInitialSetup> categoryList = new List<FAnalysis_CategoryByInitialSetup>();
                InitialSetup_FAnalysis InitialSetup_IValuationObj = iInitialSetup_FAnalysis.FindBy(x => x.UserId == UserId && x.IsActive == true).OrderByDescending(x => x.Id).First();
                if (InitialSetup_IValuationObj != null)
                {
                    InitialSetupID = InitialSetup_IValuationObj.Id;
                    categoryList = iFAnalysis_CategoryByInitialSetup.FindBy(x => x.FAnalysis_InitialSetupId == InitialSetupID).ToList();
                }
                //if (InitialSetupID != null && InitialSetupID != 0)
                //{
                //    List<FilingsTable> FilingsListInitialSetup = new List<FilingsTable>();
                //    FilingsListInitialSetup = iFilings.FindBy(x => x.CIK == cik).ToList();
                //    if (FilingsListInitialSetup != null && FilingsListInitialSetup.Count > 0)
                //    {
                //        foreach (FilingsTable Ft in FilingsListInitialSetup)
                //        {
                //            Ft.InitialSetupId = InitialSetupID;
                //        }
                //        iFilings.UpdatedMany(FilingsListInitialSetup);
                //        iFilings.Commit();
                //    }

                //}
                string edgarView = this.edgarDataRepository.GetEdgar(cik, startYear, endYear).FirstOrDefault<EdgarData>().EdgarView;
                if (edgarView == null)
                {
                    renderResult.StatusCode = 0;
                    renderResult.Result = filingsArrayList;
                    renderResult.InitialSetupId = InitialSetupID;
                }
                else
                {
                    foreach (List<Filings> filingsList in (List<List<Filings>>)JsonConvert.DeserializeObject<List<List<Filings>>>(edgarView))
                    {
                        FilingsArray filingsArray = new FilingsArray()
                        {
                            CompanyName = filingsList[0].CompanyName,
                            StatementType = filingsList[0].StatementType,
                            Filings = filingsList[0]
                        };

                        ////Remove All Null values
                        var Removedataslist = filingsList[0].Datas.FindAll(x => (x.Values == null || x.Values.Count == 0) && x.IsParentItem != true);
                        if (Removedataslist != null)
                        {
                            foreach (var dt in Removedataslist)
                            {
                                filingsList[0].Datas.Remove(dt);
                            }
                        }
                        ////////////////////////////////

                        int num = 0;
                        List<Values> valuesList1 = new List<Values>();
                        List<Values> valuesList2 = new List<Values>();
                        Values values1 = new Values();
                        foreach (Datas data in filingsList[0].Datas)
                        {
                            //  categoryList
                            if (categoryList != null && categoryList.Count > 0)
                            {
                                var assigncategory = categoryList.Find(x => x.DatasId == data.DataId);
                                if (assigncategory != null)
                                {
                                    data.Category = assigncategory.Category != null ? assigncategory.Category : data.Category;
                                }
                            }
                            if (data.Values != null && num < data.Values.Count)
                            {
                                num = data.Values.Count;
                                valuesList2 = data.Values;
                            }
                        }
                        List<MixedSubDatas_FAnalysis> mixedSubDatasList = iMixedSubDatas_FAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupID).ToList();
                        List<MixedSubValues_FAnalysis> mixedSubValuesList = new List<MixedSubValues_FAnalysis>();
                        if (mixedSubDatasList != null && mixedSubDatasList.Count > 0)
                        {
                            mixedSubValuesList = iMixedSubValues_FAnalysis.FindBy(t => mixedSubDatasList.Any(m => m.Id == t.MixedSubDatas_FAnalysisId)).ToList();
                        }
                        foreach (Datas data in filingsArray.Filings.Datas)
                        {

                            if (data.IsTally == true)
                            {
                                data.Category = "General";
                            }
                            Datas obj1 = data;
                            if (obj1.Values != null)
                            {
                                if (obj1.Values.Count != num)
                                {
                                    int index = 0;
                                    foreach (Values values2 in valuesList2)
                                    {
                                        Values item = values2;
                                        Values values3 = new Values();
                                        values3.CElementName = (string)null;
                                        values3.CLineItem = (string)null;
                                        values3.Value = (string)null;
                                        if (obj1.Values.FirstOrDefault(x => Convert.ToDateTime(x.FilingDate) == Convert.ToDateTime(item.FilingDate)) == null)
                                        {
                                            values3.FilingDate = item.FilingDate;
                                            obj1.Values.Insert(index, values3);
                                        }
                                        ++index;
                                    }
                                }
                            }
                            else
                            {
                                List<Values> valuesList3 = new List<Values>();
                                obj1.Values = valuesList3;
                                foreach (Values values2 in valuesList2)
                                    obj1.Values.Add(new Values()
                                    {
                                        CElementName = (string)null,
                                        CLineItem = (string)null,
                                        Value = (string)null,
                                        FilingDate = values2.FilingDate
                                    });
                            }

                            if (mixedSubDatasList != null && mixedSubDatasList.Count > 0)
                            {
                                List<MixedSubDatas_FAnalysis> MixedSubDatas = InitialSetupID != null ? mixedSubDatasList.FindAll(x => x.DatasId == obj1.DataId && x.InitialSetup_FAnalysisId == InitialSetupID).ToList() : null;
                                obj1.MixedSubDatas_FAnalysis = MixedSubDatas;
                                if (InitialSetupID != null)
                                    foreach (MixedSubDatas_FAnalysis mixedSubData in obj1.MixedSubDatas_FAnalysis)
                                    {
                                        MixedSubDatas_FAnalysis item = mixedSubData;
                                        List<MixedSubValues_FAnalysis> mixedValuesbyDataId = mixedSubValuesList.FindAll(x => x.MixedSubDatas_FAnalysisId == item.Id).OrderBy(x => x.FilingDate).ToList();
                                        item.MixedSubValues_FAnalysis = mixedValuesbyDataId;
                                    }
                            }
                        }
                        filingsArrayList.Add(filingsArray);
                    }
                    renderResult.StatusCode = 1;
                    renderResult.InitialSetupId = InitialSetupID;
                    renderResult.categoryList = categoryList;
                    renderResult.Result = filingsArrayList;
                }
                return Ok(renderResult);
            }
            catch (Exception ex)
            {
                return (ActionResult)this.BadRequest((object)Convert.ToString(ex.Message));
            }
        }

        [HttpPost]
        [Route("FAnalaysis_Assigncategory/{InitialSetupId}")]
        public ActionResult FAnalaysis_Assigncategory([FromBody] List<LineItemInfoViewModel> lineItemInfoViewModel, long? InitialSetupId)
        {
            try
            {
                /////Update SEQUENCE in Filing Table
                var initialSetup = iInitialSetup_FAnalysis.GetSingle(x => x.Id == InitialSetupId);
                List<FilingsTable> FilingsListSequence = new List<FilingsTable>();
                if (initialSetup != null && !string.IsNullOrEmpty(initialSetup.CIKNumber))
                {
                    FilingsListSequence = iFilings.FindBy(x => x.CIK == initialSetup.CIKNumber).ToList();

                    if (FilingsListSequence != null && FilingsListSequence.Count > 0)
                    {
                        foreach (FilingsTable Ft in FilingsListSequence)
                        {
                            if (Ft.StatementType == "INCOME")
                            {
                                Ft.Sequence = 1;
                            }
                            else if (Ft.StatementType == "BALANCE_SHEET")
                            {
                                Ft.Sequence = 2;
                            }
                            else if (Ft.StatementType == "CASH_FLOW")
                            {
                                Ft.Sequence = 3;
                            }
                            iFilings.Update(Ft);
                            iFilings.Commit();
                        }

                    }

                    List<FAnalysis_CategoryByInitialSetup> categoryList = new List<FAnalysis_CategoryByInitialSetup>();
                    FAnalysis_CategoryByInitialSetup category;
                    try
                    {
                        foreach (LineItemInfoViewModel Datas in lineItemInfoViewModel)
                        {
                            category = new FAnalysis_CategoryByInitialSetup();
                            category.DatasId = Datas.Id;
                            category.Category = Datas.Category;
                            category.FAnalysis_InitialSetupId = initialSetup != null ? initialSetup.Id : 0;
                            categoryList.Add(category);

                            // check if id exist or not in SubDatasTable
                            List<MixedSubDatas_FAnalysis> SubdatasList = iMixedSubDatas_FAnalysis.FindBy(x => x.DatasId == Datas.Id && x.InitialSetup_FAnalysisId == InitialSetupId).ToList();
                            if (SubdatasList != null && SubdatasList.Count > 0)
                            {
                                //delete all items of same DataId from MixedSubDatas and MixedSubvalues
                                foreach (MixedSubDatas_FAnalysis subdataObj in SubdatasList)
                                {
                                    var mixedSubValues = iMixedSubValues_FAnalysis.FindBy(x => x.MixedSubDatas_FAnalysisId == subdataObj.Id).ToList();
                                    if (mixedSubValues != null && mixedSubValues.Count > 0)
                                    {
                                        // delete subValues                                        
                                        iMixedSubValues_FAnalysis.DeleteMany(mixedSubValues);
                                    }
                                    iMixedSubValues_FAnalysis.Commit();
                                }
                                //delete subdatas                               
                                iMixedSubDatas_FAnalysis.DeleteMany(SubdatasList);
                                iMixedSubDatas_FAnalysis.Commit();

                            }

                            //save to mixed SubDatas and MixedSubValues
                            if (Datas.MixedSubDatas_FAnalysis != null && Datas.MixedSubDatas_FAnalysis.Count > 0)
                            {
                                foreach (MixedSubDatas_FAnalysis mixedSubDataobj in Datas.MixedSubDatas_FAnalysis)
                                {
                                    mixedSubDataobj.Id = 0;
                                    mixedSubDataobj.InitialSetup_FAnalysisId = InitialSetupId;

                                    List<MixedSubValues_FAnalysis> tempValuesList = new List<MixedSubValues_FAnalysis>();
                                    if (mixedSubDataobj.MixedSubValues_FAnalysis != null && mixedSubDataobj.MixedSubValues_FAnalysis.Count > 0)
                                        tempValuesList = mixedSubDataobj.MixedSubValues_FAnalysis;

                                    mixedSubDataobj.MixedSubValues_FAnalysis = new List<MixedSubValues_FAnalysis>();
                                    iMixedSubDatas_FAnalysis.Add(mixedSubDataobj);
                                    iMixedSubDatas_FAnalysis.Commit();

                                    if (tempValuesList != null && tempValuesList.Count > 0)
                                    {
                                        foreach (var item in tempValuesList)
                                        {
                                            item.MixedSubDatas_FAnalysisId = mixedSubDataobj.Id;
                                        }
                                        iMixedSubValues_FAnalysis.AddMany(tempValuesList);
                                        iMixedSubValues_FAnalysis.Commit();
                                    }


                                    ////if (mixedSubDataobj.MixedSubValues_FAnalysis != null && mixedSubDataobj.MixedSubValues_FAnalysis.Count > 0)
                                    ////    foreach (MixedSubValues_FAnalysis values in mixedSubDataobj.MixedSubValues_FAnalysis)
                                    ////    {
                                    ////        values.Id = 0;
                                    ////    }
                                    ////iMixedSubDatas_FAnalysis.Add(mixedSubDataobj);
                                    ////iMixedSubDatas_FAnalysis.Commit();
                                }
                            }


                        }

                    }
                    catch (Exception ss)
                    {
                        return Ok("issue in save Mixed sub datas");
                    }

                    try
                    {
                        // save category
                        var tempCategory = initialSetup != null ? iFAnalysis_CategoryByInitialSetup.FindBy(x => categoryList.Any(m => m.DatasId == x.DatasId) && x.FAnalysis_InitialSetupId == initialSetup.Id).ToList() : null;
                        if (tempCategory != null && tempCategory.Count > 0)
                        {
                            iFAnalysis_CategoryByInitialSetup.DeleteMany(tempCategory);
                        }
                        iFAnalysis_CategoryByInitialSetup.AddMany(categoryList);
                        iFAnalysis_CategoryByInitialSetup.Commit();
                        // iLineItemInfoRepository.UpdatedMany(temList);
                    }
                    catch (Exception ss)
                    {
                        return Ok("issue in save category by Initial Setup" + InitialSetupId);
                    }
                    return Ok("Updated succesfully");
                }
                else
                    return Ok("Initial Setup not exist for" + InitialSetupId);

            }
            catch (Exception ss) { return Ok("Error occured"); }
        }

    
        #endregion

        #region Market Data
    

        [HttpGet]
        [Route("GetMarketData/{UserId}")]
        public ActionResult GetMarketData(long UserId)
        {
            MarketResult renderResult = new MarketResult();
            MarketDatasViewModel MarketDatasVM;
            MarketValuesViewModel MarketValuesVM;
            List<MarketDatasViewModel> MarketDatasList = new List<MarketDatasViewModel>();
            List<MarketValuesViewModel> MarketValuesList = new List<MarketValuesViewModel>();

            try
            {
                InitialSetup_FAnalysis InitialSetup_FAnalysisObj = iInitialSetup_FAnalysis.FindBy(x => x.UserId == UserId && x.IsActive == true).OrderByDescending(x => x.Id).First();
                long? Initialsetup_FAnalysisId = null;
                if (InitialSetup_FAnalysisObj != null)
                {
                    Initialsetup_FAnalysisId = InitialSetup_FAnalysisObj.Id;


                    int Startyear = Convert.ToInt32(InitialSetup_FAnalysisObj.YearFrom);
                    int Endyear = Convert.ToInt32(InitialSetup_FAnalysisObj.YearTo);
                    int Count = Endyear - Startyear + 1;
                    for (int j = 1; j <= Count; j++)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Convert.ToString(Startyear);
                        MarketValuesVM.Value = "";
                        MarketValuesList.Add(MarketValuesVM);
                        Startyear = Startyear + 1;
                    }

                }
                List<MarketDatas> tblMarketDataListObj = Initialsetup_FAnalysisId != null ? iMarketDatas.FindBy(x => x.InitialSetup_FAnalysisId == Initialsetup_FAnalysisId).OrderBy(x=>x.Sequence).ToList() : null;
                if (tblMarketDataListObj != null && tblMarketDataListObj.Count > 0)  //when data is saved 
                {

                    //get all Historical Values
                    List<MarketValues> MarketValuesListObj = iMarketValues.FindBy(x => tblMarketDataListObj.Any(m => m.Id == x.MarketDatasId)).ToList();

                    if (tblMarketDataListObj != null && tblMarketDataListObj.Count > 0)
                    {
                        foreach (MarketDatas obj in tblMarketDataListObj)
                        {
                            MarketDatasVM = new MarketDatasViewModel();
                            MarketDatasVM = mapper.Map<MarketDatas, MarketDatasViewModel>(obj);

                            // for Historical Values
                            List<MarketValues> MarketValueList = MarketValuesListObj.FindAll(x => x.MarketDatasId == obj.Id).ToList();
                            //Incomeobj.ForcastRatioValues = tempForcastValueList;
                            MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                            foreach (var tmpobj in MarketValueList)
                            {
                                MarketValuesViewModel tempValues = mapper.Map<MarketValues, MarketValuesViewModel>(tmpobj);
                                MarketDatasVM.MarketValuesVM.Add(tempValues);
                            }
                            MarketDatasList.Add(MarketDatasVM);
                        }
                    }
                }
                else
                {
                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Share Price ($)";
                    MarketDatasVM.Sequence = 1;
                    MarketDatasVM.IsHistorical_editable = true;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "";
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Number of Shares Outstanding - Basic (Millions)";
                    MarketDatasVM.Sequence = 2;
                    MarketDatasVM.IsHistorical_editable = true;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "";
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Market Value Common Stock ($M)";
                    MarketDatasVM.Sequence = 3;
                    MarketDatasVM.IsHistorical_editable = false;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "";
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Total Value Preferred Stock ($M)";
                    MarketDatasVM.Sequence = 4;
                    MarketDatasVM.IsHistorical_editable = true;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "0";
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Market Value of Debt ($M)";
                    MarketDatasVM.Sequence = 5;
                    MarketDatasVM.IsHistorical_editable = true;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "";
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Cash and Cash Equivalent ($M)";
                    MarketDatasVM.Sequence = 6;
                    MarketDatasVM.IsHistorical_editable = false;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    IntegratedDatasFAnalysis CashnCashEquivalentDatas = iIntegratedDatasFAnalysis.GetSingle(x => x.InitialSetup_FAnalysisId == Initialsetup_FAnalysisId && x.StatementTypeId == 2 && x.LineItem.ToLower().Contains("cash and cash equivalents"));
                    var CashnCashEquivalentValueList = CashnCashEquivalentDatas != null ? iIntegratedValuesFAnalysis.FindBy(x => x.IntegratedDatasFAnalysisId == CashnCashEquivalentDatas.Id).ToList() : null;

                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        double value = CashnCashEquivalentValueList != null && CashnCashEquivalentValueList.Count > 0 && CashnCashEquivalentValueList.Find(x => x.Year == Marketitem.Year) != null ? Convert.ToDouble(CashnCashEquivalentValueList.Find(x => x.Year == Marketitem.Year).Value) : 0;
                        MarketValuesVM.Value = value.ToString();
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Cash Needed for Working Capital ($M)";
                    MarketDatasVM.Sequence = 7;
                    MarketDatasVM.IsHistorical_editable = false;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "0";   //) Only for now, later it will calcuated from new ntegrated screen.
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Enterprise Value ($M)";
                    MarketDatasVM.Sequence = 8;
                    MarketDatasVM.IsHistorical_editable = false;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "";   //) Only for now, later it will calcuated from new ntegrated screen.
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "Tax Rate (%)";
                    MarketDatasVM.Sequence = 9;
                    MarketDatasVM.IsHistorical_editable = true;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                    {
                        MarketValuesVM = new MarketValuesViewModel();
                        MarketValuesVM.Year = Marketitem.Year;
                        MarketValuesVM.Id = 0;
                        MarketValuesVM.Value = "";   //) Only for now, later it will calcuated from new ntegrated screen.
                        MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                    }
                    MarketDatasList.Add(MarketDatasVM);

                    MarketDatasVM = new MarketDatasViewModel();
                    MarketDatasVM.LineItem = "NOPLAT";
                    MarketDatasVM.Sequence = 10;
                    MarketDatasVM.IsHistorical_editable = true;
                    MarketDatasVM.IsTally = false;
                    MarketDatasVM.Category = "";
                    MarketDatasVM.InitialSetup_FAnalysisId = Initialsetup_FAnalysisId;
                    MarketDatasVM.MarketValuesVM = new List<MarketValuesViewModel>();
                    if (InitialSetup_FAnalysisObj.SourceId == 1)
                    {
                        //Data get from ROIC
                        ROICDatas NOPLAT_Datas = iROICDatas.GetSingle(x => x.LineItem == "NOPLAT" && x.InitialSetupId == InitialSetup_FAnalysisObj.InitialSetupId);
                        List<ROICValues> NOPLAT_ValuesList = NOPLAT_Datas != null ? iROICValues.FindBy(x => x.ROICDatasId == NOPLAT_Datas.Id).ToList() : null;

                        foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                        {
                            MarketValuesVM = new MarketValuesViewModel();
                            MarketValuesVM.Year = Marketitem.Year;
                            MarketValuesVM.Id = 0;
                            double value = NOPLAT_ValuesList != null && NOPLAT_ValuesList.Count > 0 && NOPLAT_ValuesList.Find(x => x.Year == Marketitem.Year) != null && !string.IsNullOrEmpty(NOPLAT_ValuesList.Find(x => x.Year == Marketitem.Year).Value) ? Convert.ToDouble(NOPLAT_ValuesList.Find(x => x.Year == Marketitem.Year).Value) : 0;
                            MarketValuesVM.Value = value.ToString("0.##");
                            MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                        }

                    }
                    else
                    {
                        foreach (MarketValuesViewModel Marketitem in MarketValuesList)
                        {
                            MarketValuesVM = new MarketValuesViewModel();
                            MarketValuesVM.Year = Marketitem.Year;
                            MarketValuesVM.Id = 0;
                            MarketValuesVM.Value = "";   //) Only for now, later it will calcuated from new ntegrated screen.
                            MarketDatasVM.MarketValuesVM.Add(MarketValuesVM);
                        }
                    }
                    MarketDatasList.Add(MarketDatasVM);

                }
                renderResult.StatusCode = 1;
                renderResult.Result = MarketDatasList;
                return (ActionResult)this.Ok(new
                {
                    renderResult,
                    InitialSetup_IValuationObj = InitialSetup_FAnalysisObj
                });

            }
            catch (Exception ex)
            {
                return BadRequest(Convert.ToString(ex.Message));

            }

        }

        [HttpPost]
        [Route("SaveMarketData")]
        public ActionResult SaveMarketData([FromBody] List<MarketDatasViewModel> MarketdatasvmList)
        {
            try
            {
                if (MarketdatasvmList != null && MarketdatasvmList.Count > 0)
                {
                    long? initialSetup_FAnalysisId = 0;
                    foreach (MarketDatasViewModel MarketDatasObj in MarketdatasvmList)
                    {
                        initialSetup_FAnalysisId = MarketDatasObj != null ? MarketDatasObj.InitialSetup_FAnalysisId : 0;
                        MarketDatas tblMarketDatasObj = new MarketDatas();
                        if (MarketDatasObj != null)
                        {
                            tblMarketDatasObj = mapper.Map<MarketDatasViewModel, MarketDatas>(MarketDatasObj);
                            if (MarketDatasObj.MarketValuesVM != null && MarketDatasObj.MarketValuesVM.Count > 0)
                            {
                                tblMarketDatasObj.MarketValues = new List<MarketValues>();
                                foreach (MarketValuesViewModel MarketValues in MarketDatasObj.MarketValuesVM)
                                {
                                    MarketValues MValue = mapper.Map<MarketValuesViewModel, MarketValues>(MarketValues);
                                    tblMarketDatasObj.MarketValues.Add(MValue);
                                }
                            }
                        }

                        if (tblMarketDatasObj.Id == 0)
                        {
                            //Save Code
                            iMarketDatas.Add(tblMarketDatasObj);
                            iMarketDatas.Commit();
                        }
                        else
                        {
                            //Update Code
                            //iMarketDatas.Update(tblMarketDatasObj);
                            //iMarketDatas.Commit();

                            iMarketValues.UpdatedMany(tblMarketDatasObj.MarketValues);
                            iMarketValues.Commit();


                        }
                    }
                    return Ok(new { message = "Data Saved Successfully", status = 200, result = true });

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

        #region Integrated Data
    

        //GET
        [HttpGet]
        [Route("GetIntegratedData_FAnalysis/{UserId}/{cik}/{startYear?}/{endYear?}")]
        public ActionResult GetIntegratedData_FAnalysis(long UserId, string cik, int? startYear = null, int? endYear = null)
        {
            IntegratedFAnalysisResult integratedResult = new IntegratedFAnalysisResult();
            bool DepreciationFlag = false;
            string AdjustedMessage = "Depreciation";
            try
            {
                List<IntegratedFilingsFAnalysisViewModel> integratedFilingsList = new List<IntegratedFilingsFAnalysisViewModel>();
                List<IntegratedDatasFAnalysisViewModel> IntegratedDatasList = new List<IntegratedDatasFAnalysisViewModel>();
                IntegratedFilingsFAnalysisViewModel integratedFiling = new IntegratedFilingsFAnalysisViewModel();
                List<IntegratedDatasFAnalysis> tempintegratedDatasListObj = new List<IntegratedDatasFAnalysis>();
                IntegratedDatasFAnalysisViewModel integratedDatasVm = new IntegratedDatasFAnalysisViewModel();
                InitialSetup_FAnalysis InitialSetup_IValuationObj = iInitialSetup_FAnalysis.FindBy(x => x.UserId == UserId && x.IsActive == true).OrderByDescending(x => x.Id).First();
                //check in the database if exist then get by DB else go with the flow
                List<IntegratedDatasFAnalysis> tblintegrateddatasListObj = InitialSetup_IValuationObj != null ? iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetup_IValuationObj.Id).ToList() : null;

                #region For SourceId=1
                if (tblintegrateddatasListObj != null && tblintegrateddatasListObj.Count > 0)
                {
                    //Skip
                }
                else
                {
                    if (InitialSetup_IValuationObj.SourceId == 1)
                    {
                        List<IntegratedDatas> tblIntegratedDatasList = InitialSetup_IValuationObj != null ? iIntegratedDatas.FindBy(x => x.InitialSetupId == InitialSetup_IValuationObj.InitialSetupId).ToList() : null;

                        if (tblIntegratedDatasList != null && tblIntegratedDatasList.Count > 0)
                        {
                            List<IntegratedValues> IntegratedValuesList = iIntegratedValues.FindBy(x => tblIntegratedDatasList.Any(m => m.Id == x.IntegratedDatasId)).ToList();

                            IntegratedDatasFAnalysis IntegratedDatasFAnalysisObj = new IntegratedDatasFAnalysis();
                            IntegratedValuesFAnalysis IntegratedValuesFAnalysisObj = new IntegratedValuesFAnalysis();
                            foreach (IntegratedDatas IntegratedDatasSaveobj in tblIntegratedDatasList)
                            {
                                IntegratedDatasFAnalysisObj = new IntegratedDatasFAnalysis();
                                IntegratedDatasFAnalysisObj.Id = 0;
                                IntegratedDatasFAnalysisObj.InitialSetup_FAnalysisId = InitialSetup_IValuationObj.Id;
                                IntegratedDatasFAnalysisObj.LineItem = IntegratedDatasSaveobj.LineItem;
                                IntegratedDatasFAnalysisObj.IsTally = IntegratedDatasSaveobj.IsTally;
                                IntegratedDatasFAnalysisObj.Sequence = IntegratedDatasSaveobj.Sequence;
                                IntegratedDatasFAnalysisObj.Category = IntegratedDatasSaveobj.Category;
                                IntegratedDatasFAnalysisObj.StatementTypeId = IntegratedDatasSaveobj.StatementTypeId;
                                IntegratedDatasFAnalysisObj.IsParentItem = IntegratedDatasSaveobj.IsParentItem;
                                IntegratedDatasFAnalysisObj.IsHistorical_editable = IntegratedDatasSaveobj.IsHistorical_editable;
                                IntegratedDatasFAnalysisObj.IntegratedValuesFAnalysis = new List<IntegratedValuesFAnalysis>();
                                iIntegratedDatasFAnalysis.Add(IntegratedDatasFAnalysisObj);
                                iIntegratedDatasFAnalysis.Commit();

                                var tmpvalueList = IntegratedValuesList != null && IntegratedValuesList.Count > 0 ? IntegratedValuesList.FindAll(x => x.IntegratedDatasId == IntegratedDatasSaveobj.Id).ToList() : null;
                                if (tmpvalueList != null && tmpvalueList.Count > 0)
                                {
                                    foreach (IntegratedValues IntegratedValuesSaveobj in tmpvalueList)
                                    {
                                        IntegratedValuesFAnalysisObj = new IntegratedValuesFAnalysis();
                                        IntegratedValuesFAnalysisObj.Id = 0;
                                        IntegratedValuesFAnalysisObj.IntegratedDatasFAnalysisId = IntegratedDatasFAnalysisObj.Id;
                                        IntegratedValuesFAnalysisObj.FilingDate = IntegratedValuesSaveobj.FilingDate;
                                        IntegratedValuesFAnalysisObj.Year = IntegratedValuesSaveobj.Year;
                                        IntegratedValuesFAnalysisObj.Value = IntegratedValuesSaveobj.Value;
                                        iIntegratedValuesFAnalysis.Add(IntegratedValuesFAnalysisObj);
                                        iIntegratedValuesFAnalysis.Commit();
                                    }
                                }

                            }

                            /////Save Cash Flow in Integrated/////
                            string edgarView = edgarDataRepository.GetEdgar(cik, startYear, endYear).FirstOrDefault<EdgarData>().EdgarView;
                            if (edgarView != null)
                            {
                                List<List<Filings>> filingsListList = (List<List<Filings>>)JsonConvert.DeserializeObject<List<List<Filings>>>(edgarView);
                                List<CategoryByInitialSetup> categoryByInitialSetups = InitialSetup_IValuationObj != null ? iCategoryByInitialSetup.FindBy(x => x.InitialSetupId == InitialSetup_IValuationObj.InitialSetupId).ToList() : null;

                                foreach (List<Filings> filingsList1 in filingsListList)
                                {

                                    if (filingsList1[0].StatementType == "CASH_FLOW")
                                    {
                                        //Remove all Null Value
                                        var Removedataslist = filingsList1[0].Datas.FindAll(x => x.Values == null || x.Values.Count == 0);
                                        if (Removedataslist != null)
                                        {
                                            foreach (var dt in Removedataslist)
                                            {
                                                filingsList1[0].Datas.Remove(dt);
                                            }
                                        }
                                        //////////////////////////////
                                        Datas datas1 = new Datas();
                                        List<Datas> DataProcessing_list = filingsList1[0].Datas.OrderBy(x => x.Sequence).ToList();
                                        Datas MaxValues_data = DataProcessing_list.First<Datas>();
                                        List<Datas> Integrated_datasList = new List<Datas>();
                                        // Find Total Mixed category
                                        List<Datas> MixedList = DataProcessing_list.FindAll(x => x.Category == "Mixed").ToList();
                                        foreach (Datas dt in DataProcessing_list)
                                        {
                                            if (dt.Category != "Mixed" || dt.IsParentItem == true)
                                            {
                                                Integrated_datasList.Add(dt);
                                            }
                                            else
                                            {
                                                foreach (Datas mixedDT in MixedList)
                                                {
                                                    if (mixedDT.DataId == dt.DataId)
                                                    {
                                                        var MixedSubdataList = iMixedSubDatas.FindBy(x => x.DatasId == dt.DataId && x.InitialSetupId == InitialSetup_IValuationObj.InitialSetupId).ToList();
                                                        if (MixedSubdataList != null && MixedSubdataList.Count > 0)
                                                        {
                                                            foreach (var item in MixedSubdataList)
                                                            {
                                                                var MixedSubValuesList = iMixedSubValues.FindBy(x => x.MixedSubDatasId == item.Id).ToList();
                                                                item.MixedSubValues = new List<MixedSubValues>();
                                                                item.MixedSubValues = MixedSubValuesList != null && MixedSubValuesList.Count > 0 ? MixedSubValuesList : new List<MixedSubValues>();

                                                            }
                                                        }

                                                        Datas Balance_datas = new Datas();
                                                        foreach (MixedSubDatas msd in MixedSubdataList)
                                                        {
                                                            Balance_datas = new Datas();
                                                            Balance_datas.LineItem = msd.LineItem;
                                                            Balance_datas.DataId = msd.DatasId;
                                                            Balance_datas.IsTally = msd.IsTally;
                                                            Balance_datas.Sequence = (Integrated_datasList.Count + 1);
                                                            Balance_datas.Category = msd.Category;
                                                            Balance_datas.IsParentItem = false;
                                                            Balance_datas.Values = new List<Values>();
                                                            foreach (MixedSubValues val in msd.MixedSubValues)
                                                            {
                                                                Values MNonoperatingValues = new Values();
                                                                MNonoperatingValues.FilingDate = val.FilingDate;
                                                                MNonoperatingValues.Value = val.Value;
                                                                Balance_datas.Values.Add(MNonoperatingValues);
                                                            }
                                                            Integrated_datasList.Add(Balance_datas);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        // filingsList1[0].Datas = Integrated_datasList;
                                        var IntegratedDatasBalanceChk = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetup_IValuationObj.Id && x.StatementTypeId == 3).ToList();
                                        // Save BalanceSheet data in IntegratedDatas and Values
                                        if (IntegratedDatasBalanceChk == null || IntegratedDatasBalanceChk.Count == 0)
                                        {
                                            foreach (var item in Integrated_datasList)
                                            {
                                                IntegratedDatasFAnalysis integratedDatasObj = new IntegratedDatasFAnalysis();
                                                integratedDatasObj.Id = 0;
                                                integratedDatasObj.Category = item.Category;
                                                integratedDatasObj.Sequence = item.Sequence;
                                                integratedDatasObj.LineItem = item.LineItem;
                                                integratedDatasObj.IsTally = item.IsTally;
                                                integratedDatasObj.IsParentItem = item.IsParentItem != null ? item.IsParentItem : false;
                                                integratedDatasObj.StatementTypeId = (int)StatementTypeEnum.CashFlowStatement;
                                                integratedDatasObj.IntegratedValuesFAnalysis = new List<IntegratedValuesFAnalysis>();

                                                //foreach (var valueobj in item.Values)
                                                //{
                                                //    IntegratedValuesFAnalysis obj = new IntegratedValuesFAnalysis();
                                                //    obj.Id = 0;
                                                //    obj.FilingDate = valueobj.FilingDate;
                                                //    obj.Value = valueobj.Value;
                                                //    DateTime Dt = new DateTime();
                                                //    Dt = Convert.ToDateTime(valueobj.FilingDate);
                                                //    obj.Year = Convert.ToString(Dt.Year);
                                                //    integratedDatasObj.IntegratedValuesFAnalysis.Add(obj);
                                                //}

                                                //////////////New Code//////////////
                                                int? strt = startYear;
                                                int? end = endYear;
                                                for (int? i = strt; i <= end; i++)
                                                {
                                                    IntegratedValuesFAnalysis obj = new IntegratedValuesFAnalysis();
                                                    obj.Id = 0;

                                                    var valueobj = item.Values.Find(x => Convert.ToDateTime(x.FilingDate).Year == i);
                                                    if (valueobj != null)
                                                    {
                                                        obj.FilingDate = valueobj.FilingDate;
                                                        obj.Value = valueobj.Value;
                                                        DateTime Dt = new DateTime();
                                                        Dt = Convert.ToDateTime(valueobj.FilingDate);
                                                        obj.Year = Convert.ToString(Dt.Year);
                                                    }
                                                    else
                                                    {
                                                        obj.FilingDate = "01-jan-" + i;
                                                        obj.Value = null;
                                                        obj.Year = Convert.ToString(i);
                                                    }
                                                    integratedDatasObj.IntegratedValuesFAnalysis.Add(obj);
                                                }
                                                //////////////////////////////////////

                                                integratedDatasObj.InitialSetup_FAnalysisId = InitialSetup_IValuationObj.Id;
                                                iIntegratedDatasFAnalysis.Add(integratedDatasObj);
                                                iIntegratedDatasFAnalysis.Commit();
                                            }
                                        }

                                    }
                                }
                            }
                            //////////////////////////////////////
                        }

                    }
                }
                #endregion

                tblintegrateddatasListObj = InitialSetup_IValuationObj != null ? iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetup_IValuationObj.Id).ToList() : null;
                if (tblintegrateddatasListObj != null && tblintegrateddatasListObj.Count > 0)
                {
                    //get all Historical Values
                    List<IntegratedValuesFAnalysis> IntegratedValuesListObj = iIntegratedValuesFAnalysis.FindBy(x => tblintegrateddatasListObj.Any(m => m.Id == x.IntegratedDatasFAnalysisId)).ToList();

                    //get all Explicit Values 
                    //List<Integrated_ExplicitValues> Integrated_explicitValuesAfterListObj = iIntegrated_ExplicitValues.FindBy(x => tblintegrateddatasListObj.Any(m => m.Id == x.IntegratedDatasId)).ToList();

                    List<FilingsTable> filingsList = new List<FilingsTable>();
                    filingsList = iFilings.FindBy(x => x.CIK == cik).OrderBy(x => x.Sequence).ToList();
                    //  filingsList = iFilings.FindBy(x => x.InitialSetupId == InitialSetup_IValuationObj.Id).ToList();

                    //check depreciation Deducted

                    Datas Depreciation_datas = new Datas();
                    foreach (FilingsTable filingsTable in filingsList)
                    {
                        if (filingsTable.StatementType == "INCOME")
                        {

                            var incomeDepreciation_datas = iDatas.GetSingle(x => x.FilingId == filingsTable.Id && x.LineItem.Contains("Depreciation"));
                            if (incomeDepreciation_datas != null)
                                break;
                        }
                        else
                        if (filingsTable.StatementType == "CASH_FLOW")
                        {
                            var Depreciation = iDatas.GetSingle(x => x.FilingId == filingsTable.Id && x.LineItem.Contains("Depreciation"));
                            if (Depreciation != null)
                            {
                                DepreciationFlag = true;
                                // AdjustedMessage=  Depreciation.LineItem ;
                            }
                        }
                    }
                    string CompanyNameRetainedEarnings = "";
                    foreach (var filing in filingsList)
                    {
                        IntegratedDatasList = new List<IntegratedDatasFAnalysisViewModel>();
                        tempintegratedDatasListObj = new List<IntegratedDatasFAnalysis>();
                        integratedFiling = new IntegratedFilingsFAnalysisViewModel();
                        integratedFiling.CompanyName = filing.CompanyName;
                        CompanyNameRetainedEarnings = filing.CompanyName;
                        integratedFiling.ReportName = filing.ReportName;
                        integratedFiling.StatementType = filing.StatementType;
                        integratedFiling.Unit = filing.Unit;
                        integratedFiling.CIK = filing.CIK;
                        if (filing.StatementType == "INCOME")
                        {
                            // add Income items to datas here
                            tempintegratedDatasListObj = tblintegrateddatasListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement).ToList();
                        }
                        else if (filing.StatementType == "BALANCE_SHEET")
                        {
                            // add Balance items to datas here
                            tempintegratedDatasListObj = tblintegrateddatasListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet).ToList();
                        }
                        else if (filing.StatementType == "CASH_FLOW")
                        {
                            integratedFiling.ReportName = "STATEMENT OF CASH FLOW";
                            integratedFiling.StatementType = "STATEMENT OF CASH FLOW";
                            tempintegratedDatasListObj = tblintegrateddatasListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.CashFlowStatement).ToList();
                        }
                        else
                        {
                            //Retained Earnings
                            integratedFiling.ReportName = "STATEMENT OF RETAINED EARNINGS";
                            integratedFiling.StatementType = "STATEMENT OF RETAINED EARNINGS";
                            tempintegratedDatasListObj = tblintegrateddatasListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.RetainedEarningsStatement).ToList();
                        }

                        if (tempintegratedDatasListObj != null && tempintegratedDatasListObj.Count > 0)
                        {

                            foreach (IntegratedDatasFAnalysis Incomeobj in tempintegratedDatasListObj)
                            {
                                integratedDatasVm = new IntegratedDatasFAnalysisViewModel();
                                integratedDatasVm = mapper.Map<IntegratedDatasFAnalysis, IntegratedDatasFAnalysisViewModel>(Incomeobj);
                                // for Historical Values
                                List<IntegratedValuesFAnalysis> tempForcastValueList = IntegratedValuesListObj.FindAll(x => x.IntegratedDatasFAnalysisId == Incomeobj.Id).ToList();
                                //Incomeobj.ForcastRatioValues = tempForcastValueList;
                                integratedDatasVm.IntegratedValuesFAnalysisVM = new List<IntegratedValuesFAnalysisViewModel>();
                                foreach (var obj in tempForcastValueList)
                                {
                                    IntegratedValuesFAnalysisViewModel tempValues = mapper.Map<IntegratedValuesFAnalysis, IntegratedValuesFAnalysisViewModel>(obj);
                                    integratedDatasVm.IntegratedValuesFAnalysisVM.Add(tempValues);
                                }

                                // for Explicit Values
                                //List<Integrated_ExplicitValues> tempForcast_ExplicitValueList = Integrated_explicitValuesAfterListObj.FindAll(x => x.IntegratedDatasId == Incomeobj.Id).ToList();
                                //integratedDatasVm.Integrated_ExplicitValuesVM = new List<Integrated_ExplicitValuesViewModel>();
                                //foreach (var obj in tempForcast_ExplicitValueList)
                                //{
                                //    Integrated_ExplicitValuesViewModel tempExplicitValues = mapper.Map<Integrated_ExplicitValues, Integrated_ExplicitValuesViewModel>(obj);
                                //    integratedDatasVm.Integrated_ExplicitValuesVM.Add(tempExplicitValues);
                                //}
                                IntegratedDatasList.Add(integratedDatasVm);
                            }
                            //IntegratedDatasList.Add();
                            integratedFiling.IntegratedDatasFAnalysisVM = IntegratedDatasList;
                            integratedFilingsList.Add(integratedFiling);
                        }
                    }

                    //////////////Retained Earnings//////////////
                    IntegratedDatasList = new List<IntegratedDatasFAnalysisViewModel>();
                    tempintegratedDatasListObj = new List<IntegratedDatasFAnalysis>();
                    integratedFiling = new IntegratedFilingsFAnalysisViewModel();
                    integratedFiling.CompanyName = CompanyNameRetainedEarnings;
                    integratedFiling.ReportName = "STATEMENT OF RETAINED EARNINGS";
                    integratedFiling.StatementType = "STATEMENT OF RETAINED EARNINGS";
                    tempintegratedDatasListObj = tblintegrateddatasListObj.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.RetainedEarningsStatement).ToList();
                    if (tempintegratedDatasListObj != null && tempintegratedDatasListObj.Count > 0)
                    {

                        foreach (IntegratedDatasFAnalysis Incomeobj in tempintegratedDatasListObj)
                        {
                            integratedDatasVm = new IntegratedDatasFAnalysisViewModel();
                            integratedDatasVm = mapper.Map<IntegratedDatasFAnalysis, IntegratedDatasFAnalysisViewModel>(Incomeobj);
                            // for Historical Values
                            List<IntegratedValuesFAnalysis> tempForcastValueList = IntegratedValuesListObj.FindAll(x => x.IntegratedDatasFAnalysisId == Incomeobj.Id).ToList();
                            //Incomeobj.ForcastRatioValues = tempForcastValueList;
                            integratedDatasVm.IntegratedValuesFAnalysisVM = new List<IntegratedValuesFAnalysisViewModel>();
                            foreach (var obj in tempForcastValueList)
                            {
                                IntegratedValuesFAnalysisViewModel tempValues = mapper.Map<IntegratedValuesFAnalysis, IntegratedValuesFAnalysisViewModel>(obj);
                                integratedDatasVm.IntegratedValuesFAnalysisVM.Add(tempValues);
                            }
                            IntegratedDatasList.Add(integratedDatasVm);
                        }
                        //IntegratedDatasList.Add();
                        integratedFiling.IntegratedDatasFAnalysisVM = IntegratedDatasList;
                        integratedFilingsList.Add(integratedFiling);
                    }
                    /////////////////////////////////////////////
                }

                integratedResult.Result = integratedFilingsList;
                integratedResult.DepreciationFlag = DepreciationFlag;
                integratedResult.AdjustedMessage = AdjustedMessage;
                integratedResult.StatusCode = 1;
                return Ok(integratedResult);
            }
            catch (Exception ss)
            {
                integratedResult.StatusCode = 0;
                integratedResult.AdjustedMessage = AdjustedMessage;
                integratedResult.DepreciationFlag = DepreciationFlag;
                integratedResult.Message = Convert.ToString(ss.Message);
                return BadRequest(integratedResult);
            }

        }

        //SAVE
        [HttpGet]
        [Route("GetIntegratedDataforAnalysis/{UserId}/{InitialsetupId}/{StatementType}/{cik}/{startYear?}/{endYear?}")]
        public ActionResult GetIntegratedDataforAnalysis(long UserId, long? InitialsetupId, string StatementType, string cik, int? startYear = null, int? endYear = null)
        {
            RenderResult renderResult = new RenderResult();
            List<FilingsArray> filingsArrayList = new List<FilingsArray>();
            try
            {
                long? InitialSetupId = null;
                InitialSetup_FAnalysis InitialSetup_IValuationObj = iInitialSetup_FAnalysis.FindBy(x => x.Id == InitialsetupId).OrderByDescending(x => x.Id).First();
                //InitialSetup_IValuation InitialSetup_IValuationObj = iInitialSetup_IValuation.FindBy(x => x.UserId == UserId && x.IsActive == true).OrderByDescending(x => x.Id).First();
                if (InitialSetup_IValuationObj != null)
                {
                    InitialSetupId = InitialSetup_IValuationObj.Id;
                }
                List<IntegratedDatasFAnalysis> integratedDatasobj = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupId).ToList();
                if (integratedDatasobj != null && integratedDatasobj.Count > 0)
                {
                    //call delete api for every statement
                    //int StatementTypeId = 0;
                    //List<string> statementTypeList = StatementType.Split(',').ToList();
                    //if (statementTypeList != null && statementTypeList.Count > 0)
                    //foreach (var item in statementTypeList)
                    //{
                    //    StatementTypeId = item.ToLower().Contains("income") ? 1 : item.ToLower().Contains("balance") ? 2 : 0;
                    //    bool flag = DeleteCalculatedDataIntegrated_ForcastRatio(UserId, InitialsetupId, StatementTypeId, cik, startYear, endYear);
                    //}

                }
                else
                {
                    string CompanyName = "";
                    string edgarView = edgarDataRepository.GetEdgar(cik, startYear, endYear).FirstOrDefault<EdgarData>().EdgarView;
                    if (edgarView == null)
                    {
                        renderResult.StatusCode = 0;
                        renderResult.Result = filingsArrayList;
                    }
                    else
                    {
                        List<List<Filings>> filingsListList = (List<List<Filings>>)JsonConvert.DeserializeObject<List<List<Filings>>>(edgarView);

                        List<FAnalysis_CategoryByInitialSetup> categoryByInitialSetups = InitialSetupId != null ? iFAnalysis_CategoryByInitialSetup.FindBy(x => x.FAnalysis_InitialSetupId == InitialSetupId).ToList() : null;


                        foreach (List<Filings> filingsList1 in filingsListList)
                        {
                            //Remove all Null Value
                            var Removedataslist = filingsList1[0].Datas.FindAll(x => x.Values == null || x.Values.Count == 0);
                            if (Removedataslist != null)
                            {
                                foreach (var dt in Removedataslist)
                                {
                                    filingsList1[0].Datas.Remove(dt);
                                }
                            }
                            //////////////////////////////

                            filingsList1[0].InitialSetupId = InitialSetupId;
                            FilingsArray filingsArray = new FilingsArray()
                            {
                                CompanyName = filingsList1[0].CompanyName,
                                StatementType = filingsList1[0].StatementType,
                                Filings = filingsList1[0]
                            };
                            CompanyName = filingsList1[0].CompanyName;
                            int num1 = 0;
                            List<Values> valuesList1 = new List<Values>();
                            List<Values> valuesList2 = new List<Values>();
                            Values values1 = new Values();


                            foreach (Datas data in filingsList1[0].Datas)
                            {
                                if (data.IsTally == true)
                                    data.Category = "General";
                                if (categoryByInitialSetups != null && categoryByInitialSetups.Count > 0)
                                {
                                    var assigncategory = categoryByInitialSetups.Find(x => x.DatasId == data.DataId);
                                    if (assigncategory != null)
                                    {
                                        data.Category = assigncategory.Category != null ? assigncategory.Category : data.Category;
                                    }
                                }
                                if (data.Values != null && num1 < data.Values.Count)
                                {
                                    num1 = data.Values.Count;
                                    valuesList2 = data.Values;
                                }
                            }


                            foreach (Datas data in filingsArray.Filings.Datas)
                            {
                                if (data.Values != null)
                                {
                                    if (data.Values.Count != num1)
                                    {
                                        int index = 0;
                                        foreach (Values values2 in valuesList2)
                                        {
                                            Values item = values2;
                                            Values values3 = new Values();
                                            values3.CElementName = (string)null;
                                            values3.CLineItem = (string)null;
                                            values3.Value = (string)null;
                                            if (data.Values.FirstOrDefault((x => Convert.ToDateTime(x.FilingDate) == Convert.ToDateTime(item.FilingDate))) == null)
                                            {
                                                values3.FilingDate = item.FilingDate;
                                                data.Values.Insert(index, values3);
                                            }
                                            ++index;
                                        }
                                    }
                                }
                                else
                                {
                                    List<Values> valuesList3 = new List<Values>();
                                    data.Values = valuesList3;
                                    foreach (Values values2 in valuesList2)
                                        data.Values.Add(new Values()
                                        {
                                            CElementName = (string)null,
                                            CLineItem = (string)null,
                                            Value = (string)null,
                                            FilingDate = values2.FilingDate
                                        });
                                }
                            }


                            filingsArrayList.Add(filingsArray);
                            if (filingsList1[0].StatementType == "INCOME")
                            {
                                Datas datas1 = new Datas();
                                List<Datas> DataProcessing_list = filingsList1[0].Datas.OrderBy(x => x.Sequence).ToList();
                                Datas MaxValues_data = DataProcessing_list.First<Datas>();
                                List<Datas> Integrated_datasList = new List<Datas>();
                                long gross_Seq = 0;
                                bool depreFlag = false;
                                bool amortizationFlag = false;

                                Datas Depreciation_datas = new Datas();
                                Datas Amortization_datas = new Datas();
                                foreach (List<Filings> filingsList2 in filingsListList)
                                {
                                    if (filingsList2[0].StatementType == "INCOME")
                                    {
                                        var incomeDepreciation_datas = filingsList2[0].Datas.Find(x => x.LineItem.Contains("Depreciation"));
                                        //Depreciation_datas
                                        if (incomeDepreciation_datas != null)
                                        {
                                            Depreciation_datas = incomeDepreciation_datas;
                                            depreFlag = true;
                                            break;
                                        }
                                    }
                                    else
                                    if (filingsList2[0].StatementType == "CASH_FLOW")
                                        Depreciation_datas = filingsList2[0].Datas.Find(x => x.LineItem.Contains("Depreciation"));
                                }
                                if (Depreciation_datas != null)
                                {
                                    //check if Depreciation exist in mixed Datas or not
                                    MixedSubDatas_FAnalysis mixedSubDatas = iMixedSubDatas_FAnalysis.GetSingle(x => x.DatasId == Depreciation_datas.DataId && x.InitialSetup_FAnalysisId == InitialSetupId && x.Category == "Operating");
                                    if (mixedSubDatas != null)
                                    {
                                        List<MixedSubValues_FAnalysis> mixedSubValuesList = new List<MixedSubValues_FAnalysis>();
                                        mixedSubValuesList = iMixedSubValues_FAnalysis.FindBy(x => x.MixedSubDatas_FAnalysisId == mixedSubDatas.Id).ToList();
                                        if (mixedSubValuesList != null && mixedSubValuesList.Count > 0)
                                        {
                                            Depreciation_datas = mapper.Map<MixedSubDatas_FAnalysis, Datas>(mixedSubDatas);
                                            Depreciation_datas.Values = new List<Values>();
                                            foreach (var mixedValues in mixedSubDatas.MixedSubValues_FAnalysis)
                                            {
                                                Values mappedValue = mapper.Map<MixedSubValues_FAnalysis, Values>(mixedValues);
                                                Depreciation_datas.Values.Add(mappedValue);
                                            }

                                        }
                                    }
                                }

                                //find aamortization
                                foreach (List<Filings> filingsList2 in filingsListList)
                                {
                                    if (filingsList2[0].StatementType == "INCOME")
                                    {
                                        var incomeAmortization_datas = filingsList2[0].Datas.Find(x => x.LineItem.ToLower().Contains("amortization"));
                                        //Amortization_datas
                                        if (incomeAmortization_datas != null)
                                        {
                                            Amortization_datas = incomeAmortization_datas;
                                            amortizationFlag = true;
                                            break;
                                        }
                                    }
                                    else
                                    if (filingsList2[0].StatementType == "CASH_FLOW")
                                        Amortization_datas = filingsList2[0].Datas.Find(x => x.LineItem.ToLower().Contains("amortization"));
                                }
                                if (Amortization_datas != null)
                                {
                                    //check if Depreciation exist in mixed Datas or not
                                    MixedSubDatas_FAnalysis mixedSubDatas = iMixedSubDatas_FAnalysis.GetSingle(x => x.DatasId == Amortization_datas.DataId && x.InitialSetup_FAnalysisId == InitialSetupId && x.Category == "Non-Operating");
                                    if (mixedSubDatas != null)
                                    {
                                        List<MixedSubValues_FAnalysis> mixedSubValuesList = new List<MixedSubValues_FAnalysis>();
                                        mixedSubValuesList = iMixedSubValues_FAnalysis.FindBy(x => x.MixedSubDatas_FAnalysisId == mixedSubDatas.Id).ToList();
                                        if (mixedSubValuesList != null && mixedSubValuesList.Count > 0)
                                        {
                                            Amortization_datas = mapper.Map<MixedSubDatas_FAnalysis, Datas>(mixedSubDatas);
                                            Amortization_datas.Values = new List<Values>();
                                            foreach (var mixedValues in mixedSubDatas.MixedSubValues_FAnalysis)
                                            {
                                                Values mappedValue = mapper.Map<MixedSubValues_FAnalysis, Values>(mixedValues);
                                                Amortization_datas.Values.Add(mappedValue);
                                            }

                                        }
                                    }
                                }


                                Datas revenue_integratedDatas = new Datas();
                                Datas costofSales_integratedDatas = new Datas();

                                //find revenue
                                string revenuesynonyms = "Net Sales%Net Revenue%Revenue%Total Revenues%Sales%Total Net Revenue%Total revenue%Total net sales%Sales to customers%Total net revenues%Total revenues (Note 4)%Revenue from Contract with Customer, Excluding Assessed Tax%Revenues%Net revenues%Revenue, net";
                                string Costofsalesynonyms = "Cost of sales%COGS%Cost of Goods Sold%Cost of Revenue%Cost of Products Sold%Total cost of revenue%Total cost of revenues%Cost of products sold, excluding amortization of intangible assets%Costs of goods sold%Cost of equipment and services revenues%Cost of revenues%Cost of revenue (COR)";
                                bool revenueflag = false;
                                bool costofsalesflag = false;
                                List<string> synonyms = revenuesynonyms.Split('%').ToList(); //convert comma seperated values to list

                                int mulBy = -1;
                                foreach (Datas datas4 in DataProcessing_list)
                                {
                                    if (datas4.IsParentItem != true)
                                    {


                                        if (datas4.LineItem.Contains("Gross"))
                                        {
                                            gross_Seq = datas4.Sequence;
                                            foreach (Values obj in datas4.Values)
                                            {
                                                //get sum of revenue and COGS
                                                Values revenue = revenue_integratedDatas != null && revenue_integratedDatas.Values != null && revenue_integratedDatas.Values.Count > 0 ? revenue_integratedDatas.Values.Find(x => x.FilingDate == obj.FilingDate) : null;
                                                Values coGS = costofSales_integratedDatas != null && costofSales_integratedDatas.Values != null && costofSales_integratedDatas.Values.Count > 0 ? costofSales_integratedDatas.Values.Find(x => x.FilingDate == obj.FilingDate) : null;
                                                double calculatedgross = 0;

                                                calculatedgross = (revenue != null && !string.IsNullOrEmpty(revenue.Value) ? Convert.ToDouble(revenue.Value) : 0) + (coGS != null && !string.IsNullOrEmpty(coGS.Value) ? Convert.ToDouble(coGS.Value) : 0);
                                                double num2 = calculatedgross != 0 ? calculatedgross : (obj.Value != null ? Convert.ToDouble(obj.Value) : 0.0);
                                                obj.Value = Convert.ToString(num2);
                                            }
                                            Integrated_datasList.Add(datas4);
                                            break;
                                        }
                                        else if (costofsalesflag == true)
                                        {
                                            //add Gross Margin
                                            gross_Seq = datas4.Sequence;
                                            ////////////
                                            Datas Gross_datas = new Datas();
                                            Gross_datas.LineItem = "Gross Margin";
                                            Gross_datas.IsTally = true;
                                            //Gross_datas.Sequence = (Integrated_datasList.Count + 1);
                                            Gross_datas.Sequence = gross_Seq;
                                            Gross_datas.Category = (string)null;
                                            Gross_datas.IsParentItem = false;
                                            Gross_datas.Values = new List<Values>();


                                            foreach (Values obj in MaxValues_data.Values)
                                            {

                                                Values valueObj = new Values();
                                                //get sum of revenue and COGS
                                                Values revenue = revenue_integratedDatas != null && revenue_integratedDatas.Values != null && revenue_integratedDatas.Values.Count > 0 ? revenue_integratedDatas.Values.Find(x => x.FilingDate == obj.FilingDate) : null;
                                                Values coGS = costofSales_integratedDatas != null && costofSales_integratedDatas.Values != null && costofSales_integratedDatas.Values.Count > 0 ? costofSales_integratedDatas.Values.Find(x => x.FilingDate == obj.FilingDate) : null;
                                                double calculatedgross = 0;

                                                calculatedgross = (revenue != null && !string.IsNullOrEmpty(revenue.Value) ? Convert.ToDouble(revenue.Value) : 0) + (coGS != null && !string.IsNullOrEmpty(coGS.Value) ? Convert.ToDouble(coGS.Value) : 0);
                                                double num2 = calculatedgross != 0 ? calculatedgross : (obj.Value != null ? Convert.ToDouble(obj.Value) : 0.0);
                                                obj.Value = Convert.ToString(num2);
                                                valueObj.FilingDate = obj.FilingDate;
                                                valueObj.Value = Convert.ToString(num2);
                                                Gross_datas.Values.Add(valueObj);
                                            }
                                            Integrated_datasList.Add(Gross_datas);
                                            break;
                                        }


                                        if (revenueflag == true)
                                        {
                                            if (costofsalesflag == false)
                                                foreach (var syn in synonyms)
                                                {
                                                    if (datas4.LineItem.ToUpper() == syn.ToUpper())
                                                    {
                                                        costofsalesflag = true;
                                                        break;
                                                    }
                                                }

                                            if (datas4.Values != null && datas4.Values.Count > 0)
                                            {
                                                if (costofsalesflag == true && depreFlag == false)
                                                {
                                                    foreach (var item in datas4.Values)
                                                    {
                                                        Values depreciationValue = Depreciation_datas != null && Depreciation_datas.Values != null && Depreciation_datas.Values.Count > 0 ? Depreciation_datas.Values.Find(x => x.FilingDate == item.FilingDate) : null;
                                                        item.Value = !string.IsNullOrEmpty(item.Value) ? Convert.ToString(mulBy * (Convert.ToDouble(item.Value) - (depreciationValue != null && !string.IsNullOrEmpty(depreciationValue.Value) ? Convert.ToDouble(depreciationValue.Value) : 0))) : null;
                                                    }
                                                    costofSales_integratedDatas = datas4;
                                                }
                                                else
                                                {
                                                    foreach (var item in datas4.Values)
                                                    {
                                                        item.Value = !string.IsNullOrEmpty(item.Value) ? Convert.ToString(mulBy * Convert.ToDouble(item.Value)) : null;
                                                    }
                                                }
                                            }



                                        }

                                        Integrated_datasList.Add(datas4);

                                        if (revenueflag == false)
                                            foreach (var syn in synonyms)
                                            {
                                                if (datas4.LineItem.ToUpper() == syn.ToUpper())
                                                {

                                                    // mulBy = -1;
                                                    revenueflag = true;

                                                    revenue_integratedDatas = datas4;
                                                    // change revenue synonyms to cost of sales synonym
                                                    synonyms = new List<string>();
                                                    synonyms = Costofsalesynonyms.Split('%').ToList(); // convert comma seperated values to list
                                                    break;
                                                }
                                            }



                                    }

                                }

                                //if (categoryByInitialSetups != null && categoryByInitialSetups.Count > 0)
                                //{
                                //    var list = categoryByInitialSetups.FindAll(x => x.Category == "Operating").ToList();
                                //}

                                foreach (Datas datas4 in DataProcessing_list.Where(x => x.Sequence >= gross_Seq).ToList().FindAll(x => x.Category == "Operating" && !x.LineItem.ToLower().Contains("depreciation") && !x.LineItem.ToLower().Contains("amortization")).OrderBy(x => x.Sequence).ToList())
                                {
                                    if (datas4.Values != null && datas4.Values.Count > 0)
                                        foreach (var item in datas4.Values)
                                        {
                                            item.Value = !string.IsNullOrEmpty(item.Value) ? Convert.ToString(mulBy * Convert.ToDouble(item.Value)) : null;
                                        }
                                    //var checkcategory = categoryByInitialSetups.Find(x => x.DatasId == datas4.DataId);
                                    Integrated_datasList.Add(datas4);
                                }

                                // Find Total Mixed category
                                List<Datas> MixedList = DataProcessing_list.FindAll(x => x.Category == "Mixed").ToList();

                                // Find Total operating and NonOperating Mixed category
                                List<MixedSubDatas_FAnalysis> MixedOperatingList = new List<MixedSubDatas_FAnalysis>();
                                List<MixedSubDatas_FAnalysis> MixedNonOperatingList = new List<MixedSubDatas_FAnalysis>();

                                foreach (Datas dt in MixedList)
                                {
                                    var MixedSubdataList = iMixedSubDatas_FAnalysis.FindBy(x => x.DatasId == dt.DataId && x.InitialSetup_FAnalysisId == InitialSetupId && !x.LineItem.ToLower().Contains("depreciation") && !x.LineItem.ToLower().Contains("amortization")).ToList();
                                    if (MixedSubdataList != null && MixedSubdataList.Count > 0)
                                    {
                                        //MixedSubValues values = new MixedSubValues();
                                        foreach (var item in MixedSubdataList)
                                        {
                                            var MixedSubValuesList = iMixedSubValues_FAnalysis.FindBy(x => x.MixedSubDatas_FAnalysisId == item.Id).ToList();
                                            item.MixedSubValues_FAnalysis = new List<MixedSubValues_FAnalysis>();
                                            item.MixedSubValues_FAnalysis = MixedSubValuesList != null && MixedSubValuesList.Count > 0 ? MixedSubValuesList : new List<MixedSubValues_FAnalysis>();
                                            if (item.Category == "Operating")
                                                MixedOperatingList.Add(item);
                                            else
                                                MixedNonOperatingList.Add(item);
                                        }
                                    }
                                }

                                //// Add Mixed Operating in Integrated_datasList
                                Datas Moperating_datas = new Datas();
                                foreach (MixedSubDatas_FAnalysis dt in MixedOperatingList)
                                {
                                    Moperating_datas = new Datas();
                                    Moperating_datas.LineItem = dt.LineItem;
                                    Moperating_datas.IsTally = dt.IsTally;
                                    Moperating_datas.Sequence = (Integrated_datasList.Count + 1);
                                    Moperating_datas.Category = dt.Category;
                                    Moperating_datas.IsParentItem = false;
                                    Moperating_datas.Values = new List<Values>();
                                    foreach (MixedSubValues_FAnalysis val in dt.MixedSubValues_FAnalysis)
                                    {
                                        Values MoperatingValues = new Values();
                                        MoperatingValues.FilingDate = val.FilingDate;
                                        // MoperatingValues.Value = val.Value;
                                        MoperatingValues.Value = !string.IsNullOrEmpty(val.Value) ? Convert.ToString(mulBy * Convert.ToDouble(val.Value)) : null;
                                        Moperating_datas.Values.Add(MoperatingValues);
                                    }
                                    Integrated_datasList.Add(Moperating_datas);
                                }
                                ////////////
                                Datas EBITDA_datas = new Datas();
                                EBITDA_datas.LineItem = "EBITDA";
                                EBITDA_datas.IsTally = true;
                                EBITDA_datas.Sequence = (Integrated_datasList.Count + 1);
                                EBITDA_datas.Category = (string)null;
                                EBITDA_datas.IsParentItem = false;
                                EBITDA_datas.Values = new List<Values>();
                                foreach (Values obj in MaxValues_data.Values)
                                {
                                    double num2 = 0.0;
                                    Values values3 = new Values();
                                    foreach (Datas datas4 in Integrated_datasList.FindAll(x => x.Sequence >= gross_Seq).ToList())
                                    {
                                        Values values4 = datas4.Values.Find(x => x.FilingDate == obj.FilingDate);
                                        num2 = num2 + (values4 != null && !string.IsNullOrEmpty(values4.Value) ? Convert.ToDouble(values4.Value) : 0);
                                    }
                                    values3.FilingDate = obj.FilingDate;
                                    values3.Value = Convert.ToString(num2);
                                    EBITDA_datas.Values.Add(values3);
                                }
                                Integrated_datasList.Add(EBITDA_datas);

                                //bind depreciation
                                if (Depreciation_datas != null)
                                {
                                    if (Depreciation_datas.Values != null && Depreciation_datas.Values.Count > 0)
                                        foreach (var item in Depreciation_datas.Values)
                                        {
                                            item.Value = !string.IsNullOrEmpty(item.Value) ? Convert.ToString(mulBy * Convert.ToDouble(item.Value)) : null;
                                        }
                                    Integrated_datasList.Add(Depreciation_datas);

                                }

                                //bind EBITA
                                Datas EBITA_datas = new Datas();
                                EBITA_datas.LineItem = "EBITA";
                                EBITA_datas.IsTally = true;
                                EBITA_datas.Sequence = (Integrated_datasList.Count + 1);
                                EBITA_datas.Category = (string)null;
                                EBITA_datas.IsParentItem = false;
                                EBITA_datas.Values = new List<Values>();
                                foreach (Values obj in MaxValues_data.Values)
                                {
                                    Values values3 = new Values();
                                    Values values4 = EBITDA_datas.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    Values Depreciation_values = Depreciation_datas != null && Depreciation_datas.Values != null ? Depreciation_datas.Values.Find(x => x.FilingDate == obj.FilingDate) : null;
                                    double num2 = (values4 == null || values4.Value == null ? 0.0 : Convert.ToDouble(values4.Value)) + (Depreciation_values == null || Depreciation_values.Value == null ? 0.0 : Convert.ToDouble(Depreciation_values.Value));
                                    values3.FilingDate = obj.FilingDate;
                                    values3.Value = Convert.ToString(num2 != 0.0 ? Convert.ToString(num2) : (string)null);
                                    EBITA_datas.Values.Add(values3);
                                }
                                Integrated_datasList.Add(EBITA_datas);

                                //calculate amortization
                                if (Amortization_datas != null)
                                {
                                    if (Amortization_datas.Values != null && Amortization_datas.Values.Count > 0)
                                        foreach (var item in Amortization_datas.Values)
                                        {
                                            item.Value = !string.IsNullOrEmpty(item.Value) ? Convert.ToString(mulBy * Convert.ToDouble(item.Value)) : null;
                                        }
                                    Integrated_datasList.Add(Amortization_datas);
                                }

                                //calculate EBIT
                                Datas EBIT_datas = new Datas();
                                EBIT_datas.LineItem = "EBIT";
                                EBIT_datas.IsTally = true;
                                EBIT_datas.Sequence = (Integrated_datasList.Count + 1);
                                EBIT_datas.Category = (string)null;
                                EBIT_datas.IsParentItem = false;
                                EBIT_datas.Values = new List<Values>();
                                foreach (Values obj in MaxValues_data.Values)
                                {
                                    Values values3 = new Values();
                                    Values values4 = EBITA_datas.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    Values values5 = Amortization_datas?.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    double num2 = (values4 == null || values4.Value == null ? 0.0 : Convert.ToDouble(values4.Value)) + (values5 == null || values5.Value == null ? 0.0 : Convert.ToDouble(values5.Value));
                                    values3.FilingDate = obj.FilingDate;
                                    values3.Value = Convert.ToString(num2 != 0.0 ? Convert.ToString(num2) : (string)null);
                                    EBIT_datas.Values.Add(values3);
                                }
                                Integrated_datasList.Add(EBIT_datas);


                                // all the Non -operating item of Raw Historical and Mixed (Non-Operating Part)
                                List<Datas> NonOperatingItemList = filingsList1[0].Datas.FindAll(x => !x.LineItem.ToLower().Contains("depreciation") && !x.LineItem.ToLower().Contains("amortization") && ( x.Category == "Non-Operating" || x.Category == "Financing")).OrderBy(x => x.Sequence).ToList();
                                foreach (Datas datas4 in NonOperatingItemList)
                                {
                                    if (!datas4.LineItem.Contains("Amortization") && !datas4.LineItem.Contains("Extraordinary"))
                                    {

                                        if (!datas4.LineItem.ToLower().Contains("Gains") || !datas4.LineItem.ToLower().Contains("income"))
                                        {
                                            if (datas4.Values != null && datas4.Values.Count > 0)
                                                foreach (var item in datas4.Values)
                                                {
                                                    item.Value = !string.IsNullOrEmpty(item.Value) ? Convert.ToString(mulBy * Convert.ToDouble(item.Value)) : null;
                                                }
                                        }
                                        Integrated_datasList.Add(datas4);

                                    }
                                }


                                //// Add Mixed Non Operating in Integrated_datasList
                                Datas MNonoperating_datas = new Datas();
                                foreach (MixedSubDatas_FAnalysis dt in MixedNonOperatingList)
                                {
                                    MNonoperating_datas = new Datas();
                                    MNonoperating_datas.LineItem = dt.LineItem;
                                    MNonoperating_datas.IsTally = dt.IsTally;
                                    MNonoperating_datas.Sequence = (Integrated_datasList.Count + 1);
                                    MNonoperating_datas.Category = dt.Category;
                                    MNonoperating_datas.IsParentItem = false;
                                    MNonoperating_datas.Values = new List<Values>();
                                    foreach (MixedSubValues_FAnalysis val in dt.MixedSubValues_FAnalysis)
                                    {
                                        Values MNonoperatingValues = new Values();
                                        MNonoperatingValues.FilingDate = val.FilingDate;
                                        //  MNonoperatingValues.Value = val.Value;
                                        MNonoperatingValues.Value = !string.IsNullOrEmpty(val.Value) ? Convert.ToString(mulBy * Convert.ToDouble(val.Value)) : null;
                                        MNonoperating_datas.Values.Add(MNonoperatingValues);
                                    }
                                    Integrated_datasList.Add(MNonoperating_datas);
                                }
                                ////////////


                                Datas EBT_datas = new Datas();
                                EBT_datas.LineItem = "EBT";
                                EBT_datas.IsTally = true;
                                EBT_datas.Sequence = (Integrated_datasList.Count + 1);
                                EBT_datas.Category = (string)null;
                                EBT_datas.IsParentItem = false;
                                EBT_datas.Values = new List<Values>();
                                foreach (Values obj in MaxValues_data.Values)
                                {
                                    double num2 = 0.0;
                                    Values values3 = new Values();
                                    NonOperatingItemList.RemoveAll(x =>
                                    {
                                        if (!x.LineItem.Contains("Amortization"))
                                            return x.LineItem.Contains("Extraordinary");
                                        return true;
                                    });
                                    foreach (Datas datas4 in NonOperatingItemList)
                                    {
                                        Values values4 = datas4.Values.Find(x => x.FilingDate == obj.FilingDate);
                                        num2 += values4 == null || values4.Value == null ? 0.0 : Convert.ToDouble(values4.Value);
                                    }
                                    Values values5 = EBIT_datas.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    double num3 = num2 + (values5 == null || values5.Value == null ? 0.0 : Convert.ToDouble(values5.Value));
                                    values3.FilingDate = obj.FilingDate;
                                    values3.Value = Convert.ToString(Convert.ToString(num3));
                                    EBT_datas.Values.Add(values3);
                                }
                                Integrated_datasList.Add(EBT_datas);


                                //Provision of Taxes
                                Datas datas11 = new Datas();
                                Datas datas12 = filingsList1[0].Datas.Find(x => x.LineItem.Contains("Provision"));
                                if (datas12 != null)
                                {
                                    if (datas12.Values != null && datas12.Values.Count > 0)
                                        foreach (var item in datas12.Values)
                                        {
                                            item.Value = !string.IsNullOrEmpty(item.Value) ? Convert.ToString(mulBy * Convert.ToDouble(item.Value)) : null;
                                        }
                                }
                                Integrated_datasList.Add(datas12);

                                // NET INCOME before extraordinary items
                                Datas NetIncomeBefore_datas = new Datas();
                                NetIncomeBefore_datas.LineItem = "NET INCOME before extraordinary items";
                                NetIncomeBefore_datas.IsTally = true;
                                NetIncomeBefore_datas.Sequence = (Integrated_datasList.Count + 1);
                                NetIncomeBefore_datas.Category = (string)null;
                                NetIncomeBefore_datas.IsParentItem = false;
                                NetIncomeBefore_datas.Values = new List<Values>();
                                foreach (Values values2 in MaxValues_data.Values)
                                {
                                    Values obj = values2;
                                    Values values3 = new Values();
                                    Values values4 = EBT_datas.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    Values values5 = datas12?.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    double num2 = (values4 == null || values4.Value == null ? 0.0 : Convert.ToDouble(values4.Value)) + (values5 == null || values5.Value == null ? 0.0 : Convert.ToDouble(values5.Value));
                                    values3.FilingDate = obj.FilingDate;
                                    values3.Value = Convert.ToString(num2 != 0.0 ? Convert.ToString(num2) : (string)null);
                                    NetIncomeBefore_datas.Values.Add(values3);
                                }
                                Integrated_datasList.Add(NetIncomeBefore_datas);
                                Datas datas14 = new Datas();
                                Datas datas15 = filingsList1[0].Datas.Find(x => x.LineItem.Contains("Extraordinary"));
                                if (datas15 != null)
                                    Integrated_datasList.Add(datas15);
                                Datas NetIncomeAfter_datas = new Datas();
                                NetIncomeAfter_datas.LineItem = "NET INCOME after extraordinary items";
                                NetIncomeAfter_datas.IsTally = true;
                                NetIncomeAfter_datas.Sequence = (Integrated_datasList.Count + 1);
                                NetIncomeAfter_datas.Category = (string)null;
                                NetIncomeAfter_datas.IsParentItem = false;
                                NetIncomeAfter_datas.Values = new List<Values>();
                                foreach (Values values2 in MaxValues_data.Values)
                                {
                                    Values obj = values2;
                                    Values values3 = new Values();
                                    Values values4 = NetIncomeBefore_datas?.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    Values values5 = datas15?.Values.Find(x => x.FilingDate == obj.FilingDate);
                                    double num2 = (values4 == null || values4.Value == null ? 0.0 : Convert.ToDouble(values4.Value)) + (values5 == null || values5.Value == null ? 0.0 : Convert.ToDouble(values5.Value));
                                    values3.FilingDate = obj.FilingDate;
                                    values3.Value = Convert.ToString(num2 != 0.0 ? Convert.ToString(num2) : (string)null);
                                    NetIncomeAfter_datas.Values.Add(values3);
                                }
                                Integrated_datasList.Add(NetIncomeAfter_datas);
                                filingsList1[0].Datas = Integrated_datasList;

                                // Save Income Statement data in IntegratedDatas and Values
                                var IntegratedDatasIncomChk = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupId && x.StatementTypeId == 1).ToList();
                                if (IntegratedDatasIncomChk == null || IntegratedDatasIncomChk.Count == 0) ///////If data is not avlbl in IntegratedData save then only.
                                {
                                    foreach (var item in Integrated_datasList)
                                    {
                                        if (item != null && !string.IsNullOrEmpty(item.LineItem) && item.IsParentItem != true)
                                        {

                                            IntegratedDatasFAnalysis integratedDatasObj = new IntegratedDatasFAnalysis();
                                            integratedDatasObj.Id = 0;
                                            integratedDatasObj.Category = item.Category;
                                            integratedDatasObj.Sequence = item.Sequence;
                                            integratedDatasObj.LineItem = item.LineItem;
                                            integratedDatasObj.IsTally = item.IsTally;
                                            integratedDatasObj.IsParentItem = item.IsParentItem != null ? item.IsParentItem : false;
                                            integratedDatasObj.StatementTypeId = (int)StatementTypeEnum.IncomeStatement;
                                            integratedDatasObj.IntegratedValuesFAnalysis = new List<IntegratedValuesFAnalysis>();

                                            foreach (var valueobj in item.Values)
                                            {
                                                IntegratedValuesFAnalysis obj = new IntegratedValuesFAnalysis();
                                                obj.Id = 0;
                                                obj.FilingDate = valueobj.FilingDate;
                                                obj.Value = valueobj.Value;
                                                DateTime Dt = new DateTime();
                                                Dt = Convert.ToDateTime(valueobj.FilingDate);
                                                obj.Year = Convert.ToString(Dt.Year);
                                                integratedDatasObj.IntegratedValuesFAnalysis.Add(obj);
                                            }
                                            integratedDatasObj.InitialSetup_FAnalysisId = InitialSetupId == 0 ? null : InitialSetupId;
                                            iIntegratedDatasFAnalysis.Add(integratedDatasObj);
                                            iIntegratedDatasFAnalysis.Commit();
                                        }
                                    }
                                }
                            }

                            else if (filingsList1[0].StatementType == "BALANCE_SHEET")
                            {
                                Datas datas1 = new Datas();
                                List<Datas> DataProcessing_list = filingsList1[0].Datas.OrderBy(x => x.Sequence).ToList();
                                Datas MaxValues_data = DataProcessing_list.First<Datas>();
                                List<Datas> Integrated_datasList = new List<Datas>();
                                // Find Total Mixed category
                                List<Datas> MixedList = DataProcessing_list.FindAll(x => x.Category == "Mixed").ToList();
                                foreach (Datas dt in DataProcessing_list)
                                {
                                    if (dt.Category != "Mixed" || dt.IsParentItem == true)
                                    {
                                        Integrated_datasList.Add(dt);
                                    }
                                    else
                                    {
                                        foreach (Datas mixedDT in MixedList)
                                        {
                                            if (mixedDT.DataId == dt.DataId)
                                            {
                                                var MixedSubdataList = iMixedSubDatas_FAnalysis.FindBy(x => x.DatasId == dt.DataId && x.InitialSetup_FAnalysisId == InitialSetupId).ToList();
                                                if (MixedSubdataList != null && MixedSubdataList.Count > 0)
                                                {
                                                    foreach (var item in MixedSubdataList)
                                                    {
                                                        var MixedSubValuesList = iMixedSubValues_FAnalysis.FindBy(x => x.MixedSubDatas_FAnalysisId == item.Id).ToList();
                                                        item.MixedSubValues_FAnalysis = new List<MixedSubValues_FAnalysis>();
                                                        item.MixedSubValues_FAnalysis = MixedSubValuesList != null && MixedSubValuesList.Count > 0 ? MixedSubValuesList : new List<MixedSubValues_FAnalysis>();

                                                    }
                                                }

                                                Datas Balance_datas = new Datas();
                                                foreach (MixedSubDatas_FAnalysis msd in MixedSubdataList)
                                                {
                                                    Balance_datas = new Datas();
                                                    Balance_datas.LineItem = msd.LineItem;
                                                    Balance_datas.DataId = msd.DatasId;
                                                    Balance_datas.IsTally = msd.IsTally;
                                                    Balance_datas.Sequence = (Integrated_datasList.Count + 1);
                                                    Balance_datas.Category = msd.Category;
                                                    Balance_datas.IsParentItem = false;
                                                    Balance_datas.Values = new List<Values>();
                                                    foreach (MixedSubValues_FAnalysis val in msd.MixedSubValues_FAnalysis)
                                                    {
                                                        Values MNonoperatingValues = new Values();
                                                        MNonoperatingValues.FilingDate = val.FilingDate;
                                                        MNonoperatingValues.Value = val.Value;
                                                        Balance_datas.Values.Add(MNonoperatingValues);
                                                    }
                                                    Integrated_datasList.Add(Balance_datas);
                                                }
                                            }
                                        }
                                    }
                                }
                                filingsList1[0].Datas = Integrated_datasList;
                                var IntegratedDatasBalanceChk = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupId && x.StatementTypeId == 2).ToList();
                                // Save BalanceSheet data in IntegratedDatas and Values
                                if (IntegratedDatasBalanceChk == null || IntegratedDatasBalanceChk.Count == 0)
                                {
                                    foreach (var item in Integrated_datasList)
                                    {
                                        IntegratedDatasFAnalysis integratedDatasObj = new IntegratedDatasFAnalysis();
                                        integratedDatasObj.Id = 0;
                                        integratedDatasObj.Category = item.Category;
                                        integratedDatasObj.Sequence = item.Sequence;
                                        integratedDatasObj.LineItem = item.LineItem;
                                        integratedDatasObj.IsTally = item.IsTally;
                                        integratedDatasObj.IsParentItem = item.IsParentItem != null ? item.IsParentItem : false;
                                        integratedDatasObj.StatementTypeId = (int)StatementTypeEnum.BalanceSheet;
                                        integratedDatasObj.IntegratedValuesFAnalysis = new List<IntegratedValuesFAnalysis>();

                                        foreach (var valueobj in item.Values)
                                        {
                                            IntegratedValuesFAnalysis obj = new IntegratedValuesFAnalysis();
                                            obj.Id = 0;
                                            obj.FilingDate = valueobj.FilingDate;
                                            obj.Value = valueobj.Value;
                                            DateTime Dt = new DateTime();
                                            Dt = Convert.ToDateTime(valueobj.FilingDate);
                                            obj.Year = Convert.ToString(Dt.Year);
                                            integratedDatasObj.IntegratedValuesFAnalysis.Add(obj);
                                        }
                                        integratedDatasObj.InitialSetup_FAnalysisId = InitialSetupId == 0 ? null : InitialSetupId;
                                        iIntegratedDatasFAnalysis.Add(integratedDatasObj);
                                        iIntegratedDatasFAnalysis.Commit();
                                    }
                                }
                            }
                            else if (filingsList1[0].StatementType == "CASH_FLOW")
                            {
                                Datas datas1 = new Datas();
                                List<Datas> DataProcessing_list = filingsList1[0].Datas.OrderBy(x => x.Sequence).ToList();
                                Datas MaxValues_data = DataProcessing_list.First<Datas>();
                                List<Datas> Integrated_datasList = new List<Datas>();
                                // Find Total Mixed category
                                List<Datas> MixedList = DataProcessing_list.FindAll(x => x.Category == "Mixed").ToList();
                                foreach (Datas dt in DataProcessing_list)
                                {
                                    if (dt.Category != "Mixed" || dt.IsParentItem == true)
                                    {
                                        Integrated_datasList.Add(dt);
                                    }
                                    else
                                    {
                                        foreach (Datas mixedDT in MixedList)
                                        {
                                            if (mixedDT.DataId == dt.DataId)
                                            {
                                                var MixedSubdataList = iMixedSubDatas_FAnalysis.FindBy(x => x.DatasId == dt.DataId && x.InitialSetup_FAnalysisId == InitialSetupId).ToList();
                                                if (MixedSubdataList != null && MixedSubdataList.Count > 0)
                                                {
                                                    foreach (var item in MixedSubdataList)
                                                    {
                                                        var MixedSubValuesList = iMixedSubValues_FAnalysis.FindBy(x => x.MixedSubDatas_FAnalysisId == item.Id).ToList();
                                                        item.MixedSubValues_FAnalysis = new List<MixedSubValues_FAnalysis>();
                                                        item.MixedSubValues_FAnalysis = MixedSubValuesList != null && MixedSubValuesList.Count > 0 ? MixedSubValuesList : new List<MixedSubValues_FAnalysis>();

                                                    }
                                                }

                                                Datas Balance_datas = new Datas();
                                                foreach (MixedSubDatas_FAnalysis msd in MixedSubdataList)
                                                {
                                                    Balance_datas = new Datas();
                                                    Balance_datas.LineItem = msd.LineItem;
                                                    Balance_datas.DataId = msd.DatasId;
                                                    Balance_datas.IsTally = msd.IsTally;
                                                    Balance_datas.Sequence = (Integrated_datasList.Count + 1);
                                                    Balance_datas.Category = msd.Category;
                                                    Balance_datas.IsParentItem = false;
                                                    Balance_datas.Values = new List<Values>();
                                                    foreach (MixedSubValues_FAnalysis val in msd.MixedSubValues_FAnalysis)
                                                    {
                                                        Values MNonoperatingValues = new Values();
                                                        MNonoperatingValues.FilingDate = val.FilingDate;
                                                        MNonoperatingValues.Value = val.Value;
                                                        Balance_datas.Values.Add(MNonoperatingValues);
                                                    }
                                                    Integrated_datasList.Add(Balance_datas);
                                                }
                                            }
                                        }
                                    }
                                }
                                // filingsList1[0].Datas = Integrated_datasList;
                                var IntegratedDatasBalanceChk = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupId && x.StatementTypeId == 3).ToList();
                                // Save BalanceSheet data in IntegratedDatas and Values
                                if (IntegratedDatasBalanceChk == null || IntegratedDatasBalanceChk.Count == 0)
                                {
                                    foreach (var item in Integrated_datasList)
                                    {
                                        IntegratedDatasFAnalysis integratedDatasObj = new IntegratedDatasFAnalysis();
                                        integratedDatasObj.Id = 0;
                                        integratedDatasObj.Category = item.Category;
                                        integratedDatasObj.Sequence = item.Sequence;
                                        integratedDatasObj.LineItem = item.LineItem;
                                        integratedDatasObj.IsTally = item.IsTally;
                                        integratedDatasObj.IsParentItem = item.IsParentItem != null ? item.IsParentItem : false;
                                        integratedDatasObj.StatementTypeId = (int)StatementTypeEnum.CashFlowStatement;
                                        integratedDatasObj.IntegratedValuesFAnalysis = new List<IntegratedValuesFAnalysis>();

                                        foreach (var valueobj in item.Values)
                                        {
                                            IntegratedValuesFAnalysis obj = new IntegratedValuesFAnalysis();
                                            obj.Id = 0;
                                            obj.FilingDate = valueobj.FilingDate;
                                            obj.Value = valueobj.Value;
                                            DateTime Dt = new DateTime();
                                            Dt = Convert.ToDateTime(valueobj.FilingDate);
                                            obj.Year = Convert.ToString(Dt.Year);
                                            integratedDatasObj.IntegratedValuesFAnalysis.Add(obj);
                                        }
                                        integratedDatasObj.InitialSetup_FAnalysisId = InitialSetupId == 0 ? null : InitialSetupId;
                                        iIntegratedDatasFAnalysis.Add(integratedDatasObj);
                                        iIntegratedDatasFAnalysis.Commit();
                                    }
                                }
                            }
                        }
                        ///////////////////////Statement of Retained Earnings///////////////////////
                        int YearCount = Convert.ToInt32(endYear) - Convert.ToInt32(startYear) + 1;
                        int StYear = Convert.ToInt32(startYear);
                        Filings FObj = new Filings();
                        List<Datas> DatasRE = new List<Datas>();
                        FObj.CompanyName = CompanyName;
                        FObj.StatementType = "STATEMENT OF RETAINED EARNINGS";
                        FObj.InitialSetupId = InitialSetupId;
                        FObj.ReportName = "STATEMENT OF RETAINED EARNINGS";
                        FObj.Unit = "";

                        Datas DObj = new Datas();
                        DObj.LineItem = "RETAINED EARNINGS (BEGINNING YEAR)";
                        DObj.Sequence = 1;
                        DObj.Category = "";
                        DObj.IsParentItem = false;
                        DObj.IsTally = false;
                        DObj.Values = new List<Values>();
                        StYear = Convert.ToInt32(startYear);
                        for (int i = 0; i < YearCount; i++)
                        {
                            Values VObj = new Values();
                            VObj.Value = null;
                            VObj.FilingDate = "01-Jan-" + Convert.ToString(StYear);
                            DObj.Values.Add(VObj);
                            StYear = StYear + 1;
                        }
                        DatasRE.Add(DObj);

                        DObj = new Datas();
                        DObj.LineItem = "Net Income";
                        DObj.Sequence = 2;
                        DObj.Category = "";
                        DObj.IsParentItem = false;
                        DObj.IsTally = false;
                        DObj.Values = new List<Values>();
                        StYear = Convert.ToInt32(startYear);
                        for (int i = 0; i < YearCount; i++)
                        {
                            Values VObj = new Values();
                            VObj.Value = null;
                            VObj.FilingDate = "01-Jan-" + Convert.ToString(StYear);
                            DObj.Values.Add(VObj);
                            StYear = StYear + 1;
                        }
                        DatasRE.Add(DObj);

                        DObj = new Datas();
                        DObj.LineItem = "Dividends Paid";
                        DObj.Sequence = 3;
                        DObj.Category = "";
                        DObj.IsParentItem = false;
                        DObj.IsTally = false;
                        DObj.Values = new List<Values>();
                        StYear = Convert.ToInt32(startYear);
                        for (int i = 0; i < YearCount; i++)
                        {
                            Values VObj = new Values();
                            VObj.Value = null;
                            VObj.FilingDate = "01-Jan-" + Convert.ToString(StYear);
                            DObj.Values.Add(VObj);
                            StYear = StYear + 1;
                        }
                        DatasRE.Add(DObj);

                        DObj = new Datas();
                        DObj.LineItem = "RETAINED EARNINGS (END YEAR)";
                        DObj.Sequence = 4;
                        DObj.Category = "";
                        DObj.IsParentItem = false;
                        DObj.IsTally = true;
                        DObj.Values = new List<Values>();
                        StYear = Convert.ToInt32(startYear);
                        for (int i = 0; i < YearCount; i++)
                        {
                            Values VObj = new Values();
                            VObj.Value = null;
                            VObj.FilingDate = "01-Jan-" + Convert.ToString(StYear);
                            DObj.Values.Add(VObj);
                            StYear = StYear + 1;
                        }
                        DatasRE.Add(DObj);

                        FObj.Datas = DatasRE;
                        List<Filings> FilingsRE = new List<Filings>();
                        FilingsRE.Add(FObj);


                        FilingsArray filingsArrayRE = new FilingsArray()
                        {
                            CompanyName = FilingsRE[0].CompanyName,
                            StatementType = "STATEMENT OF RETAINED EARNINGS",
                            Filings = FilingsRE[0]
                        };
                        filingsArrayList.Add(filingsArrayRE);
                        ////////////////////////////////////////////////////////////////////////////


                        /////////////////////////SAVE Retained Earnings/////////////////////////////
                        var IntegratedDatasREChk = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupId && x.StatementTypeId == (int)StatementTypeEnum.RetainedEarningsStatement).ToList();
                        if (IntegratedDatasREChk == null || IntegratedDatasREChk.Count == 0)
                        {
                            foreach (var item in DatasRE)
                            {
                                IntegratedDatasFAnalysis integratedDatasObj = new IntegratedDatasFAnalysis();
                                integratedDatasObj.Id = 0;
                                integratedDatasObj.Category = item.Category;
                                integratedDatasObj.Sequence = item.Sequence;
                                integratedDatasObj.LineItem = item.LineItem;
                                integratedDatasObj.IsTally = item.IsTally;
                                integratedDatasObj.IsParentItem = item.IsParentItem != null ? item.IsParentItem : false;
                                integratedDatasObj.StatementTypeId = (int)StatementTypeEnum.RetainedEarningsStatement;
                                integratedDatasObj.IntegratedValuesFAnalysis = new List<IntegratedValuesFAnalysis>();
                                foreach (var valueobj in item.Values)
                                {
                                    IntegratedValuesFAnalysis obj = new IntegratedValuesFAnalysis();
                                    obj.Id = 0;
                                    obj.FilingDate = valueobj.FilingDate;
                                    obj.Value = valueobj.Value;
                                    DateTime Dt = new DateTime();
                                    Dt = Convert.ToDateTime(valueobj.FilingDate);
                                    obj.Year = Convert.ToString(Dt.Year);
                                    integratedDatasObj.IntegratedValuesFAnalysis.Add(obj);
                                }
                                //integratedDatasObj.InitialSetupId = item.InitialSetupId != 0 ? Convert.ToInt64(item.InitialSetupId) : 0;
                                integratedDatasObj.InitialSetup_FAnalysisId = InitialSetupId == 0 ? null : InitialSetupId;
                                iIntegratedDatasFAnalysis.Add(integratedDatasObj);
                                iIntegratedDatasFAnalysis.Commit();
                            }
                        }

                        ////////////////////////////////////////////////////////////////////////////
                        renderResult.StatusCode = 1;
                        renderResult.Result = filingsArrayList;
                    }
                }
                return Ok(renderResult);

            }
            catch (Exception ex)
            {
                return BadRequest(Convert.ToString(ex.Message));
            }
        }



        #endregion

        #region Financial Analysis


        [HttpGet]
        [Route("GetFinancialAnalysis/{UserId}/{cik}/{startYear?}/{endYear?}")]
        public ActionResult GetFinancialAnalysis(long UserId, string cik, int? startYear = null, int? endYear = null)
        {
            //For noe using Forcast Ratio Models to just get the data on web page
            FinancialStatementAnalysisResult renderResult = new FinancialStatementAnalysisResult();
            List<FinancialStatementAnalysisFilingsViewModel> AnalysisFilingsList = new List<FinancialStatementAnalysisFilingsViewModel>();
            List<FinancialStatementAnalysisDatasViewModel> AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
            FinancialStatementAnalysisFilingsViewModel AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
            FinancialStatementAnalysisDatasViewModel forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
            List<FinancialStatementAnalysisValuesViewModel> forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
            FinancialStatementAnalysisValuesViewModel forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
            try
            {
                //get active InitialSetup
                InitialSetup_FAnalysis initialSetupObj = iInitialSetup_FAnalysis.GetSingle(x => x.UserId == UserId && x.IsActive == true);
                if (initialSetupObj != null)
                {
                    long? initialSetupId = Convert.ToInt64(initialSetupObj.Id);
                    List<IntegratedDatasFAnalysis> integratedDatasList = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == initialSetupObj.Id).ToList();
                    if (integratedDatasList != null && integratedDatasList.Count > 0)
                    {
                        List<IntegratedValuesFAnalysis> integratedValuesList = iIntegratedValuesFAnalysis.FindBy(x => integratedDatasList.Any(m => m.Id == x.IntegratedDatasFAnalysisId)).ToList();
                        IntegratedDatasFAnalysis revenueIntegratedObj = new IntegratedDatasFAnalysis();
                        IntegratedDatasFAnalysis costofSalesIntegratedObj = new IntegratedDatasFAnalysis();
                        List<FinancialStatementAnalysisValuesViewModel> dumyforcastRatioValueList = new List<FinancialStatementAnalysisValuesViewModel>();
                        int year = Convert.ToInt32(initialSetupObj.YearFrom);
                        for (int i = year; i <= initialSetupObj.YearTo; i++)
                        {
                            FinancialStatementAnalysisValuesViewModel dumyforcastRatioValue = new FinancialStatementAnalysisValuesViewModel();
                            dumyforcastRatioValue.Year = Convert.ToString(year);
                            dumyforcastRatioValue.Value = "";
                            dumyforcastRatioValueList.Add(dumyforcastRatioValue);
                            year = year + 1;
                        }
                        // find Revenue
                        string revenuesynonyms = "Net Sales%Net Revenue%Revenue%Total Revenues%Sales%Total Net Revenue%Total revenue%Total net sales%Sales to customers%Total net revenues%Total revenues (Note 4)%Revenue from Contract with Customer, Excluding Assessed Tax%Revenues%Net revenues%Revenue, net";
                        List<string> synonyms = revenuesynonyms.Split('%').ToList(); // convert comma seperated values to list
                        bool revenueflag = false;
                        foreach (IntegratedDatasFAnalysis integrateddatasObj in integratedDatasList)
                        {
                            if (integrateddatasObj.IsParentItem != true)
                                foreach (var syn in synonyms)
                                {
                                    if (integrateddatasObj.LineItem.ToUpper() == syn.ToUpper())
                                    {
                                        revenueIntegratedObj = integrateddatasObj;
                                        revenueflag = true;
                                        break;
                                    }
                                }
                            if (revenueflag == true)
                                break;

                        }
                        string Costofsalesynonyms = "Cost of sales%COGS%Cost of Goods Sold%Cost of Revenue%Cost of Products Sold%Total cost of revenue%Total cost of revenues%Cost of products sold, excluding amortization of intangible assets%Costs of goods sold%Cost of equipment and services revenues%Cost of revenues%Cost of revenue (COR)";
                        //Costofsalesynonyms = Costofsalesynonyms.Replace("TT", "");
                        List<string> CostofsalesSynonyms = Costofsalesynonyms.Split('%').ToList(); // convert comma seperated values to list
                        bool costofsalesflag = false;
                        foreach (IntegratedDatasFAnalysis integrateddatasObj in integratedDatasList)
                        {
                            if (integrateddatasObj.IsParentItem != true)
                                foreach (var syn in CostofsalesSynonyms)
                                {
                                    if (integrateddatasObj.LineItem.ToUpper() == syn.ToUpper())
                                    {
                                        costofSalesIntegratedObj = integrateddatasObj;
                                        costofsalesflag = true;
                                        break;
                                    }
                                }
                            if (costofsalesflag == true)
                                break;
                        }
                        //get Revenue Values
                        List<IntegratedValuesFAnalysis> revenuevaluesList = new List<IntegratedValuesFAnalysis>();
                        revenuevaluesList = integratedValuesList.FindAll(x => x.IntegratedDatasFAnalysisId == revenueIntegratedObj.Id).ToList();
                        //get Cost of Goods Sold Values
                        List<IntegratedValuesFAnalysis> costofrevenuevaluesList = new List<IntegratedValuesFAnalysis>();
                        costofrevenuevaluesList = integratedValuesList.FindAll(x => x.IntegratedDatasFAnalysisId == costofSalesIntegratedObj.Id).ToList();

                        //for PROFITABILITY RATIOS (Income statement)
                        #region PROFITABILITY RATIOS (Income statement)
                        AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
                        AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        AnalysisFiling.CompanyName = initialSetupObj.Company;
                        AnalysisFiling.ReportName = initialSetupObj.Company;
                        AnalysisFiling.StatementType = "PROFITABILITY RATIOS (Income statement)";
                        AnalysisFiling.Unit = "";
                        AnalysisFiling.CIK = initialSetupObj.CIKNumber;
                        //Gross Margin
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Gross Margin";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find Gross Margin
                        IntegratedDatasFAnalysis MatchedIntegratedItem = new IntegratedDatasFAnalysis();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem.ToLower().Contains("gross"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(RevenueValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //EBITDA Margin
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "EBITDA";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find EBITDA Margin
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBITDA");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(RevenueValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //EBITA Margin
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "EBITA";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find EBITA Margin
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBITA");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(RevenueValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //EBIT Margin
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "EBIT";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find EBIT Margin
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBIT");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(RevenueValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);
                        //NOPLAT Margin
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "NOPLAT Margin";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find NOPLAT Margin
                        
                        MarketDatas NOPLAT_Datas = iMarketDatas.GetSingle(x => x.LineItem == "NOPLAT" && x.InitialSetup_FAnalysisId == initialSetupId);
                        List<MarketValues> NOPLAT_ValuesList = NOPLAT_Datas != null ? iMarketValues.FindBy(x => x.MarketDatasId == NOPLAT_Datas.Id).ToList() : null;

                        //Use calcualte formula
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = NOPLAT_ValuesList != null && NOPLAT_ValuesList.Count > 0 ? NOPLAT_ValuesList.Find(x => x.Year == obj.Year) : null;
                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(RevenueValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }


                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);
                        //Net Profit Margin
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Net Profit Margin";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find Net Profit Margin
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "NET INCOME before extraordinary items");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(RevenueValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        AnalysisFiling.FinancialStatementAnalysisDatasVM = AnalysisDatasList;
                        AnalysisFilingsList.Add(AnalysisFiling);
                        #endregion

                        #region LIQUDITY RATIOS (Balance sheet)
                        AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
                        AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        AnalysisFiling.CompanyName = initialSetupObj.Company;
                        AnalysisFiling.ReportName = initialSetupObj.Company;
                        AnalysisFiling.StatementType = "LIQUDITY RATIOS (Balance sheet)";
                        AnalysisFiling.Unit = "";
                        AnalysisFiling.CIK = initialSetupObj.CIKNumber;
                        //Current Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Current Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find current assets sequence
                        IntegratedDatasFAnalysis currentAssetsIntegratedDatas = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToUpper() == "TOTAL CURRENT ASSETS" || x.LineItem.ToUpper() == "ASSETS,CURRENT"));

                        List<IntegratedDatasFAnalysis> itemstillCurrentAssets = currentAssetsIntegratedDatas != null && currentAssetsIntegratedDatas.Sequence > 0 ? integratedDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && x.Category == "Operating" && x.Sequence < currentAssetsIntegratedDatas.Sequence).ToList() : null;

                        //find Current liabilities Total current liabilities%Liabilities, Current
                        IntegratedDatasFAnalysis currentLiabilitiesIntegratedDatas = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToUpper() == "TOTAL CURRENT LIABILITIES" || x.LineItem.ToUpper() == "LIABILITIES, CURRENT"));

                        List<IntegratedDatasFAnalysis> itemstillCurrentLiabilities = currentLiabilitiesIntegratedDatas != null && currentAssetsIntegratedDatas != null && currentAssetsIntegratedDatas.Sequence > 0 && currentLiabilitiesIntegratedDatas.Sequence > 0 ? integratedDatasList.FindAll(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && x.Category == "Operating" && (x.Sequence > currentAssetsIntegratedDatas.Sequence && x.Sequence < currentLiabilitiesIntegratedDatas.Sequence)).ToList() : null;

                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            double currentAssets = 0;
                            double currentLiabilities = 0;
                            //Current Assets (Oparting ine items)/ Current Liabilities(Oparting line items)
                            if (itemstillCurrentAssets != null && itemstillCurrentAssets.Count > 0 && integratedValuesList != null && integratedValuesList.Count > 0)
                            {
                                foreach (IntegratedDatasFAnalysis integratedObj in itemstillCurrentAssets)
                                {
                                    IntegratedValuesFAnalysis tempValue = integratedObj != null ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedObj.Id && x.Year == obj.Year) : null;
                                    currentAssets = currentAssets + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                                }
                                if (itemstillCurrentLiabilities != null && itemstillCurrentLiabilities.Count > 0)
                                    foreach (IntegratedDatasFAnalysis integratedObj in itemstillCurrentLiabilities)
                                    {
                                        IntegratedValuesFAnalysis tempValue = integratedObj != null ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedObj.Id && x.Year == obj.Year) : null;
                                        currentLiabilities = currentLiabilities + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                                    }
                                value = currentLiabilities != 0 ? currentAssets / currentLiabilities : 0;
                                value = value * 100;
                            }


                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);
                        // Quick ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Quick ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find Receivables, Trade Accounts Receivable, Trade Receivables, Accounts Receivable-Trade
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "net receivables" || x.LineItem.ToLower() == "receivables" || x.LineItem.ToLower() == "trade accounts receivable" || x.LineItem.ToLower() == "trade receivables" || x.LineItem.ToLower() == "accounts receivable-trade"));

                        IntegratedDatasFAnalysis cashEquivalent = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && x.LineItem.ToLower().Contains("cash and cash equivalents"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            double currentLiabilities = 0;
                            //Cash and Cash Equivalents +   Net Receivables / Current Liabilities(Oparting ine items)
                            if (itemstillCurrentLiabilities != null && itemstillCurrentLiabilities.Count > 0)
                                foreach (IntegratedDatasFAnalysis integratedObj in itemstillCurrentLiabilities)
                                {
                                    IntegratedValuesFAnalysis tempcurrentValue = integratedObj != null ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedObj.Id && x.Year == obj.Year) : null;
                                    currentLiabilities = currentLiabilities + (tempcurrentValue != null && !string.IsNullOrEmpty(tempcurrentValue.Value) ? Convert.ToDouble(tempcurrentValue.Value) : 0);
                                }

                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            var CashValue = cashEquivalent != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == cashEquivalent.Id && x.Year == obj.Year) : null;

                            value = currentLiabilities != 0 ? (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) + (CashValue != null && !string.IsNullOrEmpty(CashValue.Value) ? Convert.ToDouble(CashValue.Value) : 0) / currentLiabilities : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);
                        // Cash Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Cash Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            double currentLiabilities = 0;
                            //Cash and Cash Equivalents/ Current Liabilities(Oparting ine items))
                            if (itemstillCurrentLiabilities != null && itemstillCurrentLiabilities.Count > 0)
                                foreach (IntegratedDatasFAnalysis integratedObj in itemstillCurrentLiabilities)
                                {
                                    IntegratedValuesFAnalysis tempcurrentValue = integratedObj != null ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedObj.Id && x.Year == obj.Year) : null;
                                    currentLiabilities = currentLiabilities + (tempcurrentValue != null && !string.IsNullOrEmpty(tempcurrentValue.Value) ? Convert.ToDouble(tempcurrentValue.Value) : 0);
                                }


                            var CashValue = cashEquivalent != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == cashEquivalent.Id && x.Year == obj.Year) : null;

                            value = currentLiabilities != 0 ? (CashValue != null && !string.IsNullOrEmpty(CashValue.Value) ? Convert.ToDouble(CashValue.Value) : 0) / currentLiabilities : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        AnalysisFiling.FinancialStatementAnalysisDatasVM = AnalysisDatasList;
                        AnalysisFilingsList.Add(AnalysisFiling);
                        #endregion

                        #region WORKING CAPITAL RATIOS (Income statement & Balance sheet)

                        AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
                        AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();

                        AnalysisFiling.CompanyName = initialSetupObj.Company;
                        AnalysisFiling.ReportName = initialSetupObj.Company;
                        AnalysisFiling.StatementType = "WORKING CAPITAL RATIOS (Income statement & Balance sheet)";
                        AnalysisFiling.Unit = "";
                        AnalysisFiling.CIK = initialSetupObj.CIKNumber;
                        //Accounts Recievable Days
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Accounts Recievable Days";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find Receivables, Trade Accounts Receivable, Trade Receivables, Accounts Receivable-Trade
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "net receivables" || x.LineItem.ToLower() == "receivables" || x.LineItem.ToLower() == "trade accounts receivable" || x.LineItem.ToLower() == "trade receivables" || x.LineItem.ToLower() == "accounts receivable-trade"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / (Convert.ToDouble(RevenueValue.Value) / 365)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);
                        // Accounts Payable Days
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Accounts Payable Days";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find Payables, Trade Payables, Trade Accounts Payable
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "payables" || x.LineItem.ToLower() == "trade payables" || x.LineItem.ToLower() == "trade accounts payable" || x.LineItem.ToUpper() == "ACCOUNTS PAYABLE"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var costofRevenueValue = costofrevenuevaluesList != null && costofrevenuevaluesList.Count > 0 ? costofrevenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = costofRevenueValue != null && !string.IsNullOrEmpty(costofRevenueValue.Value) && costofRevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / (Convert.ToDouble(costofRevenueValue.Value) / 365)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        // Inventory Days
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Inventories";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find Merchandise Inventories, Inventory
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "inventories" || x.LineItem.ToLower() == "merchandise inventories" || x.LineItem.ToLower() == "inventory"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var costofRevenueValue = costofrevenuevaluesList != null && costofrevenuevaluesList.Count > 0 ? costofrevenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = costofRevenueValue != null && !string.IsNullOrEmpty(costofRevenueValue.Value) && costofRevenueValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / (Convert.ToDouble(costofRevenueValue.Value) / 365)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        // Accounts Receivable Turnover
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Accounts Receivable Turnover";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find Receivables, Trade Accounts Receivable, Trade Receivables, Accounts Receivable-Trade
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "net receivables" || x.LineItem.ToLower() == "receivables" || x.LineItem.ToLower() == "trade accounts receivable" || x.LineItem.ToLower() == "trade receivables" || x.LineItem.ToLower() == "accounts receivable-trade"));

                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {


                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // NET SALES/Net Receivables
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            value = tempValue != null && !string.IsNullOrEmpty(tempValue.Value) && tempValue.Value != "0" ? ((RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) ? Convert.ToDouble(RevenueValue.Value) : 0) / Convert.ToDouble(tempValue.Value)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);
                        // Accounts Payable Turnover
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Accounts Payable Turnover";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find Payables, Trade Payables, Trade Accounts Payable
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "payables" || x.LineItem.ToLower() == "trade payables" || x.LineItem.ToLower() == "trade accounts payable" || x.LineItem.ToLower() == "accounts payable"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // Cost of Goods Sold/Accounts Payable
                            var costofRevenueValue = costofrevenuevaluesList != null && costofrevenuevaluesList.Count > 0 ? costofrevenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            value = tempValue != null && !string.IsNullOrEmpty(tempValue.Value) && tempValue.Value != "0" ? ((costofRevenueValue != null && !string.IsNullOrEmpty(costofRevenueValue.Value) ? Convert.ToDouble(costofRevenueValue.Value) : 0) / Convert.ToDouble(tempValue.Value)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        // Inventory Turnover
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Inventory Turnover";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //find Merchandise Inventories, Inventory
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "inventories" || x.LineItem.ToLower() == "merchandise inventories" || x.LineItem.ToLower() == "inventory"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // Cost of Goods Sold/Accounts Payable
                            var costofRevenueValue = costofrevenuevaluesList != null && costofrevenuevaluesList.Count > 0 ? costofrevenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            value = tempValue != null && !string.IsNullOrEmpty(tempValue.Value) && tempValue.Value != "0" ? ((costofRevenueValue != null && !string.IsNullOrEmpty(costofRevenueValue.Value) ? Convert.ToDouble(costofRevenueValue.Value) : 0) / Convert.ToDouble(tempValue.Value)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        AnalysisFiling.FinancialStatementAnalysisDatasVM = AnalysisDatasList;
                        AnalysisFilingsList.Add(AnalysisFiling);
                        #endregion

                        #region INTEREST COVERAGE RATIOS (Income statement)

                        AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
                        AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();

                        AnalysisFiling.CompanyName = initialSetupObj.Company;
                        AnalysisFiling.ReportName = initialSetupObj.Company;
                        AnalysisFiling.StatementType = "INTEREST COVERAGE RATIOS (Income statement)";
                        AnalysisFiling.Unit = "";
                        AnalysisFiling.CIK = initialSetupObj.CIKNumber;

                        //Interest Expense%Interest and Debt Expense% Interest and Other Expense%Interest expense (income), net%Interest expense, net%Interest (income) expense, net%Interest and other expense, net%Interest and other income (expense), net%Interest expense (income), net (Notes 6, 7 and 8)%Interest expense (income), net (Notes 6, 7, and 8)%Interest (income) expense, net (Notes 6, 7 and 8)%Net interest expense
                        IntegratedDatasFAnalysis integratedDatasInterestExpense = new IntegratedDatasFAnalysis();
                        integratedDatasInterestExpense = integratedDatasList.Find(x => x.LineItem.ToLower() == "interest expense");
                        if(integratedDatasInterestExpense==null)
                        {
                            integratedDatasInterestExpense = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && (x.LineItem.ToLower() == "interest expense" || x.LineItem.ToLower() == "interest and debt expense" || x.LineItem.ToLower() == "interest and other expense" || x.LineItem.ToLower() == "interest expense (income), net" || x.LineItem.ToLower() == "interest expense, net" || x.LineItem.ToLower() == "interest (income) expense, net" || x.LineItem.ToLower() == "Interest and other income (expense), net" || x.LineItem.ToLower() == "interest expense (income), net (Notes 6, 7 and 8)" || x.LineItem.ToLower() == "interest (income) expense, net (Notes 6, 7 and 8)" || x.LineItem.ToLower() == "Net interest expense") && (x.Category=="Financing" || x.Category=="Non-Operating" ));
                        }
                        //EBIT / Interest Coverage
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "EBIT / Interest Coverage";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find EBIT
                        MatchedIntegratedItem = new IntegratedDatasFAnalysis();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBIT");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var EXpenseValue = integratedDatasInterestExpense != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasInterestExpense.Id && x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = EXpenseValue != null && !string.IsNullOrEmpty(EXpenseValue.Value) && EXpenseValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(EXpenseValue.Value)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);
                        //EBITDA / Interest Coverage
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "EBITDA / Interest Coverage";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //find EBITDA
                        MatchedIntegratedItem = new IntegratedDatasFAnalysis();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBITDA");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            var EXpenseValue = integratedDatasInterestExpense != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasInterestExpense.Id && x.Year == obj.Year) : null;
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            value = EXpenseValue != null && !string.IsNullOrEmpty(EXpenseValue.Value) && EXpenseValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(EXpenseValue.Value)) : 0;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        AnalysisFiling.FinancialStatementAnalysisDatasVM = AnalysisDatasList;
                        AnalysisFilingsList.Add(AnalysisFiling);
                        #endregion

                        #region LEVERAGE RATIOS (Balance sheet, Cash flow statement & Market data)

                        AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
                        AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();

                        AnalysisFiling.CompanyName = initialSetupObj.Company;
                        AnalysisFiling.ReportName = initialSetupObj.Company;
                        AnalysisFiling.StatementType = "LEVERAGE RATIOS (Balance sheet, Cash flow statement & Market data)";
                        AnalysisFiling.Unit = "";
                        AnalysisFiling.CIK = initialSetupObj.CIKNumber;

                        //find Total stockholders' equity%Total stockholders’ equity%Total Equity%Total shareholders’ equity%Total shareholders equity%Total Shareholders' Investment%Total Stockholders’ Investment%Total shareholders’ equity%Total shareholders' equity
                        IntegratedDatasFAnalysis integratedDatasStockholdersequity = new IntegratedDatasFAnalysis();
                        integratedDatasStockholdersequity = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "total stockholders' equity" || x.LineItem.ToLower() == "total stockholders’ equity" || x.LineItem.ToLower() == "total equity" || x.LineItem.ToLower() == "total shareholders equity" || x.LineItem.ToLower() == "total shareholders' investment" || x.LineItem.ToLower() == "total stockholders’ investment" || x.LineItem.ToLower() == "total shareholders’ equity" || x.LineItem.ToLower() == "total shareholders' equity" || x.LineItem.ToLower() == "total shareholders’ equity"));

                        //find Short-Term Debt,Term Debt, Covertible Short-Term Debt, Current Debt, Short-Term Borrowings
                        IntegratedDatasFAnalysis integratedDatashorttermDebt = new IntegratedDatasFAnalysis();
                        integratedDatashorttermDebt = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "short-term debt" || x.LineItem.ToLower() == "term debt" || x.LineItem.ToLower() == "covertible short-term debt" || x.LineItem.ToLower() == "current debt" || x.LineItem.ToLower() == "short-term borrowings"));


                        List<IntegratedDatasFAnalysis> debtList = integratedDatasList.FindAll(x=>x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && x.LineItem.ToLower().Contains("debt")).ToList();

                        // find long term Debt
                        // Long - term debt % Long - term liabilities%Long-Term Debt,Term Debt, Debt, Non-Current Debt
                        IntegratedDatasFAnalysis integratedDatalongtermdebt = new IntegratedDatasFAnalysis();
                        integratedDatalongtermdebt = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToLower() == "long - term debt" || x.LineItem.ToLower() == "long - term liabilities" || x.LineItem.ToLower() == "long-term debt" || x.LineItem.ToLower() == "term debt" || x.LineItem.ToLower() == "non-current debt"));

                        //Debt-to-Equity Ratio (Book) 
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Debt-to-Equity Ratio (Book)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //(Short - Term Debt + Long - Term Debt)/ TOTAL STOCKHOLDERS' EQUITY

                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            // long term debt
                            // var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;
                            //stockholders euity
                            double DebtValue = 0;
                            foreach(IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempObj != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }
                           
                            var StockholderequityValue = integratedDatasStockholdersequity != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasStockholdersequity.Id && x.Year == obj.Year) : null;

                            value = StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) && StockholderequityValue.Value != "0" ? DebtValue / Convert.ToDouble(StockholderequityValue.Value) : 0;

                            //value = StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) && StockholderequityValue.Value != "0" ? (((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0)) / Convert.ToDouble(StockholderequityValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);



                        //Debt-to-Equity Ratio (Market)
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Debt-to-Equity Ratio (Market)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //get Markt Datas
                        List<MarketDatas> marketDatasList = iMarketDatas.FindBy(x => x.InitialSetup_FAnalysisId == initialSetupId).ToList();
                        //get Market Values
                        List<MarketValues> marketValuesList = marketDatasList != null && marketDatasList.Count > 0 ? iMarketValues.FindBy(x => marketDatasList.Any(t => t.Id == x.MarketDatasId)).ToList() : null;

                        //Market Value Common Stock ($M)
                        MarketDatas market_CommonStock = marketDatasList != null && marketDatasList.Count > 0 ? marketDatasList.Find(x => x.LineItem == "Market Value Common Stock ($M)") : null;

                        //(Short-Term Debt+ Long-Term Debt)/Market Value Common Stock ($M)
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //  totalDebt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }

                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;
                            //Market Value
                            var marketValue = market_CommonStock != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == market_CommonStock.Id && x.Year == obj.Year) : null;

                            //value = marketValue != null && !string.IsNullOrEmpty(marketValue.Value) && marketValue.Value != "0" ? (((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0)) / Convert.ToDouble(marketValue.Value)) : 0;

                            value = marketValue != null && !string.IsNullOrEmpty(marketValue.Value) && marketValue.Value != "0" ? (DebtValue / Convert.ToDouble(marketValue.Value)) : 0;


                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);


                        //Debt-to-Capital Ratio (Book)
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Debt-to-Capital Ratio (Book)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        // (Short-Term Debt+ Long-Term Debt)/ (Short-Term Debt+ Long-Term Debt+TOTAL STOCKHOLDERS' EQUITY)
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;

                            //total Debt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }


                            //stockholders euity
                            var StockholderequityValue = integratedDatasStockholdersequity != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasStockholdersequity.Id && x.Year == obj.Year) : null;

                            //double totalValue = (ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0) + (StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) ? Convert.ToDouble(StockholderequityValue.Value) : 0);

                            //value = totalValue != 0 ? (((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0)) / Convert.ToDouble(totalValue)) : 0;


                            double totalValue = DebtValue + (StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) ? Convert.ToDouble(StockholderequityValue.Value) : 0);

                            value = totalValue != 0 ? (DebtValue / Convert.ToDouble(totalValue)) : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);


                        //Debt-to-Capital Ratio (Market)
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Debt-to-Capital Ratio (Market)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        // (Short-Term Debt+ Long-Term Debt)/ (Short-Term Debt+ Long-Term Debt+Market Value Common Stock ($M))
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;

                            //total Debt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }


                            //marketValue
                            var marketValue = market_CommonStock != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == market_CommonStock.Id && x.Year == obj.Year) : null;

                            //double totalValue = (ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0) + (marketValue != null && !string.IsNullOrEmpty(marketValue.Value) ? Convert.ToDouble(marketValue.Value) : 0);

                            //value = totalValue != 0 ? (((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0)) / Convert.ToDouble(totalValue)) : 0;

                            double totalValue = DebtValue + (marketValue != null && !string.IsNullOrEmpty(marketValue.Value) ? Convert.ToDouble(marketValue.Value) : 0);

                            value = totalValue != 0 ? (DebtValue / Convert.ToDouble(totalValue)) : 0;


                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Debt-to-Enterprise Value Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Debt-to-Enterprise Value Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //Enterprise Value ($M)
                        MarketDatas EnterpriceValue_Datas = marketDatasList != null && marketDatasList.Count > 0 ? marketDatasList.Find(x => x.LineItem == "Enterprise Value ($M)") : null;

                        //(Short-Term Debt+ Long-Term Debt)/Market Value Common Stock ($M)
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;

                            //Total Debt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }


                            //Market Value
                            var EnterpriceValue = EnterpriceValue_Datas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == EnterpriceValue_Datas.Id && x.Year == obj.Year) : null;

                            //value = EnterpriceValue != null && !string.IsNullOrEmpty(EnterpriceValue.Value) && EnterpriceValue.Value != "0" ? (((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0)) / Convert.ToDouble(EnterpriceValue.Value)) : 0;
                            value = EnterpriceValue != null && !string.IsNullOrEmpty(EnterpriceValue.Value) && EnterpriceValue.Value != "0" ? (DebtValue / Convert.ToDouble(EnterpriceValue.Value)) : 0;


                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);


                        //Cashflow from Operations-to-Debt Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Cashflow from Operations-to-Debt Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //Net cash provided by operating activities
                        IntegratedDatasFAnalysis cashflow_Activities = integratedDatasList.Find(x => x.LineItem.ToLower() == "net cash provided by operating activities" && x.StatementTypeId == (int)StatementTypeEnum.CashFlowStatement);

                        //Net cash provided by operating activities/(Short-Term Debt+ Long-Term Debt)
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;

                            //Total Debt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }


                            //Market Value
                            var cashflow_ActivitiesValue = EnterpriceValue_Datas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == EnterpriceValue_Datas.Id && x.Year == obj.Year) : null;
                            //double totalDebt = (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0) + (ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0);

                            //value = totalDebt != 0 ? (cashflow_ActivitiesValue != null && !string.IsNullOrEmpty(cashflow_ActivitiesValue.Value) ? Convert.ToDouble(cashflow_ActivitiesValue.Value) : 0) / totalDebt : 0;

                            value = DebtValue != 0 ? (cashflow_ActivitiesValue != null && !string.IsNullOrEmpty(cashflow_ActivitiesValue.Value) ? Convert.ToDouble(cashflow_ActivitiesValue.Value) : 0) / DebtValue : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);


                        //Free Cash Flow-to-Debt Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Free Cash Flow-to-Debt Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value);
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Debt-to-EBITDA Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Debt-to-EBITDA Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //Total Debt / EBITDA
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBITDA");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;

                            //Total Debt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }


                            //EBITDA
                            var EBITDAValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            //value = EBITDAValue != null && !string.IsNullOrEmpty(EBITDAValue.Value) && EBITDAValue.Value != "0" ? (((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0)) / Convert.ToDouble(EBITDAValue.Value)) : 0;
                            value = EBITDAValue != null && !string.IsNullOrEmpty(EBITDAValue.Value) && EBITDAValue.Value != "0" ? (DebtValue / Convert.ToDouble(EBITDAValue.Value)) : 0;
                            
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Debt-to-EBITA Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Debt-to-EBITA Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //Debt-to-EBITA Ratio
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBITA");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;

                            //total Debt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }

                            //EBITA
                            var EBITAValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            //value = EBITAValue != null && !string.IsNullOrEmpty(EBITAValue.Value) && EBITAValue.Value != "0" ? (((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) + (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0)) / Convert.ToDouble(EBITAValue.Value)) : 0;
                            value = EBITAValue != null && !string.IsNullOrEmpty(EBITAValue.Value) && EBITAValue.Value != "0" ? (DebtValue / Convert.ToDouble(EBITAValue.Value)) : 0;


                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Equity Multiplier (Book)
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Equity Multiplier (Book)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToUpper() == "TOTAL ASSETS" || x.LineItem.ToLower() == "assets" || x.LineItem.ToLower() == "assets,total"));

                        //(Total Assets)/ TOTAL STOCKHOLDERS' EQUITY
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                           
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            //stockholders euity
                            var StockholderequityValue = integratedDatasStockholdersequity != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasStockholdersequity.Id && x.Year == obj.Year) : null;

                            value = StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) && StockholderequityValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(StockholderequityValue.Value)) : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Equity Multiplier (Market)
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Equity Multiplier (Market)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //Enterprise Value ($M)/Market Value Common Stock ($M)
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //Enterprice Value
                            var EnterpriceValue = EnterpriceValue_Datas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == EnterpriceValue_Datas.Id && x.Year == obj.Year) : null;
                            //Market Value
                            var marketValue = market_CommonStock != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == market_CommonStock.Id && x.Year == obj.Year) : null;

                            value = marketValue != null && !string.IsNullOrEmpty(marketValue.Value) && marketValue.Value != "0" ? (EnterpriceValue != null && !string.IsNullOrEmpty(EnterpriceValue.Value) ? Convert.ToDouble(EnterpriceValue.Value) : 0) / Convert.ToDouble(marketValue.Value) : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        AnalysisFiling.FinancialStatementAnalysisDatasVM = AnalysisDatasList;
                        AnalysisFilingsList.Add(AnalysisFiling);
                        #endregion

                        #region OPERATING RETURNS (Income statement & Balance sheet)
                        AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
                        AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        AnalysisFiling.CompanyName = initialSetupObj.Company;
                        AnalysisFiling.ReportName = initialSetupObj.Company;
                        AnalysisFiling.StatementType = "OPERATING RETURNS (Income statement & Balance sheet)";
                        AnalysisFiling.Unit = "";
                        AnalysisFiling.CIK = initialSetupObj.CIKNumber;

                        //Asset Turnover
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Asset Turnover";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToUpper() == "TOTAL ASSETS" || x.LineItem.ToLower() == "assets" || x.LineItem.ToLower() == "assets,total"));
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //NET SALES/TOTAL ASSETS
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;
                            //total Assets
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            value = tempValue != null && !string.IsNullOrEmpty(tempValue.Value) && tempValue.Value != "0" ? (RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) ? Convert.ToDouble(RevenueValue.Value) : 0) / Convert.ToDouble(tempValue.Value) : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Return on Equity (ROE)
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Return on Equity (ROE)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "NET INCOME before extraordinary items");
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //NET INCOME before extraordinary items/TOTAL STOCKHOLDERS' EQUITY
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            //stockholders euity
                            var StockholderequityValue = integratedDatasStockholdersequity != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasStockholdersequity.Id && x.Year == obj.Year) : null;

                            value = StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) && StockholderequityValue.Value != "0" ? (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(StockholderequityValue.Value) : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Return on Equity (ROE)
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Return on Equity (ROE)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "NET INCOME before extraordinary items");

                        IntegratedDatasFAnalysis TotalAssetsIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.BalanceSheet && (x.LineItem.ToUpper() == "TOTAL ASSETS" || x.LineItem.ToLower() == "assets" || x.LineItem.ToLower() == "assets,total"));

                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //(NET INCOME before extraordinary items- Interest Expense)/TOTAL ASSETS
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            //Total Assets
                            var totalAssetsValue = TotalAssetsIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == TotalAssetsIntegratedItem.Id && x.Year == obj.Year) : null;
                            //interest Expense
                            var EXpenseValue = integratedDatasInterestExpense != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasInterestExpense.Id && x.Year == obj.Year) : null;

                            value = totalAssetsValue != null && !string.IsNullOrEmpty(totalAssetsValue.Value) && totalAssetsValue.Value != "0" ? ((tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) - (EXpenseValue != null && !string.IsNullOrEmpty(EXpenseValue.Value) ? Convert.ToDouble(EXpenseValue.Value) : 0)) / Convert.ToDouble(totalAssetsValue.Value) : 0;

                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Return on Invested Capital (ROIC) 
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Return on Invested Capital (ROIC)";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //EBIT (1 - Tax Rate) / (TOTAL STOCKHOLDERS' EQUITY +  (Short-Term Debt+ Long-Term Debt)-Cash and Cash Equivalents)
                        IntegratedDatasFAnalysis EBITDatas = integratedDatasList.Find(x => x.LineItem.ToLower() == "EBIT");
                        IntegratedDatasFAnalysis CashnCashEquivalentDatas = integratedDatasList.Find(x => x.LineItem.ToLower() == "cash and cash equivalents");

                        MarketDatas TaxRateDatas = marketDatasList != null && marketDatasList.Count > 0 ? marketDatasList.Find(x => x.LineItem == "Tax Rate (%)") : null;

                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            // short term debt
                            //var ShorttermdebtValue = integratedDatashorttermDebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatashorttermDebt.Id && x.Year == obj.Year) : null;
                            //// long term debt
                            //var longtermdebtValue = integratedDatalongtermdebt != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatalongtermdebt.Id && x.Year == obj.Year) : null;

                            //total Debt
                            double DebtValue = 0;
                            foreach (IntegratedDatasFAnalysis tempObj in debtList)
                            {
                                var tempValue = tempObj != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == tempObj.Id && x.Year == obj.Year) : null;
                                DebtValue = DebtValue + (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0);
                            }

                            //stockholders euity
                            var StockholderequityValue = integratedDatasStockholdersequity != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasStockholdersequity.Id && x.Year == obj.Year) : null;
                            //EBIT
                            var EBITValue = EBITDatas != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == EBITDatas.Id && x.Year == obj.Year) : null;
                            //CashnCashEquivalentValues
                            var CashnCashEquivalentValues = CashnCashEquivalentDatas != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == CashnCashEquivalentDatas.Id && x.Year == obj.Year) : null;
                            //TaxValues
                            var TAXmarketValue = TaxRateDatas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == TaxRateDatas.Id && x.Year == obj.Year) : null;
                            //totalSum
                            double TotalSum = (StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) ? Convert.ToDouble(StockholderequityValue.Value) : 0) + DebtValue  - (CashnCashEquivalentValues != null && !string.IsNullOrEmpty(CashnCashEquivalentValues.Value) ? Convert.ToDouble(CashnCashEquivalentValues.Value) : 0);

                            //double TotalSum = (StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) ? Convert.ToDouble(StockholderequityValue.Value) : 0) +
                            //   ((ShorttermdebtValue != null && !string.IsNullOrEmpty(ShorttermdebtValue.Value) ? Convert.ToDouble(ShorttermdebtValue.Value) : 0) +
                            //   (longtermdebtValue != null && !string.IsNullOrEmpty(longtermdebtValue.Value) ? Convert.ToDouble(longtermdebtValue.Value) : 0))
                            //   - (CashnCashEquivalentValues != null && !string.IsNullOrEmpty(CashnCashEquivalentValues.Value) ? Convert.ToDouble(CashnCashEquivalentValues.Value) : 0);

                            value = TotalSum != 0 ? ((EBITValue != null && !string.IsNullOrEmpty(EBITValue.Value) ? Convert.ToDouble(EBITValue.Value) : 0) * (1 - ((TAXmarketValue != null && !string.IsNullOrEmpty(TAXmarketValue.Value) ? Convert.ToDouble(TAXmarketValue.Value) : 0) / 100))) / TotalSum : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);


                        AnalysisFiling.FinancialStatementAnalysisDatasVM = AnalysisDatasList;
                        AnalysisFilingsList.Add(AnalysisFiling);
                        #endregion

                        #region VALUATION RATIOS (Income statement, Balance sheet & Market data)

                        AnalysisFiling = new FinancialStatementAnalysisFilingsViewModel();
                        AnalysisDatasList = new List<FinancialStatementAnalysisDatasViewModel>();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();

                        AnalysisFiling.CompanyName = initialSetupObj.Company;
                        AnalysisFiling.ReportName = initialSetupObj.Company;
                        AnalysisFiling.StatementType = "VALUATION RATIOS (Income statement, Balance sheet & Market data)";
                        AnalysisFiling.Unit = "";
                        AnalysisFiling.CIK = initialSetupObj.CIKNumber;


                        //Market-to-Book Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Market-to-Book Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //Market Value of Equity / Book Value of Equity
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //stockholders euity
                            var StockholderequityValue = integratedDatasStockholdersequity != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == integratedDatasStockholdersequity.Id && x.Year == obj.Year) : null;
                            //Market Value
                            var marketValue = market_CommonStock != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == market_CommonStock.Id && x.Year == obj.Year) : null;

                            value = StockholderequityValue != null && !string.IsNullOrEmpty(StockholderequityValue.Value) && StockholderequityValue.Value != "0" ? ((marketValue != null && !string.IsNullOrEmpty(marketValue.Value) ? Convert.ToDouble(marketValue.Value) : 0) / Convert.ToDouble(StockholderequityValue.Value)) : 0;
                            value = value * 100;
                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Price-to-Earnings Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Price-to-Earnings Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        //Share Price ($)/(NET INCOME before extraordinary items/Number of Shares Outstanding - Basic (Millions))
                        MarketDatas SharePrice_Market = marketDatasList != null && marketDatasList.Count > 0 ? marketDatasList.Find(x => x.LineItem == "Share Price ($)") : null;
                        MarketDatas market_NumberofShares = marketDatasList != null && marketDatasList.Count > 0 ? marketDatasList.Find(x => x.LineItem == "Number of Shares Outstanding - Basic (Millions)") : null;
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "NET INCOME before extraordinary items");

                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;

                            //SharePrice_Market Value
                            var SharePrice_MarketValue = SharePrice_Market != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == SharePrice_Market.Id && x.Year == obj.Year) : null;
                            //Number of Shares Outstanding 
                            var market_NumberofSharesValue = market_NumberofShares != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == market_NumberofShares.Id && x.Year == obj.Year) : null;
                            //(NET INCOME before extraordinary items- Interest Expense)/TOTAL ASSETS
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;
                            double dividendValue = market_NumberofSharesValue != null && !string.IsNullOrEmpty(market_NumberofSharesValue.Value) && market_NumberofSharesValue.Value != "0" ? (tempValue != null && !string.IsNullOrEmpty(tempValue.Value) ? Convert.ToDouble(tempValue.Value) : 0) / Convert.ToDouble(market_NumberofSharesValue.Value) : 0;

                            value = dividendValue != 0 ? ((SharePrice_MarketValue != null && !string.IsNullOrEmpty(SharePrice_MarketValue.Value) ? Convert.ToDouble(SharePrice_MarketValue.Value) : 0) / Convert.ToDouble(dividendValue)) : 0;

                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Enterprise Value-to-EBIT Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Enterprise Value-to-EBIT Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBIT");

                        //Enterprise Value / EBIT
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //Enterprice Value
                            var EnterpriceValue = EnterpriceValue_Datas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == EnterpriceValue_Datas.Id && x.Year == obj.Year) : null;
                            //Market Value
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            value = tempValue != null && !string.IsNullOrEmpty(tempValue.Value) && tempValue.Value != "0" ? (EnterpriceValue != null && !string.IsNullOrEmpty(EnterpriceValue.Value) ? Convert.ToDouble(EnterpriceValue.Value) : 0) / Convert.ToDouble(tempValue.Value) : 0;

                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Enterprise Value-to-EBITA Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Enterprise Value-to-EBITA Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBITA");

                        //Enterprise Value / EBITA
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //Enterprice Value
                            var EnterpriceValue = EnterpriceValue_Datas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == EnterpriceValue_Datas.Id && x.Year == obj.Year) : null;
                            //EBITA Value
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            value = tempValue != null && !string.IsNullOrEmpty(tempValue.Value) && tempValue.Value != "0" ? (EnterpriceValue != null && !string.IsNullOrEmpty(EnterpriceValue.Value) ? Convert.ToDouble(EnterpriceValue.Value) : 0) / Convert.ToDouble(tempValue.Value) : 0;

                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Enterprise Value-to-EBITDA Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Enterprise Value-to-EBITDA Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();
                        MatchedIntegratedItem = integratedDatasList.Find(x => x.StatementTypeId == (int)StatementTypeEnum.IncomeStatement && x.LineItem == "EBITDA");

                        //Enterprise Value / EBITDA
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //Enterprice Value
                            var EnterpriceValue = EnterpriceValue_Datas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == EnterpriceValue_Datas.Id && x.Year == obj.Year) : null;
                            //EBITA Value
                            var tempValue = MatchedIntegratedItem != null && integratedValuesList != null && integratedValuesList.Count > 0 ? integratedValuesList.Find(x => x.IntegratedDatasFAnalysisId == MatchedIntegratedItem.Id && x.Year == obj.Year) : null;

                            value = tempValue != null && !string.IsNullOrEmpty(tempValue.Value) && tempValue.Value != "0" ? (EnterpriceValue != null && !string.IsNullOrEmpty(EnterpriceValue.Value) ? Convert.ToDouble(EnterpriceValue.Value) : 0) / Convert.ToDouble(tempValue.Value) : 0;

                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);

                        //Enterprise Value-to-Sales Ratio
                        forcastRatioDatas = new FinancialStatementAnalysisDatasViewModel();
                        forcastRatioValuesList = new List<FinancialStatementAnalysisValuesViewModel>();
                        forcastRatioDatas.InitialSetup_FAnalysisId = initialSetupId;
                        forcastRatioDatas.LineItem = "Enterprise Value-to-Sales Ratio";
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = new List<FinancialStatementAnalysisValuesViewModel>();

                        //Enterprise Value / Revenue
                        foreach (FinancialStatementAnalysisValuesViewModel obj in dumyforcastRatioValueList)
                        {
                            forcastRatioValues = new FinancialStatementAnalysisValuesViewModel();
                            forcastRatioValues.Year = obj.Year;
                            double value = 0;
                            //Enterprice Value
                            var EnterpriceValue = EnterpriceValue_Datas != null && marketValuesList != null && marketValuesList.Count > 0 ? marketValuesList.Find(x => x.MarketDatasId == EnterpriceValue_Datas.Id && x.Year == obj.Year) : null;
                            //EBITA Value
                            var RevenueValue = revenuevaluesList != null && revenuevaluesList.Count > 0 ? revenuevaluesList.Find(x => x.Year == obj.Year) : null;

                            value = RevenueValue != null && !string.IsNullOrEmpty(RevenueValue.Value) && RevenueValue.Value != "0" ? (EnterpriceValue != null && !string.IsNullOrEmpty(EnterpriceValue.Value) ? Convert.ToDouble(EnterpriceValue.Value) : 0) / Convert.ToDouble(RevenueValue.Value) : 0;

                            forcastRatioValues.Value = Convert.ToString(value.ToString("0.##"));
                            forcastRatioValuesList.Add(forcastRatioValues);
                        }
                        forcastRatioDatas.FinancialStatementAnalysisValuesVM = forcastRatioValuesList;
                        AnalysisDatasList.Add(forcastRatioDatas);


                        AnalysisFiling.FinancialStatementAnalysisDatasVM = AnalysisDatasList;
                        AnalysisFilingsList.Add(AnalysisFiling);
                        #endregion

                        renderResult.StatusCode = 1;
                        renderResult.Message = "No issue found";
                        renderResult.Result = AnalysisFilingsList;
                        return Ok(renderResult);
                    }
                    else
                    {
                        renderResult.StatusCode = 0;
                        renderResult.Message = "No data available in Integrated Financial Statement";
                        renderResult.Result = AnalysisFilingsList;
                        return Ok(renderResult);
                    }

                }
                else
                {
                    renderResult.StatusCode = 0;
                    renderResult.Message = "No data availablefor this cik";
                    renderResult.Result = AnalysisFilingsList;
                    return Ok(renderResult);
                }
            }
            catch (Exception ss)
            {
                //Exception case
                renderResult.StatusCode = 0;
                renderResult.Message = "Exception occured" + Convert.ToString(ss.Message);
                renderResult.Result = AnalysisFilingsList;
                return Ok(renderResult);
            }
            //return Ok(renderResult);
        }


        #endregion

        #region get Flag

    

        [HttpGet]
        [Route("GetFlagForFinancialAnalysis/{InitialSetupId}/{cik}")]
        public List<FAnalysisFlag> GetFlagForFinancialAnalysis(long InitialSetupId, string cik)
        {
            List<FAnalysisFlag> FlagList = new List<FAnalysisFlag>();
            FAnalysisFlag flag = new FAnalysisFlag();
            try
            {
                Array values = Enum.GetValues(typeof(FAnalysisFlagEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    FlagList.Add(new FAnalysisFlag
                    {
                        FlagName = EnumHelper.DescriptionAttr((FAnalysisFlagEnum)i),
                        Id = i,
                        FlagValue = false
                    });
                }
                FlagList = FlagList.OrderBy(x => x.Id).ToList();


                //Check for filings table
                List<FilingsTable> FilingsList = new List<FilingsTable>();
                FilingsList = iFilings.FindBy(x => x.CIK == cik).ToList();
                if (FilingsList != null && FilingsList.Count > 0)
                {
                   
                    //for Raw Historical Data
                    var Rawitem = FlagList.Find(x => x.Id == (int)FAnalysisFlagEnum.RawHistorical);
                    Rawitem.FlagValue = true;

                    //Mrket Data
                    List<MarketDatas> marketDatasList = new List<MarketDatas>();
                    marketDatasList=iMarketDatas.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupId).ToList();
                    if (marketDatasList != null && marketDatasList.Count > 0)
                    {
                        var item = FlagList.Find(x => x.Id == (int)FAnalysisFlagEnum.MarketData);
                        item.FlagValue = true;

                        //Check for Integrated 
                        List<IntegratedDatasFAnalysis> IntegratedDatasList = new List<IntegratedDatasFAnalysis>();
                        IntegratedDatasList = iIntegratedDatasFAnalysis.FindBy(x => x.InitialSetup_FAnalysisId == InitialSetupId).ToList();
                        if (IntegratedDatasList != null && IntegratedDatasList.Count > 0)
                        {
                            var integrateditem = FlagList.Find(x => x.Id == (int)FAnalysisFlagEnum.IntegratedStatement);
                            integrateditem.FlagValue = true;

                            //enable Financial Statement Analysis when Integrated Financial Statement has data in database
                            var FinancialAnalysis = FlagList.Find(x => x.Id == (int)FAnalysisFlagEnum.FinancialAnalysis);
                            FinancialAnalysis.FlagValue = true;

                            var DataProcitem = FlagList.Find(x => x.Id == (int)FAnalysisFlagEnum.DataProcessing);
                            DataProcitem.FlagValue = true;
                        }
                    }
                    else
                    {
                        List<FAnalysis_CategoryByInitialSetup> CategoryList = new List<FAnalysis_CategoryByInitialSetup>();
                        CategoryList = iFAnalysis_CategoryByInitialSetup.FindBy(x => x.FAnalysis_InitialSetupId == InitialSetupId).ToList();
                        if (CategoryList != null && CategoryList.Count > 0)
                        {
                            var DataProcitem = FlagList.Find(x => x.Id == (int)FAnalysisFlagEnum.DataProcessing);
                            DataProcitem.FlagValue = true;
                        }
                    }

                   
                   

                }
            }
            catch (Exception ss)
            {

            }
            return FlagList;

        }

              #endregion

                }
            }