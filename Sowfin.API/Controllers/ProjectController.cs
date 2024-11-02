using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Sowfin.API.Lib;
using Sowfin.API.ViewModels;
using Sowfin.Data.Abstract;
using Sowfin.Data.Common.Enum;
using Sowfin.Data.Common.Helper;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using MathNet.Numerics.Financial;
using Excel.FinancialFunctions;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Hosting;
using System.IO;

using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using ServiceStack;

namespace Sowfin.API.Controllers
{
    [Microsoft.AspNetCore.Mvc.Route("api/[controller]")]
    [ApiController]
    public class ProjectController : ControllerBase
    {
        private const string PROJECTSNAPSHOT = "Project_snapshot";
        private const string SENSITIVITYSNAPSHOT = "Sensitivity_snapshot";
        private const string SCENARIOSNAPSHOT = "Scenario_snapshot";
        //Test 3434
        //private readonly IProjectRepository iProject = null;
        private readonly IProject iProject = null;
        private readonly IValuationTechniqueRepository valuationTechniqueRepository = null;
        IProjectInputDatas iProjectInputDatas;
        IProjectInputValues iProjectInputValues;
        IDepreciationInputDatas iDepreciationInputDatas;
        IDepreciationInputValues iDepreciationInputValues;
        IProjectInputComparables iProjectInputComparables;
        ISnapshot_CapitalBudgeting iSnapshot_CapitalBudgeting;
        ISensitivityInputData iSensitivityInputData;
        IScenarioInputData iScenarioInputDatas;
        IScenarioInputValues iScenarioInputValues;
        ISensitivityInputGraphData iSensitivityInputGraphData;

        IMapper mapper;
        private readonly IWebHostEnvironment  hostingEnvironment = null;
        public ProjectController(/*IProjectRepository _iProject,*/ IValuationTechniqueRepository _evaluationTechniqueRepository,
                                 IProject _iProject, IMapper _imapper, IProjectInputDatas _iProjectInputDatas,
                                 IWebHostEnvironment  _hostingEnvironment, IProjectInputValues _iProjectInputValues,
                                 IProjectInputComparables _iProjectInputComparables, ISnapshot_CapitalBudgeting _iSnapshot_CapitalBudgeting,
                                 IDepreciationInputDatas _iDepreciationInputDatas, IDepreciationInputValues _iDepreciationInputValues,
                                 ISensitivityInputData _iSensitivityInputData, IScenarioInputData _iScenarioInputDatas, IScenarioInputValues _iScenarioInputValues,
                                 ISensitivityInputGraphData _iSensitivityInputGraphData)
        {
            hostingEnvironment = _hostingEnvironment;
            iProject = _iProject;
            iProjectInputComparables = _iProjectInputComparables;
            iProjectInputDatas = _iProjectInputDatas;
            iProjectInputValues = _iProjectInputValues;
            iDepreciationInputDatas = _iDepreciationInputDatas;
            iDepreciationInputValues = _iDepreciationInputValues;
            valuationTechniqueRepository = _evaluationTechniqueRepository;
            iSnapshot_CapitalBudgeting = _iSnapshot_CapitalBudgeting;
            iSensitivityInputData = _iSensitivityInputData;
            iScenarioInputDatas = _iScenarioInputDatas;
            iScenarioInputValues = _iScenarioInputValues;
            iSensitivityInputGraphData = _iSensitivityInputGraphData;
            mapper = _imapper;
        }

        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("AddProject")]
        public ActionResult<Object> AddProject([FromBody] ProjectsViewModel model)
        {
            try
            {
                //Projects project = new Projects
                //{
                //    Name = model.Name,
                //    Description = model.Description,
                //    BusinessUnitId = model.BusinessUnitId,
                //    ValuationMethodId = model.ValuationMethodId,
                //    BusinessUnitName = model.BusinessUnitName,
                //    ValuationMethodName = model.ValuationMethodName
                //};

                Project project = new Project
                {
                    Name = model.Name,

                    ValuationTechniqueId = model.ValuationTechniqueId,
                    Description = model.Description,
                    BusinessUnitId = model.BusinessUnitId,
                    ModifiedDate = System.DateTime.UtcNow,
                    UserId = model.UserId,
                    ApprovalFlag = 0, //Disable ApprovalFlag on Update
                    StartingYear = model.StartingYear,
                    NoOfYears = model.NoOfYears,
                    IsActive = true,
                    CreatedDate = System.DateTime.UtcNow,
                    ApprovalComment = model.ApprovalComment


                };
                if (model.Id == 0)
                {
                    project.CreatedDate = System.DateTime.UtcNow;
                    iProject.Add(project);
                    iProject.Commit();

                }
                else
                {
                    project.Id = model.Id;
                    iProject.Update(project);
                    iProject.Commit();

                }


                if (project.Id != 0)
                {
                    //check for inputs 

                    if (model.ProjectInputDatasVM != null && model.ProjectInputDatasVM.Count > 0)
                    {
                        //check if existed data
                        List<ProjectInputDatas> inputDatasList = iProjectInputDatas.FindBy(x => x.ProjectId == model.Id).ToList();
                        List<DepreciationInputDatas> depreciationInputDatasList = iDepreciationInputDatas.FindBy(x => x.ProjectId == model.Id).ToList();
                        if (inputDatasList != null && inputDatasList.Count > 0)    // if(Data is available )
                        
                        {
                            //check for Values 
                            List<ProjectInputValues> valuesList = new List<ProjectInputValues>();
                            valuesList = iProjectInputValues.FindBy(x => inputDatasList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();

                            if (valuesList != null && valuesList.Count > 0)
                            {
                                iProjectInputValues.DeleteMany(valuesList);
                                iProjectInputValues.Commit();
                            }
                            if (model.ValuationTechniqueId == 4)
                            {
                                var comparableList = iProjectInputComparables.FindBy(x => inputDatasList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();
                                if (comparableList != null && comparableList.Count > 0)
                                {
                                    iProjectInputComparables.DeleteMany(comparableList);
                                    iProjectInputComparables.Commit();
                                }
                            }

                            var DepreciationInputValues = iDepreciationInputValues.FindBy(x => inputDatasList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();
                            if (DepreciationInputValues != null && DepreciationInputValues.Count > 0)
                            {
                                iDepreciationInputValues.DeleteMany(DepreciationInputValues);
                                iDepreciationInputValues.Commit();
                            }

                            iDepreciationInputDatas.DeleteMany(depreciationInputDatasList);
                            iDepreciationInputDatas.Commit();

                            iProjectInputDatas.DeleteMany(inputDatasList);
                            iProjectInputDatas.Commit();


                            // Delete Sensitivity Analysis Inputs setting

                            List<SensitivityInputData> sensitivityInputDatasList = iSensitivityInputData.FindBy(x => x.ProjectId == model.Id).ToList();
                            if (sensitivityInputDatasList!=null && sensitivityInputDatasList.Count > 0)    // if(Data is available )
                            {
                                iSensitivityInputData.DeleteMany(sensitivityInputDatasList);
                                iSensitivityInputData.Commit();
                            }

                            // Delete Scenario Analysis Inputs setting

                            List<ScenarioInputDatas> scenarioInputDatasList = iScenarioInputDatas.FindBy(x => x.ProjectId == model.Id).ToList();
                            if (scenarioInputDatasList != null && scenarioInputDatasList.Count > 0)
                            {
                                //check for Values 
                                List<ScenarioInputValues> scenarioValuesList = new List<ScenarioInputValues>();
                                scenarioValuesList = iScenarioInputValues.FindBy(x => scenarioInputDatasList.Any(t => t.Id == x.ScenarioInputDatasId)).ToList();

                                if (scenarioValuesList != null && scenarioValuesList.Count > 0)
                                {
                                    iScenarioInputValues.DeleteMany(scenarioValuesList);
                                    iScenarioInputValues.Commit();
                                }

                                iScenarioInputDatas.DeleteMany(scenarioInputDatasList);
                                iScenarioInputDatas.Commit();
                            }

                            // Delete Snapshot

                            List<Snapshot_CapitalBudgeting> Snapshot_CapitalBudgetingInputDatasList = iSnapshot_CapitalBudgeting.FindBy(x => x.ProjectId == model.Id).ToList();
                            if (Snapshot_CapitalBudgetingInputDatasList != null && Snapshot_CapitalBudgetingInputDatasList.Count > 0)    // if(Data is available )
                            {
                                iSnapshot_CapitalBudgeting.DeleteMany(Snapshot_CapitalBudgetingInputDatasList);
                                iSnapshot_CapitalBudgeting.Commit();
                            }

                            // Delete Sensitivity Analysis Graph setting

                            List<SensitivityInputGraphData> sensitivityInputGraphDatasList = iSensitivityInputGraphData.FindBy(x => x.ProjectId == model.Id).ToList();
                            if (sensitivityInputGraphDatasList != null && sensitivityInputGraphDatasList.Count > 0)    // if(Data is available )
                            {
                                iSensitivityInputGraphData.DeleteMany(sensitivityInputGraphDatasList);
                                iSensitivityInputGraphData.Commit();
                            }

                        }
                        //add new data in inputs and output
                        // map iniProjectInputComparablesput Datas
                        ProjectInputDatas projectInputDatas;
                        inputDatasList = new List<ProjectInputDatas>();
                        ProjectInputValues values = new ProjectInputValues();
                        ProjectInputComparables comparables = new ProjectInputComparables();
                        DepreciationInputValues DepreciationValues = new DepreciationInputValues();

                        var depreciationInputDatas = model.ProjectInputDatasVM.FindAll(x => x.SubHeader.Contains("Capex & Depreciation") && x.LineItem.Contains("Capex")).ToList();

                        foreach (var DatasVM in model.ProjectInputDatasVM)
                        {

                            if (DatasVM != null)
                            {

                                //mapDatas Vm to Datas 
                                DatasVM.ProjectId = model.Id;
                                DatasVM.Id = 0;

                                projectInputDatas = new ProjectInputDatas();
                                projectInputDatas = mapper.Map<ProjectInputDatasViewModel, ProjectInputDatas>(DatasVM);

                                if (projectInputDatas.ProjectInputValues != null && projectInputDatas.ProjectInputValues.Count > 0)
                                {
                                    foreach (var value in projectInputDatas.ProjectInputValues)
                                    {
                                        value.Id = 0;
                                    }

                                }
                                if (projectInputDatas.ProjectInputComparables != null && projectInputDatas.ProjectInputComparables.Count > 0)
                                {
                                    foreach (var value in projectInputDatas.ProjectInputComparables)
                                    {
                                        value.Id = 0;
                                    }

                                }
                                if (projectInputDatas.DepreciationInputValues != null && projectInputDatas.DepreciationInputValues.Count > 0)
                                {
                                    foreach (var value in projectInputDatas.DepreciationInputValues)
                                    {
                                        value.Id = 0;
                                    }

                                }

                                inputDatasList.Add(projectInputDatas);
                            }
                        }

                        //check for data availability

                        if (inputDatasList != null && inputDatasList.Count > 0)
                        {
                            iProjectInputDatas.AddMany(inputDatasList);
                            iProjectInputDatas.Commit();
                        }


                        if (depreciationInputDatas != null && depreciationInputDatas.Count > 0)
                        {
                            int i = 1;

                            foreach (var DepreciationDatasVM in depreciationInputDatas)
                            {
                                if (DepreciationDatasVM != null)
                                {
                                    var ProjectInputDatasId = iProjectInputDatas.GetSingle(x => x.ProjectId == DepreciationDatasVM.ProjectId && x.SubHeader.Contains("Capex & Depreciation" + i) && x.LineItem.Contains("Capex"));

                                    if (DepreciationDatasVM.Method == 3)
                                    {
                                        DepreciationInputDatas Depreciation = new DepreciationInputDatas
                                        {
                                            ProjectId = model.Id,
                                            ProjectInputDatasId = ProjectInputDatasId.Id,
                                            HasMultiYear = DepreciationDatasVM.HasMultiYear,
                                            Method = (int)DepreciationDatasVM.Method,
                                            SameYear = DepreciationDatasVM.SameYear,
                                            Duration = (int)(DepreciationDatasVM.Duration ?? 0)
                                        };

                                        iDepreciationInputDatas.Add(Depreciation);
                                        iDepreciationInputDatas.Commit();
                                    }
                                    else if (DepreciationDatasVM.Method != 3)
                                    {
                                        DepreciationInputDatas Depreciation = new DepreciationInputDatas
                                        {
                                            ProjectId = model.Id,
                                            ProjectInputDatasId = ProjectInputDatasId.Id,
                                            HasMultiYear = DepreciationDatasVM.HasMultiYear,
                                            Method = (int)DepreciationDatasVM.Method,
                                            SameYear = DepreciationDatasVM.SameYear,
                                            //Duration = (int)DepreciationDatasVM.Duration
                                            Duration = (int)(DepreciationDatasVM.Duration ?? 0)
                                        };

                                        iDepreciationInputDatas.Add(Depreciation);
                                        iDepreciationInputDatas.Commit();
                                    }

                                }

                                i = i + 1;
                            }
                        }

                    }

                }

                return Ok(new
                {
                    message = model.Id == 0 ? "Project Created Sucessfully" : "Project Modified Sucessfully",
                    id = project.Id,
                    statusCode = 200
                });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString() + "/n" + ex.InnerException, statusCode = 404 });
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("GetAllProjects/{BusinessId}")]
        public ActionResult<Object> GetAllProjects(long BusinessId)
        {
            try
            {

                var projects = iProject.FindBy(s => s.BusinessUnitId == BusinessId && s.IsActive == true);
                if (projects == null)
                {
                    return NotFound(new { message = "Unable to find projects", statusCode = 404 });
                }
                return Ok(new { result = projects, statusCode = 200 });
            }
            catch (Exception)
            {
                return BadRequest(new { message = "Bad Request", statusCode = 400 });
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("GetProject/{Id}")]
        public ActionResult GetProject(long Id)
        {
            try
            {
                Project projects = iProject.GetSingle(s => s.Id == Id && s.IsActive == true);

                if (projects == null)
                {
                    return NotFound(new { message = "Project not found", statusCode = 404 });
                }
                return Ok(new { result = projects, statusCode = 200 });
            }
            catch (Exception)
            {
                return BadRequest(new { message = "Bad Request", statusCode = 400 });
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("GetProjectDetails/{Id}")]
        public ActionResult GetProjectDetails(long Id)
        {
            ProjectsViewModel result = new ProjectsViewModel();
            try
            {
                Project project = iProject.GetSingle(x => x.Id == Id && x.IsActive == true);
                if (project != null)
                {
                    //List<ProjectInputDatasViewModel> projectInputDatasVM = new List<ProjectInputDatasViewModel>();
                    //List<ProjectInputValuesViewModel> projectInputValuesVM = new List<ProjectInputValuesViewModel>();
                    //ProjectInputDatasViewModel DatasVm = new ProjectInputDatasViewModel();
                    //result = mapper.Map<Project, ProjectsViewModel>(project);
                    result = ProjectDetails(Id, project);
                }
                else
                {
                    return NotFound(new { message = "Project not found", statusCode = 404 });

                    ////check for input Datas
                    //List<ProjectInputDatas> projectInputList = new List<ProjectInputDatas>();
                    //projectInputList = iProjectInputDatas.FindBy(x => x.Id != 0 && x.ProjectId == result.Id).ToList();
                    //if (projectInputList != null && projectInputList.Count > 0)
                }


            }
            catch (Exception ss)
            {
                return BadRequest(new { message = "Bad Request" + Convert.ToString(ss.Message), statusCode = 400 });
            }
            return Ok(result);
        }

        private ProjectsViewModel ProjectDetails(long Id, Project project)
        {
            ProjectsViewModel result = new ProjectsViewModel();
            try
            {
                List<ProjectInputDatasViewModel> projectInputDatasVM = new List<ProjectInputDatasViewModel>();
                List<ProjectInputValuesViewModel> projectInputValuesVM = new List<ProjectInputValuesViewModel>();
                ProjectInputDatasViewModel DatasVm = new ProjectInputDatasViewModel();

                // Depreciation
                List<DepreciationInputDatasViewModel> depreciationInputDatasVM = new List<DepreciationInputDatasViewModel>();
                List<DepreciationInputValuesViewModel> depreciationInputValuesVM = new List<DepreciationInputValuesViewModel>();
                DepreciationInputDatasViewModel DepreciationDatasVm = new DepreciationInputDatasViewModel();

                result = mapper.Map<Project, ProjectsViewModel>(project);

                //check for input Datas
                List<ProjectInputDatas> projectInputList = new List<ProjectInputDatas>();
                projectInputList = iProjectInputDatas.FindBy(x => x.Id != 0 && x.ProjectId == result.Id).ToList();
                if (projectInputList != null && projectInputList.Count > 0)
                {
                    //get all values List
                    List<ProjectInputValues> projectInputValueList = iProjectInputValues.FindBy(x => x.Id != 0 && projectInputList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();

                    //get all comparable List
                    List<ProjectInputComparables> projectInputComparableList = iProjectInputComparables.FindBy(x => x.Id != 0 && projectInputList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();

                    // List<DepreciationInputValues> depreciationInputValueList = iDepreciationInputValues.FindBy(x => x.Id != 0 && projectInputList.Any(t => t.Id == x.ProjectInputDatasId)).ToList();

                    // List<ProjectInputValues> ValuesList = null;
                    List<ProjectInputValuesViewModel> ValuesViewModelList = new List<ProjectInputValuesViewModel>();
                    // ProjectInputValuesViewModel valuesVM = null;

                    // Map Datas List to Datas VM List
                    foreach (ProjectInputDatas datas in projectInputList)
                    {


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

                        if (DatasVm != null && DatasVm.SubHeader.Contains("Capex & Depreciation") && DatasVm.LineItem.Contains("Capex"))
                        {
                            var depreciationInputDataList = iDepreciationInputDatas.GetSingle(x => x.Id != 0 && x.ProjectId == DatasVm.ProjectId && x.ProjectInputDatasId == DatasVm.Id);
                            if (depreciationInputDataList != null)
                            {
                                DatasVm.Duration = (int)depreciationInputDataList.Duration;
                                DatasVm.SameYear = depreciationInputDataList.SameYear;
                                DatasVm.Method = (int)depreciationInputDataList.Method;
                            }
                        }

                        if (DatasVm != null && DatasVm.ProjectInputValuesVM != null && DatasVm.ProjectInputValuesVM.Count > 0)
                        {
                            //DatasVm = new ProjectInputDatasViewModel();
                            var valuesList = DatasVm.ProjectInputValuesVM;
                            DatasVm.ProjectInputValuesVM = new List<ProjectInputValuesViewModel>();
                            DatasVm.ProjectInputValuesVM = valuesList.OrderBy(x => x.Year).ToList();

                            //DatasVm = mapper.Map<ProjectInputDatas, ProjectInputDatasViewModel>(datas);
                            //if(DatasVm!=null && DatasVm.ProjectInputValuesVM!=null && DatasVm.ProjectInputValuesVM.Count>0)
                            //{
                            //var valuesList = DatasVm.ProjectInputValuesVM;
                            //DatasVm.ProjectInputValuesVM = new List<ProjectInputValuesViewModel>();
                            //DatasVm.ProjectInputValuesVM=valuesList.OrderBy(x=>x.Year).ToList();
                        }


                        projectInputDatasVM.Add(DatasVm);
                    }

                }

                //add values to datas
                if (projectInputDatasVM != null && projectInputDatasVM.Count > 0)
                {

                    //add Datas to result
                    result.ProjectInputDatasVM = new List<ProjectInputDatasViewModel>();
                    result.ProjectInputDatasVM = projectInputDatasVM;

                    //get summaryoutput
                    result.ProjectSummaryDatasVM = GetProjectOutput(result);
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



                }

                //Get Approval Status
                result.ApprovalStatus = GetApprovalStatus(Id);

                #region Snapshot
                List<Snapshot_CapitalBudgeting> Snapshot = iSnapshot_CapitalBudgeting.FindBy(s => s.ProjectId == Id && s.Active == true && s.SnapshotType == PROJECTSNAPSHOT).ToList();
                result.Snapshot_ProjectVM = new List<Snapshot_ProjectViewModel>();
                foreach (var tempSanpshot in Snapshot)
                {
                    result.Snapshot_ProjectVM.Add(mapper.Map<Snapshot_CapitalBudgeting, Snapshot_ProjectViewModel>(tempSanpshot));
                }
                #endregion


                result.CurrencyValueList = EnumHelper.GetEnumListbyName<CurrencyValueEnum>(); //get CurrencyValueList
                result.NumberCountList = EnumHelper.GetEnumListbyName<NumberCountEnum>(); //get Number countList
                result.ValueTypeList = EnumHelper.GetEnumListbyName<ValueTypeEnum>(); //get Value Type List
                result.ValuationTechniqueList = GetTechniques();  //get Valuation Technique List
                result.HeaderList = EnumHelper.GetEnumListbyName<HeadersEnum>(); //get leverage Policy List

            }

            catch (Exception ex)
            {
                //throw ex;
            }
            return result;
        }

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
                                    //value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) * (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0);
                                    value = (InterestPaidValue != null ? UnitConversion.getBasicValueforCurrency(InterestpaidLongDatas.UnitId, InterestPaidValue.Value) : 0) / (CostofDebt != null ? Convert.ToDouble(CostofDebt.Value) / 100 : 0);
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

                // return OutputDatasVMList;
            }
            return OutputDatasVMList;
        }


        private ExcelWorksheet InsertInputToExcel(List<ProjectOutputDatasViewModel> outputDatasList, decimal MarginalTax, decimal DVRatio, decimal UCC, decimal CostOfDebt, int StartingYear, int totalYears, ExcelWorksheet wsCapitalStructure)
        {
            try
            {
                //fpnd FreeCashFlow
                var FreeCashFlow = outputDatasList != null && outputDatasList.Count > 0 ? outputDatasList.Find(x => x.LineItem.Contains("Free Cash Flow")) : null;
                var leveredValueData = outputDatasList != null && outputDatasList.Count > 0 ? outputDatasList.Find(x => x.LineItem.Contains("Unlevered Value @ rU")) : null;

                wsCapitalStructure.Cells["C2"].Value = MarginalTax / 100;
                wsCapitalStructure.Cells["C2"].Style.Numberformat.Format = "0.0%";

                wsCapitalStructure.Cells["C3"].Value = DVRatio / 100;
                wsCapitalStructure.Cells["C3"].Style.Numberformat.Format = "0.0%";

                wsCapitalStructure.Cells["C4"].Value = UCC / 100;
                wsCapitalStructure.Cells["C4"].Style.Numberformat.Format = "0.0%";

                wsCapitalStructure.Cells["C5"].Value = CostOfDebt / 100;
                wsCapitalStructure.Cells["C5"].Style.Numberformat.Format = "0.0%";

                #region Old Approach
                int initial = 3;
                //var value=null;
                for (int i = 0; i < totalYears; i++)
                {
                    decimal FreeCashValue = FreeCashFlow != null && FreeCashFlow.ProjectOutputValuesVM != null && FreeCashFlow.ProjectOutputValuesVM.Count > 0 && FreeCashFlow.ProjectOutputValuesVM.Find(x => x.Year == StartingYear + i) != null ? Convert.ToDecimal(FreeCashFlow.ProjectOutputValuesVM.Find(x => x.Year == StartingYear + i).Value) : 0;
                    // Bind Free Cash Flow
                    wsCapitalStructure.Cells[7, initial + i].Value = FreeCashValue;
                    wsCapitalStructure.Cells[7, initial + i].Style.Numberformat.Format = "$#,##0.00";


                    //Levered Value
                    decimal LeveredValue = leveredValueData != null && leveredValueData.ProjectOutputValuesVM != null && leveredValueData.ProjectOutputValuesVM.Count > 0 && leveredValueData.ProjectOutputValuesVM.Find(x => x.Year == StartingYear + i) != null ? Convert.ToDecimal(leveredValueData.ProjectOutputValuesVM.Find(x => x.Year == StartingYear + i).Value) : 0;
                    wsCapitalStructure.Cells[9, initial + i].Value = LeveredValue;
                    wsCapitalStructure.Cells[9, initial + i].Style.Numberformat.Format = "$#,##0.00";
                    //Debt Capacity    
                    if (i == totalYears - 1)
                    {
                        // wsCapitalStructure.Cells[11, initial + i].Value = 0;
                    }
                    else
                    {

                        wsCapitalStructure.Cells[11, initial + i].Formula = "=($C$3*" + wsCapitalStructure.Cells[16, initial + i] + ")";
                        //wsCapitalStructure.Cells[11, initial + i].Value = DVRatio *(wsCapitalStructure.Cells[16, initial + i].Value!=null ? Convert.ToDecimal(wsCapitalStructure.Cells[16, initial + i].Value) : 0);

                    }
                    wsCapitalStructure.Cells[11, initial + i].Style.Numberformat.Format = "$#,##0.00";


                    if (i == 0)
                    {
                        //Interest Paid @ rD/////
                        // wsCapitalStructure.Cells[12, initial + i].Value = 0;
                        //Interest Tax Shield @ Tc
                        // wsCapitalStructure.Cells[13, initial + i].Value = 0;
                    }
                    else
                    {
                        //Interest Paid @ rD
                        wsCapitalStructure.Cells[12, initial + i].Formula = "=($C$5*" + wsCapitalStructure.Cells[11, initial + i - 1] + ")";
                        //wsCapitalStructure.Cells[12, initial + i].Value = CostOfDebt * Convert.ToDecimal(wsCapitalStructure.Cells[11, initial + i - 1].Value) ;
                        //Interest Tax Shield @ Tc
                        wsCapitalStructure.Cells[13, initial + i].Formula = "=($C$2*" + wsCapitalStructure.Cells[12, initial + i] + ")";
                        //wsCapitalStructure.Cells[13, initial + i].Value = MarginalTax * Convert.ToDecimal(wsCapitalStructure.Cells[12, initial + i].Value);
                    }
                    wsCapitalStructure.Cells[12, initial + i].Style.Numberformat.Format = "$#,##0.00";
                    wsCapitalStructure.Cells[13, initial + i].Style.Numberformat.Format = "$#,##0.00";



                    if (i == totalYears - 1)
                    {
                        //  wsCapitalStructure.Cells[14, initial + i].Value  = 0;
                        //  wsCapitalStructure.Cells[16, initial + i].Value  = 0;
                    }
                    else
                    {
                        //Tax Shield Value @ rU
                        // wsCapitalStructure.Cells[14, initial + i].Value = Convert.ToDecimal(0);
                        wsCapitalStructure.Cells[14, initial + i].Formula = "=(" + wsCapitalStructure.Cells[13, initial + i + 1] + "+" + wsCapitalStructure.Cells[14, initial + i + 1] + ")/(1+$C$4)";
                        // wsCapitalStructure.Cells[14, initial + i].Value = (Convert.ToDecimal(wsCapitalStructure.Cells[13, initial + i + 1].Value)+ Convert.ToDecimal(wsCapitalStructure.Cells[14, initial + i + 1].Value)) /(1+(Convert.ToDecimal(UCC)/100));

                        //Levered Value (VL = VU + T)
                        //wsCapitalStructure.Cells[16, initial + i].Value = Convert.ToDecimal(0);
                        wsCapitalStructure.Cells[16, initial + i].Formula = "=SUM(" + wsCapitalStructure.Cells[9, initial + i] + "," + wsCapitalStructure.Cells[14, initial + i] + ")";
                        // wsCapitalStructure.Cells[16, initial + i].Value = Convert.ToDecimal(wsCapitalStructure.Cells[9, initial + i].Value) + Convert.ToDecimal(wsCapitalStructure.Cells[14, initial + i].Value) ;
                    }
                    wsCapitalStructure.Cells[14, initial + i].Style.Numberformat.Format = "$#,##0.00";
                    wsCapitalStructure.Cells[16, initial + i].Style.Numberformat.Format = "$#,##0.00";



                }
                #endregion

                int lastyear = StartingYear + totalYears;
                #region  New approach
                for (int j = lastyear; j >= StartingYear; j--)
                {
                    // Free Cash Flow




                }

                #endregion

                ////
                wsCapitalStructure.Cells["C17"].Formula = "=SUM(" + wsCapitalStructure.Cells["C7"] + "," + wsCapitalStructure.Cells["C16"] + ")";
                wsCapitalStructure.Cells["C17"].Style.Numberformat.Format = "$#,##0.00";
                //wsCapitalStructure.Cells["C17"].Value =  Convert.ToDecimal(wsCapitalStructure.Cells["C7"].Value) + Convert.ToDecimal(wsCapitalStructure.Cells["C16"].Value);

            }
            catch (Exception ss)
            {

            }
            return wsCapitalStructure;
        }



        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("DeleteProject/{Id}")]
        public ActionResult<Object> DeleteProject(long Id)
        {
            try
            {
                var TblProjectObj = iProject.GetSingle(s => s.Id == Id);
                if (TblProjectObj == null)
                {
                    return NotFound(new { Message = "Record not found", StatusCode = 0 });
                }
                else
                {
                    TblProjectObj.IsActive = false;
                    iProject.Update(TblProjectObj);

                    // iProject.Delete(TblProjectObj);
                    iProject.Commit();
                    return Ok(new { Message = "Record deleted", StatusCode = 1 });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "Bad Request", statusCode = 400 });
            }
        }

        private List<ValuationTechnique> GetTechniques()
        {
            try
            {
                List<ValuationTechnique> valuationTechniques = valuationTechniqueRepository.AllIncluding().ToList();
                if (valuationTechniques != null && valuationTechniques.Count > 0)
                {
                    return valuationTechniques;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("ValuationTechniques")]
        public ActionResult<Object> GetValuationTechniques()
        {
            try
            {
                var valuationTechniques = valuationTechniqueRepository.AllIncluding();
                if (valuationTechniques == null)
                {
                    return NotFound(new { message = "No valuation techniques", statusCode = 404 });
                }

                return Ok(new { result = valuationTechniques, statusCode = 200 });
            }
            catch (Exception)
            {
                return BadRequest(new { message = "Bad Request", statusCode = 400 });
            }
        }

        //Created by Anonmous
        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("AddProjectSnapShot")]
        public ActionResult AddProjectSnapShot([FromBody] Snapshot_ProjectViewModel model)
        {
            try
            {
                Snapshot_CapitalBudgeting snapshot = new Snapshot_CapitalBudgeting
                {

                    ProjectInputs = model.ProjectInputs.ToString(),
                    Snapshot = model.Snapshot.ToString(),
                    Description = model.Description,
                    UserId = model.UserId,
                    ProjectId = model.ProjectId,
                    ValuationTechniqueId = Convert.ToInt32(model.ValuationTechniqueId),
                    SnapshotType = PROJECTSNAPSHOT,
                    CreatedAt = (DateTime)model.CreatedAt, //DateTime.Now,
                    Active = true,
                    NPV = model.NPV,
                    CNPV = model.CNPV
                };
                iSnapshot_CapitalBudgeting.Add(snapshot);
                iSnapshot_CapitalBudgeting.Commit();
                return Ok(new { result = snapshot.Id, message = "Succesfully added Snapshots", code = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("GetProjectSnapshot/{Id}")]
        public ActionResult GetProjectSnapShot(long Id)
        {
            try
            {

                var SnapShot = iSnapshot_CapitalBudgeting.GetSingle(s => s.Id == Id);
                if (SnapShot == null)
                {
                    return NotFound("snapshot not found by this Id");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("ProjectSnapshots/{ProjecId}")]
        public ActionResult<Object> CapitalBudgetingSnapshots(long ProjecId)
        {
            try
            {
                var SnapShot = iSnapshot_CapitalBudgeting.FindBy(s => s.ProjectId == ProjecId);
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

        private Approval GetApprovalStatus(long ProjectId)
        {
            Approval result = new Approval();
            try
            {
                var project = iProject.FindBy(s => s.Id == ProjectId).FirstOrDefault();
                if (project != null)
                {
                    result = new Approval
                    {
                        ProjectId = project.Id,
                        ApprovalFlag = project.ApprovalFlag != null ? (int)project.ApprovalFlag : 0,
                        ApprovalComment = project.ApprovalComment,
                        ApprovalPassword = "admin123"
                    };
                }
            }
            catch (Exception ex)
            {
                return (new Approval());
            }
            return result;
        }

        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("ApprovalStatus")]
        public ActionResult<Object> UpdateApprovalStatus([FromBody] Approval model)
        {
            try
            {
                var project = iProject.GetSingle(s => s.Id == model.ProjectId);
                if (project != null)
                {
                    project.ApprovalFlag = 1;
                    project.ApprovalComment = model.ApprovalComment;
                    iProject.Update(project);
                    iProject.Commit();
                    return Ok(new { result = project, comment = project.ApprovalComment, message = "Successfully approved", code = 200 });
                }
                else{
                    return NotFound(new { message = "Record not found", code = 404 });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
            return Ok(new { result = model.ProjectId, message = "Approval comment not saved", code = 200 });
        }


        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("ProjectScenario")]
        public ActionResult<Object> ProjectScenario([FromBody] List<Dictionary<string, object>> body, long ProjectId)
        {
            List<Dictionary<string, object>> npvOutput = new List<Dictionary<string, object>>();
            try
            {
                ProjectsViewModel input = new ProjectsViewModel();
                if (ProjectId != 0)
                {
                    Project project = iProject.GetSingle(x => x.Id == ProjectId);
                    if (project != null)
                    {
                        input = ProjectDetails(ProjectId, project);

                        for (int i = 0; i < body.Count; i++)
                        {
                            Dictionary<string, object> npvList = SensiScenarioNpv(body[i], input);
                            npvOutput.Add(npvList);
                        }

                    }
                    else{
                    return NotFound(new { message = "Project not found By this ProjectId", code = 404 });
                }
                }

            }
            catch (Exception ex)
            {
                return (BadRequest(new { message = ex.Message.ToString(), code = 400 }));
            }
            return npvOutput;
        }

        private Dictionary<string, object> SensiScenarioNpv(Dictionary<string, object> model, ProjectsViewModel projectVM)
        {
            Project_ScenarioOutput scenOutput = new Project_ScenarioOutput();
            Dictionary<string, object> npvOutput = new Dictionary<string, object>();
            double npv = 0;
            try
            {
                var objectDictionary = new Dictionary<string, object>();
                ProjectsViewModel inputProject = new ProjectsViewModel();
                model = CheckForNull(model);

                objectDictionary = ScenarioValues(model, projectVM, out scenOutput);

                ProjectsViewModel ScenarioProject = ConvertScenarioValuesIntoProjectVM(objectDictionary, projectVM);

                List<ProjectOutputDatasViewModel> summaryOutput = GetProjectOutput(ScenarioProject); //TODO

                foreach (var output in summaryOutput)
                {
                    if (output.LineItem.Contains("Net Present Value"))
                    {
                        npv = Convert.ToDouble(output.Value);
                        string tempNpvStr = "NPV" + FindUnit((int)output.ValueTypeId, (int)output.UnitId);
                        npvOutput.Add(tempNpvStr, npv);
                        break;
                    }
                }

                scenOutput.sensiSceSummaryOutput = summaryOutput;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return npvOutput;
        }

        private static object FindForRevenueCapex(object value, List<double> aggregateList, out double percentChange, int flag)
        {
            List<double> finalAggList = new List<double>();
            double total = aggregateList.Sum();
            double average = aggregateList.Count != 0 ? total / aggregateList.Count : 0;
            double changeInValue = 0;
            if (flag == 0) //Volume use sum
            {
                changeInValue = total != 0 ? Convert.ToDouble(value) / total : 0;
                foreach (double aggregateValues in aggregateList)
                {
                    finalAggList.Add(aggregateValues * changeInValue);
                }
            }
            else if (flag == 1)  /// Average unit price and unit cost   use average
            {
                changeInValue = average != 0 ? Convert.ToDouble(value) / average : 0;
                foreach (double aggregateValues in aggregateList)
                {
                    finalAggList.Add(aggregateValues * changeInValue);
                }
            }
            percentChange = changeInValue;

            return finalAggList;
        }

        private static object FindAverageForCost(object value, List<double> costList, List<double> volumeList)
        {
            List<double> finalAggList = new List<double>();
            double totalVolume = volumeList.Sum();
            double changeInValue = 0;
            List<double> total = new List<double>();
            for (int count = 0; count < costList.Count; count++)
            {
                total.Add(volumeList[count] * costList[count]);
            }

            double average = totalVolume != 0 ? Convert.ToDouble(total.Sum() / totalVolume) : 0;
            changeInValue = average != 0 ? Convert.ToDouble(value) / average : 0;
            foreach (double values in costList)
            {
                finalAggList.Add(values * changeInValue);
            }
            return finalAggList;
        }

        private static object FindAverageForDepreciation(List<double> depreciationList, double ChangeInCapex)
        {
            List<double> finalAggList = new List<double>();
            foreach (double values in depreciationList)
            {
                finalAggList.Add(values * ChangeInCapex);
            }
            return finalAggList;
        }
        private static Double ParseDouble(Object obj)
        {
            if ((obj == null) || (obj.ToString() == ""))
            {
                return 0;
            }

            return Convert.ToDouble(obj);
        }
        private static Dictionary<string, object> ScenarioValues(Dictionary<string, object> model, ProjectsViewModel project, out Project_ScenarioOutput scenOutput)
        {
            var objectDictionary = new Dictionary<string, object>();
            List<KeyValuePair<string, object>> objlist = model.ToList();
            string outputStr = "";
            try
            {
                //output 
                double marginalTax = 0;
                double waccDiscount = 0;
                double dvRatio = 0;
                double unleveredCost = 0;
                double costOfDebt = 0;
                double costOfEquity = 0;
                double interestCoverage = 0;
                double unitPricePerChange = 1;
                double unitCostPerChange = 1;
                double capexDepPerChange = 1;
                double volChangePerc = 1;
                double rDPerChange = 1;
                double sgAPerChange = 1;
                double fixedSchedulePerChange = 1;

                Project_ScenarioAnalysis inputs = new Project_ScenarioAnalysis();
                scenOutput = new Project_ScenarioOutput();

                //Copy Project InputDataValues To Project_ScenarioAnalysis
                foreach (var inputdata in project.ProjectInputDatasVM)
                {
                    if (inputdata.LineItem.Contains("Volume"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputVolumes.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("Unit Price"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputUnitPrice.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("Unit Cost"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputUnitCost.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("Marginal Tax rate"))
                    {
                        inputs.inputMarginalTax = Convert.ToDouble(inputdata.Value);
                    }
                    else if (inputdata.LineItem.Contains("Capex"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputCapex.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("Depreciation"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputDepreciation.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("R&D"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputRD.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("SG&A"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputSgA.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("NWC"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputNWC.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }
                    else if (inputdata.LineItem.Contains("Fixed Schedule & Predetermined Debt Level, Dt"))
                    {
                        foreach (var inputValue in inputdata.ProjectInputValuesVM)
                        {
                            if (inputValue.Value == null)
                                inputValue.Value = 0;
                            inputs.inputFixedSchedule.Add(Convert.ToDouble(inputValue.Value));
                        }
                    }//Valution Tech 6
                    else if (inputdata.LineItem.Contains("WACC"))
                    {
                        inputs.inputWACC = Convert.ToDouble(inputdata.Value);
                    } //rWACC
                    else if (inputdata.LineItem.Contains("D/V Ratio"))
                    {
                        inputs.inputDVRatio = Convert.ToDouble(inputdata.Value);
                    } //D/V Ratio
                    else if (inputdata.LineItem.Contains("Unlevered cost of capital"))
                    {
                        inputs.inputUnleveredCost = Convert.ToDouble(inputdata.Value);
                    } //rU
                    else if (inputdata.LineItem.Contains("Cost of Debt"))
                    {
                        inputs.inputCostofDebt = Convert.ToDouble(inputdata.Value);
                    } //rD
                    else if (inputdata.LineItem.Contains("Cost of Equity"))
                    {
                        inputs.inputCostofEquity = Convert.ToDouble(inputdata.Value);
                    }  //rE
                    else if (inputdata.LineItem.Contains("Intetrest coverage ratio =Interest expense/FCF"))
                    {
                        inputs.inputIntrestCoverageRatio = Convert.ToDouble(inputdata.Value);
                    } //k
                    else if (inputdata.LineItem.Contains("Project's Equity Cost of Capital") && inputdata.SubHeader=="")
                    {
                        inputs.inputEquityCostofCapital = Convert.ToDouble(inputdata.Value);
                    }//rE and Valuation Tech 5
                    else if (inputdata.LineItem.Contains("Project's Target Leverage (D/V) Ratio") && inputdata.SubHeader == "")
                    {
                        inputs.inputProjectDVRatio = Convert.ToDouble(inputdata.Value);
                    }// D/V ratio for Valuation Tech 4
                    else if (inputdata.LineItem.Contains("Project's Debt Cost of Capital (rD)") && inputdata.SubHeader == "")
                    {
                        inputs.inputDebtCostofCapital = Convert.ToDouble(inputdata.Value);
                    }// rD for Valuation Tech 4

                }

                //Todo Check model Key When Link To UI
                #region Volume         

                if (StrMatch(model.Keys.ToList(), "volume", out outputStr))
                {
                    objectDictionary["Volume"] = FindForRevenueCapex(model[outputStr], inputs.inputVolumes, out volChangePerc, 0); // passing flag zero to use total 
                }
                else
                {
                    objectDictionary.Add("Volume", inputs.inputVolumes);
                }
                #endregion

                #region unitPrice
                if (StrMatch(model.Keys.ToList(), "unit price", out outputStr))
                {
                    objectDictionary["UnitPrice"] = FindAverageForCost(model[outputStr], inputs.inputUnitPrice, inputs.inputVolumes); // passing flag one to use average
                }
                else
                {
                    objectDictionary.Add("UnitPrice", inputs.inputUnitPrice);
                }
                #endregion

                #region unitCost
                if (StrMatch(model.Keys.ToList(), "unit cost", out outputStr))
                {
                    objectDictionary["UnitCost"] = FindAverageForCost(model[outputStr], inputs.inputUnitCost, inputs.inputVolumes); // passing flag one to use average 
                }
                else
                {
                    objectDictionary.Add("UnitCost", inputs.inputUnitCost);
                }
                #endregion

                #region marginalTax
                if (StrMatch(model.Keys.ToList(), "marginal", out outputStr))
                {
                    marginalTax = ParseDouble(model[outputStr]);
                    objectDictionary.Add("MarginalTax", model[outputStr]);
                }
                else
                {
                    marginalTax = inputs.inputMarginalTax;
                    objectDictionary.Add("MarginalTax", marginalTax);
                }
                #endregion

                #region capex&Depreciation //Todo change Depreciation Model name change
                if (StrMatch(model.Keys.ToList(), "capex", out outputStr))
                {
                    objectDictionary["Capex"] = FindForRevenueCapex(model[outputStr], inputs.inputCapex, out capexDepPerChange, 0); // passing flag zero to use total
                    objectDictionary["TotalDepreciation"] = FindAverageForDepreciation(inputs.inputDepreciation, capexDepPerChange); // passing flag zero to use total
                }
                else
                {
                    objectDictionary.Add("Capex", inputs.inputCapex);
                    objectDictionary.Add("TotalDepreciation", inputs.inputDepreciation);
                }
                #endregion

                #region Fixed Cost
                if (StrMatch(model.Keys.ToList(), "R&D", out outputStr))
                {
                    objectDictionary["RD"] = FindForRevenueCapex(model[outputStr], inputs.inputRD, out rDPerChange, 0); // passing flag one to use average 
                }
                else
                {
                    objectDictionary.Add("RD", inputs.inputRD);
                }

                if (StrMatch(model.Keys.ToList(), "SG&A", out outputStr))
                {
                    objectDictionary["SgA"] = FindForRevenueCapex(model[outputStr], inputs.inputSgA, out sgAPerChange, 0); // passing flag one to use average 
                }
                else
                {
                    objectDictionary.Add("SgA", inputs.inputSgA);
                }
                #endregion

                #region NWC
                if (StrMatch(model.Keys.ToList(), "NWC", out outputStr))
                {
                    objectDictionary["NWC"] = FindForRevenueCapex(model[outputStr], inputs.inputNWC, out _, 1); // passing flag zero to use total 
                }
                else
                {
                    objectDictionary.Add("NWC", inputs.inputNWC);
                }
                #endregion

                #region Fixed Schedule & Predetermined Debt Level, Dt //For Valuation Technique 6
                if (StrMatch(model.Keys.ToList(), "Fixed Schedule", out outputStr))
                {
                    objectDictionary["Fixed Schedule"] = FindForRevenueCapex(model[outputStr], inputs.inputFixedSchedule, out fixedSchedulePerChange, 0); // passing flag one to use average 
                }
                else
                {
                    objectDictionary.Add("Fixed Schedule", inputs.inputFixedSchedule);
                }
                #endregion

                #region discountRate //Here Apply discout rate for all Valution Technique
                if (StrMatch(model.Keys.ToList(), "WACC", out outputStr))
                {
                    waccDiscount = ParseDouble(model[outputStr]);
                    objectDictionary.Add("WACC", model[outputStr]);
                }
                else
                {
                    waccDiscount = inputs.inputWACC;
                    objectDictionary.Add("WACC", inputs.inputWACC);
                }

                if (StrMatch(model.Keys.ToList(), "D/V Ratio", out outputStr))
                {
                    dvRatio = ParseDouble(model[outputStr]);
                    objectDictionary.Add("D/V Ratio", model[outputStr]);
                }
                else
                {
                    dvRatio = inputs.inputDVRatio;
                    objectDictionary.Add("D/V Ratio", inputs.inputDVRatio);
                }

                if (StrMatch(model.Keys.ToList(), "Unlevered cost of capital", out outputStr))
                {
                    unleveredCost = ParseDouble(model[outputStr]);
                    objectDictionary.Add("Unlevered cost of capital", model[outputStr]);
                }
                else
                {
                    unleveredCost = inputs.inputUnleveredCost;
                    objectDictionary.Add("Unlevered cost of capital", inputs.inputUnleveredCost);
                }

                if (StrMatch(model.Keys.ToList(), "Cost of Equity", out outputStr))
                {
                    costOfEquity = ParseDouble(model[outputStr]);
                    objectDictionary.Add("Cost of Equity", model[outputStr]);
                }
                else
                {
                    costOfEquity = inputs.inputCostofEquity;
                    objectDictionary.Add("Cost of Equity", inputs.inputCostofEquity);
                }

                if (StrMatch(model.Keys.ToList(), "Cost of Debt", out outputStr))
                {
                    costOfDebt = ParseDouble(model[outputStr]);
                    objectDictionary.Add("Cost of Debt", model[outputStr]);
                }
                else
                {
                    costOfDebt = inputs.inputCostofDebt;
                    objectDictionary.Add("Cost of Debt", inputs.inputCostofDebt);
                }

                if (StrMatch(model.Keys.ToList(), "coverage ratio", out outputStr))
                {
                    interestCoverage = ParseDouble(model[outputStr]);
                    objectDictionary.Add("Interest coverage ratio", model[outputStr]);
                }
                else
                {
                    interestCoverage = inputs.inputIntrestCoverageRatio;
                    objectDictionary.Add("Interest coverage ratio", inputs.inputIntrestCoverageRatio);
                }

                if (StrMatch(model.Keys.ToList(), "Project's Target Leverage", out outputStr))
                {
                    interestCoverage = ParseDouble(model[outputStr]);
                    objectDictionary.Add("Project's Target Leverage", model[outputStr]);
                }
                else
                {
                    interestCoverage = inputs.inputProjectDVRatio;
                    objectDictionary.Add("Project's Target Leverage", inputs.inputProjectDVRatio);
                }

                if (StrMatch(model.Keys.ToList(), "Project's Debt Cost of Capital", out outputStr))
                {
                    interestCoverage = ParseDouble(model[outputStr]);
                    objectDictionary.Add("Project's Debt Cost of Capital", model[outputStr]);
                }
                else
                {
                    interestCoverage = inputs.inputDebtCostofCapital;
                    objectDictionary.Add("Project's Debt Cost of Capital", inputs.inputDebtCostofCapital);
                }

                #endregion


                //assign changes value to Project_ScenarioOutput class
                //TODO need to change output According to Valution Technique for Export
                scenOutput.waccDiscount = dvRatio;
                scenOutput.waccDiscount = unleveredCost;
                scenOutput.waccDiscount = costOfDebt;
                scenOutput.waccDiscount = costOfEquity;
                scenOutput.waccDiscount = interestCoverage;

                scenOutput.waccDiscount = waccDiscount;
                scenOutput.capexDepPerChange = fixedSchedulePerChange;

                scenOutput.unitPricePerChange = unitPricePerChange;
                scenOutput.unitCostPerChange = unitCostPerChange;
                scenOutput.marginalTax = marginalTax;
                scenOutput.capexDepPerChange = capexDepPerChange;
                scenOutput.volChangePerc = volChangePerc;
                scenOutput.rDPerChange = rDPerChange;
                scenOutput.sgAPerChange = sgAPerChange;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return objectDictionary;
        }

        private static bool StrMatch(List<string> modelStr, string matchStr, out string outputStr)
        {
            foreach (string str in modelStr)
            {
                if (Regex.IsMatch(str, matchStr, RegexOptions.IgnoreCase))
                {
                    outputStr = str;
                    return true;
                }
            }

            outputStr = "";
            return false;
        }

        private Dictionary<string, object> CheckForNull(Dictionary<string, object> model)
        {
            int i = 0;
            Dictionary<string, object> tempDict = new Dictionary<string, object>();
            tempDict = model;
            foreach (var temp in model.Where(kvp => kvp.Value.ToString() == string.Empty).ToList())
            {
                model.Remove(temp.Key);
            }
            return model;
        }
        private ProjectsViewModel ConvertScenarioValuesIntoProjectVM(Dictionary<string, object> changeValues, ProjectsViewModel projectList)
        {
            ProjectsViewModel SceinarioProject = new ProjectsViewModel();
            try
            {
                double inputIndustryComparable = 0;
                List<double> inputVolumes = (List<double>)changeValues["Volume"];
                List<double> inputUnitPrice = (List<double>)changeValues["UnitPrice"];
                List<double> inputUnitCost = (List<double>)changeValues["UnitCost"];
                List<double> inputCapex = (List<double>)changeValues["Capex"];
                List<double> inputDepreciation = (List<double>)changeValues["TotalDepreciation"];
                List<double> inputNWC = (List<double>)changeValues["NWC"];
                List<double> inputrD = (List<double>)changeValues["RD"];
                List<double> inputsgA = (List<double>)changeValues["SgA"];
                List<double> inputFixedSchedule = (List<double>)changeValues["Fixed Schedule"]; //Valuation Technique 6
                double inputWACC = Convert.ToDouble(changeValues["WACC"]);
                double inputDVRatio = Convert.ToDouble(changeValues["D/V Ratio"]);
                double inputUnleveredCost = Convert.ToDouble(changeValues["Unlevered cost of capital"]);
                double inputCostOfEquity = Convert.ToDouble(changeValues["Cost of Equity"]);
                double inputCostOfDebt = Convert.ToDouble(changeValues["Cost of Debt"]);
                double inputIntrestCoverage = Convert.ToDouble(changeValues["Interest coverage ratio"]);
                double inputProjectLeverage = Convert.ToDouble(changeValues["Project's Target Leverage"]);
                double inputProjectDebtCOC = Convert.ToDouble(changeValues["Project's Debt Cost of Capital"]);

                double inputMarginalTax = Convert.ToDouble(changeValues["MarginalTax"]);

                if (projectList.ProjectInputDatasVM != null)
                {
                    foreach (var data in projectList.ProjectInputDatasVM)
                    {
                        if (data.LineItem.Contains("Volume"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                inpValue.Value = inputVolumes[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("Unit Cost"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                inpValue.Value = inputUnitCost[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("Unit Price"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                inpValue.Value = inputUnitPrice[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("Marginal Tax rate"))
                        {
                            data.Value = inputMarginalTax;
                        }
                        else if (data.LineItem.Contains("Capex"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                if (inpValue.Value != null)
                                    inpValue.Value = inputCapex[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("Depreciation"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                if (inpValue.Value != null)
                                    inpValue.Value = inputDepreciation[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("R&D"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                if (inpValue.Value != null)
                                    inpValue.Value = inputrD[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("SG&A"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                if (inpValue.Value != null)
                                    inpValue.Value = inputsgA[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("NWC"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                if (inpValue.Value != null)
                                    inpValue.Value = inputNWC[i];
                                i++;
                            }
                        }
                        else if (data.LineItem.Contains("Fixed Schedule"))
                        {
                            int i = 0;
                            foreach (var inpValue in data.ProjectInputValuesVM)
                            {
                                if (inpValue.Value != null)
                                    data.Value = inputFixedSchedule[i];
                                i++;
                            }

                        }
                        else if (data.LineItem.Contains("WACC") && (projectList.ValuationTechniqueId == 1 || projectList.ValuationTechniqueId == 3))
                        {
                            data.Value = inputWACC;
                        } //Valution Technique 1
                        else if (data.LineItem.Contains("D/V Ratio") && (projectList.ValuationTechniqueId == 2 || projectList.ValuationTechniqueId == 3 || projectList.ValuationTechniqueId == 7))
                        {
                            data.Value = inputDVRatio;
                        }
                        else if (data.LineItem.Contains("Unlevered cost of capital") && (projectList.ValuationTechniqueId == 2 || projectList.ValuationTechniqueId == 5 || projectList.ValuationTechniqueId == 6 || projectList.ValuationTechniqueId == 7 || projectList.ValuationTechniqueId == 8))
                        {
                            data.Value = inputUnleveredCost;
                        }
                        else if (data.LineItem.Contains("Cost of Equity") && (projectList.ValuationTechniqueId == 3))
                        {
                            data.Value = inputCostOfEquity;
                        }
                        else if (data.LineItem.Contains("Cost of Debt") && (projectList.ValuationTechniqueId == 2 || projectList.ValuationTechniqueId == 3 || projectList.ValuationTechniqueId == 5 || projectList.ValuationTechniqueId == 6 || projectList.ValuationTechniqueId == 7 || projectList.ValuationTechniqueId == 8))
                        {
                            data.Value = inputCostOfDebt;
                        }
                        else if (data.LineItem.Contains("Intetrest coverage ratio =Interest expense/FCF") && (projectList.ValuationTechniqueId == 5 || projectList.ValuationTechniqueId == 8))
                        {
                            data.Value = inputIntrestCoverage;
                        }
                        //else if (data.LineItem.Contains("Project's Target Leverage") && data.SubHeader=="" && (projectList.ValuationTechniqueId == 4))
                        else if (data.LineItem.Contains("Project's Target Leverage (D/V) Ratio") && data.SubHeader == "" && (projectList.ValuationTechniqueId == 4))
                        {
                            data.Value = inputProjectLeverage;
                        }
                        //else if (data.LineItem.Contains("Project's Debt Cost of Capital")   && data.SubHeader == "" && (projectList.ValuationTechniqueId == 4))
                        else if (data.LineItem.Contains("Project's Debt Cost of Capital (rD)") && data.SubHeader == "" && (projectList.ValuationTechniqueId == 4))
                        {
                            data.Value = inputProjectDebtCOC;
                        }
                        // else if (data.LineItem.Contains("Industry or Comparables Unlevered Cost of Capital") && data.SubHeader == "" && (projectList.ValuationTechniqueId == 4))
                        else if (data.LineItem.Contains("Industry or Comparables Unlevered Cost of Capital (rU-Comp)") && data.SubHeader == "" && (projectList.ValuationTechniqueId == 4))
                        {
                            inputIndustryComparable = Convert.ToDouble(data.Value);
                        }
                        //else if (data.LineItem.Contains("Project's Equity Cost of Capital") && data.SubHeader == "" && (projectList.ValuationTechniqueId == 4))
                        else if (data.LineItem.Contains("Project's Equity Cost of Capital (rE)") && data.SubHeader == "" && (projectList.ValuationTechniqueId == 4))
                        {
                            //C148+C149/(1-C149)*(C148-C150)
                            // data.Value = inputIndustryComparable + inputProjectLeverage / (1 - inputProjectLeverage) * (inputIndustryComparable - inputProjectDebtCOC);
                            data.Value = ((inputIndustryComparable / 100) + (inputProjectLeverage / 100) / (1 - (inputProjectLeverage / 100)) * ((inputIndustryComparable / 100) - (inputProjectDebtCOC / 100))) * 100;
                        }
                        //else if (data.LineItem.Contains("Project's rWACC") && (projectList.ValuationTechniqueId == 4))
                        else if (data.LineItem.Contains("WACC") && (projectList.ValuationTechniqueId == 4))
                        {
                            //C148-C149*C15*C150
                            // data.Value = inputIndustryComparable - inputProjectLeverage * inputMarginalTax * inputProjectDebtCOC;
                            data.Value = ((inputIndustryComparable / 100) - (inputProjectLeverage / 100) * (inputMarginalTax / 100) * (inputProjectDebtCOC / 100)) * 100;
                        }
                        //else if (data.LineItem.Contains("Project's rWACC") && (projectList.ValuationTechniqueId == 7))
                        else if (data.LineItem.Contains("WACC") && (projectList.ValuationTechniqueId == 7))
                        {
                            //C280-C279*$C$15*C281*(1+C280)/(1+C281)
                            // data.Value = inputUnleveredCost - inputDVRatio * inputMarginalTax * inputCostOfDebt * (1 + inputUnleveredCost) / (1 + inputCostOfDebt);
                            data.Value = ((inputUnleveredCost / 100) - (inputDVRatio / 100) * (inputMarginalTax / 100) * (inputCostOfDebt / 100) * (1 + (inputUnleveredCost / 100)) / (1 + (inputCostOfDebt / 100))) * 100;
                        }

                    }

                    SceinarioProject = projectList;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return SceinarioProject;
        }


        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("ScenarioSnapshot")]
        public ActionResult<Object> AddScenarioSnapshot([FromBody] Snapshot_ProjectViewModel model)
        {
            try
            {
                Snapshot_CapitalBudgeting snapshot = new Snapshot_CapitalBudgeting
                {

                    Snapshot = model.Snapshot.ToString(),
                    Description = model.Description,
                    SnapshotType = SCENARIOSNAPSHOT,
                    ProjectInputs = model.ProjectInputs,

                    UserId = model.UserId,
                    ProjectId = model.ProjectId,
                    ValuationTechniqueId = Convert.ToInt32(model.ValuationTechniqueId),
                    CreatedAt = (DateTime)model.CreatedAt, //DateTime.Now,
                    Active = true,
                    NPV = model.NPV,
                    CNPV = model.CNPV
                };
                iSnapshot_CapitalBudgeting.Add(snapshot);
                iSnapshot_CapitalBudgeting.Commit();

                return Ok(new { result = snapshot.Id, message = "Succesfully added Snapshots", code = 200 });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("ScenarioSnapshots/{ProjectId}")]
        public ActionResult<Object> GetScenarioSnapshots(long ProjectId)
        {
            try
            {
                var snapshot = iSnapshot_CapitalBudgeting.FindBy(s => s.SnapshotType == SCENARIOSNAPSHOT && s.ProjectId == ProjectId);
                if (snapshot == null)
                {
                    return NotFound("no snapshot by this ProjectId");
                }
                return Ok(snapshot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("ScenarioSnapshot/{Id}")]
        public ActionResult<Object> GetScenarioSnapshot(long Id)
        {
            try
            {
                var snapshot = iSnapshot_CapitalBudgeting.FindBy(s => s.SnapshotType == SCENARIOSNAPSHOT && s.Id == Id);
                if (snapshot == null)
                {
                    return NotFound("no snapshot by this Id");
                }
                return Ok(snapshot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("GetScenarioInputs/{ProjectId}")]
        public ActionResult<object> GetScenarioInputs(int ProjectId)
        {
            ScenarioInputDatasViewModel result = new ScenarioInputDatasViewModel();
            ScenarioInputDatas scenarioInput = iScenarioInputDatas.GetSingle(x => x.ProjectId == ProjectId);
            if (scenarioInput != null)
            {
                // ProjectsViewModel result = new ProjectsViewModel();
                List<ScenarioInputDatas> scenarioInputList = new List<ScenarioInputDatas>();
                scenarioInputList = iScenarioInputDatas.FindBy(x => x.Id != 0 && x.ProjectId == ProjectId).ToList();
                if (scenarioInputList != null && scenarioInputList.Count > 0)
                {
                   // List<ScenarioInputDatasViewModel> scenarioInputDatasVM = new List<ScenarioInputDatasViewModel>();
                   // ScenarioInputDatasViewModel DatasVm = new ScenarioInputDatasViewModel();
                    List<ScenarioInputValues> scenarioInputValueList = iScenarioInputValues.FindBy(x => x.Id != 0 && scenarioInputList.Any(t => t.Id == x.ScenarioInputDatasId)).ToList();

                    return scenarioInputList;

                    //   List<List<double>> completeResult = new List<List<double>>(); //This combine both range and npvlist                
                    // completeResult.Add(scenarioInputValueList);

                    // List<ScenarioInputValuesViewModel> ValuesViewModelList = new List<ScenarioInputValuesViewModel>();

                    //foreach (ScenarioInputDatas datas in scenarioInputList)
                    //{
                    //    List<ScenarioInputValues> ScenarioInputValuesList = scenarioInputValueList.FindAll(x => x.Id != 0 && x.ScenarioInputDatasId == datas.Id).ToList();
                    //    ScenarioInputValuesViewModel ScenarioInputVM = null;

                    //    DatasVm = mapper.Map<ScenarioInputDatas, ScenarioInputDatasViewModel>(datas);

                    //    if (ScenarioInputValuesList != null && ScenarioInputValuesList.Count > 0)
                    //    {
                    //        DatasVm.ScenarioInputValuesVM = new List<ScenarioInputValuesViewModel>();
                    //        foreach (var scenario in ScenarioInputValuesList)
                    //        {
                    //            ScenarioInputVM = mapper.Map<ScenarioInputValues, ScenarioInputValuesViewModel>(scenario);
                    //            DatasVm.ScenarioInputValuesVM.Add(ScenarioInputVM);
                    //        }
                    //    }

                    //   // DatasVm = mapper.Map<ScenarioInputDatas, ScenarioInputDatasViewModel>(datas);

                    //    if (DatasVm != null && DatasVm.ScenarioInputValuesVM != null && DatasVm.ScenarioInputValuesVM.Count > 0)
                    //    {
                    //        //DatasVm = new ProjectInputDatasViewModel();
                    //        var valuesList = DatasVm.ScenarioInputValuesVM;
                    //        DatasVm.ScenarioInputValuesVM = new List<ScenarioInputValuesViewModel>();

                    //        //DatasVm = mapper.Map<ProjectInputDatas, ProjectInputDatasViewModel>(datas);
                    //        //if(DatasVm!=null && DatasVm.ProjectInputValuesVM!=null && DatasVm.ProjectInputValuesVM.Count>0)
                    //        //{
                    //        //var valuesList = DatasVm.ProjectInputValuesVM;
                    //        //DatasVm.ProjectInputValuesVM = new List<ProjectInputValuesViewModel>();
                    //        //DatasVm.ProjectInputValuesVM=valuesList.OrderBy(x=>x.Year).ToList();
                    //    }
                    //    scenarioInputDatasVM.Add(DatasVm);
                    //}
                }
            }
            return result;
        }

        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("SensiScenarioVariable/{ProjectId}")]
        public ActionResult<object> SensiScenarioVariable(int ProjectId)
        {         
            ScenarioAnalysis result = new ScenarioAnalysis();
            Dictionary<string, List<List<double>>> resultDict = new Dictionary<string, List<List<double>>>();
            List<string> variable = new List<string>();
            try
            {
                #region Variables
                Project project = iProject.GetSingle(x => x.Id == ProjectId);
                if (project != null)
                {
                    var inputDatas = iProjectInputDatas.FindBy(x => x.ProjectId == ProjectId).ToList();
                    //Todo Will Change after Sensitivity will complete
                    List<double> resultList = new List<double>();
                    List<double> npvResultList = new List<double>();
                    List<List<double>> completeResult = new List<List<double>>(); //This combine both range and npvlist
                    resultList = Calculate(500);
                    npvResultList = Calculate(500);
                    completeResult.Add(resultList);
                    completeResult.Add(npvResultList);

                    foreach (var strData in inputDatas)
                    {
                        if (strData.SubHeader == "" || strData.SubHeader == "Revenue & Variable Cost Tier1" || strData.SubHeader == "Capex & Depreciation1" || strData.SubHeader == "Other Fixed Cost" || strData.SubHeader == "Working Capital" || strData.SubHeader == "Fixed Schedule")
                        {
                            if (!((strData.SubHeader == "Comparable" || strData.LineItem == "comparable") || strData.LineItem == "Depreciation" || strData.LineItem == ""))
                            {
                                resultDict.Add(strData.LineItem + FindUnit(Convert.ToInt32(strData.ValueTypeId), Convert.ToInt32(strData.UnitId)), completeResult);
                            }
                        }
                    }

                    if (project.ValuationTechniqueId == 1)
                    {
                        resultDict.Remove("D/V Ratio(%)");
                        resultDict.Remove("Unlevered cost of capital(%)");
                        resultDict.Remove("Cost of Debt");
                        resultDict.Remove("Cost of Equity");
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital(%)");
                        resultDict.Remove("Project's Target Leverage(%)");
                        resultDict.Remove("Project's Debt Cost of Capital(%)");
                        resultDict.Remove("Project's rWACC(%)");
                        resultDict.Remove("Intetrest coverage ratio =Interest expense/FCF(%)");
                    }
                    else if (project.ValuationTechniqueId == 2)
                    {
                        resultDict.Remove("WACC(%)");
                        resultDict.Remove("Cost of Equity");
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital(%)");
                        resultDict.Remove("Project's Target Leverage(%)");
                        resultDict.Remove("Project's Debt Cost of Capital(%)");
                        resultDict.Remove("Project's rWACC(%)");
                        resultDict.Remove("Intetrest coverage ratio =Interest expense/FCF(%)");
                    }
                    else if (project.ValuationTechniqueId == 3)
                    {
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital(%)");
                        resultDict.Remove("Project's Target Leverage(%)");
                        resultDict.Remove("Project's Debt Cost of Capital(%)");
                        resultDict.Remove("Project's rWACC(%)");
                        resultDict.Remove("Intetrest coverage ratio =Interest expense/FCF(%)");
                    }
                    else if (project.ValuationTechniqueId == 4)//For Valuation Technique 4 resultDict.Remove "Industry or Comparables Unlevered Cost of Capital (rU-Comp)", Cost of Equity,and "Project's rWACC"
                    {
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital (rU-Comp)(%)");
                        resultDict.Remove("Project's Equity Cost of Capital (rE)(%)");
                        resultDict.Remove("WACC(%)");

                        resultDict.Remove("D/V Ratio(%)");
                        resultDict.Remove("Unlevered cost of capital(%)");
                        resultDict.Remove("Cost of Debt");
                        resultDict.Remove("Cost of Equity");
                        resultDict.Remove("Project's rWACC(%)");
                        resultDict.Remove("Intetrest coverage ratio =Interest expense/FCF(%)");
                    }
                    else if (project.ValuationTechniqueId == 5)
                    {
                        resultDict.Remove("WACC(%)");
                        resultDict.Remove("D/V Ratio(%)");
                        resultDict.Remove("Cost of Equity");
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital(%)");
                        resultDict.Remove("Project's Target Leverage(%)");
                        resultDict.Remove("Project's Debt Cost of Capital(%)");
                        resultDict.Remove("Project's rWACC(%)");
                    }
                    else if (project.ValuationTechniqueId == 6)
                    {
                        resultDict.Remove("WACC(%)");
                        resultDict.Remove("D/V Ratio(%)");
                        resultDict.Remove("Cost of Equity");
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital(%)");
                        resultDict.Remove("Project's Target Leverage(%)");
                        resultDict.Remove("Project's Debt Cost of Capital(%)");
                        resultDict.Remove("Project's rWACC(%)");
                        resultDict.Remove("Intetrest coverage ratio =Interest expense/FCF(%)");
                    }
                    else if (project.ValuationTechniqueId == 7) //For Valuation Technique 7 resultDict.Remove "WACC"
                    {
                        resultDict.Remove("WACC(%)");

                        resultDict.Remove("Cost of Equity");
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital(%)");
                        resultDict.Remove("Project's Target Leverage(%)");
                        resultDict.Remove("Project's Debt Cost of Capital(%)");
                        resultDict.Remove("Project's rWACC(%)");
                        resultDict.Remove("Intetrest coverage ratio =Interest expense/FCF(%)");
                    }
                    else if (project.ValuationTechniqueId == 8)
                    {
                        resultDict.Remove("WACC(%)");
                        resultDict.Remove("D/V Ratio(%)");
                        resultDict.Remove("Cost of Equity");
                        resultDict.Remove("Industry or Comparables Unlevered Cost of Capital(%)");
                        resultDict.Remove("Project's Target Leverage(%)");
                        resultDict.Remove("Project's Debt Cost of Capital(%)");
                        resultDict.Remove("Project's rWACC(%)");
                    }

                    resultDict.Remove("Project's Equity Cost of Capital (rE)(%)");
                    result.variables = resultDict;
                }
                #endregion
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
            return Ok(resultDict);
        }

        private static List<double> Calculate(double value)
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

        private string FindUnit(int ValueTypeId, int UnitId)
        {
            string strUnit = "";
            if (ValueTypeId == 1)
            {
                strUnit = "(%)";
            }
            else if (ValueTypeId == 2)
            {
                List<ValueTextWrapper> unitList = EnumHelper.GetEnumDescriptions<NumberCountEnum>();
                strUnit = "(" + unitList[--UnitId].text.ToString() + ")";
            }
            else if (ValueTypeId == 3)
            {
                List<ValueTextWrapper> unitList = EnumHelper.GetEnumDescriptions<CurrencyValueEnum>();
                strUnit = "(" + unitList[--UnitId].text.ToString() + ")";
            }
            else
            {
                strUnit = "";
            }

            return strUnit;
        }

        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("OldProjectSensitivity/{ProjectId}/{Lineitem}/{difference}")]
        public ActionResult<Object> OldProjectSensitivity([FromBody] List<Dictionary<string, object>> body, long ProjectId, string Lineitem, int difference)
        {
            List<Dictionary<string, List<List<double>>>> npvOutput = new List<Dictionary<string, List<List<double>>>>();
            try
            {
                ProjectsViewModel input = new ProjectsViewModel();
                if (ProjectId != 0)
                {
                    Project project = iProject.GetSingle(x => x.Id == ProjectId);
                    if (project != null)
                    {
                        input = ProjectDetails(ProjectId, project);

                        // find Mid Value
                        //double midvalue = 0;
                        //if (input != null && input.ProjectInputDatasVM != null && input.ProjectInputDatasVM.Count > 0)
                        //{
                        //    var NPVOutput = input.ProjectInputDatasVM.Find(x => x.LineItem.Contains(Lineitem));
                        //    midvalue = Convert.ToDouble(NPVOutput.Value);
                        //}

                        for (int i = 0; i < body.Count; i++)
                        {
                            Dictionary<string, List<List<double>>> npvList = new Dictionary<string, List<List<double>>>();
                            npvList.TryAdd(Lineitem, getNPVforSensitivity(body[i], input, Lineitem, difference, true, (int)ProjectId));
                            npvOutput.Add(npvList);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                return (BadRequest(new { message = ex.Message.ToString(), code = 400 }));
            }
            return npvOutput;
        }
     
        private List<List<string>> getNPVforSensitivityInputNew(Dictionary<string, object> model, ProjectsViewModel projectVM, string LineItem, double difference, bool isdefault, long projectId)
        {
            Project_ScenarioOutput scenOutput = new Project_ScenarioOutput();
            List<List<string>> Result = new List<List<string>>();
            List<string> npvOutput = new List<string>();
            List<double> fractionValue = new List<double>();
            double npv = 0;
            try
            {
                var objectDictionary = new Dictionary<string, object>();
                ProjectsViewModel inputProject = new ProjectsViewModel();
                model = CheckForNull(model);
                //find value by name
                if (LineItem.Contains('('))
                {
                    int index = LineItem.IndexOf('(');
                    LineItem = LineItem.Substring(0, index);
                }
                double DictValue = 0;
                //DictValue = objectDictionary.TryGetValue(LineItem, out DictValue);
                if (projectVM != null && projectVM.ProjectInputDatasVM != null && projectVM.ProjectInputDatasVM.Count > 0)
                {
                    ProjectInputDatasViewModel Datas = new ProjectInputDatasViewModel();
                    //find linitem from input
                    if (LineItem.ToLower().Contains("project's target leverage") || LineItem.ToLower().Contains("project's debt cost of capital"))
                    {
                        Datas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains(LineItem.ToLower()) && x.SubHeader == "");
                    }
                    else
                        Datas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains(LineItem.ToLower()));
                    if (Datas != null)
                    {
                        if (Datas.HasMultiYear == true && Datas.ProjectInputValuesVM != null && Datas.ProjectInputValuesVM.Count > 0)
                        {
                            double Average = 0;
                            double total = Convert.ToDouble(Datas.ProjectInputValuesVM.Sum(x => x.Value));
                            //for total
                            //if (LineItem.ToLower().Contains("volume") || LineItem.ToLower().Contains("capex") || LineItem.ToLower().Contains("depreciation") || LineItem.ToLower().Contains("nwc"))
                            if (LineItem.ToLower().Contains("volume") || LineItem.ToLower().Contains("capex") || LineItem.ToLower().Contains("depreciation"))
                            {
                                DictValue = total ;
                            }
                            else if (LineItem.ToLower().Contains("nwc"))
                            {
                                DictValue = Convert.ToDouble(Datas.ProjectInputValuesVM.Average(x => x.Value));
                            }
                            else if (LineItem.ToLower().Contains("r&d") || LineItem.ToLower().Contains("sg&a") || LineItem.ToLower().Contains("fixed schedule"))
                            {
                                Average = projectVM.NoOfYears != null && projectVM.NoOfYears != 0 ? Convert.ToDouble(total / projectVM.NoOfYears) : 0;
                                DictValue = Average;
                            }
                            else if (LineItem.ToLower().Contains("unit price") || LineItem.ToLower().Contains("unit cost"))
                            {
                                // find Volume
                                var VolumeDatas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains("volume"));

                                double totalVolume = VolumeDatas.ProjectInputValuesVM != null && VolumeDatas.ProjectInputValuesVM.Count > 0 ? Convert.ToDouble(VolumeDatas.ProjectInputValuesVM.Sum(x => x.Value)) : 0;
                                double totalBasicVolume = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(VolumeDatas.UnitId, totalVolume) : 0;

                                if (LineItem.ToLower().Contains("unit price"))
                                {
                                    //find sum Revenue
                                    var Revenue = projectVM.ProjectSummaryDatasVM != null && projectVM.ProjectSummaryDatasVM.Count > 0 ? projectVM.ProjectSummaryDatasVM.Find(x => x.LineItem.ToLower() == "sales") : null;
                                    double Revenuetotal = Revenue.ProjectOutputValuesVM != null && Revenue.ProjectOutputValuesVM.Count > 0 ? Convert.ToDouble(Revenue.ProjectOutputValuesVM.Sum(x => x.BasicValue)) : 0;
                                    double temp = Revenuetotal / totalBasicVolume;
                                    Average = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(Datas.UnitId, temp) : 0;
                                }
                                else
                                {
                                    //find Sum COGS
                                    var COGSDatas = projectVM.ProjectSummaryDatasVM != null && projectVM.ProjectSummaryDatasVM.Count > 0 ? projectVM.ProjectSummaryDatasVM.Find(x => x.LineItem.ToLower() == "cogs") : null;
                                    double Revenuetotal = COGSDatas.ProjectOutputValuesVM != null && COGSDatas.ProjectOutputValuesVM.Count > 0 ? Convert.ToDouble(COGSDatas.ProjectOutputValuesVM.Sum(x => x.BasicValue)) : 0;
                                    double temp = Revenuetotal / totalBasicVolume;
                                    Average = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(Datas.UnitId, temp) : 0;
                                }
                                DictValue = Average ;
                            }
                            // add to model dictionary
                            //model.Add(LineItem, DictValue);
                        }
                        else
                        {
                            DictValue = Convert.ToDouble(Datas.Value);
                            // model.Add(LineItem, Datas.Value);
                        }
                    }
                }

                double differenceValue = 0;
                List<string> fractionList = FractionValuesNew(DictValue, difference, isdefault);
                List<double> fractionListNew = FractionValues(DictValue, difference, isdefault);

                differenceValue = fractionListNew[1] - fractionListNew[0];

                // List<double> fractionList = FractionValues(DictValue, out differenceValue , difference, isdefault);
                foreach (var Fraction in fractionList)
                {

                    if (Fraction == "0")
                    {
                        model = new Dictionary<string, object>();
                        model.Add(LineItem, Fraction);
                        var TempProjectVm = projectVM;
                        objectDictionary = new Dictionary<string, object>();
                        objectDictionary = ScenarioValues(model, TempProjectVm, out scenOutput);
                        ProjectsViewModel ScenarioProject = ConvertScenarioValuesIntoProjectVM(objectDictionary, TempProjectVm);
                        List<ProjectOutputDatasViewModel> summaryOutput = GetProjectOutput(TempProjectVm);
                        // find NPV
                        var NPVOutput = summaryOutput.Find(x => x.LineItem.Contains("Net Present Value"));
                        npv = Convert.ToDouble(NPVOutput.Value);
                        npvOutput.Add(npv.ToString());

                        Project project = iProject.GetSingle(x => x.Id == projectId);
                        if (project != null)
                        {
                            projectVM = ProjectDetails(projectId, project);
                        }
                    }
                    else
                    {
                        model = new Dictionary<string, object>();
                        //replace actual value to fraction value in model
                        //if(LineItem.Contains(""))
                        model.Add(LineItem, Fraction);
                        var TempProjectVm = projectVM;
                        objectDictionary = new Dictionary<string, object>();
                        objectDictionary = ScenarioValues(model, TempProjectVm, out scenOutput);
                        ProjectsViewModel ScenarioProject = ConvertScenarioValuesIntoProjectVM(objectDictionary, TempProjectVm);
                        List<ProjectOutputDatasViewModel> summaryOutput = GetProjectOutput(ScenarioProject);
                        // find NPV
                        var NPVOutput = summaryOutput.Find(x => x.LineItem.Contains("Net Present Value"));
                        npv = Convert.ToDouble(NPVOutput.Value);
                        //npv = NPVOutput.Value;
                        npvOutput.Add(npv.ToString());
                    }                   
                }
                Result.Add(fractionList);
                Result.Add(npvOutput);

                //for Actual Values
                List<string> actualList = new List<string>();
                actualList.Add(DictValue.ToString());
                actualList.Add(DictValue.ToString());
                actualList.Add(differenceValue.ToString());
                Result.Add(actualList);

                List<string> saved = new List<string>();
                saved.Add("0");
                Result.Add(saved);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Result;
        }

        private List<List<double>> getNPVforSensitivity(Dictionary<string, object> model, ProjectsViewModel projectVM, string LineItem, double difference, bool isdefault,long projectId)
        {
            Project_ScenarioOutput scenOutput = new Project_ScenarioOutput();
            List<List<double>> Result = new List<List<double>>();
            List<double> npvOutput = new List<double>();
            List<double> fractionValue = new List<double>();
            double npv = 0;
            try
            {
                var objectDictionary = new Dictionary<string, object>();
                ProjectsViewModel inputProject = new ProjectsViewModel();
                model = CheckForNull(model);
                //find value by name
                if (LineItem.Contains('('))
                {
                    int index = LineItem.IndexOf('(');
                    LineItem = LineItem.Substring(0, index);
                }
                double DictValue = 0;
                //DictValue = objectDictionary.TryGetValue(LineItem, out DictValue);
                if (projectVM != null && projectVM.ProjectInputDatasVM != null && projectVM.ProjectInputDatasVM.Count > 0)
                {
                    ProjectInputDatasViewModel Datas = new ProjectInputDatasViewModel();
                    //find linitem from input
                    if (LineItem.ToLower().Contains("project's target leverage")|| LineItem.ToLower().Contains("project's debt cost of capital") )
                    {
                        Datas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains(LineItem.ToLower()) && x.SubHeader=="");
                    }
                    else
                     Datas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains(LineItem.ToLower()) );
                    if (Datas != null)
                    {
                        if (Datas.HasMultiYear == true && Datas.ProjectInputValuesVM != null && Datas.ProjectInputValuesVM.Count > 0)
                        {
                            double Average = 0;
                            double total = Convert.ToDouble(Datas.ProjectInputValuesVM.Sum(x => x.Value));
                            //for total
                            //if (LineItem.ToLower().Contains("volume") || LineItem.ToLower().Contains("capex") || LineItem.ToLower().Contains("depreciation") || LineItem.ToLower().Contains("nwc"))
                            if (LineItem.ToLower().Contains("volume") || LineItem.ToLower().Contains("capex") || LineItem.ToLower().Contains("depreciation"))
                            {
                                DictValue = total;
                            }
                            else if (LineItem.ToLower().Contains("nwc"))
                            {
                                DictValue = Convert.ToDouble(Datas.ProjectInputValuesVM.Average(x => x.Value));
                            }
                            else if (LineItem.ToLower().Contains("r&d") || LineItem.ToLower().Contains("sg&a") || LineItem.ToLower().Contains("fixed schedule"))
                            {
                                Average = projectVM.NoOfYears != null && projectVM.NoOfYears != 0 ? Convert.ToDouble(total / projectVM.NoOfYears) : 0;
                                DictValue = Average;
                            }
                            else if (LineItem.ToLower().Contains("unit price") || LineItem.ToLower().Contains("unit cost"))
                            {
                                // find Volume
                                var VolumeDatas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains("volume"));

                                double totalVolume = VolumeDatas.ProjectInputValuesVM != null && VolumeDatas.ProjectInputValuesVM.Count > 0 ? Convert.ToDouble(VolumeDatas.ProjectInputValuesVM.Sum(x => x.Value)) : 0;
                                double totalBasicVolume = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(VolumeDatas.UnitId, totalVolume) : 0;

                                if (LineItem.ToLower().Contains("unit price"))
                                {
                                    //find sum Revenue
                                    var Revenue = projectVM.ProjectSummaryDatasVM != null && projectVM.ProjectSummaryDatasVM.Count > 0 ? projectVM.ProjectSummaryDatasVM.Find(x => x.LineItem.ToLower() == "sales") : null;
                                    double Revenuetotal = Revenue.ProjectOutputValuesVM != null && Revenue.ProjectOutputValuesVM.Count > 0 ? Convert.ToDouble(Revenue.ProjectOutputValuesVM.Sum(x => x.BasicValue)) : 0;
                                    double temp = Revenuetotal / totalBasicVolume;
                                    Average = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(Datas.UnitId, temp) : 0;
                                }
                                else
                                {
                                    //find Sum COGS
                                    var COGSDatas = projectVM.ProjectSummaryDatasVM != null && projectVM.ProjectSummaryDatasVM.Count > 0 ? projectVM.ProjectSummaryDatasVM.Find(x => x.LineItem.ToLower() == "cogs") : null;
                                    double Revenuetotal = COGSDatas.ProjectOutputValuesVM != null && COGSDatas.ProjectOutputValuesVM.Count > 0 ? Convert.ToDouble(COGSDatas.ProjectOutputValuesVM.Sum(x => x.BasicValue)) : 0;
                                    double temp = Revenuetotal / totalBasicVolume;
                                    Average = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(Datas.UnitId, temp) : 0;
                                }
                                DictValue = Average;
                            }
                            // add to model dictionary
                            //model.Add(LineItem, DictValue);
                        }
                        else
                        {
                            DictValue = Convert.ToDouble(Datas.Value);
                            // model.Add(LineItem, Datas.Value);
                        }
                    }
                }
                double differenceValue;
                List<double> fractionList = FractionValues(DictValue, difference, isdefault);               
                // List<double> fractionList = FractionValues(DictValue, out differenceValue , difference, isdefault);

                differenceValue = fractionList[1] - fractionList[0];

                foreach (var Fraction in fractionList)
                {

                    if (Fraction == 0)
                    {
                        model = new Dictionary<string, object>();
                        model.Add(LineItem, Fraction);
                        var TempProjectVm = projectVM;
                        objectDictionary = new Dictionary<string, object>();
                        objectDictionary = ScenarioValues(model, TempProjectVm, out scenOutput);
                        ProjectsViewModel ScenarioProject = ConvertScenarioValuesIntoProjectVM(objectDictionary, TempProjectVm);
                        List<ProjectOutputDatasViewModel> summaryOutput = GetProjectOutput(TempProjectVm);
                        // find NPV
                        var NPVOutput = summaryOutput.Find(x => x.LineItem.Contains("Net Present Value"));
                        npv = Convert.ToDouble(NPVOutput.Value);
                        npvOutput.Add(npv);

                        Project project = iProject.GetSingle(x => x.Id == projectId);
                        if (project != null)
                        {
                            projectVM = ProjectDetails(projectId, project);
                        }
                    }
                    else
                    {
                        model = new Dictionary<string, object>();
                        //replace actual value to fraction value in model
                        //if(LineItem.Contains(""))
                        model.Add(LineItem, Fraction);
                        var TempProjectVm = projectVM;
                        objectDictionary = new Dictionary<string, object>();
                        objectDictionary = ScenarioValues(model, TempProjectVm, out scenOutput);
                        ProjectsViewModel ScenarioProject = ConvertScenarioValuesIntoProjectVM(objectDictionary, TempProjectVm);
                        List<ProjectOutputDatasViewModel> summaryOutput = GetProjectOutput(ScenarioProject);

                        // find NPV
                        var NPVOutput = summaryOutput.Find(x => x.LineItem.Contains("Net Present Value"));
                        npv = Convert.ToDouble(NPVOutput.Value);
                        npvOutput.Add(npv);
                    }                   
                }
                Result.Add(fractionList);
                Result.Add(npvOutput);

                //for Actual Values
                List<Double> actualList = new List<double>();
                actualList.Add(DictValue);
                actualList.Add(DictValue);
                actualList.Add(differenceValue);
                Result.Add(actualList);

                List<double> saved = new List<double>();
                saved.Add(0);
                Result.Add(saved);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Result;
        }

        private List<List<double>> getNPVforSensitivityInputsGraph(Dictionary<string, object> model, ProjectsViewModel projectVM, string LineItem, double difference, bool isdefault, long projectId)
        {
            Project_ScenarioOutput scenOutput = new Project_ScenarioOutput();
            List<List<double>> Result = new List<List<double>>();

              List<double> npvOutput = new List<double>();
            List<double> fractionValue = new List<double>();
            double npv = 0;
            try
            {
                var objectDictionary = new Dictionary<string, object>();
                ProjectsViewModel inputProject = new ProjectsViewModel();
                model = CheckForNull(model);
                //find value by name
                if (LineItem.Contains('('))
                {
                    int index = LineItem.IndexOf('(');
                    LineItem = LineItem.Substring(0, index);
                }
                double DictValue = 0;
                //DictValue = objectDictionary.TryGetValue(LineItem, out DictValue);
                if (projectVM != null && projectVM.ProjectInputDatasVM != null && projectVM.ProjectInputDatasVM.Count > 0)
                {
                    ProjectInputDatasViewModel Datas = new ProjectInputDatasViewModel();
                    //find linitem from input
                    if (LineItem.ToLower().Contains("project's target leverage") || LineItem.ToLower().Contains("project's debt cost of capital"))
                    {
                        Datas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains(LineItem.ToLower()) && x.SubHeader == "");
                    }
                    else
                        Datas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains(LineItem.ToLower()));
                    if (Datas != null)
                    {
                        if (Datas.HasMultiYear == true && Datas.ProjectInputValuesVM != null && Datas.ProjectInputValuesVM.Count > 0)
                        {
                            double Average = 0;
                            double total = Convert.ToDouble(Datas.ProjectInputValuesVM.Sum(x => x.Value));
                            //for total
                            //if (LineItem.ToLower().Contains("volume") || LineItem.ToLower().Contains("capex") || LineItem.ToLower().Contains("depreciation") || LineItem.ToLower().Contains("nwc"))
                            if (LineItem.ToLower().Contains("volume") || LineItem.ToLower().Contains("capex") || LineItem.ToLower().Contains("depreciation"))
                            {
                                DictValue = total;
                            }
                            else if (LineItem.ToLower().Contains("nwc"))
                            {
                                DictValue = Convert.ToDouble(Datas.ProjectInputValuesVM.Average(x => x.Value));
                            }
                            else if (LineItem.ToLower().Contains("r&d") || LineItem.ToLower().Contains("sg&a") || LineItem.ToLower().Contains("fixed schedule"))
                            {
                                Average = projectVM.NoOfYears != null && projectVM.NoOfYears != 0 ? Convert.ToDouble(total / projectVM.NoOfYears) : 0;
                                DictValue = Average;
                            }
                            else if (LineItem.ToLower().Contains("unit price") || LineItem.ToLower().Contains("unit cost"))
                            {
                                // find Volume
                                var VolumeDatas = projectVM.ProjectInputDatasVM.Find(x => x.LineItem.ToLower().Contains("volume"));

                                double totalVolume = VolumeDatas.ProjectInputValuesVM != null && VolumeDatas.ProjectInputValuesVM.Count > 0 ? Convert.ToDouble(VolumeDatas.ProjectInputValuesVM.Sum(x => x.Value)) : 0;
                                double totalBasicVolume = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(VolumeDatas.UnitId, totalVolume) : 0;

                                if (LineItem.ToLower().Contains("unit price"))
                                {
                                    //find sum Revenue
                                    var Revenue = projectVM.ProjectSummaryDatasVM != null && projectVM.ProjectSummaryDatasVM.Count > 0 ? projectVM.ProjectSummaryDatasVM.Find(x => x.LineItem.ToLower() == "sales") : null;
                                    double Revenuetotal = Revenue.ProjectOutputValuesVM != null && Revenue.ProjectOutputValuesVM.Count > 0 ? Convert.ToDouble(Revenue.ProjectOutputValuesVM.Sum(x => x.BasicValue)) : 0;
                                    double temp = Revenuetotal / totalBasicVolume;
                                    Average = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(Datas.UnitId, temp) : 0;
                                }
                                else
                                {
                                    //find Sum COGS
                                    var COGSDatas = projectVM.ProjectSummaryDatasVM != null && projectVM.ProjectSummaryDatasVM.Count > 0 ? projectVM.ProjectSummaryDatasVM.Find(x => x.LineItem.ToLower() == "cogs") : null;
                                    double Revenuetotal = COGSDatas.ProjectOutputValuesVM != null && COGSDatas.ProjectOutputValuesVM.Count > 0 ? Convert.ToDouble(COGSDatas.ProjectOutputValuesVM.Sum(x => x.BasicValue)) : 0;
                                    double temp = Revenuetotal / totalBasicVolume;
                                    Average = VolumeDatas.UnitId != null ? UnitConversion.getBasicValueforNumbers(Datas.UnitId, temp) : 0;
                                }
                                DictValue = Average;
                            }
                            // add to model dictionary
                            //model.Add(LineItem, DictValue);
                        }
                        else
                        {
                            DictValue = Convert.ToDouble(Datas.Value);
                            // model.Add(LineItem, Datas.Value);
                        }
                    }
                }
                double differenceValue;
                List<double> fractionList = FractionValuesForGraphs(DictValue, difference);
                // List<double> fractionList = FractionValues(DictValue, out differenceValue , difference, isdefault);

                differenceValue = fractionList[1] - fractionList[0];

                foreach (var Fraction in fractionList)
                {

                    if(Fraction == 0)
                    {
                        model = new Dictionary<string, object>();
                        //replace actual value to fraction value in model
                        //if(LineItem.Contains(""))
                        model.Add(LineItem, Fraction);
                        var TempProjectVm = projectVM;
                        objectDictionary = new Dictionary<string, object>();
                        objectDictionary = ScenarioValues(model, TempProjectVm, out scenOutput);
                        ProjectsViewModel ScenarioProject = ConvertScenarioValuesIntoProjectVM(objectDictionary, TempProjectVm);                                           
                        List<ProjectOutputDatasViewModel> summaryOutput = GetProjectOutput(TempProjectVm);
                        // find NPV
                        var NPVOutput = summaryOutput.Find(x => x.LineItem.Contains("Net Present Value"));
                        npv = Convert.ToDouble(NPVOutput.Value);
                        npvOutput.Add(npv);

                        Project project = iProject.GetSingle(x => x.Id == projectId);
                        if (project != null)
                        {
                            projectVM = ProjectDetails(projectId, project);  
                        }
                    }
                    else
                    {
                        model = new Dictionary<string, object>();
                        //replace actual value to fraction value in model
                        //if(LineItem.Contains(""))
                        model.Add(LineItem, Fraction);
                        var TempProjectVm = projectVM;
                        objectDictionary = new Dictionary<string, object>();
                        objectDictionary = ScenarioValues(model, TempProjectVm, out scenOutput);
                        ProjectsViewModel ScenarioProject = ConvertScenarioValuesIntoProjectVM(objectDictionary, TempProjectVm);
                        List<ProjectOutputDatasViewModel> summaryOutput = GetProjectOutput(ScenarioProject);
                        // find NPV
                        var NPVOutput = summaryOutput.Find(x => x.LineItem.Contains("Net Present Value"));
                        npv = Convert.ToDouble(NPVOutput.Value);
                        //npvOutput.Add(0);
                        npvOutput.Add(npv);
                    }
                }

               // fractionList = fractionList.OrderByDescending(d => d).ToList();

                Result.Add(fractionList);
                //  Result.Add(npvOutput);
                //npvOutput = new List<double>();
                
                //npvOutput.Add(-1);
                //npvOutput.Add(-2);
                //npvOutput.Add(-3);
                //npvOutput.Add(-4);
                //npvOutput.Add(-5);
                //npvOutput.Add(1);

              //  npvOutput = npvOutput.OrderBy(d => d).ToList();
                Result.Add(npvOutput);

                //for Actual Values
                //List<Double> actualList = new List<double>();
                //actualList.Add(DictValue);
                //actualList.Add(DictValue);
                //actualList.Add(differenceValue);
                //Result.Add(actualList);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Result;
        }


        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("projectSensitivity")]
        public ActionResult<Object> getprojectSensitivity([FromBody] sensitivity sensimodel/* variableList,long ProjectId,string linItem,int difference*/)
        {            
            ProjectsViewModel input = new ProjectsViewModel();
            
            if (sensimodel.ProjectId != 0)
            {
                SensitivityInputData sensitivityInput = iSensitivityInputData.GetSingle(x => x.ProjectId == sensimodel.ProjectId);
                if (sensitivityInput != null && sensimodel.intervalFlag == false)
                {
                    List<Dictionary<string, List<List<string>>>> npvOutput = new List<Dictionary<string, List<List<string>>>>();

                    Project project = iProject.GetSingle(x => x.Id == sensimodel.ProjectId);
                    if (project != null)
                    {
                        input = ProjectDetails(sensimodel.ProjectId, project);
                        Dictionary<string, List<List<string>>> npvList = new Dictionary<string, List<List<string>>>();
                       
                        if (input != null && input.ProjectSummaryDatasVM != null && input.ProjectSummaryDatasVM.Count > 0)
                        {
                            // find NPV
                            var NPVOutput = input.ProjectSummaryDatasVM.Find(x => x.LineItem.Contains("Net Present Value"));
                            List<List<string>> tList = new List<List<string>>();
                            // List<double> tdouble = new List<double>();
                            List<string> tstring = new List<string>();
                            // tdouble.Add(Convert.ToDouble(sensitivityInput.ActualNPV));
                            // tdouble.Add(Convert.ToDouble(sensitivityInput.NewNPV));
                            tstring.Add(sensitivityInput.ActualNPV.ToString());
                            tstring.Add(sensitivityInput.NewNPV.ToString());
                            tList.Add(tstring);
                            //npvList.TryAdd(NPVOutput.LineItem, tList);
                            npvList.TryAdd("Net Present Value" + FindUnit(Convert.ToInt32(NPVOutput.ValueTypeId), Convert.ToInt32(NPVOutput.UnitId)), tList);
                            npvOutput.Add(npvList);
                        }

                        var sensitivityInputDatas = iSensitivityInputData.FindBy(x => x.Id != 0 && x.ProjectId == sensimodel.ProjectId).ToList();
                        if (sensitivityInputDatas != null && sensitivityInputDatas.Count > 0)
                        {
                            //if (!string.IsNullOrEmpty(sensimodel.linItem))
                            //{

                            //    int excecuted = 0;

                            //    foreach (var sensitivityDatasVM in sensitivityInputDatas)
                            //    {
                            //        if (sensitivityDatasVM != null)
                            //        {
                            //            if (sensimodel.linItem == (sensitivityDatasVM.LineItem + "(" + sensitivityDatasVM.LineItemUnit.ToString() + ")"))
                            //            {
                            //                excecuted = 1;

                            //                npvList = new Dictionary<string, List<List<string>>>();
                            //                Dictionary<string, object> model = new Dictionary<string, object>();
                            //                List<List<string>> Result = new List<List<string>>();

                            //                string lineItemIntervals = sensitivityDatasVM.LineItemIntervals.Replace("\r\n", "");
                            //                lineItemIntervals = lineItemIntervals.Replace(" ", "");
                            //                lineItemIntervals = lineItemIntervals.Replace("[", "");
                            //                lineItemIntervals = lineItemIntervals.Replace("]", "");
                            //                List<string> listOfLineItemIntervals = lineItemIntervals.Split(',').ToList<string>();

                            //                string lineItemNPV = sensitivityDatasVM.LineItemNPV.Replace("\r\n", "");
                            //                lineItemNPV = lineItemNPV.Replace(" ", "");
                            //                lineItemNPV = lineItemNPV.Replace("[", "");
                            //                lineItemNPV = lineItemNPV.Replace("]", "");
                            //                List<string> listOfLineItemNPV = lineItemNPV.Split(',').ToList<string>();

                            //                Result.Add(listOfLineItemIntervals);
                            //                Result.Add(listOfLineItemNPV);

                            //                ////for Actual Values
                            //                List<string> actualList = new List<string>();
                            //                actualList.Add(sensitivityDatasVM.ActualValue.ToString());
                            //                actualList.Add(sensitivityDatasVM.NewValue.ToString());
                            //                actualList.Add(sensitivityDatasVM.IntervalDifference.ToString());
                            //                Result.Add(actualList);

                            //                string aa = sensitivityDatasVM.LineItem + "(" + sensitivityDatasVM.LineItemUnit.ToString() + ")";

                            //                // npvList.TryAdd(sensitivityDatasVM.LineItem + FindUnit(Convert.ToInt32(NPVOutput.ValueTypeId), Convert.ToInt32(NPVOutput.UnitId)), Result);
                            //                npvList.TryAdd(sensitivityDatasVM.LineItem + "(" + sensitivityDatasVM.LineItemUnit.ToString() + ")", Result);
                            //                npvOutput.Add(npvList);
                            //            }
                            //        }
                            //    }

                            //    if ((!string.IsNullOrEmpty(sensimodel.linItem)) && excecuted == 0)
                            //    {
                            //        npvList = new Dictionary<string, List<List<string>>>();
                            //        Dictionary<string, object> model = new Dictionary<string, object>();

                            //        npvList.TryAdd(sensimodel.linItem, getNPVforSensitivityInputNew(model, input, sensimodel.linItem, sensimodel.difference, false/*, midvalue*/));
                            //        npvOutput.Add(npvList);
                            //    }

                            //    return npvOutput;

                            //}
                           // else
                           // {
                                //var sensitivityInputDatas = iSensitivityInputData.FindBy(x => x.Id != 0 && x.ProjectId == sensimodel.ProjectId).ToList();
                                //if (sensitivityInputDatas != null && sensitivityInputDatas.Count > 0)
                                //{
                                foreach (var sensitivityDatasVM in sensitivityInputDatas)
                                {
                                    if (sensitivityDatasVM != null)
                                    {
                                        npvList = new Dictionary<string, List<List<string>>>();
                                        Dictionary<string, object> model = new Dictionary<string, object>();
                                        List<List<string>> Result = new List<List<string>>();

                                        string lineItemIntervals = sensitivityDatasVM.LineItemIntervals.Replace("\r\n", "");
                                        lineItemIntervals = lineItemIntervals.Replace(" ", "");
                                        lineItemIntervals = lineItemIntervals.Replace("[", "");
                                        lineItemIntervals = lineItemIntervals.Replace("]", "");
                                        List<string> listOfLineItemIntervals = lineItemIntervals.Split(',').ToList<string>();

                                        string lineItemNPV = sensitivityDatasVM.LineItemNPV.Replace("\r\n", "");
                                        lineItemNPV = lineItemNPV.Replace(" ", "");
                                        lineItemNPV = lineItemNPV.Replace("[", "");
                                        lineItemNPV = lineItemNPV.Replace("]", "");
                                        List<string> listOfLineItemNPV = lineItemNPV.Split(',').ToList<string>();

                                        Result.Add(listOfLineItemIntervals);
                                        Result.Add(listOfLineItemNPV);

                                        ////for Actual Values
                                        List<string> actualList = new List<string>();                                        
                                        actualList.Add(sensitivityDatasVM.ActualValue.ToString());
                                        actualList.Add(sensitivityDatasVM.NewValue.ToString());
                                        actualList.Add(sensitivityDatasVM.IntervalDifference.ToString());
                                        Result.Add(actualList);

                                        List<string> saved = new List<string>();
                                        saved.Add("1");
                                        Result.Add(saved);

                                    //  var NPVOutput = input.ProjectInputDatasVM.Find(x => x.LineItem.Contains(sensitivityDatasVM.LineItem));

                                    //  string aa = sensitivityDatasVM.LineItem + FindUnit(Convert.ToInt32(NPVOutput.ValueTypeId), Convert.ToInt32(NPVOutput.UnitId));

                                    string sensitivityLineItem = "";

                                     if (sensitivityDatasVM.LineItemUnit == "")
                                    {
                                        sensitivityLineItem = sensitivityDatasVM.LineItem;
                                        npvList.TryAdd(sensitivityDatasVM.LineItem , Result);
                                    }
                                    else
                                    {
                                        sensitivityLineItem = sensitivityDatasVM.LineItem + "(" + sensitivityDatasVM.LineItemUnit.ToString() + ")";
                                        npvList.TryAdd(sensitivityDatasVM.LineItem + "(" + sensitivityDatasVM.LineItemUnit.ToString() + ")", Result);
                                    }

                                        //string aa = sensitivityDatasVM.LineItem + "(" + sensitivityDatasVM.LineItemUnit.ToString() + ")";

                                        // npvList.TryAdd(sensitivityDatasVM.LineItem + FindUnit(Convert.ToInt32(NPVOutput.ValueTypeId), Convert.ToInt32(NPVOutput.UnitId)), Result);
                                       // npvList.TryAdd(sensitivityDatasVM.LineItem + "(" + sensitivityDatasVM.LineItemUnit.ToString() + ")", Result);

                                        npvOutput.Add(npvList);
                                        sensimodel.variableList.Remove(sensitivityLineItem);

                                        // npvList.TryAdd(sensitivityDatasVM.LineItem, getNPVforSensitivityInputsNew(model, input, sensitivityDatasVM.LineItem, 0, true/*, midvalue*/));
                                        //npvOutput.Add(npvList);
                                        // return npvOutput;
                                    }
                                }

                                foreach (string variable in sensimodel.variableList)
                                {
                                    npvList = new Dictionary<string, List<List<string>>>();
                                    Dictionary<string, object> model = new Dictionary<string, object>();

                                    input = ProjectDetails(sensimodel.ProjectId, project);

                                    npvList.TryAdd(variable, getNPVforSensitivityInputNew(model, input, variable, 0, true, sensimodel.ProjectId/*, midvalue*/));
                                    npvOutput.Add(npvList);
                                }

                                return npvOutput;
                                //}
                           // }
                        }                                             
                    }                       
                }
                else
                {
                    try
                    {
                        List<Dictionary<string, List<List<double>>>> npvOutput = new List<Dictionary<string, List<List<double>>>>();

                        if (sensimodel.ProjectId != 0)
                        {
                            Project project = iProject.GetSingle(x => x.Id == sensimodel.ProjectId);
                            if (project != null)
                            {
                                input = ProjectDetails(sensimodel.ProjectId, project);
                                Dictionary<string, List<List<double>>> npvList = new Dictionary<string, List<List<double>>>();

                                //find NPV  for default
                                if (input != null && input.ProjectSummaryDatasVM != null && input.ProjectSummaryDatasVM.Count > 0)
                                {
                                    // find NPV
                                    var NPVOutput = input.ProjectSummaryDatasVM.Find(x => x.LineItem.Contains("Net Present Value"));
                                    // npv = Convert.ToDouble(NPVOutput.Value);
                                    List<List<double>> tList = new List<List<double>>();
                                    List<double> tdouble = new List<double>();
                                    tdouble.Add(Convert.ToDouble(NPVOutput.Value));
                                    tdouble.Add(Convert.ToDouble(NPVOutput.Value));
                                    tList.Add(tdouble);
                                    //npvList.TryAdd(NPVOutput.LineItem, tList);
                                    npvList.TryAdd(NPVOutput.LineItem + FindUnit(Convert.ToInt32(NPVOutput.ValueTypeId), Convert.ToInt32(NPVOutput.UnitId)), tList);
                                    npvOutput.Add(npvList);
                                }
                                if (!string.IsNullOrEmpty(sensimodel.linItem))
                                {
                                    npvList = new Dictionary<string, List<List<double>>>();
                                    Dictionary<string, object> model = new Dictionary<string, object>();

                                    //model.Add(linItem, 0);
                                    npvList.TryAdd(sensimodel.linItem, getNPVforSensitivity(model, input, sensimodel.linItem, sensimodel.difference, false, sensimodel.ProjectId));
                                    npvOutput.Add(npvList);
                                }
                                else
                                {
                                    //get  Variable List
                                    if (sensimodel.variableList == null || sensimodel.variableList.Count == 0)
                                    {
                                        if (input != null && input.ProjectInputDatasVM != null && input.ProjectInputDatasVM.Count > 0)
                                        {
                                            sensimodel.variableList = new List<string>();
                                            foreach (var datas in input.ProjectInputDatasVM)
                                            {
                                                sensimodel.variableList.Add(datas.LineItem);
                                            }
                                        }
                                    }
                                    foreach (string variable in sensimodel.variableList)
                                    {
                                        npvList = new Dictionary<string, List<List<double>>>();
                                        Dictionary<string, object> model = new Dictionary<string, object>();

                                        //model.Add(variable, 10);
                                        input = ProjectDetails(sensimodel.ProjectId, project);
                                        npvList.TryAdd(variable, getNPVforSensitivity(model, input, variable, 0, true/*, midvalue*/, sensimodel.ProjectId));
                                        npvOutput.Add(npvList);
                                    }
                                }
                            }
                        }
                        return npvOutput;
                    }
                    catch (Exception ex)
                    {
                        return (BadRequest(new { message = ex.Message.ToString(), code = 400 }));
                    }                    
                }
            }
                    
            return null;
        }

        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("getprojectSensitivityGraph")]
        public ActionResult<Object> getprojectSensitivityGraph([FromBody] sensitivity sensimodel/* variableList,long ProjectId,string linItem,int difference*/)
        {
            List<SensitivityInputGraphData> sensitivityInputGraphDatasList = iSensitivityInputGraphData.FindBy(x => x.ProjectId == sensimodel.ProjectId).ToList();

            if ((sensimodel.linItem == "") && (sensitivityInputGraphDatasList.Count == 0))
            {
                SensitivityInputsViewModel result = new SensitivityInputsViewModel();
                return result;
            }
            else if ((sensimodel.linItem == "") && (sensitivityInputGraphDatasList.Count > 0))
            {
                //  var sensitivityInputDatas = iSensitivityInputGraphData.FindBy(x => x.Id != 0 && x.ProjectId == sensimodel.ProjectId).ToList();
                // if (sensitivityInputDatas != null && sensitivityInputDatas.Count > 0)
                // {           

                List<Dictionary<string, List<List<string>>>> npvOutput = new List<Dictionary<string, List<List<string>>>>();
                Dictionary<string, List<List<string>>> npvList = new Dictionary<string, List<List<string>>>();

                foreach (var sensitivityInputGraphDatasVM in sensitivityInputGraphDatasList)
                    {
                        if (sensitivityInputGraphDatasVM != null)
                        {

                        npvList = new Dictionary<string, List<List<string>>>();
                        List<List<string>> Result = new List<List<string>>();

                            string lineItemIntervals = sensitivityInputGraphDatasVM.LineItemIntervals.Replace("\r\n", "");
                            lineItemIntervals = lineItemIntervals.Replace(" ", "");
                            lineItemIntervals = lineItemIntervals.Replace("[", "");
                            lineItemIntervals = lineItemIntervals.Replace("]", "");
                            List<string> listOfLineItemIntervals = lineItemIntervals.Split(',').ToList<string>();

                            string lineItemNPV = sensitivityInputGraphDatasVM.LineItemNPV.Replace("\r\n", "");
                            lineItemNPV = lineItemNPV.Replace(" ", "");
                            lineItemNPV = lineItemNPV.Replace("[", "");
                            lineItemNPV = lineItemNPV.Replace("]", "");
                            List<string> listOfLineItemNPV = lineItemNPV.Split(',').ToList<string>();

                            Result.Add(listOfLineItemIntervals);
                            Result.Add(listOfLineItemNPV);

                           npvList.TryAdd(sensitivityInputGraphDatasVM.LineItem , Result);
                           npvOutput.Add(npvList);

                    }
                }

                 return npvOutput;
                //  }
            }
            else if (sensimodel.linItem != "")
            {
                ProjectsViewModel input = new ProjectsViewModel();
                try
                {
                    List<Dictionary<string, List<List<double>>>> npvOutput = new List<Dictionary<string, List<List<double>>>>();

                    if (sensimodel.ProjectId != 0)
                    {
                        Project project = iProject.GetSingle(x => x.Id == sensimodel.ProjectId);
                        if (project != null)
                        {
                            input = ProjectDetails(sensimodel.ProjectId, project);
                            Dictionary<string, List<List<double>>> npvList = new Dictionary<string, List<List<double>>>();

                            if (!string.IsNullOrEmpty(sensimodel.linItem))
                            {
                                npvList = new Dictionary<string, List<List<double>>>();
                                Dictionary<string, object> model = new Dictionary<string, object>();

                                //model.Add(linItem, 0);
                                npvList.TryAdd(sensimodel.linItem, getNPVforSensitivityInputsGraph(model, input, sensimodel.linItem, sensimodel.difference, false, sensimodel.ProjectId));
                                npvOutput.Add(npvList);
                            }
                        }
                    }
                    return npvOutput;
                }
                catch (Exception ex)
                {
                    return (BadRequest(new { message = ex.Message.ToString(), code = 400 }));
                }
            }

            // List<SensitivityInputGraphData> inputDatasList = iSensitivityInputGraphData.FindBy(x => x.ProjectId == sensimodel.ProjectId).ToList();
            //if (inputDatasList != null && inputDatasList.Count > 0)    // if(Data is available )
            //{
            //    iSensitivityInputGraphData.DeleteMany(inputDatasList);
            //    iSensitivityInputGraphData.Commit();
            //}

            return null;
        }


        [HttpGet]
        [Microsoft.AspNetCore.Mvc.Route("SensitivityByVariable/{ProjectId}")]
        public ActionResult<Object> SensitivityByVariable(double value, long UserId, long ProjectId)
        {
            //Dictionary<string, List<List<double>>> result = new Dictionary<string, List<List<double>>>();
            List<List<double>> NPVLineItemList = new List<List<double>>();
            List<double> fractionList = FractionValues(value);
            foreach (var fraction in fractionList)
            {
                //get NPV for each Section


            }


            return NPVLineItemList;
        }

        private List<double> FractionValuesForGraphs(double Value, double difference = 0)
        {
            List<double> resultList = new List<double>();
            // List<double> resultListNew = new List<double>();

            // double cleanmidValue = Value < 0 ? Value : Convert.ToDouble(Value.ToString("0."));

            //if (Value < 100)
            //    difference = 10;
            //else if (Value < 1000)
            //    difference = 100;
            //else if (Value < 10000)
            //    difference = 1000;
            //else if (Value < 100000)
            //    difference = 10000;
            //else if (Value < 1000000)
            //    difference = 100000;
            //else if (Value < 10000000)
            //    difference = 1000000;

            //   int integerDifference = 0;

            if (Value == 0 || Value < 2)
                difference = 0.1;
            else if (Value == 2 || Value < 20)
                difference = 1;
            else if (Value == 20 || Value < 200)
                difference = 10;
            else if (Value == 200 || Value < 2000) //1K
                difference = 100;
            else if (Value == 2000 || Value < 20000) //10K
                difference = 1000;
            else if (Value == 200000 || Value < 2000000) //1M
                difference = 10000;
            else if (Value == 2000000 || Value < 20000000)
                difference = 100000;
            else if (Value == 20000000 || Value < 200000000)
                difference = 1000000;
            else if (Value == 200000000 || Value < 2000000000) //1B
                difference = 10000000;
            else if (Value == 2000000000 || Value < 20000000000) 
                difference = 100000000;
            else if (Value == 20000000000 || Value < 200000000000) 
                difference = 1000000000;
            else if (Value == 200000000000 || Value < 2000000000000) //1T
                difference = 10000000000;

            int roundedValue = (int)Math.Round(Value);
            double intervalValue = 0;
            int rightIntervalCount = 0;

            // Right & Left intervals

                if (roundedValue == 0 || roundedValue < 2)
                {

                for (double i = 0; i <= roundedValue; i = i + difference)
                {
                    resultList.Add(i);
                    intervalValue = i;
                }

                for (double i = intervalValue + difference; i <= roundedValue * 2; i = i + difference)
                {
                    resultList.Add(i);
                    rightIntervalCount = rightIntervalCount + 1;
                    intervalValue = i;
                }

                if (rightIntervalCount < 5)
                {
                    for (int i = rightIntervalCount + 1; i <= 5; i++)
                    {
                        intervalValue = intervalValue + difference;
                        resultList.Add(intervalValue);
                        // rightIntervalCount++;
                    }
                }

            }

              //  else if (roundedValue == 2 || roundedValue < 20)
              else
                {
                
                for (double i = 0; i <= roundedValue; i =  i + difference)
                {
                    resultList.Add(i);
                    intervalValue = i;
                }

                for (double i = intervalValue + difference; i <= roundedValue * 2; i = i + difference)
                {
                    resultList.Add(i);
                    rightIntervalCount = rightIntervalCount + 1;
                    intervalValue = i;
                }

                if (rightIntervalCount < 5)
                {
                    for (int i = rightIntervalCount + 1; i <= 5; i++)
                    {
                        intervalValue = intervalValue + difference;
                        resultList.Add(intervalValue);
                       // rightIntervalCount++;
                    }
                }
            }

            //else if (roundedValue == 20 || roundedValue < 200)
            //{

            //    for (double i = 0; i <= roundedValue; i = i + difference)
            //    {
            //        resultList.Add(i);
            //        intervalValue = i;
            //    }

            //    for (double i = intervalValue + difference; i <= roundedValue * 2; i = i + difference)
            //    {
            //        resultList.Add(i);
            //        rightIntervalCount = rightIntervalCount + 1;
            //        intervalValue = i;
            //    }

            //    if (rightIntervalCount < 5)
            //    {
            //        for (int i = rightIntervalCount + 1; i <= 5; i++)
            //        {
            //            intervalValue = intervalValue + 10;
            //            resultList.Add(intervalValue);
            //        }
            //    }

            //}

            //foreach (var number in resultList)
            //{
            //   resultListNew.Add((Double)Decimal.Truncate(number));
            //}

            return resultList;
        }

        private List<string> FractionValuesNew(double Value, double difference = 0, bool isdefault = false)
        {
            List<string> resultList = new List<string>();

            // Find nearest value 
            double nearestvalue;
            double cleanmidValue = Value < 0 ? Value : Convert.ToDouble(Value.ToString("0."));

            if (isdefault == true)
            {
                if (Value < 10)
                {
                    cleanmidValue = Value;
                    difference = Value < 1 ? 0.1 : 0.5;
                }
                else if (Value < 100)
                    difference = 1;
                else if (Value < 1000)
                    difference = 5;
                else if (Value < 10000)
                    difference = 100;
                else if (Value < 100000)
                    difference = 200;
                else if (Value < 1000000)
                    difference = 500;
                else if (Value < 10000000)
                    difference = 10000;
            }

            if (difference == 0)
            {
                double modulusValue = cleanmidValue % 10;

                if (modulusValue > 5)
                    nearestvalue = cleanmidValue + (10 - modulusValue);
                else
                    nearestvalue = cleanmidValue - modulusValue;

                double tempvalue = nearestvalue - 50;

                for (int i = 1; i <= 11; i++)
                {
                    resultList.Add(tempvalue.ToString());
                    tempvalue = tempvalue + 10;
                    //if (tempvalue <= 0)
                    //   break;
                }
            }
            else
            {
                double tempvalue = cleanmidValue - (difference * 5);

                for (int i = 1; i <= 11; i++)
                {
                    resultList.Add(tempvalue.ToString());
                    tempvalue = tempvalue + difference;
                    //if (tempvalue <= 0)
                    //{

                    //}
                }
            }
            // differenceValue = difference;
            return resultList;
        }

        // public List<double> FractionValues(double Value, out double differenceValue , double difference = 0, bool isdefault = false)
        private List<double> FractionValues(double Value, double difference = 0, bool isdefault = false)
        {
            List<double> resultList = new List<double>();

            // Find nearest value 
            double nearestvalue;
            double cleanmidValue = Value < 0 ? Value : Convert.ToDouble(Value.ToString("0."));

            if (isdefault == true)
            {
                if (Value < 10)
                {
                    cleanmidValue = Value;
                    difference = Value < 1 ? 0.1 : 0.5;
                }
                else if (Value < 100)
                    difference = 1;
                else if (Value < 1000)
                    difference = 5;
                else if (Value < 10000)
                    difference = 100;
                else if (Value < 100000)
                    difference = 200;
                else if (Value < 1000000)
                    difference = 500;
                else if (Value < 10000000)
                    difference = 10000;
            }

            if (difference == 0)
            {
                double modulusValue = cleanmidValue % 10;

                if (modulusValue > 5)
                    nearestvalue = cleanmidValue + (10 - modulusValue);
                else
                    nearestvalue = cleanmidValue - modulusValue;

                double tempvalue = nearestvalue - 50;

                for (int i = 1; i <= 11; i++)
                {
                    resultList.Add(tempvalue);
                    tempvalue = tempvalue + 10;
                    //if (tempvalue <= 0)
                     //   break;
                }
            }
            else
            {
                double tempvalue = cleanmidValue - (difference * 5);

                for (int i = 1; i <= 11; i++)
                {
                    resultList.Add(tempvalue);
                    tempvalue = tempvalue + difference;
                    //if (tempvalue <= 0)
                    //{

                    //}
                }
            }
           // differenceValue = difference;
            return resultList;
        }


        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("AddSensitivityGraphInputs")]
        public ActionResult AddSensitivityGraphInputs([FromBody] SensitivityInputsViewModel model)
        {
            try
            {

                List<SensitivityInputGraphData> inputDatasList = iSensitivityInputGraphData.FindBy(x => x.ProjectId == model.ProjectId).ToList();
                if (inputDatasList != null && inputDatasList.Count > 0)    // if(Data is available )
                {
                    iSensitivityInputGraphData.DeleteMany(inputDatasList);
                    iSensitivityInputGraphData.Commit();
                }

                SensitivityInputGraphData SensitivityInputGraph = new SensitivityInputGraphData
                {
                    ProjectId = model.ProjectId,
                    LineItem = model.LineItem,
                    LineItemIntervals = model.LineItemIntervals.ToString(),
                    LineItemNPV = model.LineItemNPV.ToString()
                };

                iSensitivityInputGraphData.Add(SensitivityInputGraph);
                iSensitivityInputGraphData.Commit();

                return Ok(new { message = "Succesfully added Sensitivity Graph Inputs", code = 200 });
                
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }


        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("AddSensitivityInputs")]
        public ActionResult AddSensitivityInputs([FromBody] SensitivityInputs_ProjectViewModel model)
        {
            try
            {
                //SensitivityInputData SensitivityInput = new SensitivityInputData
                //{
                //    ProjectId = model.ProjectId,
                //    ActualNPV = model.ActualNPV,
                //    NewNPV = model.NewNPV,
                //    NewValue = model.NewValue,
                //    ActualValue = model.ActualValue,
                //    IntervalDifference = model.IntervalDifference,
                //    LineItem = model.LineItem,
                //    LineItemUnit = model.LineItemUnit,
                //    NPVUnit = model.NPVUnit,
                //    LineItemIntervals = model.LineItemIntervals.ToString(),
                //    LineItemNPV = model.LineItemNPV.ToString()
                //};

                //iSensitivityInputData.Add(SensitivityInput);
                //iSensitivityInputData.Commit();

                //  return Ok(new { result = SensitivityInput.SensitivityInputValues[0].Id, message = "Succesfully added Sensitivity Input", code = 200 });

                int ProjectId = 0;

                foreach (var DatasVM in model.SensitivityInputDatasVM)
                {
                    if (DatasVM != null)
                    {
                        ProjectId = DatasVM.ProjectId;
                        break;
                    }
                }

                List<SensitivityInputData> inputDatasList = iSensitivityInputData.FindBy(x => x.ProjectId == ProjectId).ToList();               
                if (inputDatasList != null && inputDatasList.Count > 0)    // if(Data is available )
                {
                    iSensitivityInputData.DeleteMany(inputDatasList);
                    iSensitivityInputData.Commit();
                }

                    SensitivityInputData SensitivityInput;
                //   List<SensitivityInputData> inputDatasList = new List<SensitivityInputData>();

                inputDatasList = new List<SensitivityInputData>();

                //List<SensitivityInputs> inputDatasList = new List<SensitivityInputs>();
                // List<SensitivityInputs> SensitivityInputValues = new List<SensitivityInputs>();
                // SensitivityInputs values = new SensitivityInputs();

                foreach (var DatasVM in model.SensitivityInputDatasVM)
                {
                    if (DatasVM != null)
                    {
                        // model.Id = 0;

                        SensitivityInput = new SensitivityInputData();
                        // values = new SensitivityInputs();
                        // SensitivityInput.SensitivityInputValues = new SensitivityInputs();
                        //  SensitivityInput.SensitivityInputValues = new List<SensitivityInputs>();
                        //  SensitivityInput = mapper.Map<SensitivityInputs_ProjectViewModel, SensitivityInputData>(DatasVM);
                        //SensitivityInputValues[0].ProjectId = DatasVM.ProjectId;
                        // values.ProjectId = DatasVM.ProjectId;
                        // SensitivityInput.Add(values); 

                        SensitivityInput.ProjectId = DatasVM.ProjectId;
                        SensitivityInput.ActualNPV = DatasVM.ActualNPV;
                        SensitivityInput.NewNPV = DatasVM.NewNPV;
                        SensitivityInput.NewValue = DatasVM.NewValue;
                        SensitivityInput.ActualValue = DatasVM.ActualValue;
                        SensitivityInput.IntervalDifference = DatasVM.IntervalDifference;
                        SensitivityInput.LineItem = DatasVM.LineItem;
                        SensitivityInput.LineItemUnit = DatasVM.LineItemUnit;
                        SensitivityInput.NPVUnit = DatasVM.NPVUnit;
                        SensitivityInput.LineItemIntervals = DatasVM.LineItemIntervals.ToString();
                        SensitivityInput.LineItemNPV = DatasVM.LineItemNPV.ToString();

                        //SensitivityInput.SensitivityInputValues.Add(values);

                        //SensitivityInput.SensitivityInputValues[0].ProjectId = DatasVM.ProjectId;
                        //SensitivityInput.SensitivityInputValues[0].ActualNPV = DatasVM.ActualNPV;
                        //SensitivityInput.SensitivityInputValues[0].NewNPV = DatasVM.NewNPV;
                        //SensitivityInput.SensitivityInputValues[0].NewValue = DatasVM.NewValue;
                        //SensitivityInput.SensitivityInputValues[0].ActualValue = DatasVM.ActualValue;
                        //SensitivityInput.SensitivityInputValues[0].IntervalDifference = DatasVM.IntervalDifference;
                        //SensitivityInput.SensitivityInputValues[0].LineItem = DatasVM.LineItem;
                        //SensitivityInput.SensitivityInputValues[0].LineItemUnit = DatasVM.LineItemUnit;
                        //SensitivityInput.SensitivityInputValues[0].NPVUnit = DatasVM.NPVUnit;
                        //SensitivityInput.SensitivityInputValues[0].LineItemIntervals = DatasVM.LineItemIntervals.ToString();
                        //SensitivityInput.SensitivityInputValues[0].LineItemNPV = DatasVM.LineItemNPV.ToString();

                         inputDatasList.Add(SensitivityInput);
                        //inputDatasList.Add(values);

                    }
                }

                if (inputDatasList != null && inputDatasList.Count > 0)
                {
                    iSensitivityInputData.AddMany(inputDatasList);
                    iSensitivityInputData.Commit();
                }

                return Ok(new {message = "Succesfully added Sensitivity Input", code = 200 });             
           
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

        [HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("AddScenarioInputs")]
        public ActionResult AddScenarioInputs([FromBody] ScenarioInputDataViewModel model)
        {
            try
            {
                if (model.ScenarioInputDatasVM != null && model.ScenarioInputDatasVM.Count > 0)
                {
                    //check if existed data
                    List<ScenarioInputDatas> inputDatasList = iScenarioInputDatas.FindBy(x => x.ProjectId == model.ProjectId).ToList();
                    if (inputDatasList != null && inputDatasList.Count > 0) 
                    {
                        //check for Values 
                        List<ScenarioInputValues> valuesList = new List<ScenarioInputValues>();
                        valuesList = iScenarioInputValues.FindBy(x => inputDatasList.Any(t => t.Id == x.ScenarioInputDatasId)).ToList();

                        if (valuesList != null && valuesList.Count > 0)
                        {
                            iScenarioInputValues.DeleteMany(valuesList);
                            iScenarioInputValues.Commit();
                        }

                          iScenarioInputDatas.DeleteMany(inputDatasList);
                          iScenarioInputDatas.Commit();
                    }

                    ScenarioInputDatas scenarioInputDatas;
                    inputDatasList = new List<ScenarioInputDatas>();
                    ScenarioInputValues values = new ScenarioInputValues();

                    foreach (var DatasVM in model.ScenarioInputDatasVM)
                    {
                        if (DatasVM != null)
                        {
                            //DatasVM.ProjectId = model.ProjectId;
                            DatasVM.Id = 0;
                          //  DatasVM.ScenarioInputDatasId = 0;

                            scenarioInputDatas = new ScenarioInputDatas();
                            scenarioInputDatas = mapper.Map<ScenarioInputDatasViewModel, ScenarioInputDatas>(DatasVM);

                            if (scenarioInputDatas.ScenarioInputValues != null && scenarioInputDatas.ScenarioInputValues.Count > 0)
                            {
                                foreach (var value in scenarioInputDatas.ScenarioInputValues)
                                {
                                    value.Id = 0;
                                }

                            }                           
                            inputDatasList.Add(scenarioInputDatas);
                        }
                    }

                    if (inputDatasList != null && inputDatasList.Count > 0)
                    {
                        iScenarioInputDatas.AddMany(inputDatasList);
                        iScenarioInputDatas.Commit();
                    }
                }

                return Ok(new { message = "Succesfully added Sensitivity Input", code = 200 });
        }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message.ToString(), code = 400 });
            }
        }

    }
}
