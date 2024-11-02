using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using Sowfin.Data.Repositories;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels
{

    
    public class ProjectsViewModel
    {
        public ProjectsViewModel()
        {

        }
        public long ValuationMethodId { get; set; }
        public string BusinessUnitName { get; set; }
        public string ValuationMethodName { get; set; }
        public string ApprovalComment { get; set; }


        public long Id { get; set; }
        public long BusinessUnitId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int ApprovalFlag { get; set; }
        public int? ValuationTechniqueId { get; set; }
        public int? StartingYear { get; set; }
        public int? NoOfYears { get; set; }
        public int? UserId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public bool? IsActive { get; set; }
        public List<ProjectInputDatasViewModel> ProjectInputDatasVM { get; set; }
        public List<ProjectOutputDatasViewModel> ProjectSummaryDatasVM { get; set; }
        public List<SelectListItem> CurrencyValueList { get; set; }
        public List<SelectListItem> NumberCountList { get; set; }
        public List<SelectListItem> ValueTypeList { get; set; }
        public List<ValuationTechnique> ValuationTechniqueList { get; set; }
        public List<SelectListItem> HeaderList { get; set; }
        public List<DepreciationInputDatasViewModel> DepreciationInputDatasVM { get; set; }

        public List<Snapshot_ProjectViewModel> Snapshot_ProjectVM { get; set; }
        public Approval ApprovalStatus { get; set; }
       // public List<SensitivityInputs_ProjectViewModel> Sensitivity_ProjectVM { get; set; }

        public static explicit operator ProjectsViewModel(ActionResult v)
        {
            throw new NotImplementedException();
        }

         
    }
    public class ProjectInputDatasViewModel
    {
        public long Id { get; set; }
        public long? ProjectId { get; set; }
        public long? HeaderId { get; set; }
        public string SubHeader { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int? ValueTypeId { get; set; }
        public bool HasMultiYear { get; set; }
        public double? Value { get; set; }
        public int? Duration { get; set; }
        public int? Method { get; set; }
        public bool SameYear { get; set; }
        public List<ProjectInputValuesViewModel> ProjectInputValuesVM { get; set; }

        public List<ProjectInputComparablesViewModel> ProjectInputComparablesVM { get;set; }
        public List<DepreciationInputValuesViewModel> DepreciationInputValuesVM { get; set; }
    }
  public class DepreciationInputDatasViewModel
    {
        public long Id { get; set; }
        public long? ProjectId { get; set; }
        public long? ProjectInputDatasId { get; set; }
        // public string LineItem { get; set; }
        public bool HasMultiYear { get; set; }
        public long? Method { get; set; }
        // public long? DepreciationMethodValue { get; set; }
        public bool SameYear { get; set; }
        public long? Duration { get; set; }
        //public List<DepreciationInputValuesViewModel> DepreciationInputValuesVM { get; set; }
    }

    public class ProjectInputComparablesViewModel
    {
        public long Id { get; set; }
        public long? ProjectInputDatasId { get; set; }
        public string Comparable { get; set; }
        public int? count { get; set; }
        public double? Value { get; set; }
    }

    public class ProjectInputValuesViewModel
    {
        public long Id { get; set; }
        public long? ProjectInputDatasId { get; set; }
        public int? Year { get; set; }
        public double? Value { get; set; }

    }
    public class DepreciationInputValuesViewModel
    {
        public long Id { get; set; }
        public long? ProjectInputDatasId { get; set; }
        public int? Year { get; set; }
        public double? Value { get; set; }

    }

    //Output Summary
    public class ProjectOutputDatasViewModel
    {
        public long Id { get; set; }
        public long? ProjectId { get; set; }
        public long? HeaderId { get; set; }
        public string Header { get; set; }
        public string SubHeader { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int? ValueTypeId { get; set; }
        public bool HasMultiYear { get; set; }
        public double? Value { get; set; }
        public List<ProjectOutputValuesViewModel> ProjectOutputValuesVM { get; set; }
    }

    public class ProjectOutputValuesViewModel
    {
        public long Id { get; set; }
        public long? ProjectOutputDatasId { get; set; }
        public int? Year { get; set; }
        public double? Value { get; set; }
        public double? BasicValue { get; set; }

    }

    public class Project_SnapshotDatasViewModel
    {
        public long Id { get; set; }
        public long? ProjectId { get; set; }
        public long? HeaderId { get; set; }
        public string SubHeader { get; set; }
        public string LineItem { get; set; }
        public int? UnitId { get; set; }
        public int? ValueTypeId { get; set; }
        public bool HasMultiYear { get; set; }
        public double? Value { get; set; }

    }

    public class Project_SnapshotValuesViewModel
    {
        public long Id { get; set; }
        public long? Project_SnapshotDatasId { get; set; }
        public int? Year { get; set; }
        public double? Value { get; set; }

    }

    public class Snapshot_ProjectViewModel
    {
        public int? Id { get; set; }
        public string ProjectInputs { get; set; }
        public string Snapshot { get; set; }
        public string Description { get; set; }
        public int UserId { get; set; }
        public int ProjectId { get; set; }
        public int? ValuationTechniqueId { get; set; }
        public string NPV { get; set; }
        public string CNPV { get; set; }
        public DateTime? CreatedAt { get; set; }
        public bool? Active { get; set; }

    }

    public class Project_ScenarioAnalysis
    {
        public Project_ScenarioAnalysis()
        {
            inputVolumes = new List<double>();
            inputUnitPrice = new List<double>();
            inputUnitCost = new List<double>();
            inputCapex = new List<double>();
            inputDepreciation = new List<double>();
            inputNWC = new List<double>();
            inputWACC = 0;
            inputUnleveredCost = 0;
            inputCostofDebt = 0;
            inputDVRatio = 0;
            inputIntrestCoverageRatio = 0;
            inputCostofEquity = 0;
            inputMarginalTax = 0;
            inputFixedSchedule = new List<double>();
            inputSgA = new List<double>();
            inputRD = new List<double>();
            inputProjectDVRatio = 0;
            inputEquityCostofCapital = 0;
            inputDebtCostofCapital = 0;
        }
        public List<double> inputVolumes { get; set; }
        public List<double> inputUnitPrice { get; set; }
        public List<double> inputUnitCost { get; set; }
        public List<double> inputCapex { get; set; }
        public List<double> inputDepreciation { get; set; }
        public List<double> inputNWC { get; set; }
        public double inputWACC { get; set; }
        public double inputDVRatio { get; set; }
        public double inputUnleveredCost { get; set; }
        public double inputCostofDebt { get; set; }
        public double inputCostofEquity { get; set; }
        public double inputIntrestCoverageRatio { get; set; }
        public double inputMarginalTax { get; set; }
        public List<double> inputSgA { get; set; }
        public List<double> inputRD { get; set; }
        public List<double> inputFixedSchedule { get; set; }
        public double inputProjectDVRatio { get; set; }
        public double inputEquityCostofCapital { get; set; }
        public double inputDebtCostofCapital { get; set; }

    }

    public class Project_ScenarioOutput
    {
        public Project_ScenarioOutput()
        {
            sensiSceSummaryOutput = new List<ProjectOutputDatasViewModel>();
        }
        public double marginalTax { get; set; }
        public double waccDiscount { get; set; }
        public double volChangePerc { get; set; }
        public double unitPricePerChange { get; set; }
        public double unitCostPerChange { get; set; }
        public double capexDepPerChange { get; set; }
        public double rDPerChange { get; set; }
        public double sgAPerChange { get; set; }
        public List<ProjectOutputDatasViewModel> sensiSceSummaryOutput { get; set; }

    }

    public class ScenarioAnalysis
    { 
        public Dictionary<string, List<List<double>>> variables { get; set; }
        public List<Snapshot_ProjectViewModel> scenarioSnapshots { get; set; }
    }

    public class SensitivityInputs_ProjectViewModel
    {
        public int? Id { get; set; }
        public List<SensitivityInputsViewModel> SensitivityInputDatasVM { get; set; }

        //public int? Id { get; set; }
        //public int ProjectId { get; set; }
        //public double ActualNPV { get; set; }
        //public double NewNPV { get; set; }
        //public double NewValue { get; set; }
        //public double ActualValue { get; set; }
        //public double IntervalDifference { get; set; }
        //public string LineItem { get; set; }
        //public int LineItemUnit { get; set; }
        //public int NPVUnit { get; set; }
        //public object LineItemIntervals { get; set; }
        //public object LineItemNPV { get; set; }
   }

    public class SensitivityInputsViewModel
    {
        public int? Id { get; set; }
        public int ProjectId { get; set; }
        public double ActualNPV { get; set; }
        public double NewNPV { get; set; }
        public double NewValue { get; set; }
        public double ActualValue { get; set; }
        public double IntervalDifference { get; set; }
        public string LineItem { get; set; }
        public string LineItemUnit { get; set; }
        public string NPVUnit { get; set; }
        public object LineItemIntervals { get; set; }
        public object LineItemNPV { get; set; }
    }

    public class ScenarioInputDataViewModel
    {
        //public long Id { get; set; }
        public long? ProjectId { get; set; }
        //public string Name { get; set; }
        //public string Description { get; set; }
        //public double? Probability { get; set; }
        //public string LineItem { get; set; }
        //public double? NPV { get; set; }
        public List<ScenarioInputDatasViewModel> ScenarioInputDatasVM { get; set; }
    }
    public class ScenarioInputDatasViewModel
    {
        public long Id { get; set; }
       // public long? ScenarioInputDatasId { get; set; }
        public long? ProjectId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public double? Probability { get; set; }
      //  public string LineItem { get; set; }
        public double? NPV { get; set; }
        public List<ScenarioInputValuesViewModel> ScenarioInputValuesVM { get; set; }
    }
    public class ScenarioInputValuesViewModel
    {
        public long Id { get; set; }
      //  public long? ScenarioInputDatasId { get; set; }
        public string LineItem { get; set; }
        public double? LineItemValue { get; set; }
    }

}
