using Sowfin.Data.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalStructure
{
    public class MasterCostofCapitalNStructureViewModel
    {
        public long Id { get; set; }
        public int LeveragePolicyId { get; set; }
        public int SALeveragePolicyID { get; set; }
        
        public bool HasEquity { get; set; }
        public bool HasPreferredEquity { get; set; }
        public bool HasDebt { get; set; }
        public string Company { get; set; }
        public int? BetaSourceId { get; set; }
        public int CostofDebtMethodId { get; set; }
        public long? UserId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? CalculateBeta_Id { get; set; }

        public List<CapitalStructure_InputViewModel> CapitalStructureInputList { get; set; }
        public List<CapitalStructure_OutputViewModel> CapitalStructureOutputList { get; set; }
        public List<CostofCapital_InputViewModel> CostofCapitalInputList { get; set; }
        public List<CostofCapital_OutputViewModel> CostofCapitalOutputList { get; set; }
        public CalculateBetaViewModel CalculateBetaViewModel { get; set; }
        public List<Snapshot_CostofCapitalNStructureViewModel> SnapshotMasterList  { get; set; }

        public List<SelectListItem> CurrencyValueList { get; set; }
        public List<SelectListItem> NumberCountList { get; set; }
        public List<SelectListItem> ValueTypeList { get; set; }
        public List<SelectListItem> LeveragePolicyList { get; set; }
        public List<SelectListItem> HeaderList { get; set; }
        public List<SelectListItem> ScenarioAnalysisHeaderList { get; set; }
        public List<SelectListItem> BetasourceList { get; set; }
        public List<SelectListItem> SALeveragePolicyList { get; set; }
        public List<ScenarioAnalysis_OutputViewModel> ScenarioAnalysisOutputList { get; set; }
    }


    //CapitalStructure input VM
    public class CapitalStructure_InputViewModel
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public int HeaderId { get; set; }
        public string SubHeader { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }
        public string ListType { get; set; }
        public string LeverageType { get; set; }

    }

    public class CapitalStructure_OutputViewModel
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public int HeaderId { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }
    }

    public class ScenarioAnalysis_OutputViewModel
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public int HeaderId { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }
    }

    //cost of czpital
    public class CostofCapital_InputViewModel
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public int HeaderId { get; set; }
        public string SubHeader { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }

    }
         
public class CostofCapital_OutputViewModel
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public int HeaderId { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }

    }

    public class CalculateBetaViewModel
    {       
        public long Id { get; set; }
        public long? CostOfCapitals_Id { get; set; }
        public long? MasterId { get; set; }
        public long? CalculateBeta_Id { get; set; }

        public long? Frequency_Id { get; set; }
        public long? TargetMarketIndex_Id { get; set; }
        public long? TargetRiskFreeRate_Id { get; set; }
        public long? DataSource_Id { get; set; }
        public DateTime? Duration_FromDate { get; set; }
        public DateTime? Duration_toDate { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public bool Active { get; set; }
        public double? BetaValue { get; set; }
        public string FileData { get; set; }
        public string Frequency_Value { get; set; }
        public string TargetMarketIndex_Value { get; set; }
        public string TargetRiskFreeRate_Value { get; set; }
        public string DataSource_Value { get; set; }
    }


    //  snapshot  Capital Structure
    public class Snapshot_CostofCapitalNStructureViewModel
    {
        public long Id { get; set; }
        public int? MasterId { get; set; }
        public string Description { get; set; }
        public string Comment { get; set; }
        public bool Active { get; set; }
        public DateTime? CreatedDate { get; set; }
        public List<CapitalStructure_SnapshotViewModel> capitalStructure_SnapshotList { get; set; }
        public List<CostofCapital_SnapshotViewModel> CostofCapital_SnapshotList { get; set; }

    }

    public class CapitalStructure_SnapshotViewModel
    {
        public long Id { get; set; }
        public long? Snapshot_CostofCapitalNStructureId { get; set; }
        public int HeaderId { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }

    }
    public class CostofCapital_SnapshotViewModel
    {
        public long Id { get; set; }
        public long? Snapshot_CostofCapitalNStructureId { get; set; }
        public int HeaderId { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }

    }
  



}
