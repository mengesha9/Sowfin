using AutoMapper;
using Sowfin.Model;
using Sowfin.API.ViewModels.InternalValuation;
using Sowfin.API.ViewModels.FAnalysis;
using Sowfin.Model.Entities;
using Sowfin.API.ViewModels.PayoutPolicy;
using Sowfin.API.ViewModels.CapitalStructure;
using Sowfin.API.ViewModels;

namespace Sowfin.API.ViewModels.Mapping
{
    public class MappingProfile : Profile
    {
        public MappingProfile()
        {
            CreateMap<Story, StoryDetailViewModel>()
                //.ForMember(s => s.OwnerUsername, map => map.MapFrom(s => s.Owner.Username))
                .ForMember(s => s.LikesNumber, map => map.MapFrom(s => s.Likes.Count))
                .ForMember(s => s.Liked, map => map.Ignore());
            CreateMap<Story, DraftViewModel>();
            CreateMap<Story, OwnerStoryViewModel>();

            CreateMap<MixedSubDatas, Datas>();
            CreateMap<MixedSubValues, Values>();

            CreateMap<MixedSubDatas_FAnalysis, Datas>();
            CreateMap<MixedSubValues_FAnalysis, Values>();

            CreateMap<IntegratedDatas, IntegratedDatasViewModel>();
            CreateMap<IntegratedValues, IntegratedValuesViewModel>();
            CreateMap<Integrated_ExplicitValues, Integrated_ExplicitValuesViewModel>();

            CreateMap<IntegratedDatasViewModel, IntegratedDatas>();
            CreateMap<IntegratedValuesViewModel, IntegratedValues>();
            CreateMap<Integrated_ExplicitValuesViewModel, Integrated_ExplicitValues>();

            CreateMap<IntegratedDatasFAnalysis, IntegratedDatasFAnalysisViewModel>();
            CreateMap<IntegratedValuesFAnalysis, IntegratedValuesFAnalysisViewModel>();

            CreateMap<IntegratedDatasFAnalysisViewModel, IntegratedDatasFAnalysis>();
            CreateMap<IntegratedValuesFAnalysisViewModel, IntegratedValuesFAnalysis>();

            CreateMap<FinancialStatementAnalysisDatas, FinancialStatementAnalysisDatasViewModel>();
            CreateMap<FinancialStatementAnalysisValues, FinancialStatementAnalysisValuesViewModel>();

            CreateMap<FinancialStatementAnalysisDatasViewModel, FinancialStatementAnalysisDatas>();
            CreateMap<FinancialStatementAnalysisValuesViewModel, FinancialStatementAnalysisValues>();

            CreateMap<CurrentSetup, CurrentSetupViewModel>();
            CreateMap<CurrentSetupViewModel, CurrentSetup>();

            CreateMap<CurrentSetupIpDatas, CurrentSetupIpDatasViewModel>();
            CreateMap<CurrentSetupIpValues, CurrentSetupIpValuesViewModel>();
            CreateMap<CurrentSetupIpDatasViewModel, CurrentSetupIpDatas>();
            CreateMap<CurrentSetupIpValuesViewModel, CurrentSetupIpValues>();

            CreateMap<CurrentSetupSoDatas, CurrentSetupSoDatasViewModel>();
            CreateMap<CurrentSetupSoValues, CurrentSetupSoValuesViewModel>();
            CreateMap<CurrentSetupSoDatasViewModel, CurrentSetupSoDatas>();
            CreateMap<CurrentSetupSoValuesViewModel, CurrentSetupSoValues>();


            CreateMap<CurrentSetupSnapshot, CurrentSetupSnapshotViewModel>();
            CreateMap<CurrentSetupSnapshotViewModel, CurrentSetupSnapshot>();

            CreateMap<CurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasViewModel>();
            CreateMap<CurrentSetupSnapshotDatasViewModel, CurrentSetupSnapshotDatas>();

            CreateMap<CurrentSetupSnapshotValues, CurrentSetupSnapshotValuesViewModel>();
            CreateMap<CurrentSetupSnapshotValuesViewModel, CurrentSetupSnapshotValues>();

            CreateMap<CurrentSetupSoDatas, CurrentSetupSnapshotDatasViewModel>();
            CreateMap<CurrentSetupSnapshotDatasViewModel, CurrentSetupSoDatas>();
            CreateMap<CurrentSetupSoValues, CurrentSetupSnapshotValuesViewModel>();
            CreateMap<CurrentSetupSnapshotValuesViewModel, CurrentSetupSoValues>();

            CreateMap<CurrentSetupSoDatasViewModel, CurrentSetupSnapshotDatas>();
            CreateMap<CurrentSetupSoValuesViewModel, CurrentSetupSnapshotValues>();


            CreateMap<Values, IntegratedValues>().ForMember(x => x.IntegratedDatasId, opt => opt.Ignore());

            CreateMap<IntegratedDatas, ForcastRatioDatasViewModel>().ForMember(x => x.ForcastRatioValuesVM, opt => opt.Ignore());
            CreateMap<IntegratedValues, ForcastRatioValuesViewModel>().ForMember(x => x.ForcastRatioDatasId, opt => opt.Ignore());

            CreateMap<ForcastRatioDatas, ForcastRatioDatasViewModel>();
            CreateMap<ForcastRatioValues, ForcastRatioValuesViewModel>();
            CreateMap<ForcastRatio_ExplicitValues, ForcastRatio_ExplicitValuesViewModel>();

            CreateMap<ForcastRatioDatasViewModel, ForcastRatioDatas>();
            CreateMap<ForcastRatioValuesViewModel, ForcastRatioValues>();
            CreateMap<ForcastRatio_ExplicitValuesViewModel, ForcastRatio_ExplicitValues>();

            CreateMap<IntegratedDatas, ReorganizedDatasViewModel>().ForMember(x => x.ReorganizedValuesVM, opt => opt.Ignore());
            CreateMap<IntegratedValues, ReorganizedValuesViewModel>().ForMember(x => x.ReorganizedDatasId, opt => opt.Ignore());

            CreateMap<ROICDatas, ROICDatasViewModel>();
            CreateMap<ROICValues, ROICValuesViewModel>();
            CreateMap<ROIC_ExplicitValues, ROIC_ExplicitValuesViewModel>();

            CreateMap<ROICDatasViewModel, ROICDatas>();
            CreateMap<ROICValuesViewModel, ROICValues>();
            CreateMap<ROIC_ExplicitValuesViewModel, ROIC_ExplicitValues>();

            CreateMap<MarketDatasViewModel, MarketDatas>();
            CreateMap<MarketValuesViewModel, MarketValues>();

            CreateMap<MarketDatas, MarketDatasViewModel>();
            CreateMap<MarketValues, MarketValuesViewModel>();

            CreateMap<ReorganizedDatas, ReorganizedDatasViewModel>();
            CreateMap<ReorganizedValues, ReorganizedValuesViewModel>();
            CreateMap<Reorganized_ExplicitValues, Reorganized_ExplicitValuesViewModel>();

            CreateMap<InitialSetup_IValuation, Initialsetup_IValuationViewModel>();

            CreateMap<InitialSetup_FAnalysis, InitialSetup_FAnalysisViewModel>();
            CreateMap<InitialSetup_FAnalysisViewModel, InitialSetup_FAnalysis>();


            CreateMap<PayoutPolicy_ScenarioDatas, PayoutPolicy_ScenarioDatasViewModel>().ForMember(x => x.PayoutPolicy_ScenarioValuesVM, map => map.MapFrom(s => s.PayoutPolicy_ScenarioValues));
            CreateMap<PayoutPolicy_ScenarioDatasViewModel, PayoutPolicy_ScenarioDatas>().ForMember(x=>x.PayoutPolicy_ScenarioValues,map=>map.MapFrom(s=>s.PayoutPolicy_ScenarioValuesVM));

            CreateMap<PayoutPolicy_ScenarioValues, PayoutPolicy_ScenarioValuesViewModel>();
            CreateMap<PayoutPolicy_ScenarioValuesViewModel, PayoutPolicy_ScenarioValues>();

            CreateMap<CurrentSetupIpDatas, PayoutPolicy_ScenarioDatasViewModel>();
            CreateMap<CurrentSetupIpValues, PayoutPolicy_ScenarioValuesViewModel>();

            // mapping for PayoutPolicy_ScenarioOutputDatas
            CreateMap<PayoutPolicy_ScenarioOutputDatas, PayoutPolicy_ScenarioOutputDatasViewModel>().ForMember(x => x.PayoutPolicy_ScenarioOutputValuesVM, map => map.MapFrom(s => s.PayoutPolicy_ScenarioOutputValues));
            CreateMap<PayoutPolicy_ScenarioOutputDatasViewModel, PayoutPolicy_ScenarioOutputDatas>().ForMember(x => x.PayoutPolicy_ScenarioOutputValues, map => map.MapFrom(s => s.PayoutPolicy_ScenarioOutputValuesVM
            ));

            CreateMap<PayoutPolicy_ScenarioOutputValues, PayoutPolicy_ScenarioOutputValuesViewModel>();
            CreateMap<PayoutPolicy_ScenarioOutputValuesViewModel, PayoutPolicy_ScenarioOutputValues>();

            CreateMap<MasterCostofCapitalNStructureViewModel, MasterCostofCapitalNStructure>();
            CreateMap <CapitalStructure_InputViewModel, CapitalStructure_Input>();
            CreateMap<CostofCapital_InputViewModel, CostofCapital_Input>();

            CreateMap<MasterCostofCapitalNStructure, MasterCostofCapitalNStructureViewModel>();
            CreateMap<CapitalStructure_Input, CapitalStructure_InputViewModel>();
            CreateMap<CostofCapital_Input, CostofCapital_InputViewModel>();


            // cost of capital and Capital Structure Snapshot
            CreateMap<Snapshot_CostofCapitalNStructureViewModel, Snapshot_CostofCapitalNStructure>();
            CreateMap<CapitalStructure_SnapshotViewModel,CapitalStructure_Snapshot>();
            CreateMap<CostofCapital_SnapshotViewModel, CostofCapital_Snapshot>();

            CreateMap<Snapshot_CostofCapitalNStructure, Snapshot_CostofCapitalNStructureViewModel>();
            CreateMap<CapitalStructure_Snapshot, CapitalStructure_SnapshotViewModel>();
            CreateMap<CostofCapital_Snapshot, CostofCapital_SnapshotViewModel>();

            CreateMap<Project, ProjectsViewModel>();
            //CreateMap<ProjectInputDatas, ProjectInputDatasViewModel>().ForMember(x => x.ProjectInputValuesVM, map => map.MapFrom(s => s.ProjectInputValues)).ForMember(x=>x.ProjectInputComparablesVM,map=>map.MapFrom(s=>s.ProjectInputComparables));
            CreateMap<ProjectInputDatas, ProjectInputDatasViewModel>().ForMember(x => x.ProjectInputValuesVM, map => map.MapFrom(s => s.ProjectInputValues)).ForMember(x => x.ProjectInputComparablesVM, map => map.MapFrom(s => s.ProjectInputComparables)).ForMember(x => x.DepreciationInputValuesVM, map => map.MapFrom(s => s.DepreciationInputValues));
            CreateMap<ProjectInputValues, ProjectInputValuesViewModel>();
            CreateMap<ProjectInputComparables, ProjectInputComparablesViewModel>();

            CreateMap<ProjectsViewModel, Project>();
            //CreateMap<ProjectInputDatasViewModel, ProjectInputDatas>().ForMember(x => x.ProjectInputValues, map => map.MapFrom(s => s.ProjectInputValuesVM)).ForMember(x => x.ProjectInputComparables, map => map.MapFrom(s => s.ProjectInputComparablesVM));
            CreateMap<ProjectInputDatasViewModel, ProjectInputDatas>().ForMember(x => x.ProjectInputValues, map => map.MapFrom(s => s.ProjectInputValuesVM)).ForMember(x => x.ProjectInputComparables, map => map.MapFrom(s => s.ProjectInputComparablesVM)).ForMember(x => x.DepreciationInputValues, map => map.MapFrom(s => s.DepreciationInputValuesVM));
            CreateMap<ProjectInputValuesViewModel, ProjectInputValues>();
            CreateMap<ProjectInputComparablesViewModel, ProjectInputComparables>();

            //CreateMap<SensitivityInputs_ProjectViewModel, SensitivityInputData>();
            //CreateMap<SensitivityInputData, SensitivityInputs_ProjectViewModel>();

            //CreateMap<DepreciationInputDatasViewModel, DepreciationInputDatas>().ForMember(x => x.DepreciationInputValues, map => map.MapFrom(s => s.DepreciationInputValuesVM));
            //CreateMap<DepreciationInputDatas, DepreciationInputDatasViewModel>().ForMember(x => x.DepreciationInputValuesVM, map => map.MapFrom(s => s.DepreciationInputValues));

            //CreateMap<Story, StoryViewModel>()
            //.ForMember(s => s.OwnerUsername, map => map.MapFrom(s => s.Owner.Username));

            CreateMap<ScenarioInputDatas, ScenarioInputDatasViewModel>().ForMember(x => x.ScenarioInputValuesVM, map => map.MapFrom(s => s.ScenarioInputValues));
            CreateMap<ScenarioInputValues, ScenarioInputValuesViewModel>();

            CreateMap<ScenarioInputDatasViewModel, ScenarioInputDatas>().ForMember(x => x.ScenarioInputValues, map => map.MapFrom(s => s.ScenarioInputValuesVM));
            CreateMap<ScenarioInputValuesViewModel, ScenarioInputValues>();
        }
    }
}