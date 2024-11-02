using System.Text;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Sowfin.Data.Abstract;
using Sowfin.Data.Repositories;


namespace Sowfin.Data
{
    public static class SowfinDataConfigurationRegistration
    {
        public static IServiceCollection AddDataConfigurationRegistration(
        this IServiceCollection services,
        IConfiguration configuration
        )
        {


            




            services.AddScoped<IFilingRepository, FilingRepository>();
            services.AddScoped<IStatementRepository, StatementRepository>();
            services.AddScoped<ISynonymRepository, SynonymRepository>();
            services.AddScoped<IFindataRepository, FindataRepository>();
            services.AddScoped<IYearsRepository, YearsRepository>();
            services.AddScoped<IHistriocalTable, HistriocalTableRepository>();
            services.AddScoped<IEdgarDataRepository, EdgarDataRepository>();
            services.AddScoped<IProcessedDataRepository, ProcessedDataRepository>();
            services.AddScoped<ILineItemInfoRepository, LineItemInfoRepository>();

            ///from old framework
            services.AddScoped<IBusinessUnit, BusinessUnitRepository>();
            services.AddScoped<ICapitalAnalysisSnapshots, CapitalAnalysisSnapshotRepository>();
            services.AddScoped<IProjectRepository, ProjectsRepository>();
            services.AddScoped<ICapitalBudgeting, CapitalBudgetingRepository>();
            services.AddScoped<ICapitalBugetingTables, CapitalBugetingTablesRepository>();
            services.AddScoped<ISnapshots, SnapshotsRepository>();
            services.AddScoped<ICapitalStructure, CapitalStructureRepository>();
            services.AddScoped<ICostOfCapital, CostOfCapitalRepository>();
            services.AddScoped<IInitialSetup_IValuation, InitialSetup_IValuationRepository>();
            services.AddScoped<IInitialSetup_FAnalysis, InitialSetup_FAnalysisRepository>();

            services.AddScoped<IInterest_IValuation, Interest_IValuationRepository>();
            services.AddScoped<IPayoutPolicy_IValuation, PayoutPolicy_IValuationRepository>();
            services.AddScoped<ITaxRates_IValuation, TaxRates_IValuationRepository>();
            services.AddScoped<ICostOfCapital_IValuation, CostOfCapital_IValuationRepository>();
            services.AddScoped<ICalculateBeta, CalculateBetaRepository>();
            services.AddScoped<IRawHistoricalValues, RawHistoricalValuesRepository>();
            services.AddScoped<IHistoryElementMapping, HistoryElementMappingRepository>();
            services.AddScoped<IIntegratedElementMaster, IntegratedElementMasterRepository>();
            services.AddScoped<IIntegratedFinancialStmt, IntegratedFinancialStmtRepository>();
            services.AddScoped<IHistory, HistoryRepository>();
            services.AddScoped<IPayoutPolicy, PayoutPolicyRepository>();
            services.AddScoped<IUserRepository, UserRepository>();
            services.AddScoped<IValuationTechniqueRepository, ValuationTechniqueRepository>();
            services.AddScoped<IHistoryAnalysisAndForcastRatio, HistoryAnalysisAndForcastRatioRepository>();
            services.AddScoped<IForcastRatioElementMaster, ForcastRatioElementMasterRepository>();
            services.AddScoped<IExplicitPeriod_HistoryForcastRatio, ExplicitPeriod_HistoryForcastRatioRepository>();

            services.AddScoped<IMixedSubDatas, MixedSubDatasRepository>();
            services.AddScoped<IMixedSubValues, MixedSubValuesRepository>();

            services.AddScoped<IIntegratedDatas, IntegratedDatasRepository>();
            services.AddScoped<IIntegratedValues, IntegratedValuesRepository>();

            services.AddScoped<IIntegratedDatasFAnalysis, IntegratedDatasFAnalysisRepository>();
            services.AddScoped<IIntegratedValuesFAnalysis, IntegratedValuesFAnalysisRepository>();

            services.AddScoped<IFilings, FilingsRepository>();
            services.AddScoped<ICIKStatus, CIKStatusRepository>();
            services.AddScoped<IForcastRatioDatas, ForcastRatioDatasRepository>();
            services.AddScoped<IForcastRatioValues, ForcastRatioValuesRepository>();
            services.AddScoped<IForcastRatio_ExplicitValues, ForcastRatio_ExplicitValuesRepository>();

            services.AddScoped<IFinancialStatementAnalysisDatas, FinancialStatementAnalysisDatasRepository>();
            services.AddScoped<IFinancialStatementAnalysisValues, FinancialStatementAnalysisValuesRepository>();

            services.AddScoped<ICurrentSetup, CurrentSetupRepository>();
            

            services.AddScoped<ICurrentSetupIpDatas, CurrentSetupIpDatasRepository>();
            services.AddScoped<ICurrentSetupIpValues, CurrentSetupIpValuesRepository>();

            services.AddScoped<ICurrentSetupSoDatas, CurrentSetupSoDatasRepository>();
            services.AddScoped<ICurrentSetupSoValues, CurrentSetupSoValuesRepository>();

            services.AddScoped<ICurrentSetupSnapshot, CurrentSetupSnapshotRepository>();
            services.AddScoped<ICurrentSetupSnapshotDatas, CurrentSetupSnapshotDatasRepository>();
            services.AddScoped<ICurrentSetupSnapshotValues, CurrentSetupSnapshotValuesRepository>();

            services.AddScoped<IReorganizedDatas, ReorganizedDatasRepository>();
            services.AddScoped<IReorganizedValues, ReorganizedValuesRepository>();
            services.AddScoped<IReorganized_ExplicitValues, Reorganized_ExplicitValuesRepository>();

            services.AddScoped<IIVSensitivity, IVSensitivityRepository>();
            services.AddScoped<IIVScenario, IVScenarioRepository>();

            services.AddScoped<IROICDatas, ROICDatasRepository>();
            services.AddScoped<IROICValues, ROICValuesRepository>();
            services.AddScoped<IROIC_ExplicitValues, ROIC_ExplicitValuesRepository>();

            services.AddScoped<IMarketDatas, MarketDatasRepository>();
            services.AddScoped<IMarketValues, MarketValuesRepository>();

            services.AddScoped<IMixedSubDatas_FAnalysis, MixedSubDatas_FAnalysisRepository>();
            services.AddScoped<IMixedSubValues_FAnalysis, MixedSubValues_FAnalysisRepository>();

            services.AddScoped<IAssetsEquityDatas, AssetsEquityDatasRepository>();
            services.AddScoped<IValuationDatas, ValuationDatasRepository>();
            services.AddScoped<IDatas, DatasRepository>();
            services.AddScoped<IValues, ValuesRepository>();
            services.AddScoped<IIntegrated_ExplicitValues, Integrated_ExplicitValuesRepository>();

            services.AddScoped<ICategoryByInitialSetup, CategoryByInitialSetupRepository>();
            services.AddScoped<IFAnalysis_CategoryByInitialSetup, FAnalysis_CategoryByInitialSetupRepository>();

            services.AddScoped<IPayoutPolicy_ScenarioDatas, PayoutPolicy_ScenarioDatasRepository>();
            services.AddScoped<IPayoutPolicy_ScenarioValues, PayoutPolicy_ScenarioValuesRepository>();
            
            services.AddScoped<IPayoutPolicy_ScenarioOutputDatas, PayoutPolicy_ScenarioOutputDatasRepository>();
            services.AddScoped<IPayoutPolicy_ScenarioOutputValues, PayoutPolicy_ScenarioOutputValuesRepository>();

            services.AddScoped<IMasterCostofCapitalNStructure, MasterCostofCapitalNStructureRepository>();
            services.AddScoped<ICapitalStructure_Input, CapitalStructure_InputRepository>();
            services.AddScoped<ICapitalStructure_Output, CapitalStructure_OutputRepository>();
            services.AddScoped<ICostofCapital_Input, CostofCapital_InputRepository>();
            services.AddScoped<ICostofCapital_Output, CostofCapital_OutputRepository>();

            services.AddScoped<ISnapshot_CostofCapitalNStructure, Snapshot_CostofCapitalNStructureRepository>();
            services.AddScoped<ICapitalStructure_Snapshot, CapitalStructure_SnapshotRepository>();
            services.AddScoped<ICostofCapital_Snapshot, CostofCapital_SnapshotRepository>();

            services.AddScoped<IProject, ProjectRepository>();
            services.AddScoped<IProjectInputDatas, ProjectInputDatasRepository>();
            services.AddScoped<IProjectInputValues, ProjectInputValuesRepository>();

            services.AddScoped<IDepreciationInputDatas, DepreciationInputDatasRepository>();
            services.AddScoped<IDepreciationInputValues, DepreciationInputValuesRepository>();

            services.AddScoped<IProjectInputComparables, ProjectInputComparablesRepository>();
            
            services.AddScoped<IProject_SnapshotDatas, Project_SnapshotDatasRepository>();
            services.AddScoped<IProject_SnapshotValues, Project_SnapshotValuesRepository>();
            services.AddScoped<ISnapshot_CapitalBudgeting, Project_SnapshotRepository>();
            services.AddScoped<ISensitivityInputData, SensitivityInputDataRepository>();

            services.AddScoped<IScenarioInputData, ScenarioInputDataRepository>();
            services.AddScoped<IScenarioInputValues, ScenarioInputValuesRepository>();

            services.AddScoped<ISensitivityInputGraphData, SensitivityInputGraphDataRepository>();
            services.AddScoped<ICapitalStructureScenarioSnapshot, CapitalStructureScenarioSnapshotRepository>();

        
            return services;

        }

    }
}