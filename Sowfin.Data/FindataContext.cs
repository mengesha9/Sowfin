using Sowfin.Model;
using Microsoft.EntityFrameworkCore;
using Sowfin.Model.Entities;
using System.Linq;

namespace Sowfin.Data
{
    public class FindataContext : DbContext
    {

        // Key less tables 
        public DbSet<EdgarData> EdgarData { get; set; }
        public DbSet<EdgarDataByCategory> EdgarDataByCategory { get; set; }
        public DbSet<ProcessedData> ProcessedData { get; set; }



        //public DbSet<Filing> Filing { get; set; }
        public DbSet<CapitalAnalysisSnapshot> CapitalAnalysisSnapshot { get; set; }
        public DbSet<Findata> FinData { get; set; }
        public DbSet<HistoricalTable> HistoricalTable { get; set; }
        public DbSet<LineItemInfo> LineItemInfo { get; set; }
        public DbSet<BusinessUnit> BusinessUnits { get; set; }
        public DbSet<CapitalAnalysis> CapitalAnalyses { get; set; }
        public DbSet<CapitalBudgeting> CapitalBudgetings { get; set; }
        public DbSet<CapitalBugetingTables> CapitalBugetingTables { get; set; }
        public DbSet<CapitalStructure> CapitalStructures { get; set; }
        public DbSet<CostOfCapital> CostOfCapitals { get; set; }
        public DbSet<Integrated_ExplicitValues> Integrated_ExplicitValues { get; set; }
        public DbSet<CalculateBeta> CalculateBeta { get; set; }
        public DbSet<MixedSubDatas_FAnalysis> MixedSubDatas_FAnalysis { get; set; }
        public DbSet<MixedSubValues_FAnalysis> MixedSubValues_FAnalysis { get; set; }
        public DbSet<MixedSubDatas> MixedSubDatas { get; set; }
        public DbSet<MixedSubValues> MixedSubValues { get; set; }
        public DbSet<ForcastRatioDatas> ForcastRatioDatas { get; set; }
        public DbSet<ForcastRatioValues> ForcastRatioValues { get; set; }
        public DbSet<ForcastRatio_ExplicitValues> ForcastRatio_ExplicitValues { get; set; }

        public DbSet<FinancialStatementAnalysisDatas> FinancialStatementAnalysisDatas { get; set; }
        public DbSet<FinancialStatementAnalysisValues> FinancialStatementAnalysisValues { get; set; }

        public DbSet<CurrentSetup> CurrentSetup { get; set; }
        public DbSet<CurrentSetupIpDatas> CurrentSetupIpDatas { get; set; }
        public DbSet<CurrentSetupIpValues> CurrentSetupIpValues { get; set; }

        public DbSet<CurrentSetupSoDatas> CurrentSetupSoDatas { get; set; }
        public DbSet<CurrentSetupSoValues> CurrentSetupSoValues { get; set; }


        public DbSet<CurrentSetupSnapshot> CurrentSetupSnapshot { get; set; }
        public DbSet<CurrentSetupSnapshotDatas> CurrentSetupSnapshotDatas { get; set; }
        public DbSet<CurrentSetupSnapshotValues> CurrentSetupSnapshotValues { get; set; }


        public DbSet<CIKStatus> CIKStatus { get; set; }
        public DbSet<FAnalysis_CategoryByInitialSetup> FAnalysis_CategoryByInitialSetup { get; set; }
        
        public DbSet<ReorganizedDatas> ReorganizedDatas { get; set; }
        public DbSet<ReorganizedValues> ReorganizedValues { get; set; }
        public DbSet<Reorganized_ExplicitValues> Reorganized_ExplicitValues { get; set; }

        public DbSet<IVSensitivity> IVSensitivity { get; set; }
        public DbSet<IVScenario> IVScenario { get; set; }


        public DbSet<ROICDatas> ROICDatas { get; set; }
        public DbSet<ROICValues> ROICValues { get; set; }
        public DbSet<ROIC_ExplicitValues> ROIC_ExplicitValues { get; set; }

        public DbSet<MarketDatas> MarketDatas { get; set; }
        public DbSet<MarketValues> MarketValues { get; set; }


        public DbSet<AssetsEquityDatas> AssetsEquityDatas { get; set; }
        public DbSet<ValuationDatas> ValuationDatas { get; set; }

        public DbSet<IntegratedDatas> IntegratedDatas { get; set; }
        public DbSet<IntegratedValues> IntegratedValues { get; set; }

        public DbSet<IntegratedDatasFAnalysis> IntegratedDatasFAnalysis { get; set; }
        public DbSet<IntegratedValuesFAnalysis> IntegratedValuesFAnalysis { get; set; }


        //  public DbSet<Synonyms> Synonyms { get; set; }
        public DbSet<RawHistoricalValues> RawHistoricalValues { get; set; }
        public DbSet<HistoryElementMapping> HistoryElementMapping { get; set; }
        public DbSet<IntegratedElementMaster> IntegratedElementMaster { get; set; }
        public DbSet<IntegratedFinancialStmt> IntegratedFinancialStmt { get; set; }
        public DbSet<History> Histories { get; set; }
        public DbSet<CostOfCapital_IValuation> CostOfCapital_IValuation { get; set; }
        public DbSet<InitialSetup_IValuation> InitialSetup_IValuation { get; set; }

        public DbSet<InitialSetup_FAnalysis> InitialSetup_FAnalysis { get; set; }

        public DbSet<Interest_IValuation> Interest_IValuation { get; set; }
        public DbSet<PayoutPolicy_IValuation> PayoutPolicy_IValuation { get; set; }
        public DbSet<TaxRates_IValuation> TaxRates_IValuation { get; set; }
        public DbSet<HistoryAnalysisAndForcastRatio> HistoryAnalysisAndForcastRatio { get; set; }
        public DbSet<ForcastRatioElementMaster> ForcastRatioElementMaster { get; set; }
        public DbSet<ExplicitPeriod_HistoryForcastRatio> ExplicitPeriod_HistoryForcastRatio { get; set; }
        public DbSet<PayoutPolicy> PayoutPolicies { get; set; }
        public DbSet<Projects> Projects { get; set; }
        public DbSet<Snapshots> Snapshots { get; set; }
        public DbSet<User> Users { get; set; }
        public DbSet<FilingsTable> Filings { get; set; }
        public DbSet<ValuationTechnique> ValuationTechniques { get; set; }
        public DbSet<CategoryByInitialSetup> CategoryByInitialSetup { get; set; }

        public DbSet<PayoutPolicy_ScenarioDatas> PayoutPolicy_ScenarioDatas { get; set; }
        public DbSet<PayoutPolicy_ScenarioValues> PayoutPolicy_ScenarioValues { get; set; } 
        public DbSet<PayoutPolicy_ScenarioOutputDatas> PayoutPolicy_ScenarioOutputDatas { get; set; }
        public DbSet<PayoutPolicy_ScenarioOutputValues> PayoutPolicy_ScenarioOutputValues { get; set; }
        public DbSet<MasterCostofCapitalNStructure> MasterCostofCapitalNStructure { get; set; }
        public DbSet<CapitalStructure_Input> CapitalStructure_Input { get; set; }
        public DbSet<CapitalStructure_Output> CapitalStructure_Output { get; set; }
        public DbSet<CostofCapital_Input> CostofCapital_Input { get; set; }
        public DbSet<CostofCapital_Output> CostofCapital_Output { get; set; }
        public DbSet<Snapshot_CostofCapitalNStructure> Snapshot_CostofCapitalNStructure { get; set; }
        public DbSet<CapitalStructure_Snapshot> CapitalStructure_Snapshot { get; set; }
        public DbSet<CostofCapital_Snapshot> CostofCapital_Snapshot { get; set; }

        public DbSet<Project> Project { get; set; }
        public DbSet<ProjectInputDatas> ProjectInputDatas { get; set; }
        public DbSet<ProjectInputValues> ProjectInputValues { get; set; }

        public DbSet<DepreciationInputDatas> DepreciationInputDatas { get; set; }
        public DbSet<DepreciationInputValues> DepreciationInputValues { get; set; }

        public DbSet<ProjectInputComparables> ProjectInputComparables { get; set; }        
        public DbSet<Project_SnapshotDatas> Project_SnapshotDatas { get; set; }
        public DbSet<Project_SnapshotValues> Project_SnapshotValues { get; set; }
        public DbSet<Snapshot_CapitalBudgeting> Snapshot_CapitalBudgetings { get; set; }
         public DbSet<SensitivityInputData> SensitivityInputData { get; set; }
        // public DbSet<SensitivityInputs> SensitivityInputs { get; set; }

        public DbSet<ScenarioInputDatas> ScenarioInputDatas { get; set; }
        public DbSet<ScenarioInputValues> ScenarioInputValues { get; set; }

        public DbSet<SensitivityInputGraphData> SensitivityGraphInputData { get; set; }
        public DbSet<CapitalStructureScenarioSnapshot> CapitalStructureScenarioSnapshot { get; set; }

        // public FindataContext() { }
        public FindataContext(DbContextOptions<FindataContext> options) : base(options)
        {
            Database.Migrate();// change EnsureCreated to Migrate
        }



        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            foreach (var relationship in modelBuilder.Model.GetEntityTypes().SelectMany(e => e.GetForeignKeys()))
            {
                relationship.DeleteBehavior = DeleteBehavior.Restrict;
            }

            ConfigureModelBuilderForFiling(modelBuilder);

            ConfigureModelBuilderForStatement(modelBuilder);

            ConfigureModelBuilderForSynonym(modelBuilder);

            ConfigureModelBuilderForFindata(modelBuilder);

            ConfigureModelBuilderForYear(modelBuilder);

            ConfigureModelBuilderForHistoricalTable(modelBuilder);

            ConfigureModelBuilderForEdgarData(modelBuilder);

            ConfigureModelBuilderForEdgarIntegratedData(modelBuilder);

            ConfigureModelBuilderForEdgarDataByCategory(modelBuilder);

            ConfigureModelBuilderForProcessedData(modelBuilder);

            ConfigureModelBuilderLineItemIfo(modelBuilder);

            /////
            ConfigureModelBuilderBusinessUnits(modelBuilder);

            ConfigureModelBuilderForCapitalAnalysis(modelBuilder);

            ConfigureModelBuilderForCapitalBudgetings(modelBuilder);

            ConfigureModelBuilderForCapitalBugetingTables(modelBuilder);

            ConfigureModelBuilderForCapitalStructures(modelBuilder);

            ConfigureModelBuilderForCostOfCapitals(modelBuilder);
            ConfigureModelBuilderForIntegrated_ExplicitValues(modelBuilder);

            ConfigureModelBuilderForMixedSubDatas_FAnalysis(modelBuilder);
            ConfigureModelBuilderForMixedSubValues_FAnalysis(modelBuilder);
            ConfigureModelBuilderForMixedSubDatas(modelBuilder);

            ConfigureModelBuilderForMixedSubValues(modelBuilder);

            ConfigureModelBuilderForCIKStatus(modelBuilder);

            ConfigureModelBuilderForFAnalysis_CategoryByInitialSetup(modelBuilder);

            ConfigureModelBuilderForForcastRatioDatas(modelBuilder);

            ConfigureModelBuilderForForcastRatioValues(modelBuilder);

            ConfigureModelBuilderForForcastRatio_ExplicitValues(modelBuilder);


            ConfigureModelBuilderForFinancialStatementAnalysisDatas(modelBuilder);

            ConfigureModelBuilderForFinancialStatementAnalysisValues(modelBuilder);

            ConfigureModelBuilderForCurrentSetup(modelBuilder);

            ConfigureModelBuilderForCurrentSetupIpDatas(modelBuilder);
            ConfigureModelBuilderForCurrentSetupIpValues(modelBuilder);

            ConfigureModelBuilderForCurrentSetupSoDatas(modelBuilder);
            ConfigureModelBuilderForCurrentSetupSoValues(modelBuilder);


            ConfigureModelBuilderForCurrentSetupSnapshot(modelBuilder);

            ConfigureModelBuilderForCurrentSetupSnapshotDatas(modelBuilder);
            ConfigureModelBuilderForCurrentSetupSnapshotValues(modelBuilder);

            ConfigureModelBuilderForReorganizedDatas(modelBuilder);
            ConfigureModelBuilderForReorganizedValues(modelBuilder);
            ConfigureModelBuilderForReorganized_ExplicitValues(modelBuilder);

            ConfigureModelBuilderForIVSensitivity(modelBuilder);
            ConfigureModelBuilderForIVScenario(modelBuilder);

            ConfigureModelBuilderForROICDatas(modelBuilder);
            ConfigureModelBuilderForROICValues(modelBuilder);
            ConfigureModelBuilderForROIC_ExplicitValues(modelBuilder);

            ConfigureModelBuilderForMarketDatas(modelBuilder);
            ConfigureModelBuilderForMarketValues(modelBuilder);



            ConfigureModelBuilderForAssetsEquityDatas(modelBuilder);
            ConfigureModelBuilderForIntegratedDatas(modelBuilder);
            ConfigureModelBuilderForIntegratedDatasFAnalysis(modelBuilder);
            ConfigureModelBuilderForValuationDatas(modelBuilder);
            ConfigureModelBuilderForIntegratedValues(modelBuilder);
            ConfigureModelBuilderForIntegratedValuesFAnalysis(modelBuilder);


            ConfigureModelBuilderForCalculateBeta(modelBuilder);
            // ConfigureModelBuilderForDatas(modelBuilder);
            // ConfigureModelBuilderForValues(modelBuilder);

            ConfigureModelBuilderForRawHistoricalValues(modelBuilder);

            ConfigureModelBuilderForHistoryElementMapping(modelBuilder);

            ConfigureModelBuilderForIntegratedElementMaster(modelBuilder);

            ConfigureModelBuilderForIntegratedFinancialStmt(modelBuilder);

            ConfigureModelBuilderForHistory(modelBuilder);

            ConfigureModelBuilderForCostOfCapital_IValuation(modelBuilder);

            ConfigureModelBuilderForInitialSetup_IValuation(modelBuilder);
            ConfigureModelBuilderForInitialSetup_FAnalysis(modelBuilder);


            ConfigureModelBuilderForInterest_IValuation(modelBuilder);

            ConfigureModelBuilderForPayoutPolicy_IValuation(modelBuilder);
            ConfigureModelBuilderForFilings(modelBuilder);

            ConfigureModelBuilderForTaxRates_IValuation(modelBuilder);

            ConfigureModelBuilderForHistoryAnalysisAndForcastRatio(modelBuilder);

            ConfigureModelBuilderForForcastRatioElementMaster(modelBuilder);

            ConfigureModelBuilderForExplicitPeriod_HistoryForcastRatio(modelBuilder);

            ConfigureModelBuilderForPayoutPolicies(modelBuilder);

            ConfigureModelBuilderForProjects(modelBuilder);

            ConfigureModelBuilderForSnapshots(modelBuilder);

            ConfigureModelBuilderForUsers(modelBuilder);

            ConfigureModelBuilderForValuationTechniques(modelBuilder);

            ConfigureModelBuilderForCategoryByInitialSetup(modelBuilder);

            ConfigureModelBuilderForPayoutPolicy_ScenarioDatas(modelBuilder);

            ConfigureModelBuilderForPayoutPolicy_ScenarioValues(modelBuilder);

            ConfigureModelBuilderForPayoutPolicy_ScenarioOutputDatas(modelBuilder);

            ConfigureModelBuilderForPayoutPolicy_ScenarioOutputValues(modelBuilder);

            ConfigureModelBuilderForMasterCostofCapitalNStructure(modelBuilder);

            ConfigureModelBuilderForCapitalStructure_Input(modelBuilder);

            ConfigureModelBuilderForCapitalStructure_Output(modelBuilder);

            ConfigureModelBuilderForCostofCapital_Input(modelBuilder);

            ConfigureModelBuilderForCostofCapital_Output(modelBuilder);

            ConfigureModelBuilderForSnapshot_CostofCapitalNStructure(modelBuilder);

            ConfigureModelBuilderForCapitalStructure_Snapshot(modelBuilder);

            ConfigureModelBuilderForCostofCapital_Snapshot(modelBuilder);

            ConfigureModelBuilderForProject(modelBuilder);

            ConfigureModelBuilderForProjectInputDatas(modelBuilder);

            ConfigureModelBuilderForDepreciationInputDatas(modelBuilder);
            ConfigureModelBuilderForDepreciationInputValues(modelBuilder);

            ConfigureModelBuilderForProjectInputValues(modelBuilder);
            ConfigureModelBuilderForProjectInputComparables(modelBuilder);
            

            ConfigureModelBuilderForProject_SnapshotDatas(modelBuilder);

            ConfigureModelBuilderForProject_SnapshotValues(modelBuilder);

            ConfigureModelBuilderForSnapshot_CapitalBudgeting(modelBuilder);

            ConfigureModelBuilderForSensitivityInputData(modelBuilder);

            ConfigureModelBuilderForScenarioInputDatas(modelBuilder);
            ConfigureModelBuilderForScenarioInputValues(modelBuilder);

            ConfigureModelBuilderForSensitivityInputGraphData(modelBuilder);

            ConfigureModelBuilderForCapitalStructureScenarioSnapshot(modelBuilder);
            ConfigureModelBuilderForCapitalAnalysisSnapshot(modelBuilder);
        }

        void ConfigureModelBuilderForCapitalAnalysisSnapshot(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalAnalysisSnapshot>().ToTable("capitalAnalysisSnapshot");
        }

        void ConfigureModelBuilderForFiling(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Filing>().ToTable("filing");
        }

        void ConfigureModelBuilderForStatement(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Statement>().ToTable("statement");
        }

        void ConfigureModelBuilderForSynonym(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Synonyms>().ToTable("Synonyms");
        }

        void ConfigureModelBuilderForFindata(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Findata>().ToTable("temp_filing");
        }

        void ConfigureModelBuilderForYear(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Years>();
        }

        void ConfigureModelBuilderForHistoricalTable(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<HistoricalTable>().ToTable("histriocal_table");
        }

        void ConfigureModelBuilderLineItemIfo(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<LineItemInfo>().ToTable("Datas");
        }

        void ConfigureModelBuilderForRawHistoricalValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<RawHistoricalValues>().ToTable("Values");
        }

        void ConfigureModelBuilderForEdgarData(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<EdgarData>().HasNoKey(); // changed from modlebuilder.query<EdgarData>() 
        }

        void ConfigureModelBuilderForEdgarIntegratedData(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IntegratedView>().HasNoKey(); // changed from modlebuilder.query<IntegratedView>() 
        }

        void ConfigureModelBuilderForEdgarDataByCategory(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<EdgarDataByCategory>().HasNoKey(); // changed from modlebuilder.query<EdgarDataByCategory>() 
        }

        void ConfigureModelBuilderForProcessedData(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ProcessedData>().HasNoKey(); // changed from modlebuilder.query<ProcessedData>()
        }

        void ConfigureModelBuilderBusinessUnits(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<BusinessUnit>();
        }

        void ConfigureModelBuilderForCapitalAnalysis(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalAnalysis>();
        }

        void ConfigureModelBuilderForCapitalBudgetings(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalBudgeting>();
        }

        void ConfigureModelBuilderForCapitalBugetingTables(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalBugetingTables>();
        }

        void ConfigureModelBuilderForCapitalStructures(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalStructure>();
        }

        void ConfigureModelBuilderForCostOfCapitals(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CostOfCapital>();
        }

        void ConfigureModelBuilderForIntegrated_ExplicitValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Integrated_ExplicitValues>();
        }
        //void ConfigureModelBuilderForDatas(ModelBuilder modelBuilder)
        //{
        //    modelBuilder.Entity<DatasTable>().ToTable("Datas"); ;
        //}
        void ConfigureModelBuilderForCalculateBeta(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CalculateBeta>();
        }
        //void ConfigureModelBuilderForValues(ModelBuilder modelBuilder)
        //{
        //    modelBuilder.Entity<ValuesTable>().ToTable("Values");
        //}

        void ConfigureModelBuilderForMixedSubDatas_FAnalysis(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MixedSubDatas_FAnalysis>();
        }
        void ConfigureModelBuilderForMixedSubValues_FAnalysis(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MixedSubValues_FAnalysis>();
        }
        void ConfigureModelBuilderForMixedSubDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MixedSubDatas>();
        }
        void ConfigureModelBuilderForCIKStatus(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CIKStatus>();
        }
        void ConfigureModelBuilderForFAnalysis_CategoryByInitialSetup(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<FAnalysis_CategoryByInitialSetup>();
        }

        void ConfigureModelBuilderForForcastRatioDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ForcastRatioDatas>();
        }
        void ConfigureModelBuilderForForcastRatioValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ForcastRatioValues>();
        }
        void ConfigureModelBuilderForForcastRatio_ExplicitValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ForcastRatio_ExplicitValues>();
        }

        void ConfigureModelBuilderForFinancialStatementAnalysisDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<FinancialStatementAnalysisDatas>();
        }
        void ConfigureModelBuilderForFinancialStatementAnalysisValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<FinancialStatementAnalysisValues>();
        }

        void ConfigureModelBuilderForCurrentSetup(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetup>();
        }

        void ConfigureModelBuilderForCurrentSetupIpDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetupIpDatas>();
        }
        void ConfigureModelBuilderForCurrentSetupIpValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetupIpValues>();
        }
        void ConfigureModelBuilderForCurrentSetupSnapshot(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetupSnapshot>();
        }
        void ConfigureModelBuilderForCurrentSetupSnapshotDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetupSnapshotDatas>();
        }
        void ConfigureModelBuilderForCurrentSetupSnapshotValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetupSnapshotValues>();
        }



        void ConfigureModelBuilderForCurrentSetupSoDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetupSoDatas>();
        }
        void ConfigureModelBuilderForCurrentSetupSoValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CurrentSetupSoValues>();
        }

        void ConfigureModelBuilderForReorganizedDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ReorganizedDatas>();
        }
        void ConfigureModelBuilderForIVSensitivity(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IVSensitivity>();
        }
        void ConfigureModelBuilderForIVScenario(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IVScenario>();
        }
        
        void ConfigureModelBuilderForReorganizedValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ReorganizedValues>();
        }
        void ConfigureModelBuilderForReorganized_ExplicitValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Reorganized_ExplicitValues>();
        }

        void ConfigureModelBuilderForROICDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ROICDatas>();
        }

        void ConfigureModelBuilderForROICValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ROICValues>();
        }

        void ConfigureModelBuilderForMarketDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MarketDatas>();
        }
        void ConfigureModelBuilderForMarketValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MarketValues>();
        }
        void ConfigureModelBuilderForROIC_ExplicitValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ROIC_ExplicitValues>();
        }
        void ConfigureModelBuilderForAssetsEquityDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<AssetsEquityDatas>();
        }
        void ConfigureModelBuilderForValuationDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ValuationDatas>();
        }
        void ConfigureModelBuilderForIntegratedDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IntegratedDatas>();
        }
        void ConfigureModelBuilderForIntegratedDatasFAnalysis(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IntegratedDatasFAnalysis>();
        }
        void ConfigureModelBuilderForIntegratedValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IntegratedValues>();
        }

        void ConfigureModelBuilderForIntegratedValuesFAnalysis(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IntegratedValuesFAnalysis>();
        }

        void ConfigureModelBuilderForMixedSubValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MixedSubValues>();
        }

        void ConfigureModelBuilderForHistoryElementMapping(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<HistoryElementMapping>();
        }
        void ConfigureModelBuilderForIntegratedElementMaster(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IntegratedElementMaster>();
        }
        void ConfigureModelBuilderForIntegratedFinancialStmt(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<IntegratedFinancialStmt>();
        }
        void ConfigureModelBuilderForHistory(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<History>();
        }
        void ConfigureModelBuilderForCostOfCapital_IValuation(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CostOfCapital_IValuation>();
        }
        void ConfigureModelBuilderForInitialSetup_IValuation(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<InitialSetup_IValuation>();
        }

        void ConfigureModelBuilderForInitialSetup_FAnalysis(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<InitialSetup_FAnalysis>();
        }

        void ConfigureModelBuilderForInterest_IValuation(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Interest_IValuation>();
        }
        void ConfigureModelBuilderForPayoutPolicy_IValuation(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PayoutPolicy_IValuation>();
        }
        void ConfigureModelBuilderForFilings(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<FilingsTable>().ToTable("Filings");
        } //IntegratedFilings
        void ConfigureModelBuilderForTaxRates_IValuation(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<TaxRates_IValuation>();
        }
        void ConfigureModelBuilderForHistoryAnalysisAndForcastRatio(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<HistoryAnalysisAndForcastRatio>();
        }
        void ConfigureModelBuilderForForcastRatioElementMaster(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ForcastRatioElementMaster>();
        }
        void ConfigureModelBuilderForExplicitPeriod_HistoryForcastRatio(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ExplicitPeriod_HistoryForcastRatio>();
        }
        void ConfigureModelBuilderForPayoutPolicies(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PayoutPolicy>();
        }

        void ConfigureModelBuilderForProjects(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Projects>();
        }
        void ConfigureModelBuilderForSnapshots(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Snapshots>();
        }
        void ConfigureModelBuilderForUsers(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<User>();
        }

        void ConfigureModelBuilderForValuationTechniques(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ValuationTechnique>();
        }
        void ConfigureModelBuilderForCategoryByInitialSetup(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CategoryByInitialSetup>();
        }
        void ConfigureModelBuilderForPayoutPolicy_ScenarioDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PayoutPolicy_ScenarioDatas>();
        }

        void ConfigureModelBuilderForPayoutPolicy_ScenarioValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PayoutPolicy_ScenarioValues>();
        }

        void ConfigureModelBuilderForPayoutPolicy_ScenarioOutputDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PayoutPolicy_ScenarioOutputDatas>();
        }
        void ConfigureModelBuilderForPayoutPolicy_ScenarioOutputValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PayoutPolicy_ScenarioOutputValues>();
        }
        void ConfigureModelBuilderForMasterCostofCapitalNStructure(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MasterCostofCapitalNStructure>();
        }
        void ConfigureModelBuilderForCapitalStructure_Input(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalStructure_Input>();
        }
        void ConfigureModelBuilderForCapitalStructure_Output(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalStructure_Output>();
        }
        void ConfigureModelBuilderForCostofCapital_Input(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CostofCapital_Input>();
        }
        void ConfigureModelBuilderForCostofCapital_Output(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CostofCapital_Output>();
        }

        void ConfigureModelBuilderForSnapshot_CostofCapitalNStructure(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Snapshot_CostofCapitalNStructure>();
        }
        void ConfigureModelBuilderForCapitalStructure_Snapshot(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalStructure_Snapshot>();
        }
        void ConfigureModelBuilderForCostofCapital_Snapshot(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CostofCapital_Snapshot>();
        }

        void ConfigureModelBuilderForProject(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Project>();
        }
        void ConfigureModelBuilderForProjectInputDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ProjectInputDatas>();
        }
        void ConfigureModelBuilderForDepreciationInputDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<DepreciationInputDatas>();
        }

        void ConfigureModelBuilderForProjectInputValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ProjectInputValues>();
        }
        void ConfigureModelBuilderForDepreciationInputValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<DepreciationInputValues>();
        }

        void ConfigureModelBuilderForProjectInputComparables(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ProjectInputComparables>();
        }

        
        void ConfigureModelBuilderForProject_SnapshotDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Project_SnapshotDatas>();
        }
        void ConfigureModelBuilderForProject_SnapshotValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Project_SnapshotValues>();
        }

        void ConfigureModelBuilderForSnapshot_CapitalBudgeting(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Snapshot_CapitalBudgeting>();
        }

        void ConfigureModelBuilderForSensitivityInputData(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<SensitivityInputData>();
        }
        //void ConfigureModelBuilderForSensitivityInputData(ModelBuilder modelBuilder)
        //{
        //    modelBuilder.Entity<SensitivityInputs>();
        //}

        void ConfigureModelBuilderForScenarioInputDatas(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ScenarioInputDatas>();
        }
        void ConfigureModelBuilderForScenarioInputValues(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ScenarioInputValues>();
        }
        void ConfigureModelBuilderForSensitivityInputGraphData(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<SensitivityInputGraphData>();
        }

        void ConfigureModelBuilderForCapitalStructureScenarioSnapshot(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CapitalStructureScenarioSnapshot>();
        }

    }
}
