using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class ProjectRepository : EntityBaseRepository2<Project>, IProject
    {
        public ProjectRepository(FindataContext context) : base(context) { }
    }
    public class ProjectInputDatasRepository : EntityBaseRepository2<ProjectInputDatas>, IProjectInputDatas
    {
        public ProjectInputDatasRepository(FindataContext context) : base(context) { }
    }
    public class ProjectInputValuesRepository : EntityBaseRepository2<ProjectInputValues>, IProjectInputValues
    {
        public ProjectInputValuesRepository(FindataContext context) : base(context) { }
    }
    
         public class ProjectInputComparablesRepository : EntityBaseRepository2<ProjectInputComparables>, IProjectInputComparables
    {
        public ProjectInputComparablesRepository(FindataContext context) : base(context) { }
    }
    public class Project_SnapshotDatasRepository : EntityBaseRepository2<Project_SnapshotDatas>, IProject_SnapshotDatas
    {
        public Project_SnapshotDatasRepository(FindataContext context) : base(context) { }
    }
    public class Project_SnapshotValuesRepository : EntityBaseRepository2<Project_SnapshotValues>, IProject_SnapshotValues
    {
        public Project_SnapshotValuesRepository(FindataContext context) : base(context) { }
    }
    //old
    public class ProjectsRepository : EntityBaseRepository2<Projects>, IProjectRepository
    {
        public ProjectsRepository(FindataContext context) : base(context) { }
    }

    public class Project_SnapshotRepository : EntityBaseRepository2<Snapshot_CapitalBudgeting>, ISnapshot_CapitalBudgeting
    {
        public Project_SnapshotRepository(FindataContext context) : base(context) { }
    }

    public class DepreciationInputDatasRepository : EntityBaseRepository2<DepreciationInputDatas>, IDepreciationInputDatas
    {
        public DepreciationInputDatasRepository(FindataContext context) : base(context) { }
    }
    public class DepreciationInputValuesRepository : EntityBaseRepository2<DepreciationInputValues>, IDepreciationInputValues
    {
        public DepreciationInputValuesRepository(FindataContext context) : base(context) { }
    }
    public class SensitivityInputDataRepository : EntityBaseRepository2<SensitivityInputData>, ISensitivityInputData
    {
        public SensitivityInputDataRepository(FindataContext context) : base(context) { }
    }
    //public class SensitivityInputDataRepository : EntityBaseRepository2<SensitivityInputs>, ISensitivityInputData
    //{
    //    public SensitivityInputDataRepository(FindataContext context) : base(context) { }
    //}
    public class ScenarioInputDataRepository : EntityBaseRepository2<ScenarioInputDatas>, IScenarioInputData
    {
        public ScenarioInputDataRepository(FindataContext context) : base(context) { }
    }
    public class ScenarioInputValuesRepository : EntityBaseRepository2<ScenarioInputValues>, IScenarioInputValues
    {
        public ScenarioInputValuesRepository(FindataContext context) : base(context) { }
    }
    public class SensitivityInputGraphDataRepository : EntityBaseRepository2<SensitivityInputGraphData>, ISensitivityInputGraphData
    {
        public SensitivityInputGraphDataRepository(FindataContext context) : base(context) { }
    }

    public class CapitalStructureScenarioSnapshotRepository : EntityBaseRepository2<CapitalStructureScenarioSnapshot>, ICapitalStructureScenarioSnapshot
    {
        public CapitalStructureScenarioSnapshotRepository(FindataContext context) : base(context) { }
    }
}
