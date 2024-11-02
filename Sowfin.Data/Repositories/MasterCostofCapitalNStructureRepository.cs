using Sowfin.Data;
using Sowfin.Data.Repositories;
using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Repositories
{
    //Cost of capital and Capital Structure  Master
      public class MasterCostofCapitalNStructureRepository : EntityBaseRepository2<MasterCostofCapitalNStructure>, IMasterCostofCapitalNStructure
    {
        public MasterCostofCapitalNStructureRepository(FindataContext context) : base(context) { }

    }

    //capital Structure input
    public class CapitalStructure_InputRepository : EntityBaseRepository2<CapitalStructure_Input>, ICapitalStructure_Input
    {
        public CapitalStructure_InputRepository(FindataContext context) : base(context) { }

    }

    //capital Structure Output
    public class CapitalStructure_OutputRepository : EntityBaseRepository2<CapitalStructure_Output>, ICapitalStructure_Output
    {
        public CapitalStructure_OutputRepository(FindataContext context) : base(context) { }

    }

    //cost of capital input
    public class CostofCapital_InputRepository : EntityBaseRepository2<CostofCapital_Input>, ICostofCapital_Input
    {
        public CostofCapital_InputRepository(FindataContext context) : base(context) { }

    }

    //cost of capital output
    public class CostofCapital_OutputRepository : EntityBaseRepository2<CostofCapital_Output>, ICostofCapital_Output
    {
        public CostofCapital_OutputRepository(FindataContext context) : base(context) { }

    }

    //  Master Cost of Capital and capital Structure 
    public class Snapshot_CostofCapitalNStructureRepository : EntityBaseRepository2<Snapshot_CostofCapitalNStructure>, ISnapshot_CostofCapitalNStructure
    {
        public Snapshot_CostofCapitalNStructureRepository(FindataContext context) : base(context) { }

    }

    //capital Structure Snapshot
    public class CapitalStructure_SnapshotRepository : EntityBaseRepository2<CapitalStructure_Snapshot>, ICapitalStructure_Snapshot
    {
        public CapitalStructure_SnapshotRepository(FindataContext context) : base(context) { }

    }


    //  Cost of capital Snapshot
    public class CostofCapital_SnapshotRepository : EntityBaseRepository2<CostofCapital_Snapshot>, ICostofCapital_Snapshot
    {
        public CostofCapital_SnapshotRepository(FindataContext context) : base(context) { }

    }




}
