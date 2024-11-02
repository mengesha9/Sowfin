using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Abstract
{
     public  interface IMasterCostofCapitalNStructure : IEntityBaseRepository<MasterCostofCapitalNStructure>
    {
    }
    public interface ICapitalStructure_Input : IEntityBaseRepository<CapitalStructure_Input>
    {
    }
    public interface ICapitalStructure_Output : IEntityBaseRepository<CapitalStructure_Output>
    {
    }

    public interface ICostofCapital_Input : IEntityBaseRepository<CostofCapital_Input>
    {
    }

    public interface ICostofCapital_Output : IEntityBaseRepository<CostofCapital_Output>
    {
    }


    //snapshot
    public interface ISnapshot_CostofCapitalNStructure : IEntityBaseRepository<Snapshot_CostofCapitalNStructure>
    {
    }

    public interface ICapitalStructure_Snapshot : IEntityBaseRepository<CapitalStructure_Snapshot>
    {
    }

    public interface ICostofCapital_Snapshot : IEntityBaseRepository<CostofCapital_Snapshot>
    {
    }
}
