using Sowfin.Data.Abstract;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Data.Abstract
{
    public interface IBusinessUnit : IEntityBaseRepository<BusinessUnit>
    {
        
    }
    public interface IProject : IEntityBaseRepository<Project>
    {
    }

    public interface IProjectInputDatas : IEntityBaseRepository<ProjectInputDatas>
    {
    }

    public interface IProjectInputValues : IEntityBaseRepository<ProjectInputValues>
    {
    }

    public interface IProjectInputComparables : IEntityBaseRepository<ProjectInputComparables>
    {
    }


    

    public interface IProject_SnapshotDatas : IEntityBaseRepository<Project_SnapshotDatas>
    {
    }

    public interface IProject_SnapshotValues : IEntityBaseRepository<Project_SnapshotValues>
    {
    }

    public interface IDepreciationInputDatas : IEntityBaseRepository<DepreciationInputDatas>
    {
    }
    public interface IDepreciationInputValues : IEntityBaseRepository<DepreciationInputValues>
    {
    }

}
