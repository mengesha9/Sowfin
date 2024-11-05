using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;


namespace Sowfin.Data.Abstract
{
    public interface IEntityBaseRepository<T>  where T : class
    {
        Task<List<T>>  GetAll();
        void Add(T entity);
        void AddMany(IEnumerable<T> entities);
        void Delete(T entity);
        void DeleteWhere(Expression<Func<T, bool>> predicate);
        void Commit();
        IEnumerable<T> AllIncluding(params Expression<Func<T, object>>[] includeProperties);
        T GetSingle(Expression<Func<T, bool>> predicate);
        T GetSingle(Expression<Func<T, bool>> predicate, params Expression<Func<T, object>>[] includeProperties);

        // get Latest single
        T GetLatestSingle(Expression<Func<T, bool>> predicate);
        T GetLatestSingle(Expression<Func<T, bool>> predicate, params Expression<Func<T, object>>[] includeProperties);
        IEnumerable<T> FindBy(Expression<Func<T, bool>> predicate);
        void Update(T entity);
        void UpdatedMany(IEnumerable<T> entities);
        void DeleteMany(IEnumerable<T> entities);

        //entity state change for child
        //void SetChildEntityState(T entity, EntityState entityState);


    }
}