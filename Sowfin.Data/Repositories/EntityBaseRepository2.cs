using Sowfin.Data.Abstract;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace Sowfin.Data.Repositories
{
    public class EntityBaseRepository2<T> : IEntityBaseRepository<T>
            where T : class
    {
        private readonly FindataContext _context;

        public EntityBaseRepository2(FindataContext context)
        {
            _context = context;
        }

        public virtual void Add(T entity)
        {
            EntityEntry dbEntityEntry = _context.Entry<T>(entity);
            Console.WriteLine(entity);
            _context.Set<T>().Add(entity);

        }


        public virtual void AddMany(IEnumerable<T> entities)
        {
            _context.Set<T>().AddRange(entities);
            Commit();

        }

        public async Task<List<T>> GetAll()
        {
            return await _context.Set<T>().ToListAsync();
        }

        public virtual void Delete(T entity)
        {
            EntityEntry dbEntityEntry = _context.Entry<T>(entity);
            dbEntityEntry.State = EntityState.Deleted;
        }

        public virtual void DeleteWhere(Expression<Func<T, bool>> predicate)
        {
            IEnumerable<T> entities = _context.Set<T>().Where(predicate);

            foreach (var entity in entities)
            {
                _context.Entry<T>(entity).State = EntityState.Deleted;
            }
        }

        public virtual void DeleteMany(IEnumerable<T> entities)
        {
            foreach (var entity in entities)
            {
                _context.Entry<T>(entity).State = EntityState.Deleted;
            }
            //_context.Set<T>().AddRange(entities);
           // Commit();

        }

        public virtual void Commit()
        {
            _context.SaveChanges();
        }

        // public virtual IEnumerable<T> GetAll()
        // {
        //     return _context.Set<T>().AsEnumerable();
        // }

        public virtual int Count()
        {
            return _context.Set<T>().Count();
        }
        public virtual IEnumerable<T> AllIncluding(params Expression<Func<T, object>>[] includeProperties)
        {
            IQueryable<T> query = _context.Set<T>();

            foreach (var includeProperty in includeProperties)
            {
                query = query.Include(includeProperty);
            }
            return query.AsEnumerable();
        }

        public T GetSingle(Expression<Func<T, bool>> predicate)
        {
            return _context.Set<T>().FirstOrDefault(predicate);
        }

        public T GetSingle(Expression<Func<T, bool>> predicate, params Expression<Func<T, object>>[] includeProperties)
        {
            IQueryable<T> query = _context.Set<T>();
            foreach (var includeProperty in includeProperties)
            {
                query = query.Include(includeProperty);
            }

            return query.Where(predicate).FirstOrDefault();
        }

        //method to get latest single
        public T GetLatestSingle(Expression<Func<T, bool>> predicate)
        {
            return _context.Set<T>().LastOrDefault(predicate);
        }

        public T GetLatestSingle(Expression<Func<T, bool>> predicate, params Expression<Func<T, object>>[] includeProperties)
        {
            IQueryable<T> query = _context.Set<T>();
            foreach (var includeProperty in includeProperties)
            {
                query = query.Include(includeProperty);
            }
            return query.Where(predicate).LastOrDefault();
        }

        public virtual IEnumerable<T> FindBy(Expression<Func<T, bool>> predicate)
        {
            return _context.Set<T>().Where(predicate);
        }

        public virtual void Update(T entity)
        {
            EntityEntry dbEntityEntry = _context.Entry<T>(entity);
            dbEntityEntry.State = EntityState.Modified;
        }

        public virtual void UpdatedMany(IEnumerable<T> entities)
        {
            _context.Set<T>().UpdateRange(entities);
            Commit();
        }


        // entity state change for child
        //public virtual void SetChildEntityState(T entity, EntityState entityState)
        //{
        //    _context.Entry(entity).State = entityState;
        //    //EntityEntry dbEntityEntry = _context.Entry<T>(entity);
        //    //dbEntityEntry.State = entityState;
        //}
        //void SetMixedSubValuesEntityState(MixedSubValues mixedSubValuesObj, EntityState entityState)
        //{
        //    _context.Entry(mixedSubValuesObj).State = entityState;
        //}
    }
}