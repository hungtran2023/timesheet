using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data
{
    /// <summary>
    /// Interface for Repository Base, use for almost DAL actions
    /// If you need to implement other actions, extend a new interface to implement this
    /// Example: IDirectoryRepository : IRepository
    /// </summary>
    public interface IRepository : IDisposable
    {

    }

    public interface IRepository<TEntity> : IRepository where TEntity : Entity
    {
        bool AddByType(dynamic value);
        bool AddByType(IEnumerable<dynamic> value);
        TEntity Add(TEntity target);
        void AddAll(TEntity[] us);
        TEntity Update(TEntity target);
        bool UpdateAll(IEnumerable<TEntity> entities);
        TEntity Update(Int32 id, Func<TEntity, bool> pred);
        void Delete(Int32 id);
        void Delete(TEntity target);
        void DeleteAll(Int32[] ids);
        void DeleteAll(TEntity[] targets);
        void Active(TEntity target);
        IQueryable GetListByType(String typeName);
        TEntity Find(Int32 id);
        TEntity FindByCriteria(Expression<Func<TEntity, bool>> exp);
        IEnumerable<TEntity> FindAll();
        IEnumerable<TEntity> FindAllByCriteria(Expression<Func<TEntity, bool>> exp);
        void SubmitChanges();
    }
}
