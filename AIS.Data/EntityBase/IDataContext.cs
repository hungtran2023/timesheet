using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data
{
    public interface IDataContext
    {
        TEntity Add<TEntity>(TEntity entity) where TEntity : class;
        void Delete<TEntity>(TEntity entity) where TEntity : class;
        TEntity Find<TEntity>(Int32 id) where TEntity : Entity;
        TEntity Update<TEntity>(TEntity entity) where TEntity : class;
        IQueryable<TEntity> FindAll<TEntity>() where TEntity : class;
        void SaveChanges();
    }
}
