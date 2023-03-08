using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data
{
    #region IDataContext
    public partial class LeaveManagementContext : IDataContext
    {
        public TEntity Add<TEntity>(TEntity entity) where TEntity : class
        {
            return this.Set<TEntity>().Add(entity);
        }

        public void Delete<TEntity>(TEntity entity) where TEntity : class
        {
            this.Set<TEntity>().Attach(entity);
            this.Set<TEntity>().Remove(entity);
        }

        public TEntity Find<TEntity>(Int32 id) where TEntity :  Entity
        {
            var target = Set<TEntity>().Find(id);
            return target;
        }

        public IQueryable<TEntity> FindAll<TEntity>() where TEntity : class
        {
            return this.Set<TEntity>().AsQueryable();
        }
        public TEntity Update<TEntity>(TEntity entity) where TEntity : class
        {
            SaveChanges();
            return entity;
        }
        public new void SaveChanges()
        {
            base.SaveChanges();
        }

        
    }
    #endregion
}
