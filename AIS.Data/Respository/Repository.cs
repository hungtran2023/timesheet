using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;

namespace AIS.Data
{

    public class Repository<TEntity> : IRepository<TEntity> where TEntity : Entity
    {
        private bool disposed = false;
        protected LeaveManagementContext Context { get; set; }
        public Repository(LeaveManagementContext context)
        {
            Context = context;
        }
        public virtual TEntity Add(TEntity target)
        {
            this.Context.Add(target);
            this.Context.SaveChanges();
            return target;
        }

        public virtual void AddAll(TEntity[] Entities)
        {
            foreach (var item in Entities)
            {
                this.Context.Add(item);
            }
            try
            {
                this.Context.SaveChanges();
            }
            catch (Exception ex)
            {
                
            }
           
        }

        public virtual TEntity Update(TEntity entity)
        {
            var updater = Find(entity.ObjId);
            updater.Copy(entity);
            this.Context.SaveChanges();
            return entity;
        }

        public virtual bool UpdateAll(IEnumerable<TEntity> entities)
        {
            foreach (var item in entities)
            {
                this.Context.Update<TEntity>(item);
            }
            this.Context.SaveChanges();
            return true;
        }

        public virtual TEntity Update(Int32 id, Func<TEntity, bool> pred)
        {
            if (pred == null)
            {
                throw new ArgumentNullException();
            }
            TEntity saving = Find(id);
            if (pred(saving))
            {
                try
                {
                    this.Context.SaveChanges();
                }
                catch (Exception ex)
                {
                }
            }
            else
            {
                throw new System.Data.DataException();
            }
            return saving;
        }

        public virtual void Delete(TEntity target)
        {
            //this.Entities.Remove(u);
            this.Context.Delete(target);
            this.Context.SaveChanges();
        }

        public virtual void Delete(Int32 id)
        {
            var target = Find(id);
            if (target != null)
            {
                Delete(target);
            }
        }

        public virtual void DeleteAll(Int32[] ids)
        {
            IEnumerable<TEntity> deleting = FindAllByCriteria(c => ids.Contains(c.ObjId));

            foreach (var item in deleting)
            {
                this.Context.Delete(item);
            }

            this.Context.SaveChanges();
        }

        public virtual void DeleteAll(TEntity[] us)
        {
            foreach (var item in us)
            {
                this.Context.Delete(item);
            }

            this.Context.SaveChanges();
        }

        public virtual void Active(TEntity u)
        {

        }

        public virtual TEntity Find(Int32 id)
        {
            return this.Context.Find<TEntity>(id);
        }

        public virtual TEntity FindByCriteria(Expression<Func<TEntity, bool>> exp)
        {
            return this.Context.FindAll<TEntity>().Where(exp).FirstOrDefault();
        }

        public IEnumerable<TEntity> FindAll()
        {
            return this.Context.FindAll<TEntity>();
        }
        public IEnumerable<TEntity> FindAllByCriteria(Expression<Func<TEntity, bool>> exp)
        {
            return FindAll().Where(exp.Compile()).AsEnumerable();
        }
        public void SubmitChanges()
        {
            this.Context.SaveChanges();
        }
        public IQueryable GetListByType(String typeName) {
            IQueryable result = null;
            try
            {
                result = this.Context.Set(Type.GetType(typeName)).AsQueryable();
            }
            catch (Exception)
            {
            }
            return result;
        }
        public bool AddByType(dynamic value) {
            try
            {
                this.Context.Set(value.GetType()).Add(value);
                Context.SaveChanges();
            }
            catch (Exception)
            {
            }
            return false;
        }
        public bool AddByType(IEnumerable<dynamic> value)
        {
            try
            {
                this.Context.Set(value.First().GetType()).AddRange(value);
                Context.SaveChanges();
                return true;
            }
            catch (Exception)
            {
            }
            return false;
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                Context.Dispose();
            }
            disposed = true;
        }
    }
}
