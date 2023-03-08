using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using AIS.Data;

namespace AIS.Domain.Base
{
    public abstract class Service<TEntity> : IService<TEntity> where TEntity : Entity
    {
        protected readonly IUnitOfWork _iUnitofWork;

        public Service(IUnitOfWork unitofwork)
        {
            _iUnitofWork = unitofwork;
        }

        protected IRepository<TEntity> GetRepository()
        {
            return _iUnitofWork.Repository<TEntity>();
        }
        public virtual TEntity Add(TEntity target)
        {
            return GetRepository().Add(target);
        }

        public virtual void AddAll(TEntity[] targets)
        {
            GetRepository().AddAll(targets);
        }

        public virtual TEntity Update(TEntity target)
        {
            return GetRepository().Update(target);

        }

        public virtual void Delete(Int32 id)
        {
            var found = FindById(id);
            if (found != null)
                GetRepository().Delete(found);
            else
                throw new NullReferenceException();
        }

        public virtual void Delete(TEntity target)
        {
            Delete(target.ObjId);
        }

        public virtual void DeleteNow(TEntity target)
        {
            GetRepository().Delete(target);
        }

        public void Active(TEntity target)
        {
            GetRepository().Active(target);
        }

        public void DeleteAll(Int32[] ids)
        {
            if (ids != null && ids.Length > 0)
            {
                GetRepository().DeleteAll(ids);
            }
        }

        public void DeleteAll(TEntity[] targets)
        {
            if (targets != null && targets.Length > 0)
            {
                GetRepository().DeleteAll(targets);
            }
        }

        public virtual TEntity FindById(Int32 id)
        {
            return GetRepository().Find(id);
        }

        public TEntity FindByCriteria(Expression<Func<TEntity, bool>> exp)
        {
            return GetRepository().FindByCriteria(exp);
        }

        public virtual bool CheckExisInt32entity(string identity)
        {
            return true;
        }

        public virtual bool CheckExisInt32entity(TEntity current, string identity)
        {
            return true;
        }

        public virtual IEnumerable<TEntity> FindByName(string name)
        {
            return null;
        }

        public virtual IEnumerable<TEntity> FindAll()
        {
            return GetRepository().FindAll();
        }

        public IEnumerable<TEntity> FindAllByCriteria(Expression<Func<TEntity, bool>> exp)
        {
            return GetRepository().FindAllByCriteria(exp);
        }

        public void SubmitChanges()
        {
            GetRepository().SubmitChanges();
        }

        public bool UpdateAll(IEnumerable<TEntity> entites)
        {
            GetRepository().UpdateAll(entites);
            return true;
        }
        public bool AddByType(dynamic type)
        {
            return GetRepository().AddByType(type);
        }
        public bool AddByType(List<dynamic> type)
        {
            return GetRepository().AddByType(type);
        }

        public IQueryable GetListByType(string typeName)
        {
            return GetRepository().GetListByType(typeName);
        }
    }
}
