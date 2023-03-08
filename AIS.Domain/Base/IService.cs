using System;
using System.Collections.Generic;
using System.Linq;
using AIS.Data;
using System.Linq.Expressions;

namespace AIS.Domain.Base
{
    public interface IService {
    }

    public interface IService<TEntity> : IService  where TEntity : Entity
    {
        TEntity Add(TEntity target);
        void AddAll(TEntity[] target);
        TEntity Update(TEntity target);
        bool UpdateAll(IEnumerable<TEntity> entites);
        void Delete(Int32 id);
        void Delete(TEntity target);
        void DeleteNow(TEntity target);
        void DeleteAll(Int32[] ids);
        void DeleteAll(TEntity[] target);
        TEntity FindById(Int32 id);
        TEntity FindByCriteria(Expression<Func<TEntity, bool>> exp);
        bool CheckExisInt32entity(string identity);
        bool CheckExisInt32entity(TEntity current, string identity);
        IEnumerable<TEntity> FindAll();
        IEnumerable<TEntity> FindByName(string name);
        IEnumerable<TEntity> FindAllByCriteria(Expression<Func<TEntity, bool>> exp);
        void SubmitChanges();
        bool AddByType(dynamic type);
        bool AddByType(List<dynamic> type);
        IQueryable GetListByType(String typeName);
    }

}
