using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data
{
    public class UnitOfWork : IUnitOfWork
    {
        private IDataContext Context { get; set; }
        private Dictionary<string, object> repositories;
        private bool disposed = false;

        public UnitOfWork(IDataContext context)
        {
            Context = context;
            if (repositories == null)
            {
                repositories = new Dictionary<string, object>();
            }
        }

        private object GetInstance(Type type)
        {
            var repositoryInstance = Activator.CreateInstance(
                            typeof(Repository<>).MakeGenericType(type), Context);
            return repositoryInstance;
        }
        

        public IRepository<TEntity> Repository<TEntity>() where TEntity : Entity
        {
            var type = typeof(TEntity);
            if (!repositories.ContainsKey(type.Name))
            {
                repositories.Add(type.Name, GetInstance(type));
            }
            return (IRepository<TEntity>)repositories[type.Name];
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern.
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                if (repositories != null && repositories.Count > 0)
                {
                    foreach (var item in repositories)
                    {
                        if (item.Value != null)
                            ((IRepository)item.Value).Dispose();
                    }
                }
            }
            disposed = true;
        }
    }
}
