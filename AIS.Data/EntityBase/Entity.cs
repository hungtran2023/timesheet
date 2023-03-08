using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data
{

    public abstract class Entity
    {
        public virtual Int32 ObjId { get { throw new NotImplementedException(); } }

        public void Copy(Entity entity)
        {
            var prop = entity.GetType().GetProperties().Where(t => !t.GetType().IsGenericType && t.Name != "Id" && t.Name != "ObjId" && !t.PropertyType.Namespace.Contains("AIS"));
            foreach (var item in prop)
            {
                this.GetType().GetProperty(item.Name).SetValue(this, item.GetValue(entity));
            }
        }
    }
}
