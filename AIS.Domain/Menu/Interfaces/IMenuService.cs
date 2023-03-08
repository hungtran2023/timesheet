using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Domain.Menu
{
    public interface IMenuService {
        List<Menu> GetMenu(int Id);
    }
}
