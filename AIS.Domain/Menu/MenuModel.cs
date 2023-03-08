using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Domain.Menu
{
    public class Menu
    {
        public String Name { get; set; }
        public String Link { get; set; }
        public List<Menu> Children { get; set; }
        public int? ParentId { get; set; }
        public int OrderLevel { get; set; }
        public Menu()
        {
            Children = new List<Menu>();
        }

        public Menu(String Name, String Link)
        {
            this.Name = Name;
            this.Link = Link;
            Children = new List<Menu>();
        }

        public Menu(String Name , String Link,int? ParentId , int OrderLevel)
        {
            this.Name = Name;
            this.Link = Link;
            this.ParentId = ParentId;
            this.OrderLevel = OrderLevel;
            Children = new List<Menu>();
        }

        public Menu(String Name, String Link, int OrderLevel, List<Menu> Children)
        {
            this.Name = Name;
            this.Link = Link;
            this.OrderLevel = OrderLevel;
            this.Children = Children;
        }
    }
}
