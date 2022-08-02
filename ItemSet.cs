using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Daniel_s_UI_prototype
{
    class ItemSet
    {
        private IList<string> items;
        private float support = 0;
        private IList<string> ids;

        public IList<string> Ids { get => ids; set => ids = value; }
        public IList<string> Items { get => items; set => items = value; }
        public float Support { get => support; set => support = value; }
    }
}
