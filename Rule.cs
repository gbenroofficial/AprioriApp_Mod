using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Windows.Forms;
using System.Collections;
using System.Reflection;


namespace Daniel_s_UI_prototype
{
    class Rule
    {
        private ItemSet Antecedent;
        private ItemSet Consequent;
        private float Lift = 0;
        private float Confidence = 0;
        private float Support = 0;

        public ItemSet antecedent { get => Antecedent; set => Antecedent = value; }

        public IList<string> getAnte { get => Antecedent.Items; set => Antecedent.Items = value; }
        public IList<string> getConse { get => Consequent.Items; set => Consequent.Items = value; }
        public ItemSet consequent { get => Consequent; set => Consequent = value; }
        public float lift { get => Lift; set => Lift = value; }
        public float confidence { get => Confidence; set => Confidence = value; }
        public float support { get => Support; set => Support = value; }
    }
}
