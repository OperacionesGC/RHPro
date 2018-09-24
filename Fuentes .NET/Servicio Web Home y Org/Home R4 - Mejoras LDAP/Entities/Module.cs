using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Entities
{
    [Serializable]
    public class Module
    {
        public Guid Id { get; set; }
        public string MenuTitle { get; set; }
        public string MenuDetail { get; set; }
        public string Action { get; set; }
        public string MenuObjective { get; set; }
        public string MenuObjectiveDetail { get; set; }
        public string LinkManual { get; set; }
        public string LinkDvd { get; set; }

        public string MenuName { get; set; }
        public int  pos { get; set; }
    }
}
