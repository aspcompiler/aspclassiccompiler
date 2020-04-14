using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace Dlrsoft.Asp
{
    public class AspHostConfiguration
    {
        public IList<Assembly> Assemblies { get; set; }
        public bool Trace { get; set; }
    }
}
