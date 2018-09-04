using System;
using System.Collections.Generic;
//using System.Text;
using System.Reflection;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
using Microsoft.Scripting.Utils;
#else
using System.Linq.Expressions;
#endif
using System.Runtime.CompilerServices;

namespace Dlrsoft.VBScript.Runtime
{
    // This class is needed to canonicalize InvokeMemberBinders in Sympl.  See
    // the comment above the GetXXXBinder methods at the end of the Sympl class.
    //
    public class InvokeMemberBinderKey
    {
        string _name;
        CallInfo _info;

        public InvokeMemberBinderKey(string name, CallInfo info)
        {
            _name = name;
            _info = info;
        }

        public string Name { get { return _name; } }
        public CallInfo Info { get { return _info; } }

        public override bool Equals(object obj)
        {
            InvokeMemberBinderKey key = obj as InvokeMemberBinderKey;
            // Don't lower the name.  Sympl is case-preserving in the metadata
            // in case some DynamicMetaObject ignores ignoreCase.  This makes
            // some interop cases work, but the cost is that if a Sympl program
            // spells ".foo" and ".Foo" at different sites, they won't share rules.
            return key != null && key._name == _name && key._info.Equals(_info);
        }

        public override int GetHashCode()
        {
            // Stolen from DLR sources when it overrode GetHashCode on binders.
            return 0x28000000 ^ _name.GetHashCode() ^ _info.GetHashCode();
        }

    } //InvokeMemberBinderKey
}
