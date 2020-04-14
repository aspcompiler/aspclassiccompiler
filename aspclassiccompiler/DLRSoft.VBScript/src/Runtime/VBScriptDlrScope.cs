using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting.Runtime;

using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif

namespace Dlrsoft.VBScript.Runtime
{
    // This class represents Sympl modules or globals scopes.  We derive from
    // DynamicObject for an easy IDynamicMetaObjectProvider implementation, and
    // just hold onto the DLR internal scope object.  Languages typically have
    // their own scope so that 1) it can flow around an a dynamic object and 2)
    // to dope in any language-specific behaviors, such as case-INsensitivity.
    //
    public sealed class VBScriptDlrScope : DynamicObject
    {
        private readonly Scope _scope;

        public VBScriptDlrScope(Scope scope)
        {
            _scope = scope;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            return _scope.TryGetVariable(SymbolTable.StringToCaseInsensitiveId(binder.Name),
                                     out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            _scope.SetVariable(SymbolTable.StringToId(binder.Name), value);
            return true;
        }
    }
}
