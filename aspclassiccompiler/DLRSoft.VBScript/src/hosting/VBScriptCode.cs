using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Microsoft.Scripting.Runtime;

using System.Dynamic;
#if USE35
// Needed for type language implementers need to support DLR Hosting, such as
// ScriptCode.
using Microsoft.Scripting.Ast;
using Microsoft.Scripting.Utils;
#else
using System.Linq.Expressions;
#endif
using Microsoft.Scripting;
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScript.Hosting
{
    class VBScriptCode : ScriptCode
    {
        //private readonly Expression<Action<VBScript, IDynamicMetaObjectProvider>> _lambda;
        private readonly VBScript _vbscript;
        private Action<VBScript, IDynamicMetaObjectProvider> _compiledLambda;

        public VBScriptCode(
             VBScript vbscript,
             Action<VBScript, IDynamicMetaObjectProvider> lambda,
             SourceUnit sourceUnit)
             : base(sourceUnit) {
             _compiledLambda = lambda;
            _vbscript = vbscript;
        }

        public override object Run() {
            return Run(new Scope());
        }

        public override object Run(Scope scope) {
            //if (_compiledLambda == null) {
            //    _compiledLambda = _lambda.Compile();
            //}
            //LC Load the system by default
            //RuntimeHelpers.VBScriptImport(_vbscript, module, new string[] { "System" }, new string[] { }, new string[] { });
            //RuntimeHelpers.VBScriptImport(_vbscript, module, new string[] { "VBScript", "Runtime" }, new string[] { "BuiltInFunctions" }, new string[] { "BuiltIn" });

            //if (this.SourceUnit.Kind == SourceCodeKind.File)
            //{
            //    // Simple way to convey script rundir for RuntimeHelpers.SymplImport
            //    // to load .sympl files relative to the current script file.
            //    DynamicObjectHelpers.SetMember(module, "__file__",
            //                                   Path.GetFullPath(this.SourceUnit.Path));
            //}
            _compiledLambda(_vbscript, scope);
            return null;
        }
    }
}
