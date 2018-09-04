using System;
using System.Collections.Generic;
using System.Text;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif
using System.Runtime.CompilerServices;

namespace Dlrsoft.VBScript.Runtime
{
    public interface ITrace
    {
        //void TraceDebugInfo(SymbolDocumentInfo docInfo, int startLine, int startColumn, int endLine, int endColumn);
        void TraceDebugInfo(string source, int startLine, int startColumn, int endLine, int endColumn);
        //void RegisterRuntimeVariables(string[] varNames, IRuntimeVariables runtimeVariables);
    }
}
