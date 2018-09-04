using System;
using System.Collections.Generic;
using System.Text;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;

namespace Dlrsoft.VBScript.Runtime
{
    public class TraceHelper : ITrace
    {
        private static TraceSource _ts = new TraceSource("Dlrsoft.VBScript");
        //private SymbolDocumentInfo _docInfo;
        private string _source;
        private int _startLine;
        private int _startColumn;
        private int _endLine;
        private int _endColumn;

        #region ITrace Members

        //public void TraceDebugInfo(SymbolDocumentInfo docInfo, int startLine, int startColumn, int endLine, int endColumn)
        public void TraceDebugInfo(string source, int startLine, int startColumn, int endLine, int endColumn)
        {
            //_docInfo = docInfo;
            _source = source;
            _startLine = startLine;
            _startColumn = startColumn;
            _endLine = endLine;
            _endColumn = endColumn;

            //_ts.TraceData(TraceEventType.Information, 1, Thread.CurrentThread.ManagedThreadId, docInfo.FileName, startLine, startColumn, endLine, endColumn);
            _ts.TraceData(TraceEventType.Information, 1, Thread.CurrentThread.ManagedThreadId, source, startLine, startColumn, endLine, endColumn);
        }

        public string Source { get { return _source; } }
        public int StartLine { get { return _startLine; } }
        public int StartColumn { get { return _startColumn; } }
        public int EndLine { get { return _endLine; } }
        public int EndColumn { get { return _endColumn; } }

        #endregion
    }
}
