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
using Microsoft.Scripting.Utils;

namespace Dlrsoft.VBScript.Compiler
{
    // AnalysisScope holds identifier information so that we can do name binding
    // during analysis.  It manages a map from names to ParameterExprs so ET
    // definition locations and reference locations can alias the same variable.
    //
    // These chain from inner most BlockExprs, through LambdaExprs, to the root
    // which models a file or top-level expression.  The root has non-None
    // ModuleExpr and RuntimeExpr, which are ParameterExprs.
    //
    internal class AnalysisScope
    {
        private AnalysisScope _parent;
        private string _name;
        // Need runtime for interning VBScript constants at code gen time.
        private VBScript _runtime;
        private ParameterExpression _runtimeParam;
        private ParameterExpression _moduleParam;
        private ParameterExpression _traceParam;
        private ISourceMapper _mapper;
        // Need IsLambda when support return to find tightest closing fun.
        private bool _isLambda = false;
        private bool _isLambdaBody = false;
        private bool _isDoLoop = false;
        private bool _isForLoop = false;
        private bool _isWith = false;
        private bool _isOptionExplicitOn = false;
        private bool _isOnErrorResumeNextOn = false;
        private LabelTarget _loopBreak = null;
        private LabelTarget _continueBreak = null;
        private LabelTarget _methodExit = null;
        private Dictionary<string, ParameterExpression> _names;
        private Set<string> _functionTable;
        private IList<VBScriptSyntaxError> _errors;
        private Expression _withExpression;
        private IDictionary<string, SymbolDocumentInfo> _docInfos;

        public AnalysisScope(AnalysisScope parent, string name)
            : this(parent, name, null, null, null, null) { }

        public AnalysisScope(AnalysisScope parent,
                              string name,
                              VBScript runtime,
                              ParameterExpression runtimeParam,
                              ParameterExpression moduleParam,
                              ISourceMapper mapper)
        {
            _parent = parent;
            _name = name;
            _runtime = runtime;
            _runtimeParam = runtimeParam;
            _moduleParam = moduleParam;
            _mapper = mapper;

            if (_moduleParam != null)
            {
                _functionTable = new Set<string>(StringComparer.InvariantCultureIgnoreCase);
                _errors = new List<VBScriptSyntaxError>();
                _docInfos = new Dictionary<string, SymbolDocumentInfo>();
            }

            _names = new Dictionary<string, ParameterExpression>();
        }

        public AnalysisScope Parent { get { return _parent; } }

        public ParameterExpression ModuleExpr { get { return _moduleParam; } }

        public ParameterExpression RuntimeExpr { get { return _runtimeParam; } }

        public VBScript Runtime { get { return _runtime; } }

        public ISourceMapper SourceMapper { get { return _mapper; } }

        public string Name { get { return _name; } }

        public bool IsModule { get { return _moduleParam != null; } }

        public bool IsOptionExplicitOn
        {
            get { return _isOptionExplicitOn; }
            set { _isOptionExplicitOn = value; }
        }

        public bool IsOnErrorResumeNextOn
        {
            get { return _isOnErrorResumeNextOn; }
            set { _isOnErrorResumeNextOn = value; }
        }

        public bool IsLambda
        {
            get { return _isLambda; }
            set { _isLambda = value; }
        }


        public bool IsLambdaBody
        {
            get { return _isLambdaBody; }
            set { _isLambdaBody = value; }
        }

        public bool IsDoLoop
        {
            get { return _isDoLoop; }
            set { _isDoLoop = value; }
        }

        public bool IsForLoop
        {
            get { return _isForLoop; }
            set { _isForLoop = value; }
        }

        public bool IsWith
        {
            get { return _isWith; }
            set { _isWith = value; }
        }

        public Expression WithExpression
        {
            get { return _withExpression;  }
            set { _withExpression = value; }
        }

        public Expression NearestWithExpression
        {
            get
            {
                var curScope = this;
                while (!curScope.IsWith && !curScope.IsLambdaBody && !curScope.IsModule) //With cannot across those boundaries
                {
                    curScope = curScope.Parent;
                }
                return curScope.WithExpression;
            }
        }

        public Expression ErrExpression
        {
            get
            {
                return this.ModuleScope.Names[Dlrsoft.VBScript.VBScript.ERR_PARAMETER];
            }
        }

        public Expression TraceExpression
        {
            get
            {
                return this.ModuleScope.Names[Dlrsoft.VBScript.VBScript.TRACE_PARAMETER];
            }
        }

        public LabelTarget LoopBreak
        {
            get { return _loopBreak; }
            set { _loopBreak = value; }
        }

        public LabelTarget LoopContinue
        {
            get { return _continueBreak; }
            set { _continueBreak = value; }
        }

        public LabelTarget MethodExit
        {
            get { return _methodExit; }
            set { _methodExit = value; }
        }

        public Dictionary<string, ParameterExpression> Names
        {
            get { return _names; }
            set { _names = value; }
        }

        public ParameterExpression GetModuleExpr()
        {
            var curScope = this;
            while (!curScope.IsModule)
            {
                curScope = curScope.Parent;
            }
            return curScope.ModuleExpr;
        }


        public AnalysisScope ModuleScope
        {
            get
            {
                var curScope = this;
                while (!curScope.IsModule)
                {
                    curScope = curScope.Parent;
                }
                return curScope;
            }
        }

        public AnalysisScope VariableScope
        {
            get
            {
                var curScope = this;
                while (!curScope.IsModule && !curScope.IsLambdaBody)
                {
                    curScope = curScope.Parent;
                }
                return curScope;
            }
        }

        public Set<string> FunctionTable
        {
            get
            {
                return ModuleScope._functionTable;
            }
        }

        public IList<VBScriptSyntaxError> Errors
        {
            get { return _errors; }
        }

        public VBScript GetRuntime()
        {
            var curScope = this;
            while (curScope.Runtime == null)
            {
                curScope = curScope.Parent;
            }
            return curScope.Runtime;
        }

        //Can only call it on ModuleScope
        public SymbolDocumentInfo GetDocumentInfo(string path)
        {
            SymbolDocumentInfo docInfo;
            if (_docInfos.ContainsKey(path))
            {
                docInfo = _docInfos[path];
            }
            else
            {
                docInfo = Expression.SymbolDocument(path);
                _docInfos.Add(path, docInfo);
            }
            return docInfo;
        }

        public void GetVariablesInScope(List<string> names, List<ParameterExpression> parameters)
        {
            AnalysisScope scope = this;
            while (!scope.IsModule)
            {
                foreach (string name in _names.Keys)
                {
                    names.Add(name);
                    parameters.Add(_names[name]);
                }
                scope = scope.Parent;
            }

            //Now we are dealing with module scope. We want to exclude lambda variables
            Set<string> functionTable = FunctionTable;
            foreach (string name in _names.Keys)
            {
                if (!functionTable.Contains(name))
                {
                    names.Add(name);
                    parameters.Add(_names[name]);
                }
            }
        }
    } //AnalysisScope}
}