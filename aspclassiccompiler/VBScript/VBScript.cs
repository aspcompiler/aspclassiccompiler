using System;
using System.Collections.Generic;
using System.Text;
using VB = Dlrsoft.VBScript.Parser;
using Microsoft.Scripting;
using Microsoft.Scripting.Runtime;
//using Microsoft.Scripting.Debugging.CompilerServices;
//using Microsoft.Scripting.Debugging;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
using Microsoft.Scripting.Utils;
#else
using System.Linq;
using System.Linq.Expressions;
#endif

using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.IO;
using Dlrsoft.VBScript.Binders;
using Dlrsoft.VBScript.Compiler;
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScript
{
    public class VBScript
    {
        public const string ERR_PARAMETER = "err";
        public const string TRACE_PARAMETER = "__trace";

        private IList<Assembly> _assemblies;
        private ExpandoObject _globals = new ExpandoObject();
        private Scope _dlrGlobals;

        public VBScript(IList<Assembly> assms, Scope dlrGlobals)
        {
            _assemblies = assms;
            _dlrGlobals = dlrGlobals;
            AddAssemblyNamesAndTypes();
        }

		// _addNamespacesAndTypes builds a tree of ExpandoObjects representing
		// .NET namespaces, with TypeModel objects at the leaves.  Though Sympl is
		// case-insensitive, we store the names as they appear in .NET reflection
		// in case our globals object or a namespace object gets passed as an IDO
		// to another language or library, where they may be looking for names
		// case-sensitively using EO's default lookup.
		//
        public void AddAssemblyNamesAndTypes() {
            foreach (var assm in _assemblies) {
                foreach (var typ in assm.GetExportedTypes()) {
                    string[] names = typ.FullName.Split('.');
                    var table = _globals;
                    for (int i = 0; i < names.Length - 1; i++) {
                        string name = names[i].ToLower();
                        if (DynamicObjectHelpers.HasMember(
                               (IDynamicMetaObjectProvider)table, name)) {
                            // Must be Expando since only we have put objs in
                            // the tables so far.
                            table = (ExpandoObject)(DynamicObjectHelpers
                                                       .GetMember(table, name));
                        } else {
                            var tmp = new ExpandoObject();
                            DynamicObjectHelpers.SetMember(table, name, tmp);
                            table = tmp;
                        }
                    }
                    DynamicObjectHelpers.SetMember(table, names[names.Length - 1],
                                                   new TypeModel(typ));
                }
            }
        }

		// ExecuteFile executes the file in a new module scope and stores the
		// scope on Globals, using either the provided name, globalVar, or the
		// file's base name.  This function returns the module scope.
		//
        public IDynamicMetaObjectProvider ExecuteFile(string filename) {
            return ExecuteFile(filename, null);
        }

        public IDynamicMetaObjectProvider ExecuteFile(string filename,
                                                      string globalVar) {
            var moduleEO = CreateScope();
            ExecuteFileInScope(filename, moduleEO);

            globalVar = globalVar ?? Path.GetFileNameWithoutExtension(filename);
            DynamicObjectHelpers.SetMember(this._globals, globalVar, moduleEO);

            return moduleEO;
        }

		// ExecuteFileInScope executes the file in the given module scope.  This
		// does NOT store the module scope on Globals.  This function returns
		// nothing.
		//
        public void ExecuteFileInScope(string filename,
                                       IDynamicMetaObjectProvider moduleEO) {
            var f = new StreamReader(filename);
            // Simple way to convey script rundir for RuntimeHelpes.SymplImport
            // to load .sympl files.
            DynamicObjectHelpers.SetMember(moduleEO, "__file__", 
                                           Path.GetFullPath(filename));
            try {
                var moduleFun = ParseFileToLambda(filename, f);
                var d = moduleFun;
                d(this, moduleEO);
            } finally {
                f.Close();
            }
        }

        internal Action<VBScript, IDynamicMetaObjectProvider>
                 ParseFileToLambda(string filename, TextReader reader) {
            var scanner = new VB.Scanner(reader);
            var errorTable = new List<VB.SyntaxError>();

            var block = new VB.Parser().ParseScriptFile(scanner, errorTable);

            if (errorTable.Count > 0)
            {
                List<VBScriptSyntaxError> errors = new List<VBScriptSyntaxError>();
                foreach (VB.SyntaxError error in errorTable)
                {
                    errors.Add(new VBScriptSyntaxError(
                        filename,
                        SourceUtil.ConvertSpan(error.Span),
                        (int)error.Type,
                        error.Type.ToString())
                    );
                }
                throw new VBScriptCompilerException(errors);
            }

            VBScriptSourceCodeReader sourceReader = reader as VBScriptSourceCodeReader;
            ISourceMapper mapper = null;
            if (sourceReader != null)
            {
                mapper = sourceReader.SourceMapper;
            }

            var scope = new AnalysisScope(
                null,
                filename,
                this,
                Expression.Parameter(typeof(VBScript), "vbscriptRuntime"),
                Expression.Parameter(typeof(IDynamicMetaObjectProvider), "fileModule"),
                mapper);
            
            //Generate function table
            List<Expression> body = new List<Expression>();

            //Add the built in globals
            ParameterExpression err = Expression.Parameter(typeof(ErrObject), ERR_PARAMETER);
            scope.Names.Add(ERR_PARAMETER, err);
            body.Add(
                Expression.Assign(
                    err,
                    Expression.New(typeof(ErrObject))
                )
            );

            if (Trace)
            {
                ParameterExpression trace = Expression.Parameter(typeof(ITrace), TRACE_PARAMETER);
                scope.Names.Add(TRACE_PARAMETER, trace);
                body.Add(
                    Expression.Assign(
                        trace,
                        Expression.Convert(
                            Expression.Dynamic(
                               scope.GetRuntime().GetGetMemberBinder(TRACE_PARAMETER),
                               typeof(object),
                               scope.GetModuleExpr()
                            ),
                            typeof(ITrace)
                        )
                    )
                );
            }

            //Put module variables and functions into the scope
            VBScriptAnalyzer.AnalyzeFile(block, scope);

            //Generate the module level code other than the methods:
            if (block.Statements != null)
            {
                foreach (var s in block.Statements)
                {
                    if (s is VB.MethodDeclaration)
                    {
                        //Make sure methods are created first before being executed
                        body.Insert(0, VBScriptGenerator.GenerateExpr(s, scope));
                    }
                    else
                    {
                        Expression stmt = VBScriptGenerator.GenerateExpr(s, scope);
                        if (scope.VariableScope.IsOnErrorResumeNextOn)
                        {
                            stmt = VBScriptGenerator.WrapTryCatchExpression(stmt, scope);
                        }
                        Expression debugInfo = null;
                        Expression clearDebugInfo = null;
                        if (Trace && s is VB.Statement && !(s is VB.BlockStatement))
                        {
                            debugInfo = VBScriptGenerator.GenerateDebugInfo(s, scope, out clearDebugInfo);
                            body.Add(debugInfo);
                        }

                        body.Add(stmt);

                        if (clearDebugInfo != null)
                        {
                            body.Add(clearDebugInfo);
                        }
                    }
                }
            }

            body.Add(Expression.Constant(null)); //Stop anything from returning

            if (scope.Errors.Count > 0)
            {
                throw new VBScriptCompilerException(scope.Errors);
            }

            //if (Debug)
            //{
            //    Expression registerRuntimeVariables = VBScriptGenerator.GenerateRuntimeVariablesExpression(scope);

            //    body.Insert(0, registerRuntimeVariables);
            //}

            var moduleFun = Expression.Lambda<Action<VBScript, IDynamicMetaObjectProvider>>(
                Expression.Block(
                    scope.Names.Values,
                    body),
                scope.RuntimeExpr,
                scope.ModuleExpr);

            //if (!Debug)
            //{
                return moduleFun.Compile();
            //}
            //else
            //{
            //    Expression<Action<VBScript, IDynamicMetaObjectProvider>> lambda =  (Expression<Action<VBScript, IDynamicMetaObjectProvider>>)DebugContext.TransformLambda(moduleFun);
            //    return lambda.Compile();
            //}
        }

        // Execute a single expression parsed from string in the provided module
        // scope and returns the resulting value.
        //
        public void ExecuteExpr(string expr_str,
                                  IDynamicMetaObjectProvider moduleEO) {
            var moduleFun = ParseExprToLambda(new StringReader(expr_str));
            var d = moduleFun;
            d(this, moduleEO);
        }

        internal Action<VBScript, IDynamicMetaObjectProvider>
                 ParseExprToLambda(TextReader reader) {
            var scanner = new VB.Scanner(reader);
            var errorTable = new List<VB.SyntaxError>();
            var ast = new VB.Parser().ParseScriptFile(scanner, errorTable);

            VBScriptSourceCodeReader sourceReader = reader as VBScriptSourceCodeReader;
            ISourceMapper mapper = null;
            if (sourceReader != null)
            {
                mapper = sourceReader.SourceMapper;
            }

            var scope = new AnalysisScope(
                null,
                "__snippet__",
                this,
                Expression.Parameter(typeof(VBScript), "vbscriptRuntime"),
                Expression.Parameter(typeof(IDynamicMetaObjectProvider), "fileModule"),
                mapper
            );

            List<Expression> body = new List<Expression>();
            body.Add(Expression.Convert(VBScriptGenerator.GenerateExpr(ast, scope), typeof(object)));

            if (scope.Errors.Count > 0)
            {
                throw new VBScriptCompilerException(scope.Errors);
            }

            var moduleFun = Expression.Lambda<Action<VBScript, IDynamicMetaObjectProvider>>(
                Expression.Block(body),
                scope.RuntimeExpr,
                scope.ModuleExpr
            );
            return moduleFun.Compile();
        }


        public IDynamicMetaObjectProvider Globals { get { return _globals; } }
        public IDynamicMetaObjectProvider DlrGlobals { get { return _dlrGlobals; } }
        public bool Trace { get; set; }

        public static ExpandoObject CreateScope() {
            return new ExpandoObject();
        }
        

        /////////////////////////
        // Canonicalizing Binders
        /////////////////////////

        // We need to canonicalize binders so that we can share L2 dynamic
        // dispatch caching across common call sites.  Every call site with the
        // same operation and same metadata on their binders should return the
        // same rules whenever presented with the same kinds of inputs.  The
        // DLR saves the L2 cache on the binder instance.  If one site somewhere
        // produces a rule, another call site performing the same operation with
        // the same metadata could get the L2 cached rule rather than computing
        // it again.  For this to work, we need to place the same binder instance
        // on those functionally equivalent call sites.

        private Dictionary<string, VBScriptGetMemberBinder> _getMemberBinders =
            new Dictionary<string, VBScriptGetMemberBinder>();
        public VBScriptGetMemberBinder GetGetMemberBinder (string name) {
            lock (_getMemberBinders) {
                // Don't lower the name.  Sympl is case-preserving in the metadata
                // in case some DynamicMetaObject ignores ignoreCase.  This makes
                // some interop cases work, but the cost is that if a Sympl program
                // spells ".foo" and ".Foo" at different sites, they won't share rules.
                if (_getMemberBinders.ContainsKey(name))
                    return _getMemberBinders[name];
                var b = new VBScriptGetMemberBinder(name);
                _getMemberBinders[name] = b;
                return b;
            }
        }

        private Dictionary<string, VBScriptSetMemberBinder> _setMemberBinders =
            new Dictionary<string, VBScriptSetMemberBinder>();
        public VBScriptSetMemberBinder GetSetMemberBinder (string name) {
            lock (_setMemberBinders) {
                // Don't lower the name.  Sympl is case-preserving in the metadata
                // in case some DynamicMetaObject ignores ignoreCase.  This makes
                // some interop cases work, but the cost is that if a Sympl program
                // spells ".foo" and ".Foo" at different sites, they won't share rules.
                if (_setMemberBinders.ContainsKey(name))
                    return _setMemberBinders[name];
                var b = new VBScriptSetMemberBinder(name);
                _setMemberBinders[name] = b;
                return b;
            }
        }

        private Dictionary<CallInfo, VBScriptInvokeBinder> _invokeBinders =
            new Dictionary<CallInfo, VBScriptInvokeBinder>();
        public VBScriptInvokeBinder GetInvokeBinder (CallInfo info) {
            lock (_invokeBinders) {
                if (_invokeBinders.ContainsKey(info))
                    return _invokeBinders[info];
                var b = new VBScriptInvokeBinder(info);
                _invokeBinders[info] = b;
                return b;
            }
        }

        private Dictionary<InvokeMemberBinderKey, VBScriptInvokeMemberBinder>
            _invokeMemberBinders =
                new Dictionary<InvokeMemberBinderKey, VBScriptInvokeMemberBinder>();
        public VBScriptInvokeMemberBinder GetInvokeMemberBinder
                (InvokeMemberBinderKey info) {
            lock (_invokeMemberBinders) {
                if (_invokeMemberBinders.ContainsKey(info))
                    return _invokeMemberBinders[info];
                var b = new VBScriptInvokeMemberBinder(info.Name, info.Info);
                _invokeMemberBinders[info] = b;
                return b;
            }
        }

        private Dictionary<CallInfo, VBScriptCreateInstanceBinder> 
            _createInstanceBinders =
                new Dictionary<CallInfo, VBScriptCreateInstanceBinder>();
        public VBScriptCreateInstanceBinder GetCreateInstanceBinder(CallInfo info) {
            lock (_createInstanceBinders) {
                if (_createInstanceBinders.ContainsKey(info))
                    return _createInstanceBinders[info];
                var b = new VBScriptCreateInstanceBinder(info);
                _createInstanceBinders[info] = b;
                return b;
            }
        }

        private Dictionary<CallInfo, VBScriptGetIndexBinder> _getIndexBinders =
            new Dictionary<CallInfo, VBScriptGetIndexBinder>();
        public VBScriptGetIndexBinder GetGetIndexBinder(CallInfo info) {
            lock (_getIndexBinders) {
                if (_getIndexBinders.ContainsKey(info))
                    return _getIndexBinders[info];
                var b = new VBScriptGetIndexBinder(info);
                _getIndexBinders[info] = b;
                return b;
            }
        }

        private Dictionary<CallInfo, VBScriptSetIndexBinder> _setIndexBinders =
            new Dictionary<CallInfo, VBScriptSetIndexBinder>();
        public VBScriptSetIndexBinder GetSetIndexBinder(CallInfo info) {
            lock (_setIndexBinders) {
                if (_setIndexBinders.ContainsKey(info))
                    return _setIndexBinders[info];
                var b = new VBScriptSetIndexBinder(info);
                _setIndexBinders[info] = b;
                return b;
            }
        }

        private Dictionary<ExpressionType, VBScriptBinaryOperationBinder>
            _binaryOperationBinders =
                new Dictionary<ExpressionType, VBScriptBinaryOperationBinder>();
        public VBScriptBinaryOperationBinder GetBinaryOperationBinder
                (ExpressionType op) {
            lock (_binaryOperationBinders) {
                if (_binaryOperationBinders.ContainsKey(op))
                    return _binaryOperationBinders[op];
                var b = new VBScriptBinaryOperationBinder(op);
                _binaryOperationBinders[op] = b;
                return b;
            }
        }

        private Dictionary<ExpressionType, VBScriptUnaryOperationBinder>
            _unaryOperationBinders =
                new Dictionary<ExpressionType, VBScriptUnaryOperationBinder>();
        public VBScriptUnaryOperationBinder GetUnaryOperationBinder
                (ExpressionType op) {
            lock (_unaryOperationBinders) {
                if (_unaryOperationBinders.ContainsKey(op))
                    return _unaryOperationBinders[op];
                var b = new VBScriptUnaryOperationBinder(op);
                _unaryOperationBinders[op] = b;
                return b;
            }
        }

        //private DebugContext _debugContext;
        //public DebugContext DebugContext
        //{
        //    get
        //    {
        //        if (_debugContext == null)
        //        {
        //            _debugContext = DebugContext.CreateInstance();
        //            ITracePipeline pipeLine = TracePipeline.CreateInstance(_debugContext);
        //            pipeLine.TraceCallback = new VBScriptTraceCallbackListener();
        //        }

        //        return _debugContext;
        //    }
        //}
    } 
}
