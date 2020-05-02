using System;
using System.Collections.Generic;
using System.Text;
using VB = Dlrsoft.VBScript.Parser;
using Dlrsoft.VBScript.Runtime;
using Microsoft.Scripting;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq;
using System.Linq.Expressions;
#endif

using System.Reflection;
using System.Collections;

namespace Dlrsoft.VBScript.Compiler
{
    internal static class VBScriptGenerator
    {
        private static Set<string> _builtinFunctions;
        private static Set<string> _builtinConstants;

        public static bool IsBuiltInConstants(string name)
        {
            ensureBuiltinConstants();
            return _builtinConstants.Contains(name.ToLower());
        }

        public static bool IsBuiltInFunction(string name)
        {
            ensureBuiltinFunctions();

            return _builtinFunctions.Contains(name.ToLower());
        }

        public static Expression GenerateExpr(VB.Tree expr, AnalysisScope scope)
        {
            try
            {
                if (expr is VB.ImportsDeclaration)
                {
                    return GenerateImportExpr((VB.ImportsDeclaration)expr, scope);
                }
                else if (expr is VB.NameImport)
                {
                    return GenerateImportExpr((VB.NameImport)expr, scope);
                }
                else if (expr is VB.AliasImport)
                {
                    return GenerateImportExpr((VB.AliasImport)expr, scope);
                }
                if (expr is VB.CallStatement)
                {
                    return GenerateCallStmtExpr((VB.CallStatement)expr, scope);
                }
                else if (expr is VB.LocalDeclarationStatement)
                {
                    return GenerateDeclarationExpr((VB.LocalDeclarationStatement)expr, scope);
                }
                else if (expr is VB.FunctionDeclaration)
                {
                    return GenerateFunctionExpr((VB.FunctionDeclaration)expr, scope);
                }
                else if (expr is VB.SubDeclaration)
                {
                    return GenerateSubExpr((VB.SubDeclaration)expr, scope);
                }
                //else if (expr is SymplLambdaExpr)
                //{
                //    return GenerateLambdaExpr((SymplLambdaExpr)expr, scope);
                //}
                else if (expr is VB.CallOrIndexExpression)
                {
                    return GenerateCallOrIndexExpr((VB.CallOrIndexExpression)expr, scope);
                }
                else if (expr is VB.SimpleNameExpression)
                {
                    return GenerateIdExpr((VB.SimpleNameExpression)expr, scope);
                }
                //else if (expr is SymplQuoteExpr)
                //{
                //    return GenerateQuoteExpr((SymplQuoteExpr)expr, scope);
                //}
                else if (expr is VB.NothingExpression)
                {
                    return Expression.Constant(null);
                }
                else if (expr is VB.LiteralExpression)
                {
                    return Expression.Constant(((VB.LiteralExpression)expr).Value);
                }
                else if (expr is VB.AssignmentStatement)
                {
                    return GenerateAssignExpr((VB.AssignmentStatement)expr, scope);
                }
                //else if (expr is SymplEqExpr)
                //{
                //    return GenerateEqExpr((SymplEqExpr)expr, scope);
                //}
                else if (expr is VB.IfBlockStatement)
                {
                    return GenerateIfExpr((VB.IfBlockStatement)expr, scope);
                }
                else if (expr is VB.LineIfStatement)
                {
                    return GenerateIfExpr((VB.LineIfStatement)expr, scope);
                }
                else if (expr is VB.QualifiedExpression)
                {
                    return GenerateDottedExpr((VB.QualifiedExpression)expr, scope);
                }
                else if (expr is VB.NewExpression)
                {
                    return GenerateNewExpr((VB.NewExpression)expr, scope);
                }
                else if (expr is VB.ForBlockStatement)
                {
                    return GenerateForBlockExpr((VB.ForBlockStatement)expr, scope);
                }
                else if (expr is VB.ForEachBlockStatement)
                {
                    return GenerateForEachBlockExpr((VB.ForEachBlockStatement)expr, scope);
                }
                else if (expr is VB.WhileBlockStatement)
                {
                    return GenerateWhileBlockExpr((VB.WhileBlockStatement)expr, scope);
                }
                else if (expr is VB.DoBlockStatement)
                {
                    return GenerateDoBlockExpr((VB.DoBlockStatement)expr, scope);
                }
                else if (expr is VB.WithBlockStatement)
                {
                    return GenerateWithBlockExpr((VB.WithBlockStatement)expr, scope);
                }
                else if (expr is VB.ExitStatement)
                {
                    return GenerateBreakExpr((VB.ExitStatement)expr, scope);
                }
                //else if (expr is SymplEltExpr)
                //{
                //    return GenerateEltExpr((SymplEltExpr)expr, scope);
                //}
                else if (expr is VB.BinaryOperatorExpression)
                {
                    return GenerateBinaryExpr((VB.BinaryOperatorExpression)expr, scope);
                }
                else if (expr is VB.UnaryOperatorExpression)
                {
                    return GenerateUnaryExpr((VB.UnaryOperatorExpression)expr, scope);
                }
                else if (expr is VB.SelectBlockStatement)
                {
                    return GenerateSelectBlockExpr((VB.SelectBlockStatement)expr, scope);
                }
                else if (expr is VB.BlockStatement)
                {
                    return GenerateBlockExpr(((VB.BlockStatement)expr).Statements, scope);
                }
                else if (expr is VB.EmptyStatement)
                {
                    return Expression.Empty();
                }
                else if (expr is VB.OptionDeclaration)
                {
                    scope.ModuleScope.IsOptionExplicitOn = true;
                    return Expression.Empty();
                }
                else if (expr is VB.ParentheticalExpression)
                {
                    return GenerateExpr(((VB.ParentheticalExpression)expr).Operand, scope);
                }
                else if (expr is VB.ReDimStatement)
                {
                    return GenerateRedimExpr((VB.ReDimStatement)expr, scope);
                }
                else if (expr is VB.OnErrorStatement)
                {
                    return GenerateOnErrorStatement((VB.OnErrorStatement)expr, scope);
                }
                else if (expr is ExpressionExpression)
                {
                    return ((ExpressionExpression)expr).Expression;
                }
                else
                {
                    VBScriptSyntaxError error = new VBScriptSyntaxError(
                        scope.ModuleScope.Name,
                        SourceUtil.ConvertSpan(expr.Span),
                        (int)VB.SyntaxErrorType.NotImplemented,
                        string.Format("{0} is not yet implemented.", expr.GetType().FullName)
                    );

                    scope.ModuleScope.Errors.Add(error);
                    return Expression.Default(typeof(object));
                }
            }
            catch (Exception ex)
            {
                VBScriptSyntaxError error = new VBScriptSyntaxError(
                    scope.ModuleScope.Name,
                    SourceUtil.ConvertSpan(expr.Span),
                    (int)VB.SyntaxErrorType.Unexpected,
                    string.Format("Unexpected error in {0}: {1}", expr.GetType().FullName, ex.Message)
                );

                scope.ModuleScope.Errors.Add(error);
                return Expression.Default(typeof(object));
            }
        }

        public static Expression GenerateImportExpr(VB.ImportsDeclaration importDesc, AnalysisScope scope)
        {
            if (!scope.IsModule)
            {
                throw new InvalidOperationException(
                    "Import expression must be a top level expression.");
            }
            List<Expression> exprs = new List<Expression>();
            foreach (VB.Import statement in importDesc.ImportMembers)
            {
                Expression expr = GenerateExpr(statement, scope);
                exprs.Add(expr);
            }
            return Expression.Block(exprs);
        }

        public static Expression GenerateImportExpr(VB.NameImport import,
                                                    AnalysisScope scope)
        {
            if (!scope.IsModule)
            {
                throw new InvalidOperationException(
                    "Import expression must be a top level expression.");
            }
            return Expression.Call(
                typeof(RuntimeHelpers).GetMethod("VBScriptImport"),
                scope.RuntimeExpr,
                scope.ModuleExpr,
                Expression.Constant(new string[] {
                    ((VB.SimpleName)((VB.NamedTypeName)import.TypeName).Name).Name
                }),
                Expression.Constant(new string[]{}),
                Expression.Constant(new string[]{}));
        }

        public static Expression GenerateImportExpr(VB.AliasImport import,
                                                     AnalysisScope scope)
        {
            if (!scope.IsModule)
            {
                throw new InvalidOperationException(
                    "Import expression must be a top level expression.");
            }

            string alias = ((VB.SimpleName)import.Name).Name;
            VB.NamedTypeName namedTypeName = (VB.NamedTypeName)import.AliasedTypeName;
            List<string> names = new List<string>();
            string[] typeNames;
            if (namedTypeName.Name is VB.QualifiedName)
            {
                VB.QualifiedName qualifiedName = (VB.QualifiedName)namedTypeName.Name;
                typeNames = new string[] {qualifiedName.Name.Name};

                VB.Name name = qualifiedName.Qualifier;
                while (name is VB.QualifiedName)
                {
                    names.Insert(0, ((VB.QualifiedName)name).Name.Name);
                    name = ((VB.QualifiedName)name).Qualifier;
                }
                names.Insert(0, ((VB.SimpleName)name).Name);
            }
            else
            {
                typeNames = new string[] { };
                names.Insert(0, ((VB.SimpleName)namedTypeName.Name).Name);
            }

            return Expression.Call(
                typeof(RuntimeHelpers).GetMethod("VBScriptImport"),
                scope.RuntimeExpr,
                scope.ModuleExpr,
                Expression.Constant(names.ToArray()),
                Expression.Constant(typeNames),
                Expression.Constant(new string[] {alias})
            );
        }

        public static Expression GenerateFunctionExpr(VB.FunctionDeclaration func,
                                                          AnalysisScope scope)
        {
            if (!scope.IsModule)
            {
                throw new InvalidOperationException(
                    "Use Defmethod or Lambda when not defining top-level function.");
            }
            string funcName = func.Name.Name.ToLower();
            Expression lambda = GenerateLambdaDef(func.Parameters, func.Statements, scope, funcName, false);
            //ParameterExpression p = Expression.Parameter(typeof(object), funcName);
            //scope.Names.Add(funcName, p);
            ParameterExpression p = scope.Names[funcName];
            return Expression.Assign(p, lambda);
        }

        public static Expression GenerateSubExpr(VB.SubDeclaration sub,
                                                          AnalysisScope scope)
        {
            if (!scope.IsModule)
            {
                throw new InvalidOperationException(
                    "Use Defmethod or Lambda when not defining top-level function.");
            }
            string subName = sub.Name.Name.ToLower();
            Expression lambda = GenerateLambdaDef(sub.Parameters, sub.Statements, scope, subName, true);
            //ParameterExpression p = Expression.Parameter(typeof(object), subName);
            //scope.Names.Add(subName, p);
            ParameterExpression p = scope.Names[subName];
            return Expression.Assign(p, lambda);
        }

        public static Expression GenerateDeclarationExpr(VB.LocalDeclarationStatement stmt, AnalysisScope scope)
        {
            List<Expression> expressions = new List<Expression>();

            bool isConst = false;

            if (stmt.Modifiers != null)
            {
                if (stmt.Modifiers.ModifierTypes == VB.ModifierTypes.Const)
                {
                    isConst = true;
                }
            }

            foreach (VB.VariableDeclarator d in stmt.VariableDeclarators)
            {
                Expression initializer = null;

                if (d.Initializer != null)
                {
                    VB.ExpressionInitializer exprInitializer = d.Initializer as VB.ExpressionInitializer;

                    if (exprInitializer != null)
                    {
                        initializer = GenerateExpr(exprInitializer.Expression, scope);
                    }
                    else
                    {
                        new NotImplementedException(d.Initializer.GetType().Name + " is not yet implemented.");
                    }
                }

                foreach (VB.VariableName v in d.VariableNames)
                {
                    string name = v.Name.Name.ToLower();
                    Expression initExpression = null;
                    if (initializer != null)
                    {
                        initExpression = initializer;
                    }
                    else
                    {
                        if (v.ArrayType == null || v.ArrayType.Arguments.Count == 0)
                        {
                            initExpression = Expression.Constant(null);
                        }
                        else
                        {
                            List<Expression> args = GenerateArgumentList(v.ArrayType.Arguments, scope);
                            Expression converted = ConvertToIntegerArrayExpression(args);

                            initExpression = Expression.Call(
                                typeof(HelperFunctions).GetMethod("Redim"),
                                Expression.Constant(typeof(object)),
                                converted
                            );
                            //initExpression = Expression.Convert(
                            //        Expression.NewArrayBounds(
                            //            typeof(object),
                            //            GenerateArgumentList(v.ArrayType.Arguments, scope)
                            //        ),
                            //        typeof(object)
                            //    );
                        }
                    }

                    ParameterExpression p;
                    if (scope.IsModule)
                    {
                        //Module variable already put into scope in the analysis phase
                        p = scope.Names[name];
                    }
                    else
                    {
                        p = Expression.Parameter(typeof(object), name);
                        scope.VariableScope.Names.Add(name, p);
                    }
                    expressions.Add(Expression.Assign(p, Expression.Convert(initExpression, p.Type)));
                }
            }
            if (expressions.Count > 0)
                return Expression.Block(expressions);
            else
                return Expression.Empty();
        }

        public static Expression GenerateRedimExpr(VB.ReDimStatement redim, AnalysisScope scope)
        {
            List<Expression> expressions = new List<Expression>();
            foreach (VB.Expression variable in redim.Variables)
            {
                VB.CallOrIndexExpression vd = (VB.CallOrIndexExpression)variable;
                List<Expression> args = GenerateArgumentList(vd.Arguments, scope);
                Expression converted = ConvertToIntegerArrayExpression(args);

                Expression arrayExp;

                if (!redim.IsPreserve)
                {
                    //arrayExp = Expression.Convert(
                    //    Expression.NewArrayBounds(
                    //        typeof(object),
                    //        converted
                    //        ),
                    //    typeof(object)
                    //);
                    arrayExp = Expression.Call(
                        typeof(HelperFunctions).GetMethod("Redim"),
                        Expression.Constant(typeof(object)),
                        converted
                    );

                }
                else
                {
                    arrayExp = Expression.Call(
                        typeof(HelperFunctions).GetMethod("RedimPreserve"),
                        GenerateExpr(vd.TargetExpression, scope),
                        converted
                    );
                }

                Expression initExpression = GenerateSimpleNameAssignExpr(
                    (VB.SimpleNameExpression)vd.TargetExpression,
                    arrayExp,
                    scope
                );
                expressions.Add(initExpression);
            }
            return Expression.Block(expressions);
        }

        // Returns a dynamic InvokeMember or Invoke expression, depending on the
        // Function expression.
        //
        public static Expression GenerateCallStmtExpr(
                VB.CallStatement expr, AnalysisScope scope)
        {
            return GenerateCallOrIndexExpr(
                new VB.CallOrIndexExpression(
                    expr.TargetExpression,
                    expr.Arguments,
                    expr.Span
                ),
                scope
            );
        }

        // Returns a chain of GetMember and InvokeMember dynamic expressions for
        // the dotted expr.
        //
        public static Expression GenerateDottedExpr(VB.QualifiedExpression expr,
                                                    AnalysisScope scope)
        {
            Expression curExpr = null;
            if (expr.Qualifier != null)
            {
                curExpr = GenerateExpr(expr.Qualifier, scope);
            }
            else
            {
                curExpr = scope.NearestWithExpression;
            }
            
            curExpr = Expression.Dynamic(
                scope.GetRuntime()
                     .GetGetMemberBinder(expr.Name.Name),
                typeof(object),
                curExpr
            );
            //    }
            //    else if (e is SymplFunCallExpr)
            //    {
            //        var call = (SymplFunCallExpr)e;
            //        List<Expression> args = new List<Expression>();
            //        args.Add(curExpr);
            //        args.AddRange(call.Arguments.Select(a => GenerateExpr(a, scope)));

            //        curExpr = Expression.Dynamic(
            //            // Dotted exprs must be simple invoke members, a.b.(c ...) 
            //            scope.GetRuntime().GetInvokeMemberBinder(
            //                new InvokeMemberBinderKey(
            //                    ((SymplIdExpr)call.Function).IdToken.Name,
            //                    new CallInfo(call.Arguments.Length))),
            //            typeof(object),
            //            args
            //        );
            //    }
            //    else
            //    {
            //        throw new InvalidOperationException(
            //            "Internal: dotted must be IDs or Funs.");
            //    }
            //}
            return curExpr;
        }

        public static Expression GenerateQualifiedNameExpr(VB.QualifiedName expr,
                                                    AnalysisScope scope)
        {
            Expression curExpr;
            if (expr.Qualifier is VB.SimpleName)
                curExpr = GenerateSimpleNameExpr((VB.SimpleName)expr.Qualifier, scope);
            else
                curExpr = GenerateQualifiedNameExpr((VB.QualifiedName)expr.Qualifier, scope);

            return Expression.Dynamic(
                scope.GetRuntime()
                     .GetGetMemberBinder(expr.Name.Name),
                typeof(object),
                curExpr
            );
        }

        // GenerateAssignExpr handles IDs, indexing, and member sets.  IDs are either
        // lexical or dynamic exprs on the module scope.  Everything
        // else is dynamic.
        //
        public static Expression GenerateAssignExpr(VB.AssignmentStatement expr,
                                                    AnalysisScope scope)
        {
            var val = GenerateExpr(expr.SourceExpression, scope);

            if (!expr.IsSetStatement)
            {
                val = RuntimeHelpers.GetDefaultValue(val);
            }

            if (expr.TargetExpression is VB.SimpleNameExpression)
            {
                var idExpr = (VB.SimpleNameExpression)expr.TargetExpression;
                return GenerateSimpleNameAssignExpr(idExpr, val, scope);
            }
            else if (expr.TargetExpression is VB.CallOrIndexExpression)
            {
                //If LHS is CallOrIndexExpression, it must be an index expression
                var callOrIndex = (VB.CallOrIndexExpression)(expr.TargetExpression);
                var args = new List<Expression>();
                args.Add(GenerateExpr(callOrIndex.TargetExpression, scope));
                args.AddRange(callOrIndex.Arguments.Select(e => GenerateExpr(e.Expression, scope)));
                args.Add(Expression.Convert(val, typeof(object)));
                // Trusting MO convention to return stored values.
                return Expression.Dynamic(
                           scope.GetRuntime().GetSetIndexBinder(
                                   new CallInfo(callOrIndex.Arguments.Count)),
                           typeof(object),
                           args);
            }
            else if (expr.TargetExpression is VB.QualifiedExpression)
            {
                // For now, one dot only.  Later, pick oflast dotted member
                // access (like GenerateFunctionCall), and use a temp and block.
                var dottedExpr = (VB.QualifiedExpression)(expr.TargetExpression);
                var id = (VB.SimpleName)(dottedExpr.Name);

                Expression qualifier = null;
                if (dottedExpr.Qualifier != null)
                {
                    qualifier = GenerateExpr(dottedExpr.Qualifier, scope);
                }
                else
                {
                    qualifier = scope.NearestWithExpression;
                }


                // Trusting MOs convention to return stored values.
                return Expression.Dynamic(
                           scope.GetRuntime().GetSetMemberBinder(id.Name),
                           typeof(object),
                           qualifier,
                           val
                );
            }

            throw new InvalidOperationException("Invalid left hand side type.");
        }

        public static Expression GenerateCallOrIndexExpr(VB.CallOrIndexExpression expr, AnalysisScope scope)
        {
            List<Expression> args = GenerateArgumentList(expr.Arguments, scope);
            if (expr.TargetExpression is VB.SimpleNameExpression)
            {
                string name = ((VB.SimpleNameExpression)expr.TargetExpression).Name.Name;
                //Call the built-in function it is one
                if (IsBuiltInFunction(name))
                {
                    int argCount = expr.Arguments == null? 0 : expr.Arguments.Count;
                    args.Insert(0, Expression.Constant(new TypeModel(typeof(BuiltInFunctions))));
                    return Expression.Dynamic(
                            scope.GetRuntime().GetInvokeMemberBinder(
                                new InvokeMemberBinderKey(
                                    name,
                                    new CallInfo(argCount))
                            ),
                            typeof(object),
                            args
                        );
                }
                else if (scope.FunctionTable.Contains(name))
                {
                    var fun = GenerateSimpleNameExpr(((VB.SimpleNameExpression)expr.TargetExpression).Name, scope);
                    List<Type> argTypes = new List<Type>();
                    foreach (Expression arg in args)
                    {
                        argTypes.Add(arg.Type.MakeByRefType());
                    }
                    argTypes.Insert(0, typeof(object)); //delegate itself
                    argTypes.Add(typeof(object)); //return type

                    args.Insert(0, fun);
                    // Use DynExpr so that I don't always have to have a delegate to call,
                    // such as what happens with IPy interop.
                    int argCount = expr.Arguments == null ? 0 : expr.Arguments.Count;

                    return Expression.MakeDynamic(
                        Microsoft.Scripting.Actions.DynamicSiteHelpers.MakeCallSiteDelegate(argTypes.ToArray()),
                        scope.GetRuntime().GetInvokeBinder(new CallInfo(argCount)),
                        args
                        );
                }
                else
                {
                    ////If not a function, it must be an array
                    args.Insert(0, GenerateSimpleNameExpr(((VB.SimpleNameExpression)expr.TargetExpression).Name, scope));
                    return Expression.Dynamic(
                        scope.GetRuntime().GetGetIndexBinder(
                            new CallInfo(args.Count)
                        ),
                        typeof(object),
                        args
                    );
                }
            }
            else if (expr.TargetExpression is VB.CallOrIndexExpression)
            {
                Expression objExpr = GenerateCallOrIndexExpr((VB.CallOrIndexExpression)expr.TargetExpression, scope);
                ////What followed must be an array as VBScript does not have function return a delegate
                args.Insert(0, objExpr);
                return Expression.Dynamic(
                    scope.GetRuntime().GetGetIndexBinder(
                        new CallInfo(args.Count)
                    ),
                    typeof(object),
                    args
                );
            }
            else //Qualified Expression
            {
                VB.QualifiedExpression qualifiedExpr = (VB.QualifiedExpression)expr.TargetExpression;
                Expression objExpr;
                if (qualifiedExpr.Qualifier is VB.QualifiedExpression)
                {
                    objExpr = GenerateDottedExpr(
                        (VB.QualifiedExpression)qualifiedExpr.Qualifier,
                        scope
                    );
                }
                else if (qualifiedExpr.Qualifier is VB.SimpleNameExpression)
                {
                    objExpr = GenerateSimpleNameExpr(((VB.SimpleNameExpression)qualifiedExpr.Qualifier).Name, scope);
                }
                else if (qualifiedExpr.Qualifier is VB.CallOrIndexExpression)
                {
                    objExpr = GenerateCallOrIndexExpr((VB.CallOrIndexExpression)qualifiedExpr.Qualifier, scope);
                }
                else //null
                {
                    objExpr = scope.NearestWithExpression;
                    if (objExpr == null)
                        throw new Exception("Missing With statement");
                }

                //LC 12/16/2009 Try the byref type
                List<Type> argTypes = new List<Type>();
                foreach (Expression arg in args)
                {
                    argTypes.Add(arg.Type.MakeByRefType());
                }
                argTypes.Insert(0, typeof(object)); //target itself
                argTypes.Add(typeof(object)); //return type
                
                args.Insert(0, objExpr);

                // last expr must be an id
                var lastExpr = qualifiedExpr.Name;
                int argCount = 0;
                if (expr.Arguments != null)
                    argCount = expr.Arguments.Count;

                return Expression.MakeDynamic(
                    Microsoft.Scripting.Actions.DynamicSiteHelpers.MakeCallSiteDelegate(argTypes.ToArray()),
                    scope.GetRuntime().GetInvokeMemberBinder(
                        new InvokeMemberBinderKey(
                            lastExpr.Name,
                            new CallInfo(argCount))),
                    args
                );

                //return Expression.Dynamic(
                //    scope.GetRuntime().GetInvokeMemberBinder(
                //        new InvokeMemberBinderKey(
                //            lastExpr.Name,
                //            new CallInfo(argCount))),
                //    typeof(object),
                //    args
                //);
            }
        }

        // Return an Expression for referencing the ID.  If we find the name in the
        // scope chain, then we just return the stored ParamExpr.  Otherwise, the
        // reference is a dynamic member lookup on the root scope, a module object.
        //
        public static Expression GenerateIdExpr(VB.SimpleNameExpression expr,
                                                AnalysisScope scope)
        {
            string name = expr.Name.Name;

            if (IsBuiltInFunction(name) || scope.FunctionTable.Contains(name))
            {
                return GenerateCallOrIndexExpr(
                    new VB.CallOrIndexExpression(
                        expr,
                        null,
                        expr.Span),
                    scope
                );
            }
            else
            {
                return GenerateSimpleNameExpr(expr.Name, scope);
            }
        }

        public static Expression GenerateOnErrorStatement(VB.OnErrorStatement onError, AnalysisScope scope)
        {
            if (onError.OnErrorType == VB.OnErrorType.Next)
            {
                scope.VariableScope.IsOnErrorResumeNextOn = true;
            }
            else if(onError.OnErrorType == VB.OnErrorType.Zero)
            {
                scope.VariableScope.IsOnErrorResumeNextOn = false;
            }
            return Expression.Call(
                        scope.ErrExpression,
                        typeof(ErrObject).GetMethod("Clear")
                    );
        }

        // GenerateLetStar returns a Block with vars, each initialized in the order
        // they appear.  Each var's init expr can refer to vars initialized before it.
        // The Block's body is the Let*'s body.
        //
        //public static Expression GenerateLetStarExpr(SymplLetStarExpr expr,
        //                                              AnalysisScope scope)
        //{
        //    var letscope = new AnalysisScope(scope, "let*");
        //    // Generate bindings.
        //    List<Expression> inits = new List<Expression>();
        //    List<ParameterExpression> varsInOrder = new List<ParameterExpression>();
        //    foreach (var b in expr.Bindings)
        //    {
        //        // Need richer logic for mvbind
        //        var v = Expression.Parameter(typeof(object), b.Variable.Name);
        //        varsInOrder.Add(v);
        //        inits.Add(
        //            Expression.Assign(
        //                v,
        //                Expression.Convert(GenerateExpr(b.Value, letscope), v.Type))
        //        );
        //        // Add var to scope after analyzing init value so that init value
        //        // references to the same ID do not bind to his uninitialized var.
        //        letscope.Names[b.Variable.Name.ToLower()] = v;
        //    }
        //    List<Expression> body = new List<Expression>();
        //    foreach (var e in expr.Body)
        //    {
        //        body.Add(GenerateExpr(e, letscope));
        //    }
        //    // Order of vars to BlockExpr don't matter semantically, but may as well
        //    // keep them in the order the programmer specified in case they look at the
        //    // Expr Trees in the debugger or for meta-programming.
        //    inits.AddRange(body);
        //    return Expression.Block(typeof(object), varsInOrder.ToArray(), inits);
        //}

        // GenerateBlockExpr returns a Block with the body exprs.
        //
        public static Expression GenerateBlockExpr(VB.StatementCollection stmts,
                                                    AnalysisScope scope)
        {
            List<Expression> body = new List<Expression>();
            if (stmts != null)
            {
                foreach (VB.Statement s in stmts)
                {
                    Expression stmt = VBScriptGenerator.GenerateExpr(s, scope);
                    if (scope.VariableScope.IsOnErrorResumeNextOn)
                    {
                        stmt = WrapTryCatchExpression(stmt, scope);
                    }
                    Expression debugInfo = null;
                    Expression clearDebugInfo = null;
                    if (scope.GetRuntime().Trace && s is VB.Statement && !(s is VB.BlockStatement))
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
                //return Expression.Block(typeof(object), body);
                return Expression.Block(typeof(void), body);
            }
            else
            {
                //return Expression.Constant(null);
                return Expression.Empty();
            }
        }

        // GenerateQuoteExpr converts a list, literal, or id expr to a runtime quoted
        // literal and returns the Constant expression for it.
        //
        //public static Expression GenerateQuoteExpr(SymplQuoteExpr expr,
        //                                            AnalysisScope scope)
        //{
        //    return Expression.Constant(MakeQuoteConstant(
        //                                   expr.Expr, scope.GetRuntime()));
        //}

        //private static object MakeQuoteConstant(object expr, Sympl symplRuntime)
        //{
        //    if (expr is SymplListExpr)
        //    {
        //        SymplListExpr listexpr = (SymplListExpr)expr;
        //        int len = listexpr.Elements.Length;
        //        var exprs = new object[len];
        //        for (int i = 0; i < len; i++)
        //        {
        //            exprs[i] = MakeQuoteConstant(listexpr.Elements[i], symplRuntime);
        //        }
        //        return Cons._List(exprs);
        //    }
        //    else if (expr is IdOrKeywordToken)
        //    {
        //        return symplRuntime.MakeSymbol(((IdOrKeywordToken)expr).Name);
        //    }
        //    else if (expr is LiteralToken)
        //    {
        //        return ((LiteralToken)expr).Value;
        //    }
        //    else
        //    {
        //        throw new InvalidOperationException(
        //            "Internal: quoted list has -- " + expr.ToString());
        //    }
        //}

        //public static Expression GenerateEqExpr(SymplEqExpr expr,
        //                                        AnalysisScope scope)
        //{
        //    var mi = typeof(RuntimeHelpers).GetMethod("SymplEq");
        //    return Expression.Call(mi, Expression.Convert(
        //                                   GenerateExpr(expr.Left, scope),
        //                                   typeof(object)),
        //                           Expression.Convert(
        //                               GenerateExpr(expr.Right, scope),
        //                               typeof(object)));
        //}

        //public static Expression GenerateConsExpr(SymplConsExpr expr,
        //                                          AnalysisScope scope)
        //{
        //    var mi = typeof(RuntimeHelpers).GetMethod("MakeCons");
        //    return Expression.Call(mi, Expression.Convert(
        //                                   GenerateExpr(expr.Left, scope),
        //                                   typeof(object)),
        //                           Expression.Convert(
        //                               GenerateExpr(expr.Right, scope),
        //                               typeof(object)));
        //}

        //public static Expression GenerateListCallExpr(SymplListCallExpr expr,
        //                                              AnalysisScope scope)
        //{
        //    var mi = typeof(Cons).GetMethod("_List");
        //    int len = expr.Elements.Length;
        //    var args = new Expression[len];
        //    for (int i = 0; i < len; i++)
        //    {
        //        args[i] = Expression.Convert(GenerateExpr(expr.Elements[i], scope),
        //                                     typeof(object));
        //    }
        //    return Expression.Call(mi, Expression
        //                                   .NewArrayInit(typeof(object), args));
        //}

        public static Expression GenerateIfExpr(VB.IfBlockStatement ifBlock, AnalysisScope scope)
        {
            Expression elseBlock = null;
            if (ifBlock.ElseBlockStatement != null)
            {
                elseBlock = GenerateExpr((VB.BlockStatement)ifBlock.ElseBlockStatement, scope);
            }
            else
            {
                elseBlock = Expression.Empty();
            }

            if (ifBlock.ElseIfBlockStatements != null && ifBlock.ElseIfBlockStatements.Count > 0)
            {
                for (int i = ifBlock.ElseIfBlockStatements.Count - 1; i > -1; i--)
                {
                    VB.ElseIfBlockStatement elseifBock = (VB.ElseIfBlockStatement)ifBlock.ElseIfBlockStatements.get_Item(i);
                    elseBlock = Expression.Condition(
                            WrapBooleanTest(GenerateExpr(elseifBock.ElseIfStatement.Expression, scope)),
                            GenerateBlockExpr(elseifBock.Statements, scope),
                            elseBlock,
                            typeof(void));
                }
            }
            Expression test = WrapBooleanTest(GenerateExpr(ifBlock.Expression, scope));
            Expression ifblock = GenerateBlockExpr(ifBlock.Statements, scope);

            return Expression.Condition(
                       test,
                       ifblock,
                       elseBlock,
                       typeof(void));
        }

        public static Expression GenerateIfExpr(VB.LineIfStatement ifstmt, AnalysisScope scope)
        {
            Expression elseBlock = null;
            if (ifstmt.ElseStatements != null)
            {
                elseBlock = GenerateBlockExpr(ifstmt.ElseStatements, scope);
            }
            else
            {
                elseBlock = Expression.Empty();
            }

            Expression test = WrapBooleanTest(GenerateExpr(ifstmt.Expression, scope));
            Expression ifblock = GenerateBlockExpr(ifstmt.IfStatements, scope);

            return Expression.Condition(
                       test,
                       ifblock,
                       elseBlock,
                       typeof(void));
        }


        public static Expression GenerateSelectBlockExpr(VB.SelectBlockStatement selectBlock,
                                                    AnalysisScope scope)
        {
            ParameterExpression tmp = Expression.Parameter(typeof(object));
            Expression alt = null;
            if (selectBlock.CaseElseBlockStatement != null)
            {
                alt = GenerateExpr((VB.BlockStatement)selectBlock.CaseElseBlockStatement, scope);
            }
            else
            {
                alt = Expression.Constant(false);
            }

            if (selectBlock.CaseBlockStatements != null && selectBlock.CaseBlockStatements.Count > 0)
            {
                for (int i = selectBlock.CaseBlockStatements.Count - 1; i >= 0; i--)
                {
                    VB.CaseBlockStatement caseStmt = (VB.CaseBlockStatement)selectBlock.CaseBlockStatements.get_Item(i);
                    Expression condition = null;
                    for (int j = caseStmt.CaseStatement.CaseClauses.Count - 1; j >= 0; j--)
                    {
                        VB.RangeCaseClause caseClause = (VB.RangeCaseClause)caseStmt.CaseStatement.CaseClauses.get_Item(j);
                        //Expression oneCase = Expression.Equal(
                        //    tmp, 
                        //    Expression.Convert(
                        //        GenerateExpr(caseClause.RangeExpression, scope),
                        //        typeof(object)
                        //    )
                        //);
                        Expression oneCase = WrapBooleanTest(
                            Expression.Dynamic(
                                scope.GetRuntime().GetBinaryOperationBinder(ExpressionType.Equal),
                                typeof(object),
                                new Expression[] {
                                    tmp,
                                    GenerateExpr(caseClause.RangeExpression, scope)
                                }
                            )
                        );
                        
                        if (condition == null)
                        {
                            condition = oneCase;
                        }
                        else
                        {
                            condition = Expression.OrElse(oneCase, condition);
                        }
                    }
                    alt = Expression.Condition(
                            condition,
                            GenerateBlockExpr(caseStmt.Statements, scope),
                            alt,
                            typeof(void));
                }

                return Expression.Block(
                    new ParameterExpression[] {tmp},
                    Expression.Assign(
                        tmp, 
                        RuntimeHelpers.EnsureObjectResult(
                            GenerateExpr(selectBlock.Expression, scope)
                        )
                    ),
                    alt
                );
            }
            else
            {
                throw new Exception("At least one case block needed.");
            }
        }

        public static Expression GenerateForBlockExpr(VB.ForBlockStatement forBlock,
                                                  AnalysisScope scope)
        {
            var loopscope = new AnalysisScope(scope, "loop ");
            loopscope.IsForLoop = true; // needed for exit
            loopscope.LoopBreak = Expression.Label("loop break");

            Expression loopVariable;
            if (forBlock.ControlExpression is VB.SimpleNameExpression)
            {
                loopVariable = FindIdDef(((VB.SimpleNameExpression)forBlock.ControlExpression).Name.Name, scope, true);
            }
            else
            {
                loopVariable = GenerateExpr(forBlock.ControlExpression, loopscope);
            }
            Expression body = GenerateBlockExpr(forBlock.Statements, loopscope);

            LoopExpression myLoop = Expression.Loop(
                Expression.Block(
                    Expression.IfThen(
                        WrapBooleanTest(
                            GenerateExpr(
                                new VB.BinaryOperatorExpression(
                                    forBlock.ControlExpression,
                                    VB.OperatorType.GreaterThan,
                                    forBlock.ToLocation,
                                    forBlock.UpperBoundExpression,
                                    forBlock.UpperBoundExpression.Span
                                ),
                                loopscope
                            )
                        ),
                        Expression.Goto(loopscope.LoopBreak)
                    ),
                    body,
                    GenerateExpr(
                        new VB.AssignmentStatement(
                            forBlock.ControlExpression,
                            forBlock.LowerBoundExpression.Span.Start,
                            new VB.BinaryOperatorExpression(
                                forBlock.ControlExpression,
                                VB.OperatorType.Plus,
                                forBlock.ToLocation,
                                forBlock.StepExpression ?? new VB.IntegerLiteralExpression(1, VB.IntegerBase.Decimal, VB.TypeCharacter.None, forBlock.UpperBoundExpression.Span),
                                forBlock.LowerBoundExpression.Span
                            ),
                            forBlock.LowerBoundExpression.Span,
                            null
                        ),
                        loopscope
                    )
                ),
                loopscope.LoopBreak
            );

            return Expression.Block(
                GenerateExpr(
                    new VB.AssignmentStatement(
                        forBlock.ControlExpression,
                        forBlock.EqualsLocation,
                        forBlock.LowerBoundExpression,
                        forBlock.LowerBoundExpression.Span,
                        null
                    ),
                    loopscope
                ),
                myLoop
            );
        }

        public static Expression GenerateForEachBlockExpr(VB.ForEachBlockStatement forBlock,
                                                    AnalysisScope scope)
        {
            var loopscope = new AnalysisScope(scope, "loop ");
            loopscope.IsForLoop = true; // needed for break and continue
            loopscope.LoopBreak = Expression.Label("loop break");
            Expression body = GenerateBlockExpr(forBlock.Statements, loopscope);
            Expression variable;
            if (forBlock.ControlExpression is VB.SimpleNameExpression)
            {
                variable = FindIdDef(((VB.SimpleNameExpression)forBlock.ControlExpression).Name.Name, scope, true);
            }
            else
            {
                variable = GenerateExpr(forBlock.ControlExpression, loopscope);
            }
            Expression enumerable = GenerateExpr(forBlock.CollectionExpression, scope);
            ParameterExpression temp = Expression.Variable(typeof(IEnumerator), "$enumerator");

            return Expression.Block(
                new ParameterExpression[] { temp},
                Expression.Assign(temp,
                  Expression.Call(
                    Expression.Convert(
                      enumerable,
                      typeof(IEnumerable)
                    ),
                    typeof(IEnumerable).GetMethod("GetEnumerator")
                  )
                ),
                Expression.Loop(
                    Expression.Block(
                        Expression.Condition(
                            Expression.Call(
                                temp,
                                typeof(IEnumerator).GetMethod("MoveNext")
                            ),
                            Expression.Empty(),
                            Expression.Break(loopscope.LoopBreak),
                            typeof(void)
                        ),
                        GenerateExpr(
                            new VB.AssignmentStatement(
                                forBlock.ControlExpression,
                                forBlock.InLocation,
                                new ExpressionExpression(
                                    Expression.Convert(
                                        Expression.Property(
                                            temp,
                                            typeof(IEnumerator).GetProperty("Current")
                                        ),
                                        variable.Type
                                    ),
                                    forBlock.ControlExpression.Span
                                ),
                                forBlock.ControlExpression.Span,
                                null,
                                true
                                ),
                            loopscope
                        ),
                        body
                    ),
                    loopscope.LoopBreak)
                );
        }

        public static Expression GenerateWhileBlockExpr(VB.WhileBlockStatement whileBlock,
                                                  AnalysisScope scope)
        {
            var breakTarget = Expression.Label("loop break");
            var body = GenerateBlockExpr(whileBlock.Statements, scope);
            return Expression.Loop(
                Expression.Block(
                    Expression.IfThenElse(
                        WrapBooleanTest(
                            GenerateExpr(whileBlock.Expression, scope)
                        ),
                        Expression.Empty(),
                        Expression.Goto(breakTarget)
                    ),
                    body
                ),
                breakTarget
            );
        }

        public static Expression GenerateWithBlockExpr(VB.WithBlockStatement withBlock,
                                                AnalysisScope scope)
        {
            var withScope = new AnalysisScope(scope, "with");
            withScope.IsWith = true;
            withScope.WithExpression = GenerateExpr(withBlock.Expression, scope);
            return GenerateBlockExpr(withBlock.Statements, withScope); 
        }

        public static Expression GenerateDoBlockExpr(VB.DoBlockStatement doBlock,
                                          AnalysisScope scope)
        {
            //Todo: Need to take care While/Until at beginning/end
            var loopscope = new AnalysisScope(scope, "loop ");
            loopscope.IsDoLoop = true; // needed for break and continue
            loopscope.LoopBreak = Expression.Label("loop break");
            loopscope.LoopContinue = Expression.Label("loop continue");
            var body = GenerateBlockExpr(doBlock.Statements, loopscope);

            if (doBlock.Expression != null)
            {
                if (doBlock.IsWhile) //do while ... loop
                {
                    return Expression.Loop(
                        Expression.Block(
                            Expression.IfThenElse(
                                WrapBooleanTest(GenerateExpr(doBlock.Expression, loopscope)),
                                Expression.Empty(),
                                Expression.Goto(loopscope.LoopBreak)
                            ),
                            body
                        ),
                        loopscope.LoopBreak,
                        loopscope.LoopContinue
                    );
                }
                else // do until ... loop
                {
                    return Expression.Loop(
                        Expression.Block(
                            Expression.IfThen(
                                WrapBooleanTest(GenerateExpr(doBlock.Expression, loopscope)),
                                Expression.Goto(loopscope.LoopBreak)
                            ),
                            body
                        ),
                        loopscope.LoopBreak,
                        loopscope.LoopContinue
                    );
                }
            }
            else
            {
                if (doBlock.EndStatement.IsWhile) //do ... loop while
                {
                    return Expression.Loop(
                        Expression.Block(
                            body,
                            Expression.IfThenElse(
                                WrapBooleanTest(GenerateExpr(doBlock.EndStatement.Expression, loopscope)),
                                Expression.Empty(),
                                Expression.Goto(loopscope.LoopBreak)
                            )
                        ),
                        loopscope.LoopBreak,
                        loopscope.LoopContinue
                    );
                }
                else // do until ... loop
                {
                    return Expression.Loop(
                        Expression.Block(
                            body,
                            Expression.IfThen(
                                WrapBooleanTest(GenerateExpr(doBlock.EndStatement.Expression, loopscope)),
                                Expression.Goto(loopscope.LoopBreak)
                            )
                        ),
                        loopscope.LoopBreak,
                        loopscope.LoopContinue
                    );
                }
            }
        }

        public static Expression GenerateBreakExpr(VB.ExitStatement expr,
                                        AnalysisScope scope)
        {
            var exitScope = _findFirstScope(scope, expr.ExitType);
            if (exitScope == null)
                throw new InvalidOperationException("Cannot find exit target.");
            LabelTarget target = null;
            switch (expr.ExitType)
            {
                case Dlrsoft.VBScript.Parser.BlockType.Function:
                    string name = exitScope.Name;
                    return Expression.Break(exitScope.MethodExit, FindIdDef(name, scope));

                case Dlrsoft.VBScript.Parser.BlockType.Sub:
                    target = exitScope.MethodExit;
                    break;
                case Dlrsoft.VBScript.Parser.BlockType.Do:
                case Dlrsoft.VBScript.Parser.BlockType.For:
                    target = exitScope.LoopBreak;    
                    break;
            }
            return Expression.Break(target);
        }

        public static Expression GenerateNewExpr(VB.NewExpression expr,
                                                AnalysisScope scope)
        {

            VB.Name target = ((VB.NamedTypeName)expr.Target).Name;
            Expression targetExpr;
            if (target is VB.QualifiedName)
            {
                targetExpr = GenerateQualifiedNameExpr((VB.QualifiedName)target, scope);
            }
            else
            {
                targetExpr = GenerateSimpleNameExpr((VB.SimpleName)target, scope);
            }
            //args.Add(targetExpr);
            //args.AddRange(expr.Arguments.Select(a => GenerateExpr(a, scope)));
            List<Expression> args = GenerateArgumentList(expr.Arguments, scope);
            args.Insert(0, targetExpr);

            return Expression.Dynamic(
                scope.GetRuntime().GetCreateInstanceBinder(
                                     new CallInfo(expr.Arguments.Count)),
                typeof(object),
                args
            );
        }

        public static Expression GenerateBinaryExpr(VB.BinaryOperatorExpression expr,
                                                   AnalysisScope scope)
        {

            // The language has the following special logic to handle And and Or
            // x And y == if x then y
            // x Or y == if x then x else (if y then y)
            ExpressionType op;
            switch(expr.Operator)
            {
                case VB.OperatorType.Concatenate:
                    return Expression.Call(
                        typeof(HelperFunctions).GetMethod("Concatenate"),
                        RuntimeHelpers.EnsureObjectResult(GenerateExpr(expr.LeftOperand, scope)),
                        RuntimeHelpers.EnsureObjectResult(GenerateExpr(expr.RightOperand, scope))
                        );
                case VB.OperatorType.Plus:
                    op = ExpressionType.Add;
                    break;
                case VB.OperatorType.Minus:
                    op = ExpressionType.Subtract;
                    break;
                case VB.OperatorType.Multiply:
                    op = ExpressionType.Multiply;
                    break;
                case VB.OperatorType.Divide:
                    op = ExpressionType.Divide;
                    break;
                case VB.OperatorType.IntegralDivide:
                    op = ExpressionType.Divide;
                    break;
                case VB.OperatorType.Modulus:
                    op = ExpressionType.Modulo;
                    break;
                case VB.OperatorType.Equals:
                    op = ExpressionType.Equal;
                    break;
                case VB.OperatorType.NotEquals:
                    op = ExpressionType.NotEqual;
                    break;
                case VB.OperatorType.LessThan:
                    op = ExpressionType.LessThan;
                    break;
                case VB.OperatorType.GreaterThan:
                    op = ExpressionType.GreaterThan;
                    break;
                case VB.OperatorType.LessThanEquals:
                    op = ExpressionType.LessThanOrEqual;
                    break;
                case VB.OperatorType.GreaterThanEquals:
                    op = ExpressionType.GreaterThanOrEqual;
                    break;
                case VB.OperatorType.Is:
                    op = ExpressionType.Equal;
                    break;
                case VB.OperatorType.And:
                    op = ExpressionType.And;
                    break;
                case VB.OperatorType.Or:
                    op = ExpressionType.Or;
                    break;
                case VB.OperatorType.Xor:
                    op = ExpressionType.ExclusiveOr;
                    break;
                case VB.OperatorType.Power:
                    op = ExpressionType.Power;
                    break;
                default:
                    throw new InvalidOperationException("Unknown binary operator " + expr.Operator);
            }
            return Expression.Dynamic(
                scope.GetRuntime().GetBinaryOperationBinder(op),
                typeof(object),
                GenerateExpr(expr.LeftOperand, scope),
                GenerateExpr(expr.RightOperand, scope)
            );
        }

        public static Expression GenerateUnaryExpr(VB.UnaryOperatorExpression expr,
                                                  AnalysisScope scope)
        {
            ExpressionType op;
            switch (expr.Operator)
            {
                case VB.OperatorType.Negate:
                    op = ExpressionType.Negate;
                    break;
                case VB.OperatorType.Not:
                    op = ExpressionType.Not;
                    break;
                default:
                    throw new InvalidOperationException("Unknown unary operator " + expr.Operator);

            }
            return Expression.Dynamic(
                scope.GetRuntime().GetUnaryOperationBinder(op),
                typeof(object),
                GenerateExpr(expr.Operand, scope)
            );
        }

        private static void ensureBuiltinConstants()
        {
            lock (typeof(VBScriptGenerator))
                if (_builtinConstants == null)
                {
                    _builtinConstants = new Set<string>(StringComparer.InvariantCultureIgnoreCase);
                    FieldInfo[] fis = typeof(BuiltInConstants).GetFields(BindingFlags.Public | BindingFlags.Static);
                    foreach (FieldInfo fi in fis)
                    {
                        _builtinConstants.Add(fi.Name);
                    }
                }
        }

        private static object getBuiltinConstant(string name)
        {
            FieldInfo fi = typeof(BuiltInConstants).GetField(name, BindingFlags.Static | BindingFlags.Public | BindingFlags.IgnoreCase);
            return fi.GetValue(null);
        }

        private static void ensureBuiltinFunctions()
        {
            lock (typeof(VBScriptGenerator))
                if (_builtinFunctions == null)
                {
                    _builtinFunctions = new Set<string>(StringComparer.InvariantCultureIgnoreCase);
                    MethodInfo[] mis = typeof(BuiltInFunctions).GetMethods(BindingFlags.Public | BindingFlags.Static);
                    foreach (MethodInfo mi in mis)
                    {
                        _builtinFunctions.Add(mi.Name);
                    }
                }
        }

        private static Expression GenerateSimpleNameExpr(VB.SimpleName simpleName, AnalysisScope scope)
        {
            string name = simpleName.Name;
            if (IsBuiltInConstants(name))
            {
                return Expression.Constant(getBuiltinConstant(name));
            }

            var param = FindIdDef(name, scope);
            if (param != null)
            {
                return param;
            }
            else
            {
                return Expression.Dynamic(
                   scope.GetRuntime().GetGetMemberBinder(name),
                   typeof(object),
                   scope.GetModuleExpr()
                );
            }
        }

        private static Expression FindIdDef(string name, AnalysisScope scope)
        {
            return FindIdDef(name, scope, false);
        }

        // FindIdDef returns the ParameterExpr for the name by searching the scopes,
        // or it returns None.
        //
        private static Expression FindIdDef(string name, AnalysisScope scope, bool generateIfNotFound)
        {
            var curscope = scope;
            name = name.ToLower();
            ParameterExpression res;
            while (curscope != null)
            {
                if (curscope.Names.TryGetValue(name, out res))
                {
                    return res;
                }
                else
                {
                    curscope = curscope.Parent;
                }
            }

            if (generateIfNotFound)
            {
                if (scope.ModuleScope.IsOptionExplicitOn)
                    throw new InvalidOperationException(string.Format("Must declare variable {0}.", name));

                res = Expression.Parameter(typeof(object), name);
                scope.ModuleScope.Names.Add(name, res);
                return res;
            }

            if (scope == null)
            {
                throw new InvalidOperationException(
                    "Got bad AnalysisScope chain with no module at end.");
            }

            return null;
        }

        private static List<Expression> GenerateArgumentList(VB.ArgumentCollection vbargs, AnalysisScope scope)
        {
            List<Expression> args = new List<Expression>();
            if (vbargs != null)
            {
                //args.AddRange(vbargs.Select(a => GenerateExpr(a.Expression, scope)));
                foreach (VB.Argument a in vbargs)
                {
                    Expression argExp;
                    if (a != null && a.Expression != null)
                    {
                        argExp = GenerateExpr(a.Expression, scope);
                    }
                    else
                    {
                        argExp = Expression.Constant(null);
                    }
                    args.Add(argExp);
                }
            }
            return args;
        }

        private static Expression GenerateSimpleNameAssignExpr(VB.SimpleNameExpression idExpr, Expression val, AnalysisScope scope)
        {
            string varName = idExpr.Name.Name;
            var param = FindIdDef(varName, scope, true);
            return Expression.Assign(
                       param,
                       Expression.Convert(val, param.Type));
        }

        private static LambdaExpression GenerateLambdaDef
        (VB.ParameterCollection parms, VB.StatementCollection body,
         AnalysisScope scope, string name, bool isSub)
        {
            var funscope = new AnalysisScope(scope, name);
            var bodyscope = new AnalysisScope(funscope, name + " body");
            bodyscope.IsLambdaBody = true;
            funscope.IsLambda = true;  // needed for return support.
            var returnParameter = Expression.Parameter(typeof(object), name); //to support assign return value to func name
            if (isSub)
            {
                funscope.MethodExit = Expression.Label();
            }
            else
            {
                funscope.MethodExit = Expression.Label(typeof(object));
                bodyscope.Names[name] = returnParameter;
            }

            var paramsInOrder = new List<ParameterExpression>();
            int parmCount = 0;
            if (parms != null)
            {
                parmCount = parms.Count;
                foreach (var p in parms)
                {
                    Type paramType;
                    if (p.Modifiers != null && p.Modifiers.get_Item(0).ModifierType == Dlrsoft.VBScript.Parser.ModifierTypes.ByVal)
                    {
                        paramType = typeof(object);
                    }
                    else
                    {
                        paramType = typeof(object).MakeByRefType();
                    }
                    var pe = Expression.Parameter(paramType, p.VariableName.Name.Name);
                    paramsInOrder.Add(pe);
                    funscope.Names[p.VariableName.Name.Name.ToLower()] = pe;
                }
            }

            Expression bodyexpr = GenerateBlockExpr(body, bodyscope);

            // Set up the Type arg array for the delegate type.  Must include
            // the return type as the last Type, which is object for Sympl defs.
            var funcTypeArgs = new List<Type>();
            for (int i = 0; i < parmCount + 1; i++)
            {
                funcTypeArgs.Add(typeof(object));
            }

            Expression lastExpression;
            if (isSub)
            {
                //Return void
                lastExpression = Expression.Empty();
                //funcTypeArgs.Add(typeof(void));
            }
            else
            {
                lastExpression = returnParameter; //to return the function value
            }

            List<ParameterExpression> locals = new List<ParameterExpression>();
            foreach (ParameterExpression local in bodyscope.Names.Values)
            {
                locals.Add(local);
            }

            LambdaExpression lambda;
            //if (scope.GetRuntime().Debug)
            //{
            //    Expression registerRuntimeVariables = GenerateRuntimeVariablesExpression(bodyscope);

            //    lambda = Expression.Lambda(
            //           Expression.Label(
            //                funscope.MethodExit,
            //                Expression.Block(locals, registerRuntimeVariables, bodyexpr, lastExpression)
            //           ),
            //           paramsInOrder);
            //}
            //else
            //{
                lambda = Expression.Lambda(
                       Expression.Label(
                            funscope.MethodExit,
                            Expression.Block(locals, bodyexpr, lastExpression)
                       ),
                       paramsInOrder);
            //}
            //if (scope.GetRuntime().Debug)
            //{
            //    return scope.GetRuntime().DebugContext.TransformLambda(lambda);
            //}
            //else
            //{
                return lambda;
            //}
        }

        //public static Expression GenerateRuntimeVariablesExpression(AnalysisScope bodyscope)
        //{
        //    List<string> namesInScope = new List<string>();
        //    List<ParameterExpression> parametersInScope = new List<ParameterExpression>();
        //    bodyscope.GetVariablesInScope(namesInScope, parametersInScope);

        //    //Expression 
        //    RuntimeVariablesExpression runtimeVariables = Expression.RuntimeVariables(parametersInScope);
        //    Expression traceHelper = getTraceHelper(bodyscope);
        //    Expression registerRuntimeVariables = Expression.Call(
        //        traceHelper,
        //        typeof(ITrace).GetMethod("RegisterRuntimeVariables"),
        //        Expression.Constant(namesInScope.ToArray()),
        //        runtimeVariables
        //    );
        //    return registerRuntimeVariables;
        //}

        private static Expression WrapBooleanTest(Expression expr)
        {
            var tmp = Expression.Parameter(typeof(object), "testtmp");
            return Expression.Block(
                new ParameterExpression[] { tmp },
                new Expression[] 
                        {Expression.Assign(tmp, Expression
                                                  .Convert(expr, typeof(object))),
                         Expression.Condition(
                             Expression.TypeIs(tmp, typeof(bool)), 
                             Expression.Convert(tmp, typeof(bool)),
                             Expression.Call(typeof(BuiltInFunctions).GetMethod("CBool"), tmp))});
        }

        private static Expression ConvertToIntegerArrayExpression(List<Expression> args)
        {
            List<Expression> converted = new List<Expression>();
            foreach (Expression arg in args)
            {
                if (arg.Type == typeof(int))
                {
                    converted.Add(arg);
                }
                else
                {
                    converted.Add(Expression.Convert(arg, typeof(int)));
                }
            }
            //return converted;
            return Expression.NewArrayInit(typeof(int), converted);
        }

        internal static Expression WrapTryCatchExpression(Expression stmt, Dlrsoft.VBScript.Compiler.AnalysisScope scope)
        {
            ParameterExpression exception = Expression.Parameter(typeof(Exception));

            return Expression.TryCatch(
                Expression.Block(
                    stmt,
                    Expression.Empty()
                ),
                Expression.Catch(
                    exception,
                    Expression.Call(
                        typeof(HelperFunctions).GetMethod("SetError"),
                        scope.ErrExpression,
                        exception
                    )
                )
            );
        }

        internal static Expression GenerateDebugInfo(VB.Tree stmt, AnalysisScope scope, out Expression clearDebugInfo)
        {
            ISourceMapper mapper = scope.ModuleScope.SourceMapper;
            DocSpan docSpan = mapper.Map(SourceUtil.ConvertSpan(stmt.Span));
            SourceLocation start = docSpan.Span.Start;
            SourceLocation end = docSpan.Span.End;
            SymbolDocumentInfo docInfo = scope.ModuleScope.GetDocumentInfo(docSpan.Uri);

            //Expression debugInfo = Expression.DebugInfo(docInfo, start.Line, start.Column, end.Line, end.Column);
            //clearDebugInfo = Expression.ClearDebugInfo(docInfo);
            //return debugInfo;

            Expression traceHelper = getTraceHelper(scope);

            Expression debugInfo = Expression.Call(
                traceHelper,
                typeof(ITrace).GetMethod("TraceDebugInfo"),
                Expression.Constant(docInfo.FileName),
                Expression.Constant(start.Line),
                Expression.Constant(start.Column),
                Expression.Constant(end.Line),
                Expression.Constant(end.Column)
            );
            clearDebugInfo = Expression.Empty();
            return debugInfo;
        }

        private static Expression getTraceHelper(AnalysisScope scope)
        {
            Expression traceHelper = scope.TraceExpression;
            return traceHelper;
        }

        // _findFirstLoop returns the first loop AnalysisScope or None.
        //
        private static AnalysisScope _findFirstScope(AnalysisScope scope, VB.BlockType blockType)
        {
            var curscope = scope;
            while (curscope != null)
            {
                switch (blockType)
                {
                    case Dlrsoft.VBScript.Parser.BlockType.Sub:
                    case Dlrsoft.VBScript.Parser.BlockType.Function:
                        if (curscope.IsLambda)
                        {
                            return curscope;
                        }
                        break;
                    case Dlrsoft.VBScript.Parser.BlockType.Do:
                        if (curscope.IsDoLoop)
                        {
                            return curscope;
                        }
                        break;
                    case Dlrsoft.VBScript.Parser.BlockType.For:
                        if (curscope.IsForLoop)
                        {
                            return curscope;
                        }
                        break;
                }
                curscope = curscope.Parent;

            }
            return null;
        }

        /// <summary>
        /// Wrap around a DLR expression to inject into the VB expression tree
        /// </summary>
        class ExpressionExpression : VB.Expression
        {
            Expression _expression;

            public ExpressionExpression(Expression expr, VB.Span span)
                : base(VB.TreeType.AddressOfExpression, span)
            {
                _expression = expr;
            }

            public Expression Expression
            {
                get { return _expression; }
            }
        }
    }
}
