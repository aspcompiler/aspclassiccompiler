using System;
using System.Collections.Generic;
//using System.Text;
using System.Reflection;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif
using Debug = System.Diagnostics.Debug;
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScript.Binders
{
    public class VBScriptBinaryOperationBinder : BinaryOperationBinder
    {
        public VBScriptBinaryOperationBinder(ExpressionType operation)
            : base(operation)
        {
        }

        public override DynamicMetaObject FallbackBinaryOperation(
                    DynamicMetaObject target, DynamicMetaObject arg,
                    DynamicMetaObject errorSuggestion)
        {
            // Defer if any object has no value so that we evaulate their
            // Expressions and nest a CallSite for the InvokeMember.
            if (!target.HasValue || !arg.HasValue)
            {
                return Defer(target, arg);
            }
            var restrictions = target.Restrictions.Merge(arg.Restrictions);

            if (target.Value == null)
            {
                restrictions = restrictions.Merge(
                    BindingRestrictions.GetInstanceRestriction(target.Expression, null)
                );
            }
            else
            {
                restrictions = restrictions.Merge(BindingRestrictions.GetTypeRestriction(
                    target.Expression, target.LimitType));
            }

            if (arg.Value == null)
            {
                restrictions = restrictions.Merge(BindingRestrictions.GetInstanceRestriction(
                    arg.Expression, null));
            }
            else
            {
                restrictions = restrictions.Merge(BindingRestrictions.GetTypeRestriction(
                    arg.Expression, arg.LimitType));
            }

            if (this.Operation == ExpressionType.Add)
            {
                if (target.LimitType == typeof(string) && arg.LimitType == typeof(string))
                {
                    return new DynamicMetaObject(
                      RuntimeHelpers.EnsureObjectResult(Expression.Call(
                        typeof(string).GetMethod("Concat", new Type[] { typeof(string), typeof(string) }),
                        Expression.Convert(target.Expression, target.LimitType),
                        Expression.Convert(arg.Expression, arg.LimitType))),
                      restrictions
                    );
                }
                if (target.LimitType == typeof(DateTime))
                {
                    return new DynamicMetaObject(
                       RuntimeHelpers.EnsureObjectResult( Expression.Call(
                            typeof(BuiltInFunctions).GetMethod("DateAdd"),
                            Expression.Constant("d"),
                            Expression.Convert(arg.Expression, typeof(object)),
                            target.Expression)),
                            restrictions
                        );
                }
                else if (arg.LimitType == typeof(DateTime))
                {
                    return new DynamicMetaObject(
                        Expression.Call(
                            typeof(BuiltInFunctions).GetMethod("DateAdd"),
                            Expression.Constant("d"),
                            Expression.Convert(target.Expression, typeof(object)),
                            arg.Expression),
                            restrictions
                        );
                }
                else if (!target.LimitType.IsPrimitive || !arg.LimitType.IsPrimitive)
                {
                    DynamicMetaObject[] args = new DynamicMetaObject[] { target, arg };
                    List<MethodInfo> res = RuntimeHelpers.GetExtensionMethods("op_Addition", target, args);
                    if (res.Count > 0)
                    {
                        MethodInfo mi = null;
                        if (res.Count > 1)
                        {
                            //If more than one results found, attempt overload resolution
                            mi = RuntimeHelpers.ResolveOverload(res, args);
                        }
                        else
                        {
                            mi = res[0];
                        }

                        // restrictions and conversion must be done consistently.
                        var callArgs = RuntimeHelpers.ConvertArguments(args, mi.GetParameters());

                        return new DynamicMetaObject(
                            RuntimeHelpers.EnsureObjectResult(
                                Expression.Call(null, mi, callArgs)
                            ),
                            restrictions
                        );
                    }
                }
            }

            if (target.LimitType.IsPrimitive && arg.LimitType.IsPrimitive)
            {
                Type targetType;
                Expression first;
                Expression second;
                if (target.LimitType == arg.LimitType || target.LimitType.IsAssignableFrom(arg.LimitType))
                {
                    targetType = target.LimitType;
                    first = Expression.Convert(target.Expression, targetType);
                    //if variable is object type, need to convert twice (unbox + convert)
                    second = Expression.Convert(
                                Expression.Convert(arg.Expression, arg.LimitType),
                                targetType
                             );
                }
                else
                {
                    targetType = arg.LimitType;
                    first = Expression.Convert(
                                Expression.Convert(target.Expression, target.LimitType),
                                targetType
                             );
                    second = Expression.Convert(arg.Expression, targetType);
                }

                return new DynamicMetaObject(
                    RuntimeHelpers.EnsureObjectResult(
                      Expression.MakeBinary(
                        this.Operation,
                        first,
                        second
                      )
                    ),
                    restrictions
                );
            }
            else
            {
                DynamicMetaObject[] args = null;
                MethodInfo mi = null;

                mi = typeof(HelperFunctions).GetMethod("BinaryOp");
                Expression op = Expression.Constant(this.Operation);
                DynamicMetaObject mop = new DynamicMetaObject(Expression.Constant(this.Operation), BindingRestrictions.Empty, this.Operation);

                args = new DynamicMetaObject[] { mop, target, arg };
                // restrictions and conversion must be done consistently.
                var callArgs = RuntimeHelpers.ConvertArguments(args, mi.GetParameters());

                return new DynamicMetaObject(
                    RuntimeHelpers.EnsureObjectResult(
                        Expression.Call(
                            mi,
                            callArgs
                        )
                    ),
                    restrictions
                );
            }
        }
    }
}
