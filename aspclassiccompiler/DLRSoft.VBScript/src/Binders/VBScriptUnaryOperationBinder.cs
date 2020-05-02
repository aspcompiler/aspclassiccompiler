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
using Dlrsoft.VBScript.Runtime;
using Debug = System.Diagnostics.Debug;

namespace Dlrsoft.VBScript.Binders
{
    public class VBScriptUnaryOperationBinder : UnaryOperationBinder
    {
        public VBScriptUnaryOperationBinder(ExpressionType operation)
            : base(operation)
        {
        }

        public override DynamicMetaObject FallbackUnaryOperation(
                   DynamicMetaObject target, DynamicMetaObject errorSuggestion)
        {
            // Defer if any object has no value so that we evaulate their
            // Expressions and nest a CallSite for the InvokeMember.
            if (!target.HasValue)
            {
                return Defer(target);
            }

            var restrictions = target.Restrictions;
            if (target.Value == null)
            {
                restrictions = restrictions.Merge(
                    BindingRestrictions.GetInstanceRestriction(target.Expression, null)
                );
            }
            else
            {
                restrictions = restrictions.Merge(
                    BindingRestrictions.GetTypeRestriction(
                        target.Expression, target.LimitType));
            }

            if (this.Operation == ExpressionType.Not)
            {
                if (target.LimitType != typeof(bool))
                {
                    MethodInfo mi = typeof(HelperFunctions).GetMethod("Not");

                    return new DynamicMetaObject(
                       RuntimeHelpers.EnsureObjectResult(
                         Expression.Call(mi, target.Expression)
                       ),
                       restrictions);
                }
            }
            else if (this.Operation == ExpressionType.Negate)
            {
                if (!target.LimitType.IsPrimitive || target.LimitType == typeof(bool))
                {
                    MethodInfo mi = typeof(HelperFunctions).GetMethod("Negate");

                    return new DynamicMetaObject(
                       RuntimeHelpers.EnsureObjectResult(
                         Expression.Call(mi, target.Expression)
                       ),
                       restrictions);
                }
            }

            return new DynamicMetaObject(
                RuntimeHelpers.EnsureObjectResult(
                  Expression.MakeUnary(
                    this.Operation,
                    Expression.Convert(target.Expression, target.LimitType),
                    target.LimitType)),
                restrictions);
        }
    }
}
