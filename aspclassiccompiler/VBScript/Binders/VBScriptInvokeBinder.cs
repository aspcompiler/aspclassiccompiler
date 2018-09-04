using System;
using System.Collections.Generic;
//using System.Text;
using System.Reflection;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
#if !SILVERLIGHT
using Microsoft.Scripting.ComInterop;
#endif
#else
using System.Linq;
using System.Linq.Expressions;
using Microsoft.Scripting.ComInterop;
#endif
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScript.Binders
{
    public class VBScriptInvokeBinder : InvokeBinder
    {
        public VBScriptInvokeBinder(CallInfo callinfo)
            : base(callinfo)
        {
        }

        public override DynamicMetaObject FallbackInvoke(
                DynamicMetaObject targetMO, DynamicMetaObject[] argMOs,
                DynamicMetaObject errorSuggestion)
        {
#if !SILVERLIGHT
            // First try COM binding.
            DynamicMetaObject result;

            if (ComBinder.TryBindInvoke(this, targetMO, argMOs, out result))
            {
                return result;
            }
#endif
            // Defer if any object has no value so that we evaulate their
            // Expressions and nest a CallSite for the InvokeMember.
            if (!targetMO.HasValue || argMOs.Any((a) => !a.HasValue))
            {
                var deferArgs = new DynamicMetaObject[argMOs.Length + 1];
                for (int i = 0; i < argMOs.Length; i++)
                {
                    deferArgs[i + 1] = argMOs[i];
                }
                deferArgs[0] = targetMO;
                return Defer(deferArgs);
            }
            // Find our own binding.
            if (targetMO.LimitType.IsSubclassOf(typeof(Delegate)))
            {
                var parms = targetMO.LimitType.GetMethod("Invoke").GetParameters();
                if (parms.Length == argMOs.Length)
                {
                    // Don't need to check if argument types match parameters.
                    // If they don't, users get an argument conversion error.
                    var callArgs = RuntimeHelpers.ConvertArguments(argMOs, parms);
                    var expression = Expression.Invoke(
                        Expression.Convert(targetMO.Expression, targetMO.LimitType),
                        callArgs);
                    return new DynamicMetaObject(
                        RuntimeHelpers.EnsureObjectResult(expression),
                        BindingRestrictions.GetTypeRestriction(targetMO.Expression,
                                                               targetMO.LimitType));
                }
            }
            return errorSuggestion ??
                RuntimeHelpers.CreateThrow(
                    targetMO, argMOs,
                    BindingRestrictions.GetTypeRestriction(targetMO.Expression,
                                                           targetMO.LimitType),
                    typeof(InvalidOperationException),
                    "Wrong number of arguments for function -- " +
                    targetMO.LimitType.ToString() + " got " + argMOs.ToString());

        }
    }
}
