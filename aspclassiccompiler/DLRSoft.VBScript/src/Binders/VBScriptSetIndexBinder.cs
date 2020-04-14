using System;
using System.Collections.Generic;
//using System.Text;
using System.Reflection;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
using Microsoft.Scripting.ComInterop;
#else
using System.Linq;
using System.Linq.Expressions;
using Microsoft.Scripting.ComInterop;
#endif
using Debug = System.Diagnostics.Debug;
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScript.Binders
{
    public class VBScriptSetIndexBinder : SetIndexBinder
    {
        public VBScriptSetIndexBinder(CallInfo callinfo)
            : base(callinfo)
        {
        }

        public override DynamicMetaObject FallbackSetIndex(
                   DynamicMetaObject target, DynamicMetaObject[] indexes,
                   DynamicMetaObject value, DynamicMetaObject errorSuggestion)
        {
#if !SILVERLIGHT
            // First try COM binding.
            DynamicMetaObject result;
            if (ComBinder.TryBindSetIndex(this, target, indexes, value, out result))
            {
                return result;
            }
# endif
            // Defer if any object has no value so that we evaulate their
            // Expressions and nest a CallSite for the InvokeMember.
            if (!target.HasValue || indexes.Any((a) => !a.HasValue) ||
                !value.HasValue)
            {
                var deferArgs = new DynamicMetaObject[indexes.Length + 2];
                for (int i = 0; i < indexes.Length; i++)
                {
                    deferArgs[i + 1] = indexes[i];
                }
                deferArgs[0] = target;
                deferArgs[indexes.Length + 1] = value;
                return Defer(deferArgs);
            }
            // Find our own binding.
            Expression valueExpr = value.Expression;
            //we convert a value of TypeModel to Type.
            if (value.LimitType == typeof(TypeModel))
            {
                valueExpr = RuntimeHelpers.GetRuntimeTypeMoFromModel(value).Expression;
            }
            Debug.Assert(target.HasValue && target.LimitType != typeof(Array));
            Expression setIndexExpr;
            Expression indexingExpr = RuntimeHelpers.GetIndexingExpression(
                                                        target, indexes);
            if (valueExpr.Type != indexingExpr.Type  && !indexingExpr.Type.IsAssignableFrom(valueExpr.Type))
            {
                valueExpr = RuntimeHelpers.ConvertExpression(valueExpr, indexingExpr.Type);
            }
            
            setIndexExpr = Expression.Assign(indexingExpr, valueExpr);

            //TODO review the restriction because the value may be important for proper conversion
            //Also need to handle the null value
            BindingRestrictions restrictions =
                 RuntimeHelpers.GetTargetArgsRestrictions(target, indexes, false);
            return new DynamicMetaObject(
                RuntimeHelpers.EnsureObjectResult(setIndexExpr),
                restrictions);

        }
    }
}
