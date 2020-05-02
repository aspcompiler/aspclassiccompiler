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
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScript.Binders
{
    public class VBScriptGetIndexBinder : GetIndexBinder
    {
        public VBScriptGetIndexBinder(CallInfo callinfo)
            : base(callinfo)
        {
        }

        public override DynamicMetaObject FallbackGetIndex(
                     DynamicMetaObject target, DynamicMetaObject[] indexes,
                     DynamicMetaObject errorSuggestion)
        {
#if !SILVERLIGHT
            // First try COM binding.
            DynamicMetaObject result;
            if (ComBinder.TryBindGetIndex(this, target, indexes, out result))
            {
                return result;
            }
#endif
            // Defer if any object has no value so that we evaulate their
            // Expressions and nest a CallSite for the InvokeMember.
            if (!target.HasValue || indexes.Any((a) => !a.HasValue))
            {
                var deferArgs = new DynamicMetaObject[indexes.Length + 1];
                for (int i = 0; i < indexes.Length; i++)
                {
                    deferArgs[i + 1] = indexes[i];
                }
                deferArgs[0] = target;
                return Defer(deferArgs);
            }

            var restrictions = RuntimeHelpers.GetTargetArgsRestrictions(
                                                  target, indexes, false);

            if (target.HasValue && target.Value == null)
            {
                return errorSuggestion?? RuntimeHelpers.CreateThrow(target,
                    indexes,
                    restrictions,
                    typeof(NullReferenceException),
                    "Object reference not set to an instance of an object.");
            }
            var indexingExpr = RuntimeHelpers.EnsureObjectResult(
                                  RuntimeHelpers.GetIndexingExpression(target,
                                                                       indexes));
            return new DynamicMetaObject(indexingExpr, restrictions);
        }
    }
}
