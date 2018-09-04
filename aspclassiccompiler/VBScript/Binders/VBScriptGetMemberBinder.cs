using System;
using System.Collections.Generic;
//using System.Text;
using System.Reflection;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
using Microsoft.Scripting.ComInterop;
#else
using System.Linq.Expressions;
using Microsoft.Scripting.ComInterop;
#endif
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScript.Binders
{
    // VBScriptGetMemberBinder is used for general dotted expressions for fetching
    // members.
    //
    public class VBScriptGetMemberBinder : GetMemberBinder
    {
        public VBScriptGetMemberBinder(string name)
            : base(name, true)
        {
        }

        public override DynamicMetaObject FallbackGetMember(
                DynamicMetaObject targetMO, DynamicMetaObject errorSuggestion)
        {
#if !SILVERLIGHT
            // First try COM binding.
            DynamicMetaObject result;
            if (ComBinder.TryBindGetMember(this, targetMO, out result))
            {
                return result;
            }
#endif
            // Defer if any object has no value so that we evaulate their
            // Expressions and nest a CallSite for the InvokeMember.
            if (!targetMO.HasValue) return Defer(targetMO);
            // Find our own binding.
            var flags = BindingFlags.IgnoreCase | BindingFlags.Static |
                        BindingFlags.Instance | BindingFlags.Public;
            var members = targetMO.LimitType.GetMember(this.Name, flags);
            if (members.Length == 1)
            {
                return new DynamicMetaObject(
                    RuntimeHelpers.EnsureObjectResult(
                      Expression.MakeMemberAccess(
                        Expression.Convert(targetMO.Expression,
                                           members[0].DeclaringType),
                        members[0])),
                    // Don't need restriction test for name since this
                    // rule is only used where binder is used, which is
                    // only used in sites with this binder.Name.
                    BindingRestrictions.GetTypeRestriction(targetMO.Expression,
                                                           targetMO.LimitType));
            }
            else
            {
                return errorSuggestion ??
                    RuntimeHelpers.CreateThrow(
                        targetMO, null,
                        BindingRestrictions.GetTypeRestriction(targetMO.Expression,
                                                               targetMO.LimitType),
                        typeof(MissingMemberException),
                        "cannot bind member, " + this.Name +
                            ", on object " + targetMO.Value.ToString());
            }
        }
    }
}
