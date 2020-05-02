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
    // VBScriptSetMemberBinder is used for general dotted expressions for setting
    // members.
    //
    public class VBScriptSetMemberBinder : SetMemberBinder
    {
        public VBScriptSetMemberBinder(string name)
            : base(name, true)
        {
        }

        public override DynamicMetaObject FallbackSetMember(
                DynamicMetaObject targetMO, DynamicMetaObject value,
                DynamicMetaObject errorSuggestion)
        {
#if !SILVERLIGHT
            // First try COM binding.
            DynamicMetaObject result;
            if (ComBinder.TryBindSetMember(this, targetMO, value, out result))
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
                MemberInfo mem = members[0];
                Expression val;
                // Should check for member domain type being Type and value being
                // TypeModel, similar to ConvertArguments, and building an
                // expression like GetRuntimeTypeMoFromModel.
                if (mem.MemberType == MemberTypes.Property)
                    val = Expression.Convert(value.Expression,
                                             ((PropertyInfo)mem).PropertyType);
                else if (mem.MemberType == MemberTypes.Field)
                    val = Expression.Convert(value.Expression,
                                             ((FieldInfo)mem).FieldType);
                else
                    return (errorSuggestion ??
                            RuntimeHelpers.CreateThrow(
                                targetMO, null,
                                BindingRestrictions.GetTypeRestriction(
                                    targetMO.Expression,
                                    targetMO.LimitType),
                                typeof(InvalidOperationException),
                                "Sympl only supports setting Properties and " +
                                "fields at this time."));
                return new DynamicMetaObject(
                    // Assign returns the stored value, so we're good for Sympl.
                    RuntimeHelpers.EnsureObjectResult(
                      Expression.Assign(
                        Expression.MakeMemberAccess(
                            Expression.Convert(targetMO.Expression,
                                               members[0].DeclaringType),
                            members[0]),
                        val)),
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
                         "IDynObj member name conflict.");
            }
        }
    }
}
