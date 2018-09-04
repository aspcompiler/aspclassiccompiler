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
using Debug = System.Diagnostics.Debug;

namespace Dlrsoft.VBScript.Binders
{
    // VBScriptInvokeMemberBinder is used for general dotted expressions in function
    // calls for invoking members.
    //
    public class VBScriptInvokeMemberBinder : InvokeMemberBinder
    {
        public VBScriptInvokeMemberBinder(string name, CallInfo callinfo)
            : base(name, true, callinfo)
        { // true = ignoreCase
        }

        public override DynamicMetaObject FallbackInvokeMember(
                DynamicMetaObject targetMO, DynamicMetaObject[] args,
                DynamicMetaObject errorSuggestion)
        {

#if !SILVERLIGHT
            // First try COM binding.
            DynamicMetaObject result;
            if (ComBinder.TryBindInvokeMember(this, targetMO, args, out result))
            {
                return result;
            }
#endif
            // Defer if any object has no value so that we evaulate their
            // Expressions and nest a CallSite for the InvokeMember.
            if (!targetMO.HasValue || args.Any((a) => !a.HasValue))
            {
                var deferArgs = new DynamicMetaObject[args.Length + 1];
                for (int i = 0; i < args.Length; i++)
                {
                    deferArgs[i + 1] = args[i];
                }
                deferArgs[0] = targetMO;
                return Defer(deferArgs);
            }
            // Find our own binding.
            // Could consider allowing invoking static members from an instance.
            var flags = BindingFlags.IgnoreCase | BindingFlags.Instance |
                        BindingFlags.Public;
            var members = targetMO.LimitType.GetMember(this.Name, flags);

            var restrictions = RuntimeHelpers.GetTargetArgsRestrictions(
                                                  targetMO, args, false);

            // Assigned a Null
            if (targetMO.HasValue && targetMO.Value == null)
            {
                return RuntimeHelpers.CreateThrow(targetMO, 
                    args, 
                    restrictions, 
                    typeof(NullReferenceException), 
                    "Object reference not set to an instance of an object.");
            }

            if ((members.Length == 1) && (members[0] is PropertyInfo ||
                                          members[0] is FieldInfo))
            {
                // NEED TO TEST, should check for delegate value too
                var mem = members[0];
                var target = new DynamicMetaObject(
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
                
                //If no arguments, to allow scenario like Request.QueryString()
                if (args == null || args.Length == 0)
                {
                    return target;
                }

                return new DynamicMetaObject(
                    RuntimeHelpers.GetIndexingExpression(target, args),
                    restrictions
                );
                // Don't test for eventinfos since we do nothing with them now.
            }
            else
            {
                bool isExtension = false;

                // Get MethodInfos with right arg counts.
                var mi_mems = members.
                    Select(m => m as MethodInfo).
                    Where(m => m is MethodInfo &&
                               ((MethodInfo)m).GetParameters().Length ==
                                   args.Length);
                // Get MethodInfos with param types that work for args.  This works
                // except for value args that need to pass to reftype params. 
                // We could detect that to be smarter and then explicitly StrongBox
                // the args.
                List<MethodInfo> res = new List<MethodInfo>();
                foreach (var mem in mi_mems)
                {
                    if (RuntimeHelpers.ParametersMatchArguments(
                                           mem.GetParameters(), args))
                    {
                        res.Add(mem);
                    }
                }

                List<DynamicMetaObject> argList = new List<DynamicMetaObject>(args);

                //Try extension methods if no methods found
                if (res.Count == 0)
                {
                    isExtension = true;
                    argList.Insert(0, targetMO);

                    res = RuntimeHelpers.GetExtensionMethods(this.Name, targetMO, argList.ToArray());
                }

                // False below means generate a type restriction on the MO.
                // We are looking at the members targetMO's Type.
                if (res.Count == 0)
                {
                    return errorSuggestion ??
                        RuntimeHelpers.CreateThrow(
                            targetMO, args, restrictions,
                            typeof(MissingMemberException),
                            string.Format("Can't bind member invoke {0}.{1}({2})", targetMO.RuntimeType.Name, this.Name, args.ToString()));
                }

                //If more than one results found, attempt overload resolution
                MethodInfo mi = null;
                if (res.Count > 1)
                {
                    mi = RuntimeHelpers.ResolveOverload(res, argList.ToArray());
                }
                else
                {
                    mi = res[0];
                }

                // restrictions and conversion must be done consistently.
                var callArgs = RuntimeHelpers.ConvertArguments(
                                                 argList.ToArray(), mi.GetParameters());
                
                if (isExtension)
                {
                    return new DynamicMetaObject(
                       RuntimeHelpers.EnsureObjectResult(
                         Expression.Call(
                            null,
                            mi, callArgs)),
                       restrictions);
                }
                else
                {
                    return new DynamicMetaObject(
                       RuntimeHelpers.EnsureObjectResult(
                         Expression.Call(
                            Expression.Convert(targetMO.Expression,
                                               targetMO.LimitType),
                            mi, callArgs)),
                       restrictions);
                }
                // Could hve tried just letting Expr.Call factory do the work,
                // but if there is more than one applicable method using just
                // assignablefrom, Expr.Call throws.  It does not pick a "most
                // applicable" method or any method.
            }
        }

        public override DynamicMetaObject FallbackInvoke(
                DynamicMetaObject targetMO, DynamicMetaObject[] args,
                DynamicMetaObject errorSuggestion)
        {
            var argexprs = new Expression[args.Length + 1];
            for (int i = 0; i < args.Length; i++)
            {
                argexprs[i + 1] = args[i].Expression;
            }
            argexprs[0] = targetMO.Expression;
            // Just "defer" since we have code in SymplInvokeBinder that knows
            // what to do, and typically this fallback is from a language like
            // Python that passes a DynamicMetaObject with HasValue == false.
            return new DynamicMetaObject(
                           Expression.Dynamic(
                // This call site doesn't share any L2 caching
                // since we don't call GetInvokeBinder from Sympl.
                // We aren't plumbed to get the runtime instance here.
                               new VBScriptInvokeBinder(new CallInfo(args.Length)),
                               typeof(object), // ret type
                               argexprs),
                // No new restrictions since SymplInvokeBinder will handle it.
                           targetMO.Restrictions.Merge(
                               BindingRestrictions.Combine(args)));
        }
    }
}
