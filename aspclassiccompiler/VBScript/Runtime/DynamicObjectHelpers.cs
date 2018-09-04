using System;
using System.Collections.Generic;
//using System.Text;
using System.Reflection;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
using Microsoft.Scripting.Utils;
#else
using System.Linq.Expressions;
#endif
using System.Runtime.CompilerServices;

namespace Dlrsoft.VBScript.Runtime
{
    //#####################################################
    //# Dynamic Helpers for HasMember, GetMember, SetMember
    //#####################################################

    // DynamicObjectHelpers provides access to IDynObj members given names as
    // data at runtime.  When the names are known at compile time (o.foo), then
    // they get baked into specific sites with specific binders that encapsulate
    // the name.  We need this in python because hasattr et al are case-sensitive.
    //
    class DynamicObjectHelpers
    {

        static private object _sentinel = new object();
        static internal object Sentinel { get { return _sentinel; } }

        internal static bool HasMember(IDynamicMetaObjectProvider o,
                                       string name)
        {
            return (DynamicObjectHelpers.GetMember(o, name) !=
                    DynamicObjectHelpers.Sentinel);
            //Alternative impl used when EOs had bug and didn't call fallback ...
            //var mo = o.GetMetaObject(Expression.Parameter(typeof(object), null));
            //foreach (string member in mo.GetDynamicMemberNames()) {
            //    if (string.Equals(member, name, StringComparison.OrdinalIgnoreCase)) {
            //        return true;
            //    }
            //}
            //return false;
        }

        static private Dictionary<string,
                                  CallSite<Func<CallSite, object, object>>>
            _getSites = new Dictionary<string,
                                       CallSite<Func<CallSite, object, object>>>();

        internal static object GetMember(IDynamicMetaObjectProvider o,
                                         string name)
        {
            CallSite<Func<CallSite, object, object>> site;
            if (!DynamicObjectHelpers._getSites.TryGetValue(name, out site))
            {
                site = CallSite<Func<CallSite, object, object>>
                               .Create(new DoHelpersGetMemberBinder(name));
                DynamicObjectHelpers._getSites[name] = site;
            }
            return site.Target(site, o);
        }

        static private Dictionary<string,
                                  CallSite<Action<CallSite, object, object>>>
            _setSites = new Dictionary<string,
                                       CallSite<Action<CallSite, object, object>>>();

        internal static void SetMember(IDynamicMetaObjectProvider o, string name,
                                       object value)
        {
            CallSite<Action<CallSite, object, object>> site;
            if (!DynamicObjectHelpers._setSites.TryGetValue(name, out site))
            {
                site = CallSite<Action<CallSite, object, object>>
                          .Create(new DoHelpersSetMemberBinder(name));
                DynamicObjectHelpers._setSites[name] = site;
            }
            site.Target(site, o, value);
        }

    }

    class DoHelpersGetMemberBinder : GetMemberBinder
    {

        internal DoHelpersGetMemberBinder(string name) : base(name, true) { }

        public override DynamicMetaObject FallbackGetMember(
                DynamicMetaObject target, DynamicMetaObject errorSuggestion)
        {
            return errorSuggestion ??
                   new DynamicMetaObject(
                           Expression.Constant(DynamicObjectHelpers.Sentinel),
                           target.Restrictions.Merge(
                               BindingRestrictions.GetTypeRestriction(
                                   target.Expression, target.LimitType)));

        }
    }

    class DoHelpersSetMemberBinder : SetMemberBinder
    {
        internal DoHelpersSetMemberBinder(string name) : base(name, true) { }

        public override DynamicMetaObject FallbackSetMember(
                DynamicMetaObject target, DynamicMetaObject value,
                DynamicMetaObject errorSuggestion)
        {
            return errorSuggestion ??
                   RuntimeHelpers.CreateThrow(
                       target, null, BindingRestrictions.Empty,
                       typeof(MissingMemberException),
                              "If IDynObj doesn't support setting members, " +
                              "DOHelpers can't do it for the IDO.");
        }
    }
}
