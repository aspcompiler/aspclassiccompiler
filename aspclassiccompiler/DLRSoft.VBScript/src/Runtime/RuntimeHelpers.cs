using System;
using System.Collections.Generic;
//using System.Text;
using System.Reflection;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
using Microsoft.Scripting.Utils;
#else
using System.Linq;
using System.Linq.Expressions;
#endif
using System.Runtime.CompilerServices;
using Path = System.IO.Path;
using File = System.IO.File;
using Directory = System.IO.Directory;
using Debug = System.Diagnostics.Debug;
#if SILVERLIGHT
using w = System.Windows;
using System.Windows.Resources;
#endif

namespace Dlrsoft.VBScript.Runtime
{

    // RuntimeHelpers is a collection of functions that perform operations at
    // runtime of Sympl code, such as performing an import or eq.
    //
    public static class RuntimeHelpers {
        // VBScriptImport takes the runtime and module as context for the import.
        // It takes a list of names, what, that either identify a (possibly dotted
        // sequence) of names to fetch from Globals or a file name to load.  Names
        // is a list of names to fetch from the final object that what indicates
        // and then set each name in module.  Renames is a list of names to add to
        // module instead of names.  If names is empty, then the name set in
        // module is the last name in what.  If renames is not empty, it must have
        // the same cardinality as names.
        //
        public static object VBScriptImport(VBScript runtime, IDynamicMetaObjectProvider module,
                                         string[] what, string[] names,
                                         string[] renames)
        {
            // Get object or file scope.
            object value = null;
            if (what.Length == 1)
            {
                string name = what[0];
                if (DynamicObjectHelpers.HasMember(runtime.Globals, name))
                {
                    value = DynamicObjectHelpers.GetMember(runtime.Globals, name);
                    // Since runtime.Globals has Sympl's reflection of namespaces and
                    // types, we pick those up first above and don't risk hitting a
                    // NamespaceTracker for assemblies added when we initialized Sympl.
                    // The next check will correctly look up case-INsensitively for
                    // globals the host adds to ScriptRuntime.Globals.
                }
                else if (DynamicObjectHelpers.HasMember(runtime.DlrGlobals, name))
                {
                    value = DynamicObjectHelpers.GetMember(runtime.DlrGlobals, name);
                }
                else
                {
                    throw new ArgumentException(
                        "Import: can't find name in globals -- " + name);
                }
            }
            else
            {
                // What has more than one name, must be Globals access.
                value = runtime.Globals;
                // For more correctness and generality, shouldn't assume all
                // globals are dynamic objects, or that a look up like foo.bar.baz
                // cascades through all dynamic objects.
                // Would need to manually create a CallSite here with Sympl's
                // GetMemberBinder, and think about a caching strategy per name.
                foreach (string name in what)
                {
                    value = DynamicObjectHelpers.GetMember(
                                (IDynamicMetaObjectProvider)value, name);
                }
            }
            // Assign variables in module.
            if (names.Length == 0)
            {
                if (renames.Length == 0)
                {
                    DynamicObjectHelpers.SetMember((IDynamicMetaObjectProvider)module,
                                                   what[what.Length - 1], value);
                }
                else
                {
                    DynamicObjectHelpers.SetMember((IDynamicMetaObjectProvider)module,
                                                   renames[0], value);
                }
            }
            else
            {
                if (renames.Length == 0) renames = names;
                for (int i = 0; i < names.Length; i++)
                {
                    string name = names[i];
                    string rename = renames[i];
                    DynamicObjectHelpers.SetMember(
                        (IDynamicMetaObjectProvider)module, rename,
                        DynamicObjectHelpers.GetMember(
                             (IDynamicMetaObjectProvider)value, name));
                }
            }
            return null;
        } // SymplImport

        // Uses of the 'eq' keyword form in Sympl compile to a call to this
        // helper function.
        //
        public static bool SymplEq (object x, object y) {
            if (x == null)
                return y == null;
            else if (y == null)
                return x == null;
            else {
                var xtype = x.GetType();
                var ytype = y.GetType();
                if (xtype.IsPrimitive && xtype != typeof(string) &&
                    ytype.IsPrimitive && ytype != typeof(string))
                    return x.Equals(y);
                else
                    return object.ReferenceEquals(x, y);
            }
        }

        //////////////////////////////////////////////////
        // Array Utilities (slicing) and some LINQ helpers
        //////////////////////////////////////////////////

        public static T[] RemoveFirstElt<T>(IList<T> list) {
            // Make array ...
            if (list.Count == 0) {
                return new T[0];
            }
            T[] res = new T[list.Count];
            list.CopyTo(res, 0);
            // Shift result
            return ShiftLeft(res, 1);
        }

        public static T[] RemoveFirstElt<T>(T[] array) {
            return ShiftLeft(array, 1);
        }

        private static T[] ShiftLeft<T>(T[] array, int count) {
            //ContractUtils.RequiresNotNull(array, "array");
            if (count < 0) throw new ArgumentOutOfRangeException("count");
            T[] result = new T[array.Length - count];
            System.Array.Copy(array, count, result, 0, result.Length);
            return result;
        }

        public static T[] RemoveLast<T>(T[] array) {
            //ContractUtils.RequiresNotNull(array, "array");
            System.Array.Resize(ref array, array.Length - 1);
            return array;
        }

#if USE35
        // Need to reproduce these helpers from DLR codeplex-only sources and
        // from LINQ functionality to avoid referencing System.Core.dll and
        // Microsoft.Scripting.Core.dll for internal building.

        internal static IEnumerable<U> Select<T, U>(this IEnumerable<T> enumerable, Func<T, U> select) {
            foreach (T t in enumerable) {
                yield return select(t);
            }
        }

        internal static IEnumerable<T> Where<T>(this IEnumerable<T> enumerable, Func<T, bool> where) {
            foreach (T t in enumerable) {
                if (where(t)) {
                    yield return t;
                }
            }
        }

        internal static bool Any<T>(this IEnumerable<T> source, Func<T, bool> predicate) {
            foreach (T element in source) {
                if (predicate(element)) {
                    return true;
                }
            }
            return false;
        }

        internal static T[] ToArray<T>(this IEnumerable<T> enumerable) {
            var c = enumerable as ICollection<T>;
            if (c != null) {
                var result = new T[c.Count];
                c.CopyTo(result, 0);
                return result;
            }
            return new List<T>(enumerable).ToArray();
        }

        internal static TSource Last<TSource>(this IList<TSource> list) {
            if (list == null) throw new ArgumentNullException("list");
            int count = list.Count;
            if (count > 0) return list[count - 1];
            throw new ArgumentException("list is empty");
        }

        internal static bool Contains<TSource>(this IEnumerable<TSource> source, TSource value) {
            IEqualityComparer<TSource> comparer = EqualityComparer<TSource>.Default;
            if (source == null) throw new ArgumentNullException("source");
            foreach (TSource element in source)
                if (comparer.Equals(element, value)) return true;
            return false;
        }
#endif
      
  
        ///////////////////////////////////////
        // Utilities used by binders at runtime
        ///////////////////////////////////////

        // ParamsMatchArgs returns whether the args are assignable to the parameters.
        // We specially check for our TypeModel that wraps .NET's RuntimeType, and
        // elsewhere we detect the same situation to convert the TypeModel for calls.
        //
        // Consider checking p.IsByRef and returning false since that's not CLS.
        //
        // Could check for a.HasValue and a.Value is None and
        // ((paramtype is class or interface) or (paramtype is generic and
        // nullable<t>)) to support passing nil anywhere.
        //
        public static bool ParametersMatchArguments(ParameterInfo[] parameters,
                                                    DynamicMetaObject[] args) {
            // We only call this after filtering members by this constraint.
            Debug.Assert(args.Length == parameters.Length,
                         "Internal: args are not same len as params?!");
            for (int i = 0; i < args.Length; i++) {
                var paramType = parameters[i].ParameterType;
                // We consider arg of TypeModel and param of Type to be compatible.
                if (paramType == typeof(Type) &&
                    (args[i].LimitType == typeof(TypeModel))) {
                    continue;
                }
                //LC OK to assign null to any reference type
                if (!paramType.IsValueType && args[i].Value == null)
                    continue;

                if (!paramType
                    // Could check for HasValue and Value==null AND
                    // (paramtype is class or interface) or (is generic
                    // and nullable<T>) ... to bind nullables and null.
                        .IsAssignableFrom(args[i].LimitType)) {
                    return false;
                }
            }
            return true;
        }

        public static int[] GetMatchRank(ParameterInfo[] parameters,
                                                    DynamicMetaObject[] args)
        {
            Debug.Assert(args.Length == parameters.Length,
                         "Internal: args are not same len as params?!");
            int[] rank = new int[parameters.Length];
            for (int i = 0; i < args.Length; i++)
            {
                var paramType = parameters[i].ParameterType;
                // We consider arg of TypeModel and param of Type to be compatible.
                if (paramType == typeof(Type) &&
                    (args[i].LimitType == typeof(TypeModel)))
                {
                    rank[i] = 0;
                }
                else if (paramType == args[i].LimitType)
                {
                    rank[i] = 0;
                }
                else
                {
                    rank[i] = 1;
                }
            }
            return rank;
        }

        public static MethodInfo ResolveOverload(IList<MethodInfo> mis, DynamicMetaObject[] args)
        {
            MethodInfo best = mis[0];
            int[] bestRank = GetMatchRank(mis[0].GetParameters(), args);
            foreach (MethodInfo mi in mis)
            {
                int[] rank = GetMatchRank(mi.GetParameters(), args);
                if (Compare(rank, bestRank) < 0)
                {
                    best = mi;
                    bestRank = rank;
                }
            }
            return best;
        }

        private static int Compare(int[] a, int[] b)
        {
            for (int i = 0; i < a.Length; i++)
            {
                int result = a[i].CompareTo(b[i]);
                if (result == 0)
                    continue;

                return result;
            }
            return 0;
        }

        // Returns a DynamicMetaObject with an expression that fishes the .NET
        // RuntimeType object from the TypeModel MO.
        //
        public static DynamicMetaObject GetRuntimeTypeMoFromModel(
                                              DynamicMetaObject typeModelMO) {
            Debug.Assert((typeModelMO.LimitType == typeof(TypeModel)),
                         "Internal: MO is not a TypeModel?!");
            // Get tm.ReflType
            var pi = typeof(TypeModel).GetProperty("ReflType");
            Debug.Assert(pi != null);
            return new DynamicMetaObject(
                Expression.Property(
                    Expression.Convert(typeModelMO.Expression, typeof(TypeModel)),
                    pi),
                typeModelMO.Restrictions.Merge(
                    BindingRestrictions.GetTypeRestriction(
                        typeModelMO.Expression, typeof(TypeModel)))//,
                // Must supply a value to prevent binder FallbackXXX methods
                // from infinitely looping if they do not check this MO for
                // HasValue == false and call Defer.  After Sympl added Defer
                // checks, we could verify, say, FallbackInvokeMember by no
                // longer passing a value here.
                //((TypeModel)typeModelMO.Value).ReflType
            );
        }

        // Returns list of Convert exprs converting args to param types.  If an arg
        // is a TypeModel, then we treat it special to perform the binding.  We need
        // to map from our runtime model to .NET's RuntimeType object to match.
        //
        // To call this function, args and pinfos must be the same length, and param
        // types must be assignable from args.
        //
        // NOTE, if using this function, then need to use GetTargetArgsRestrictions
        // and make sure you're performing the same conversions as restrictions.
        //
        public static Expression[] ConvertArguments(
                                 DynamicMetaObject[] args, ParameterInfo[] ps) {
            Debug.Assert(args.Length == ps.Length,
                         "Internal: args are not same len as params?!");
            Expression[] callArgs = new Expression[args.Length];
            for (int i = 0; i < args.Length; i++) {
                Expression argExpr = args[i].Expression;
                if (args[i].LimitType == typeof(TypeModel) && 
                    ps[i].ParameterType == typeof(Type)) {
                    // Get arg.ReflType
                    argExpr = GetRuntimeTypeMoFromModel(args[i]).Expression;
                }
                Type paramType;
                if (ps[i].ParameterType.IsByRef)
                {
                    paramType = ps[i].ParameterType.GetElementType();
                }
                else
                {
                    paramType = ps[i].ParameterType;
                }
                if (argExpr.Type != paramType)
                {
                    argExpr = Expression.Convert(argExpr, paramType);
                }
                callArgs[i] = argExpr;
            }
            return callArgs;
        }

        public static Expression ConvertExpression(Expression expr, Type type)
        {
            if (type == expr.Type)
            {
                return expr;
            }
            else if (type == typeof(string))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CStr"),
                    expr
                );
            }
            else if (type == typeof(int))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CLng"),
                    expr
                );
            }
            else if (type == typeof(double))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CDbl"),
                    expr
                );
            }
            else if (type == typeof(bool))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CBool"),
                    expr
                );
            }
            else if (type == typeof(decimal))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CCur"),
                    expr
                );
            }
            else if (type == typeof(DateTime))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CDate"),
                    expr
                );
            }
            else if (type == typeof(short))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CInt"),
                    expr
                );
            }
            else if (type == typeof(float))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CSng"),
                    expr
                );
            }
            else if (type == typeof(byte))
            {
                return Expression.Call(
                    typeof(BuiltInFunctions).GetMethod("CByte"),
                    expr
                );
            }
            else
            {
                return Expression.Convert(expr, type);
            }
        }
        
        // GetTargetArgsRestrictions generates the restrictions needed for the
        // MO resulting from binding an operation.  This combines all existing
        // restrictions and adds some for arg conversions.  targetInst indicates
        // whether to restrict the target to an instance (for operations on type
        // objects) or to a type (for operations on an instance of that type).
        //
        // NOTE, this function should only be used when the caller is converting
        // arguments to the same types as these restrictions.
        //
        public static BindingRestrictions GetTargetArgsRestrictions(
                DynamicMetaObject target, DynamicMetaObject[] args,
                bool instanceRestrictionOnTarget){
            // Important to add existing restriction first because the
            // DynamicMetaObjects (and possibly values) we're looking at depend
            // on the pre-existing restrictions holding true.
            var restrictions = target.Restrictions.Merge(BindingRestrictions
                                                            .Combine(args));
            //LC Use instance restriction is value is null
            if (instanceRestrictionOnTarget || target.Value == null)
            {
                restrictions = restrictions.Merge(
                    BindingRestrictions.GetInstanceRestriction(
                        target.Expression,
                        target.Value
                    ));
            } else {
                restrictions = restrictions.Merge(
                    BindingRestrictions.GetTypeRestriction(
                        target.Expression,
                        target.LimitType
                    ));
            }
            for (int i = 0; i < args.Length; i++) {
                BindingRestrictions r;
                if (args[i].HasValue && args[i].Value == null) {
                    r = BindingRestrictions.GetInstanceRestriction(
                            args[i].Expression, null);
                } else {
                    r = BindingRestrictions.GetTypeRestriction(
                            args[i].Expression, args[i].LimitType);
                }
                restrictions = restrictions.Merge(r);
            }
            return restrictions;
        }

        // Return the expression for getting target[indexes]
        //
        // Note, callers must ensure consistent restrictions are added for
        // the conversions on args and target.
        //
        public static Expression GetIndexingExpression(
                                      DynamicMetaObject target,
                                      DynamicMetaObject[] indexes) {
            //Debug.Assert(target.HasValue && target.LimitType != typeof(Array));

            // ARRAY
            if (target.LimitType.IsArray)
            {
                var indexExpressions = indexes.Select(
                    i => Expression.Convert(i.Expression, i.LimitType))
                    .ToArray();

                return Expression.ArrayAccess(
                    Expression.Convert(target.Expression,
                                       target.LimitType),
                    indexExpressions
                );
             // INDEXER
            } else {
                var props = target.LimitType.GetProperties();
                var indexers = props.
                    Where(p => p.GetIndexParameters().Length > 0).ToArray();
                indexers = indexers.
                    Where(idx => idx.GetIndexParameters().Length == 
                                 indexes.Length).ToArray();

                var res = new List<PropertyInfo>();
                foreach (var idxer in indexers) {
                    if (RuntimeHelpers.ParametersMatchArguments(
                                          idxer.GetIndexParameters(), indexes)) {
                        // all parameter types match
                        res.Add(idxer);
                    }
                }
                if (res.Count == 0) {
                    return Expression.Throw(
                        Expression.New(
                            typeof(MissingMemberException)
                                .GetConstructor(new Type[] { typeof(string) }),
                            Expression.Constant(
                               "Can't bind because there is no matching indexer.")
                        )
                    );
                }
                return Expression.MakeIndex(
                    Expression.Convert(target.Expression, target.LimitType),
                    res[0], ConvertArguments(indexes, res[0].GetIndexParameters()));
            }
        }

        // CreateThrow is a convenience function for when binders cannot bind.
        // They need to return a DynamicMetaObject with appropriate restrictions
        // that throws.  Binders never just throw due to the protocol since
        // a binder or MO down the line may provide an implementation.
        //
        // It returns a DynamicMetaObject whose expr throws the exception, and 
        // ensures the expr's type is object to satisfy the CallSite return type
        // constraint.
        //
        // A couple of calls to CreateThrow already have the args and target
        // restrictions merged in, but BindingRestrictions.Merge doesn't add 
        // duplicates.
        //
        public static DynamicMetaObject CreateThrow
                (DynamicMetaObject target, DynamicMetaObject[] args,
                 BindingRestrictions moreTests,
                 Type exception, params object[] exceptionArgs) {
            Expression[] argExprs = null;
            Type[] argTypes = Type.EmptyTypes;
            int i;
            if (exceptionArgs != null) {
                i = exceptionArgs.Length;
                argExprs = new Expression[i];
                argTypes = new Type[i];
                i = 0;
                foreach (object o in exceptionArgs) {
                    Expression e = Expression.Constant(o);
                    argExprs[i] = e;
                    argTypes[i] = e.Type;
                    i += 1;
                }
            }
            ConstructorInfo constructor = exception.GetConstructor(argTypes);
            if (constructor == null) {
                throw new ArgumentException(
                    "Type doesn't have constructor with a given signature");
            }
            return new DynamicMetaObject(
                Expression.Throw(
                    Expression.New(constructor, argExprs),
                    // Force expression to be type object so that DLR CallSite
                    // code things only type object flows out of the CallSite.
                    typeof(object)),
                target.Restrictions.Merge(BindingRestrictions.Combine(args))
                                   .Merge(moreTests));
        }

        // EnsureObjectResult wraps expr if necessary so that any binder or
        // DynamicMetaObject result expression returns object.  This is required
        // by CallSites.
        //
        public static Expression EnsureObjectResult (Expression expr) {
            if (! expr.Type.IsValueType)
                return expr;
            if (expr.Type == typeof(void))
                return Expression.Block(
                           expr, Expression.Default(typeof(object)));
            else
                return Expression.Convert(expr, typeof(object));
        }

        public static Expression GetDefaultValue(Expression target)
        {
            return Expression.Call(
                typeof(HelperFunctions).GetMethod("GetDefaultPropertyValue", new Type[] { typeof(object) }),
                EnsureObjectResult(target)
            );
        }

        public static List<MethodInfo> GetExtensionMethods(string name, DynamicMetaObject targetMO, DynamicMetaObject[] args)
        {
            List<MethodInfo> res = new List<MethodInfo>();
#if !SILVERLIGHT
            foreach (Assembly a in AppDomain.CurrentDomain.GetAssemblies())
            {
#else
            foreach (w.AssemblyPart ap in w.Deployment.Current.Parts)
            {
                StreamResourceInfo sri = w.Application.GetResourceStream(new Uri(ap.Source, UriKind.Relative)); 
                Assembly a = new w.AssemblyPart().Load(sri.Stream); 
#endif
                Type[] types;
                try
                {
                    types = a.GetTypes();
                }
                catch (ReflectionTypeLoadException ex)
                {
                    types = ex.Types;
                }

                foreach (Type t in types)
                {
                    if (t != null && t.IsPublic && t.IsAbstract && t.IsSealed)
                    {
                        foreach (MethodInfo mem in t.GetMember(name, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public | BindingFlags.IgnoreCase))
                        {
                            //Should test using mem.IsDefined(attr) but typeof(ExtenssionAttribute) in Microsoft.Scription.Extensions and System.Core are infact two different types
                            //Type attr = typeof(System.Runtime.CompilerServices.ExtensionAttribute);
                            if (mem.GetParameters().Length == args.Length
                                && mem.GetParameters()[0].ParameterType == targetMO.RuntimeType)
                            {
                                if (RuntimeHelpers.ParametersMatchArguments(
                                                       mem.GetParameters(), args))
                                {
                                    res.Add(mem);
                                }
                            }
                        }
                    }
                }
            }

            return res;
        }

        public static bool IsNumericType(Type t)
        {
            if (t == typeof(int) || t == typeof(long) ||  t == typeof(float) || t == typeof(double) || t == typeof(decimal) || t == typeof(short) || t == typeof(sbyte))
                return true;

            return false;
        }
    } // RuntimeHelpers

} // namespace
