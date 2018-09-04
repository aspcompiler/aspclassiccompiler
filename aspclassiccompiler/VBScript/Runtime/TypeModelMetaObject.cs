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

namespace Dlrsoft.VBScript.Runtime
{
    public class TypeModelMetaObject : DynamicMetaObject
    {
        private TypeModel _typeModel;

        public TypeModel TypeModel { get { return _typeModel; } }
        public Type ReflType { get { return _typeModel.ReflType; } }

        // Constructor takes ParameterExpr to reference CallSite, and a TypeModel
        // that the new TypeModelMetaObject represents.
        //
        public TypeModelMetaObject(Expression objParam, TypeModel typeModel)
            : base(objParam, BindingRestrictions.Empty, typeModel)
        {
            _typeModel = typeModel;
        }

        public override DynamicMetaObject BindGetMember(GetMemberBinder binder)
        {
            var flags = BindingFlags.IgnoreCase | BindingFlags.Static |
                        BindingFlags.Public;
            // consider BindingFlags.Instance if want to return wrapper for
            // inst members that is callable.
            var members = ReflType.GetMember(binder.Name, flags);
            if (members.Length == 1)
            {
                return new DynamicMetaObject(
                    // We always access static members for type model objects, so the
                    // first argument in MakeMemberAccess should be null.
                    RuntimeHelpers.EnsureObjectResult(
                        Expression.MakeMemberAccess(
                            null,
                            members[0])),
                    // Don't need restriction test for name since this
                    // rule is only used where binder is used, which is
                    // only used in sites with this binder.Name.
                    this.Restrictions.Merge(
                        BindingRestrictions.GetInstanceRestriction(
                            this.Expression,
                            this.Value)
                    )
                );
            }
            else
            {
                return binder.FallbackGetMember(this);
            }
        }

        // Because we don't ComboBind over several MOs and operations, and no one
        // is falling back to this function with MOs that have no values, we
        // don't need to check HasValue.  If we did check, and HasValue == False,
        // then would defer to new InvokeMemberBinder.Defer().
        //
        public override DynamicMetaObject BindInvokeMember(
                InvokeMemberBinder binder, DynamicMetaObject[] args)
        {
            var flags = BindingFlags.IgnoreCase | BindingFlags.Static |
                        BindingFlags.Public;
            var members = ReflType.GetMember(binder.Name, flags);
            if ((members.Length == 1) && (members[0] is PropertyInfo ||
                                          members[0] is FieldInfo))
            {
                // NEED TO TEST, should check for delegate value too
                var mem = members[0];
                throw new NotImplementedException();
                //return new DynamicMetaObject(
                //    Expression.Dynamic(
                //        new SymplInvokeBinder(new CallInfo(args.Length)),
                //        typeof(object),
                //        args.Select(a => a.Expression).AddFirst(
                //               Expression.MakeMemberAccess(this.Expression, mem)));

                // Don't test for eventinfos since we do nothing with them now.
            }
            else
            {
                // Get MethodInfos with right arg counts.
                var mi_mems = members.
                    Select(m => m as MethodInfo).
                    Where(m => m is MethodInfo &&
                               ((MethodInfo)m).GetParameters().Length ==
                                   args.Length);
                // Get MethodInfos with param types that work for args.  This works
                // for except for value args that need to pass to reftype params. 
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
                // True below means generate an instance restriction on the MO.
                // We are only looking at the members defined in this Type instance.
                var restrictions = RuntimeHelpers.GetTargetArgsRestrictions(
                    this, args, true);

                if (res.Count == 0)
                {
                    //LC Try to find method that support params
                    //Limited to the case params is the only argument and it is and object[]
                    //To bind something like BuiltInFunction.Array
                    MethodInfo mi = ReflType.GetMethod(binder.Name, new Type[] { typeof(object[]) });

                    if (mi != null)
                    {
                        return new DynamicMetaObject(
                          RuntimeHelpers.EnsureObjectResult(
                              Expression.Call(
                                mi,
                                Expression.NewArrayInit(typeof(object), args.Select(a => a.Expression))
                              )
                          ),
                          restrictions);
                    }

                    // Sometimes when binding members on TypeModels the member
                    // is an intance member since the Type is an instance of Type.
                    // We fallback to the binder with the Type instance to see if
                    // it binds.  The SymplInvokeMemberBinder does handle this.
                    var typeMO = RuntimeHelpers.GetRuntimeTypeMoFromModel(this);
                    var result = binder.FallbackInvokeMember(typeMO, args, null);
                    return result;
                }
                // restrictions and conversion must be done consistently.
                var callArgs =
                    RuntimeHelpers.ConvertArguments(
                    args, res[0].GetParameters());

                return new DynamicMetaObject(
                   RuntimeHelpers.EnsureObjectResult(
                       Expression.Call(res[0], callArgs)),
                   restrictions);
                // Could hve tried just letting Expr.Call factory do the work,
                // but if there is more than one applicable method using just
                // assignablefrom, Expr.Call throws.  It does not pick a "most
                // applicable" method or any method.

            }
        }

        public override DynamicMetaObject BindCreateInstance(
                   CreateInstanceBinder binder, DynamicMetaObject[] args)
        {
            var constructors = ReflType.GetConstructors();
            var ctors = constructors.
                Where(c => c.GetParameters().Length == args.Length);
            List<ConstructorInfo> res = new List<ConstructorInfo>();
            foreach (var c in ctors)
            {
                if (RuntimeHelpers.ParametersMatchArguments(c.GetParameters(),
                                                            args))
                {
                    res.Add(c);
                }
            }
            if (res.Count == 0)
            {
                // Binders won't know what to do with TypeModels, so pass the
                // RuntimeType they represent.  The binder might not be Sympl's.
                return binder.FallbackCreateInstance(
                                  RuntimeHelpers.GetRuntimeTypeMoFromModel(this),
                                  args);
            }
            // For create instance of a TypeModel, we can create a instance 
            // restriction on the MO, hence the true arg.
            var restrictions = RuntimeHelpers.GetTargetArgsRestrictions(
                                this, args, true);
            var ctorArgs =
                RuntimeHelpers.ConvertArguments(
                args, res[0].GetParameters());
            return new DynamicMetaObject(
                // Creating an object, so don't need EnsureObjectResult.
                Expression.New(res[0], ctorArgs),
               restrictions);
        }
    }//TypeModelMetaObject
}
