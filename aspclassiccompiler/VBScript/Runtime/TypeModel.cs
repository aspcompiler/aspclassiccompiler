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
    ////////////////////////////////////
    // TypeModel and TypeModelMetaObject
    ////////////////////////////////////

    // TypeModel wraps System.Runtimetypes. When Sympl code encounters
    // a type leaf node in Sympl.Globals and tries to invoke a member, wrapping
    // the ReflectionTypes in TypeModels allows member access to get the type's
    // members and not ReflectionType's members.
    //
    public class TypeModel : IDynamicMetaObjectProvider
    {
        private Type _reflType;

        public TypeModel(Type type)
        {
            _reflType = type;
        }

        public Type ReflType { get { return _reflType; } }

        DynamicMetaObject IDynamicMetaObjectProvider
                              .GetMetaObject(Expression parameter)
        {
            return new TypeModelMetaObject(parameter, this);
        }
    }
}
