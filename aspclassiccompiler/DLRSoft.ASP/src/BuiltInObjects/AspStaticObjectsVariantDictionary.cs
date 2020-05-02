using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ASPTypeLibrary;
using System.Web;

namespace Dlrsoft.Asp.BuiltInObjects
{
    public class AspStaticObjectsVariantDictionary : IVariantDictionary
    {
        private HttpStaticObjectsCollection _state;

        public AspStaticObjectsVariantDictionary(HttpStaticObjectsCollection state)
        {
            _state = state;
        }

        #region IVariantDictionary Members

        public int Count
        {
            get { return _state.Count; }
        }

        public System.Collections.IEnumerator GetEnumerator()
        {
            return _state.GetEnumerator();
        }

        public void Remove(object VarKey)
        {
            throw new NotImplementedException();
        }

        public void RemoveAll()
        {
            throw new NotImplementedException();
        }

        public object get_Key(object VarKey)
        {
            throw new NotImplementedException();
        }

        public void let_Item(object VarKey, object pvar)
        {
            throw new NotImplementedException();
        }

        public object this[object VarKey]
        {
            get
            {
                return _state[Convert.ToString(VarKey)];
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        #endregion
    }
}
