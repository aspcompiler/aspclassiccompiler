using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ASPTypeLibrary;
using System.Web;

namespace Dlrsoft.Asp.BuiltInObjects
{
    public class AspVariantDictionary : IVariantDictionary
    {
        private HttpApplicationState _state;

        public AspVariantDictionary(HttpApplicationState state)
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
            _state.Remove(Convert.ToString(VarKey));
        }

        public void RemoveAll()
        {
            _state.RemoveAll();
        }

        public object get_Key(object VarKey)
        {
            return _state.GetKey(Convert.ToInt32(VarKey));
        }

        public void let_Item(object VarKey, object pvar)
        {
            if (VarKey is int)
            {
                VarKey = _state.GetKey((int)VarKey);

            }
            
            _state[Convert.ToString(VarKey)] = pvar;
        }

        public object this[object VarKey]
        {
            get
            {
                if (VarKey is int)
                {
                    return _state[(int)VarKey];
                }
                else
                {
                    return _state[Convert.ToString(VarKey)];
                }
            }
            set
            {
                let_Item(VarKey, value);
            }
        }

        #endregion
    }
}
