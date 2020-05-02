using System;
using System.Collections.Generic;
using System.Text;
using System.Web;
using ASPTypeLibrary;

namespace Dlrsoft.Asp
{
    class AspReadCookie : IReadCookie
    {
        private HttpCookie _cookie = null;

        public AspReadCookie(HttpCookie cookie)
        {
            _cookie = cookie;
        }

        #region IReadCookie Members

        public int Count
        {
            get { return _cookie == null ? 0 : _cookie.Values.Count; }
        }

        public System.Collections.IEnumerator GetEnumerator()
        {
            return _cookie.Values.GetEnumerator();
        }

        public bool HasKeys
        {
            get { return _cookie == null ? false : _cookie.HasKeys; }
        }

        public object get_Key(object VarKey)
        {
            throw new NotImplementedException();
        }

        public object this[object key]
        {
            get 
            {
                if (_cookie == null) return null;

                return _cookie[(string)key]; 
            }
        }
        #endregion

        public override string ToString()
        {
            if (_cookie == null) return "";

            StringBuilder sb = new StringBuilder();
            bool first = true;
            foreach (string key in _cookie.Values.Keys)
            {
                if (first)
                {
                    first = false;
                }
                else
                {
                    sb.Append('&');
                }
                sb.Append(key);
                sb.Append('=');
                sb.Append(_cookie[key]);
            }
            return sb.ToString();
        }
    }
}
