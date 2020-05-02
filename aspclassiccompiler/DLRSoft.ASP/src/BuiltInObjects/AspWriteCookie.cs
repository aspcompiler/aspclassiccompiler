using System;
using System.Collections;
using System.Web;
using ASPTypeLibrary;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// A cookie for use with asp scripts
	/// </summary>
	public class AspWriteCookie : IWriteCookie
	{
		private HttpCookie _cookie=null;

		public AspWriteCookie(HttpCookie cookie)
		{
			_cookie=cookie;
		}

        #region IWriteCookie Members

        public IEnumerator GetEnumerator()
        {
            //return new CookieEnumerator(_cookie);
            return _cookie.Values.GetEnumerator();
        }

        public string Domain
        {
            set { _cookie.Domain = value; }
        }

        public DateTime Expires
        {
            set { _cookie.Expires = value; }
        }

        public bool HasKeys
        {
            get { return _cookie == null ? false : _cookie.HasKeys; }
        }

        public string Path
        {
            set { _cookie.Path = value; }
        }

        public bool Secure
        {
            set { _cookie.Secure = value; }
        }

        public string this[object Key]
        {
            set { _cookie[Key.ToString()] = value; }
        }

        #endregion
    }
}
