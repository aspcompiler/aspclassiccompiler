using System;
using System.Collections;
using System.Web;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// A cookie for use with asp scripts
	/// </summary>
	public class AspCookie : ICollection
	{
		private HttpCookie _cookie=null;
		public HttpCookie HttpCookie
		{
			get {return _cookie;}
			set {_cookie=value;}
		}

		public AspCookie(HttpCookie cookie)
		{
			_cookie=cookie;
		}

		public string this[string subkey]
		{
			get {return _cookie==null?"":_cookie[subkey];}
			set {_cookie[subkey]=value;}
		}

		public bool HasKeys
		{
			get {return _cookie==null?false:_cookie.HasKeys;}
		}

		public string Value
		{
			get {return _cookie==null?"":_cookie.Value;}
			set {_cookie.Value=value;}
		}

		public DateTime Expires
		{
			get {return _cookie.Expires;}
			set {_cookie.Expires=value;}
		}

		public string Domain
		{
			get {return _cookie.Domain;}
			set {_cookie.Domain=value;}
		}

		public string Name
		{
			get {return _cookie==null?"":_cookie.Name;}
			set {_cookie.Name=value;}
		}

		public string Path
		{
			get {return _cookie==null?"":_cookie.Path;}
			set {_cookie.Path=value;}
		}

		public bool Secure
		{
			get {return _cookie==null?false:_cookie.Secure;}
			set {_cookie.Secure=value;}
		}

		// ICollection implementation
		public int Count
		{
			get {return _cookie==null?0:_cookie.Values.Count;}
		}

		public bool IsSynchronized
		{
			get {return false;}
		}

		public object SyncRoot
		{
			get {return _cookie;}
		}

		public void CopyTo(Array array, int index)
		{
			if (_cookie.HasKeys)
				_cookie.Values.CopyTo(array, index);
			else
				array.SetValue(_cookie.Value, index);
		}

		// IEnumerable implementation
		public IEnumerator GetEnumerator()
		{
			return new CookieEnumerator(_cookie);
		}

		public class CookieEnumerator : IEnumerator
		{
			private IEnumerator _cookieenum=null;
			public CookieEnumerator(HttpCookie cookie)
			{
				_cookieenum=cookie.Values.Keys.GetEnumerator();
			}

			public object Current
			{
				get {return _cookieenum.Current;}
			}

			public bool MoveNext()
			{
				return _cookieenum.MoveNext();
			}

			public void Reset()
			{
				_cookieenum.Reset();
			}
		}
	}
}
