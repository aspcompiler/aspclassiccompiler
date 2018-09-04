using System;
using System.Collections;
using System.Collections.Specialized;
using System.Text;
using System.Web;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// A collection object that can properly represent an ASP form or querystring object
	/// </summary>
	public class AspForm : ICollection, IEnumerable
	{
		private HttpContext _context;

		public AspForm():this(HttpContext.Current)
		{
		}

		public AspForm(HttpContext context)
		{
			_context=context;
		}

		public void CopyTo(Array dest, int index)
		{
			_context.Request.Form.CopyTo(dest, index);
		}

		public int Count
		{
			get {return _context.Request.Form.Count;}
		}

		public bool IsSynchronized
		{
			get {return false;}
		}

		public object SyncRoot
		{
			get {return ((ICollection)(_context.Request.Form)).SyncRoot;}
		}

		public IEnumerator GetEnumerator()
		{
			return _context.Request.Form.GetEnumerator();
		}

		public NameObjectCollectionBase.KeysCollection Keys
		{
			get {return _context.Request.Form.Keys;}
		}

		public AspFormValue this[string key]
		{
			get {return new AspFormValue(_context, key);}
		}

		public override string ToString()
		{
			return (string)this;
		}

		public static implicit operator string(AspForm collection)
		{
			if (collection.Count==0)
				return "";
			StringBuilder sb=new StringBuilder(collection.Count*4);
			foreach(string key in collection.Keys)
			{
				sb.Append(key);
				sb.Append('=');
				sb.Append((string)collection[key]);
				sb.Append('&');
			}
			sb.Remove(sb.Length-1, 1);
			return sb.ToString();
		}
	}
}
