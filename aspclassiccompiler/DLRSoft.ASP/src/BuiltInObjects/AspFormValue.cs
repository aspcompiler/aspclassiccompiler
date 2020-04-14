using System;
using System.Collections;
using System.Collections.Specialized;
using System.Text;
using System.Web;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// A collection object that can properly represent an ASP form or querystring value object
	/// </summary>
	public class AspFormValue:StringCollection
	{
		private HttpContext _context;

		public AspFormValue():this(HttpContext.Current, "")
		{
		}

		public AspFormValue(HttpContext context, string fieldName)
		{
			_context=context;
			if (_context.Request.Form[fieldName]!=null)
				foreach(string element in _context.Request.Form[fieldName].Split(','))
					base.Add(element);
		}

		public override string ToString()
		{
			return (string)this;
		}

		public static implicit operator string(AspFormValue collection)
		{
			if (collection.Count==0)
				return "";
			StringBuilder sb=new StringBuilder(collection.Count*2);
			foreach(string key in collection)
			{
				sb.Append(key);
				sb.Append(',');
			}
			sb.Remove(sb.Length-1, 1);
			return sb.ToString();
		}
	}
}
