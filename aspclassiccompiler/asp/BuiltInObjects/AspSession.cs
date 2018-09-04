using System;
using System.Collections;
using System.Collections.Specialized;
using System.Web;
using System.Web.SessionState;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// The session object that is accessible from ASP code
	/// </summary>
	public class AspSession
	{
		private HttpContext _context;

		public AspSession()
		{
			_context=HttpContext.Current;
		}

		public AspSession(HttpContext context)
		{
			_context=context;
		}

		public object this[string key]
		{
			get {return _context.Session[key];}
			set {_context.Session[key]=value;}
		}

		public object StaticObjects
		{
			get {return _context.Session["__Session.StaticObjects"];}
		}

		public HttpSessionState Contents
		{
			get {return _context.Session.Contents;}
		}

		public void Abandon()
		{
			_context.Session.Abandon();
		}

		public int CodePage
		{
			get {return _context.Session.CodePage;}
			set {_context.Session.CodePage=value;}
		}

		public int LCID
		{
			get {return _context.Session.LCID;}
			set {_context.Session.LCID=value;}
		}

		public string SessionID
		{
			get {return _context.Session.SessionID;}
		}

		public int TimeOut
		{
			get {return _context.Session.Timeout;}
			set {_context.Session.Timeout=value;}
		}
	}
}
