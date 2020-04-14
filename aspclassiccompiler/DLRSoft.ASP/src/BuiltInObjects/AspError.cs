using System;
using ASPTypeLibrary;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// The asperror object that is accessible from ASP code
	/// </summary>
	public class AspError : IASPError
	{
		private string _aspcode="";
		private int _number=0;
		private string _source="";
		private string _category="";
		private string _file="";
		private int _line=0;
		private int _column=0;
		private string _description="";
		private string _aspdescription="";

		public AspError()
		{
		}

		public AspError(string aspcode, int number, string source, string category, string file, int line, int column, string description, string aspdescription)
		{
			_aspcode=aspcode;
			_number=number;
			_source=source;
			_category=category;
			_file=file;
			_line=line;
			_column=column;
			_description=description;
			_aspdescription=aspdescription;
		}

		public string ASPCode
		{
			get {return _aspcode;}
		}

		public int Number
		{
			get {return _number;}
		}

		public string Source
		{
			get {return _source;}
		}

		public string Category
		{
			get {return _category;}
		}

		public string File
		{
			get {return _file;}
		}

		public int Line
		{
			get {return _line;}
		}

		public int Column
		{
			get {return _column;}
		}

		public string Description
		{
			get {return _description;}
		}

		public string ASPDescription
		{
			get {return _aspdescription;}
		}
	}
}
