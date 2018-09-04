# Features and Limitations

## VBScript.net compiler features:
- For 1.0 release, we plan to implement VBScript 4.0 (see [http://msdn.microsoft.com/en-us/library/4y5y7bh5(VS.85).aspx](VBScript Version Information)). Features like defining classes in VBScript and the Eval/Execute/ExecuteGlobal are not implemented, although DLR does allow the implementation of these features. We plan to implement these features in a later release. However, if you would like to make these a higher priority, please vote in the Issue Tracker.
	- You can still use 5.x Regular Expression features since we are reusing Microsoft VBScript regular expression COM objects through COM interop. However, you must have Microsoft VBScript 5.x installed on your system.
	- The 5.0 Escape, GetLocal, SetLocal, Timer and Unescape functions are supported.
- Interop with .net framework
	- All the types in VBScript.net compiler are natively CLR Types. Please also see [http://msdn.microsoft.com/en-us/library/9e7a57cf(VS.85).aspx](VBScript Data Types). The following are the mappings between the types in VBScript and the CLR types used VBScript.net:
|| VBScript type || VBScript.net type||
| Empty | null |
| Null | System.DbNull |
| Boolean | System.Boolean |
| Byte | System.Byte |
| Integer | System.Int16 |
| Currency | System.Decimal |
| Long | System.Int32 |
| Single | System.Single |
| Double | System.Double |
| Date | System.DateTime |
| String | System.String |
| Object | System.Object |
| Variant | System.Object |
| Error | VBScript.net internal helper object |
	- Imports statement. Note that the behavior of the Imports statement in VBScript.net is different in VB.NET. See below:
		- Top namespace import. For example, if you use {{imports System}}, VBScript.net would know that the "System" is a namespace rather than a variable. However, when you use a class, you still have to specify the full namespace, for example, {{ System.Console.WriteLine("Hello World") }}. The reason is we do not do exhaustive searches in all imported namespaces to reduce the compiling time as a scripting language. If you do not want to type the entire namespace, you might use Alias Import.
		- Alias Import. You might define an alias for a namespace or a class, for example:
{{
imports c=System.Console
c.WriteLine("hello world")
}}

## VBScript.net compiler limitations
- Can assign to a constant. This should not happen.
 Imp and Eqv operators are not implemented. Not sure if anyone uses it. Submit an issue if you need it.
- LoadPicture function not implemented.
- There are many other more subtle incompatibilities (such as automatic widening of numeric type) that we have to document separately.

## ASP Classic Compiler Features:
- Implemented as an asp.net http module. See [Installation](Installation) page for setup and configuration.

## ASP Classic Compiler limitations:
- Global.asa not supported. If you want to add data to application or session state, use global.asax instead.
- TypeLibrary declaration using syntax like {{<!--METADATA TYPE="typelib" uuid="00000206-0000-0010-8000-00AA006D2EA4" -->}} is not implemented. Use include file adovbs.inc instead.
- ASP @ Directives ignored. @TRANSACTION declarative in ASP page is ignored. ASP page runs in a COM+ context. ASP Classic Compiler runs pages in asp.net worker process.
- COM objects invoked by asp page cannot access built-in asp objects, request, response and session, etc. This is because ASP Classic Compiler does not put these objects into COM+ context. As a result, COM objects like MSWC.BrowserType would not work with ASP Classic Compiler.
