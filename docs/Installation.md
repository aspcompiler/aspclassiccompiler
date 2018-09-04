# Installation

## Build for .Net framework 2.0, 3.0, 3.5
* Components that need to be copied to the bin dir of your web application
|| Component Name || Description ||
| Dlrsoft.asp.dll | ASP Classic Compiler runtime |
| Dlrsoft.VBParser.dll | VBScript Parser |
| Dlrsoft.VBScript.dll | VBScript.net compiler |
| Microsoft.Dynamic.dll | Microsoft Dynamic Language Runtime |
| Microsoft.Scripting.Core.dll | Microsoft Dynamic Language Runtime |
| Microsoft.Scripting.dll | Microsoft Dynamic Language Runtime |
| Microsoft.Scripting.ExtensionAttribute.dll | Microsoft Dynamic Language Runtime |
* Configure the ASP Classic Compiler Http Handler. Add the following line to the httpHandlers section of web.config file:
{{
    <add verb="**" path="**.asp" validate="false" type="Dlrsoft.Asp.AspHandler, Dlrsoft.Asp"/>
}}
Note that if you want to use a different file extension for ASP Classic Compiler, just modify the path attribute.
* (optional) Configure ISAPI extensions. This step is required only if you want to use ASP Classic Compiler to handle asp pages in IIS. It is not required in Cassini.
	* IIS 5/6: In IIS Manager, go to Web Site Properties, on the Home Directory tab, click Configuration.... In Mappings tab, you will notice that asp is currently mapped to C:\WINDOWS\system32\inetsrv\asp.dll. Change it to executable that .aspx extension is using, e.g., c:\windows\microsoft.net\framework\v2.0.50727\aspnet_isapi.dll.
Note that if you use a different file extension for ASP Classic Compiler, just create a new ISAPI extension mapping.