Starting with change set 30470, we added an ASP.NET MVC sample that uses ASP Classic Compiler as a view engine. This sample was ported from the [NerdDinner sample](http://nerddinner.codeplex.com). We replaced all the webform views in the Dinner directory with ASP views. This should make the migration from Classic ASP to ASP.NET MVC with ASP view to a full ASP.NET MVC application very easy.

Most of the porting is quite trivial. I only need to mention a few things:
# Since ASP view does not support master pages, we use Html.RenderPartial instead.
# It is possible to mix the ASP view engine and the webform view engine. It is even possible to mix the view in a single page. This is demonstrated in the /Home/Index.asp view. We use ASP as the view but it uses the Login control as it's partial view.
# C# supports the using block so that the object is disposed at the end of the block. The Dinner/Delete view uses HtmlHelper.BeginForm() in this way. We have to use slightly different syntax in delete.asp to dispose the object so that it will generate the {</form>} tag.
# C# supports anonymous class and object initializer. VBScript does not have equivalent syntax. So our code in Dinner/RSVPStatus.asc is slightly longer. Note that we use the imports statement at the beginning to import the namespace so that we can instantiate the AjaxOptions object later. We want to come out with our own anonymous class.

## Built-in objects in ASP page when used as ASP.NET MVC view

The following are the built-in objects supported in the ASP view pages:
* context - HttpContext
* request - HttpRequest
* session - Session wrapper
* server - Server wrapper
* application - Application wrapper
* response - Response wrapper
* writer - TextWriter. Same as the TextWrite in webform view.
* viewcontext - ViewContext. Same as the webform view.
* viewdata - ViewDataDictionary. Same as the webform view.
* model - Model. Same as the webform view.
* tempdata - TempDataDictionary. Same as the webform view.
* ajax - AjaxHelper. Same as the webform view.
* html - HtmlHelper. Same as the webform view.
* url - UrlHelper. Same as the webform view.

## Include ASP view engine.

It is fairly simply to use the ASP view engine:
# Add reference to the dlls described in [Installation](Installation) in your ASP.NET MVC project.
# In global.asax.cs, add {"using Dlrsoft.Asp.Mvc"}.
# In Application_Start procedure, add {"ViewEngines.Engines.Add(new WebFormViewEngine());"}

The ASP view engine locate views in the views directory using the following patterms /Views/Controller/Action.asp and Views/Shared/Action.asp.

When locating partial views, the ASP view engine uses two additional patterns: /Views/Controller/Action.asc and Views/Shared/Action.asc.