<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title>Host a Dinner</title>
    <link href="../../Content/Site.css" rel="stylesheet" type="text/css" />
  	<meta content="Nerd, Dinner, Geek, Luncheon, Dweeb, Breakfast, Technology, Bar, Beer, Wonk" name="keywords" /> 
	<meta name="description" content="Host and promote your own Nerd Dinner free!" /> 

    <script src="/Scripts/jquery-1.2.6.js" type="text/javascript"></script>    
</head>

<body>
    <div class="page">

        <% Html.RenderPartial("header") %>

        <div id="main">
            <% Html.RenderPartial("DinnerForm") %>
        </div>
    </div>
</body>
</html>


