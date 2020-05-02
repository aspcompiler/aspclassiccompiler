<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title>Find a Dinner</title>
    <link href="../../Content/Site.css" rel="stylesheet" type="text/css" />
  	<meta content="Nerd, Dinner, Geek, Luncheon, Dweeb, Breakfast, Technology, Bar, Beer, Wonk" name="keywords" /> 
	<meta name="description" content="Host and promote your own Nerd Dinner free!" /> 

    <script src="/Scripts/jquery-1.2.6.js" type="text/javascript"></script>    
</head>

<body>
    <div class="page">
    
        <% Html.RenderPartial("header") %>
        
        <div id="main">
<script src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6.2" type="text/javascript"></script>
<script src="/Scripts/Map.js" type="text/javascript"></script>

<h2>Find a Dinner</h2>

<div id="mapDivLeft">

    <div id="searchBox">
        Enter your location: <%= Html.TextBox("Location") %> or <%= Html.ActionLink("View All Upcoming Dinners", "Index", "Dinners") %>.
        <input id="search" type="submit" value="Search" />
    </div>

    <div id="theMap">
    </div>

</div>

<div id="mapDivRight">
    <div id="dinnerList"></div>
</div>

<script type="text/javascript">

  $(document).ready(function() {
    LoadMap();
  });

  $("#search").click(function(evt) {
    var where = jQuery.trim($("#Location").val());
    if (where.length < 1)
      return;

    FindDinnersGivenLocation(where);
  });

</script>
        </div>
    </div>
</body>
</html>


