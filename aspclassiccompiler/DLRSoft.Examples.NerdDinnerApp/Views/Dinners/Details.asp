<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title><%= Html.Encode(Model.Title) %></title>
    <link href="../../Content/Site.css" rel="stylesheet" type="text/css" />
  	<meta content="Nerd, Dinner, Geek, Luncheon, Dweeb, Breakfast, Technology, Bar, Beer, Wonk" name="keywords" /> 
	<meta name="description" content="Host and promote your own Nerd Dinner free!" /> 

    <script src="/Scripts/jquery-1.2.6.js" type="text/javascript"></script>    
</head>

<body>
    <div class="page">

        <% Html.RenderPartial("header") %>

        <div id="main">
            <div id="dinnerDiv">

                <h2><%= Html.Encode(Model.Title) %></h2>
                <p>
                    <strong>When:</strong> 
                    <%= Model.EventDate.ToShortDateString() %> 

                    <strong>@</strong>
                    <%= Model.EventDate.ToShortTimeString() %>
                </p>
                <p>
                    <strong>Where:</strong> 
                    <%= Html.Encode(Model.Address) %>,
                    <%= Html.Encode(Model.Country) %>
                </p>
                 <p>
                    <strong>Description:</strong> 
                    <%= Html.Encode(Model.Description) %>
                </p>       
                <p>
                    <strong>Organizer:</strong> 
                    <%= Html.Encode(Model.HostedBy) %>
                    (<%= Html.Encode(Model.ContactPhone) %>)
                </p>
            
                <% Html.RenderPartial("RSVPStatus"); %>
                <% Html.RenderPartial("EditAndDeleteLinks"); %>    
                
            </div>
    
            <div id="mapDiv">
                <% Html.RenderPartial("map"); %>    
            </div>
        </div>
    </div>
</body>
</html>


