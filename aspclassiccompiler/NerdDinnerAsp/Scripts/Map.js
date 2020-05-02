/// <reference path="jquery-1.2.6.js" />

var map = null;
var points = [];
var shapes = [];
var center = null;

function LoadMap(latitude, longitude, onMapLoaded) {
    map = new VEMap('theMap');
    options = new VEMapOptions();
    options.EnableBirdseye = false;

    // Makes the control bar less obtrusize.
    map.SetDashboardSize(VEDashboardSize.Small);
    
    if (onMapLoaded != null)
        map.onLoadMap = onMapLoaded;

    if (latitude != null && longitude != null) {
        center = new VELatLong(latitude, longitude);
    }

    map.LoadMap(center, null, null, null, null, null, null, options);
}

function LoadPin(LL, name, description) {
    var shape = new VEShape(VEShapeType.Pushpin, LL);

    //Make a nice Pushpin shape with a title and description
    shape.SetTitle("<span class=\"pinTitle\"> " + escape(name) + "</span>");
    if (description !== undefined) {
        shape.SetDescription("<p class=\"pinDetails\">" + 
        escape(description) + "</p>");
    }
    map.AddShape(shape);
    points.push(LL);
    shapes.push(shape);
}

function FindAddressOnMap(where) {
    var numberOfResults = 20;
    var setBestMapView = true;
    var showResults = true;

    map.Find("", where, null, null, null,
           numberOfResults, showResults, true, true,
           setBestMapView, callbackForLocation);
}

function callbackForLocation(layer, resultsArray, places,
            hasMore, VEErrorMessage) {
            
    clearMap();
    
    if (places == null) 
        return;

    //Make a pushpin for each place we find
    $.each(places, function(i, item) {
        var description = "";
        if (item.Description !== undefined) {
            description = item.Description;
        }
        var LL = new VELatLong(item.LatLong.Latitude,
                        item.LatLong.Longitude);
                        
        LoadPin(LL, item.Name, description);
    });

    //Make sure all pushpins are visible
    if (points.length > 1) {
        map.SetMapView(points);
    }

    //If we've found exactly one place, that's our address.
    if (points.length === 1) {
        $("#Latitude").val(points[0].Latitude);
        $("#Longitude").val(points[0].Longitude);
    }
}

function clearMap() {
    map.Clear();
    points = [];
    shapes = [];
}

function FindDinnersGivenLocation(where) {
    map.Find("", where, null, null, null, null, null, false,
       null, null, callbackUpdateMapDinners);
}

function callbackUpdateMapDinners(layer, resultsArray, places, hasMore, VEErrorMessage) {
    $("#dinnerList").empty();
    clearMap();
    var center = map.GetCenter();

    $.post("/Search/SearchByLocation", { latitude: center.Latitude, 
                                         longitude: center.Longitude
    }, function(dinners) {
        $.each(dinners, function(i, dinner) {

            var LL = new VELatLong(dinner.Latitude, dinner.Longitude, 0, null);

            var RsvpMessage = "";

            if (dinner.RSVPCount == 1)
                RsvpMessage = "" + dinner.RSVPCount + " RSVP";
            else
                RsvpMessage = "" + dinner.RSVPCount + " RSVPs";

            // Add Pin to Map
            LoadPin(LL, '<a href="/Dinners/Details/' + dinner.DinnerID + '">'
                        + dinner.Title + '</a>',
                        "<p>" + dinner.Description + "</p>" + RsvpMessage);

            //Add a dinner to the <ul> dinnerList on the right
            $('#dinnerList').append($('<li/>')
                            .attr("class", "dinnerItem")
                            .append($('<a/>').attr("href", 
                                      "/Dinners/Details/" + dinner.DinnerID)
                            .html(dinner.Title)).append(" ("+RsvpMessage+")"));
        });

        // Adjust zoom to display all the pins we just added.
	    if (points.length > 1) {
        	    map.SetMapView(points);
	    }

        // Display the event's pin-bubble on hover.
        $(".dinnerItem").each(function(i, dinner) {
            $(dinner).hover(
                function() { map.ShowInfoBox(shapes[i]); },
                function() { map.HideInfoBox(shapes[i]); }
            );
        });
    }, "json");
}
