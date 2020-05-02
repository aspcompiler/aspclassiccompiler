<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Azure Store Products</title>
    <link rel="shortcut icon" href="../../Content/images/favicon.ico" />
    <link href="../../Content/Site.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <div class="header-container">
        <div class="nav-login">
            <ul>
            <%if Request.IsAuthenticated then %>
                <li class="first">User:<span class="identity"><%=Html.Encode(context.User.Identity.Name)%></span></li>&nbsp;
                <li><%= Html.ActionLink("Logout", "LogOff", "Account") %></li>
            <% else %>
                    <li class="first"><%= Html.ActionLink("Register", "Register", "Account") %></li>
            <%end if %>
            </ul>
        </div>
        <div class="logo">Azure Store</div>
        <div class="clear"></div>
    </div>
    <div class="poster-container-no-image">
        <div class="poster-inner"> </div>
    </div>
    <div class="nav-main">
        <ul>
            <% if Html.IsCurrentAction("Index", "Home") then %>
            <li class="first active">
            <% else %>
            <li class="first">
            <% end if %>
                <%=Html.ActionLink("Products", "Index", "Home")%>
            </li>
            <% if Html.IsCurrentAction("Checkout", "Home") then %>
            <li class="active">
            <% else %>
            <li class="">
            <% end if %>
                <%=Html.ActionLink("Checkout", "Checkout", "Home")%>
            </li>
        </ul>
    </div>
    
    <div class="content-container">
        <div class="content-container-inner">
            <div class="content-main">
               <%if context.User.IsInRole("Home") then %>
                <h1><span class="product-category">Home</span> Products</h1>
               <% else %>
                <h1><span class="product-category">Enterprise</span> Products</h1>
               <% end if %>
                <label for="items">Select a product from the list:</label>
                <%  
                dim product
                dim myForm = Html.BeginForm("Add", "Home") %>
                    <select name="selectedItem" class="product-list" id="items" size="4">
                    <% for each product in ViewData.Model %>
                        <option value="<%=product%>"><%=product%></option>
                    <% next %>
                    </select>
                    <a href="javascript:document.forms[0].submit();">Add item to cart</a>
                <% myForm.dispose() %>
            </div>
            <div class="clear" />
        </div>
    </div>
    
    <div class="footer">
        <div class="nav-footer">
            <ul>
                <li class="first"><%=Html.ActionLink("Products", "Index", "Home")%></li>
                <li><%=Html.ActionLink("Checkout", "Checkout", "Home")%></li>
            </ul>
            <p class="copyright">Azure Store</p>
        </div>
    </div>
    
</body>
</html>
