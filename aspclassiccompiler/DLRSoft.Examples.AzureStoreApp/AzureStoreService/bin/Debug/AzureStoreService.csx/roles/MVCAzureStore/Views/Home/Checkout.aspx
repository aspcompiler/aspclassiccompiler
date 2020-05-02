<%@ Page Title="" Language="C#" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage<IEnumerable<string>>" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
Azure Store Check Out
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <h1>Your Order</h1>
    <label for="cart">You have selected the following products:</label>
    <% using (Html.BeginForm("Remove", "Home")) { %>
        <select name="selectedItem" class="product-list" id="items" size="4">
        <% foreach (string product in ViewData.Model)
           { %>
            <option value="<%=product%>"><%=product%></option>
        <% } %>
        </select>
        <a href="javascript:document.forms[0].submit();">Remove product from cart</a>
    <% } %>
</asp:Content>
