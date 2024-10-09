<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="Home.aspx.vb" Inherits="TDConnection.Home" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server"> <br /><br />
   
    &nbsp;<asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False">
            <Columns>
                <asp:BoundField DataField="Server Info" HeaderText="Server Info" />
                <asp:BoundField DataField="Value" HeaderText="Value" />
                
            </Columns>
        </asp:GridView><br />
        
  &nbsp;<asp:Button ID="LogoutButton" runat="server" Text="Logout" Width="150px" />&nbsp;&nbsp;&nbsp;
    &nbsp;<br />
  
    
</asp:Content>