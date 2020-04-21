<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="Default.aspx.vb" Inherits="TDConnection._Default1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <div align="left" style="font-size:larger; height: 250px; width: 979px;">
        <asp:RadioButtonList ID="RadioButtonList1" runat="server" Height="82px" Width="491px">
            <asp:ListItem Value="DefectList.aspx">Defect List</asp:ListItem>
            <asp:ListItem Value="TestList.aspx">Test List</asp:ListItem>
        </asp:RadioButtonList>
        <br />
    
        &nbsp;<asp:Button ID="Button1" runat="server" Text="Select" Height="32px" Width="80px" />
    </div>
</asp:Content>
