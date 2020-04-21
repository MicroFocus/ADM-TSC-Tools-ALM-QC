<%@ Page Title="Tests List" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="TestList.aspx.vb" Inherits="TDConnection.TestList"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server"> <br /><br />
   
    &nbsp;<asp:Button ID="RunButton" runat="server" Text="Run" Width="150px" />&nbsp;&nbsp;&nbsp;
    &nbsp;<br />
    &nbsp;<asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False">
            <Columns>
                <asp:BoundField DataField="Test ID" HeaderText="Test ID" />
                <asp:BoundField DataField="Execution Status" HeaderText="Execution Status" />
                <asp:BoundField DataField="Created By" HeaderText="Created By" />
                <asp:BoundField DataField="Test Name" HeaderText="Test Name" />
                <asp:BoundField DataField="Test Type" HeaderText="Test Type" />
            </Columns>
        </asp:GridView>
        
&nbsp;Log:<br />
    &nbsp;<asp:TextBox ID="LogText" runat="server" Height="113px" Width="1137px" TextMode="MultiLine"></asp:TextBox> 
        <br />
    
</asp:Content>

