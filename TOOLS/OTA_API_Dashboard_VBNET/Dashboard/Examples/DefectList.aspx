<%@ Page Title="Defects List" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="DefectList.aspx.vb" Inherits="TDConnection.DefectList"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server"> <br /><br />
   
    &nbsp;<asp:Button ID="RunButton" runat="server" Text="Run" Width="150px" />&nbsp;&nbsp;&nbsp;
    &nbsp;<br />
    &nbsp;<asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False">
            <Columns>
                <asp:BoundField DataField="Defect ID" HeaderText="Defect ID" />
                <asp:BoundField DataField="Status" HeaderText="Status" />
                <asp:BoundField DataField="Detected By" HeaderText="Detected By" />
                <asp:BoundField DataField="Summary" HeaderText="Summary" />
            </Columns>
        </asp:GridView>
        
&nbsp;Log:<br />
    &nbsp;<asp:TextBox ID="LogText" runat="server" Height="113px" Width="1137px" TextMode="MultiLine"></asp:TextBox> 
        <br />
    
</asp:Content>

