<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="TDConnection._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
        
            
            <div style="height: 100px; width: 900px">


        <div style="height: 267px; width: 906px">
             <div style="height: 100px; width: 900px">
                <asp:Image ID="Image1" runat="server" Height="99px" ImageUrl="~/SaaS Banner.png" Width="900px" />

        </div>
        
            
           
        <asp:table ID="Table1" runat="server" Height="272px" Width="900px" BackColor="#99CCFF">

            <asp:TableRow Width="400">
                <asp:TableCell Width="200">
                    <asp:Label ID="Label1" runat="server" Text="Server URL:"  Width="200" ></asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="200">
                    <asp:TextBox ID="ServerURLTextbox" runat="server"  Width="500"  Text="https://almastqcdemo15.saas.microfocus.com/qcbin/"></asp:TextBox>
                </asp:TableCell>
                 <asp:TableCell Width="200">
                    <asp:Button ID="InitializeButton" runat="server" OnClick="InitializeButton_Click"  Width="150" Text="Initialize Server" />
                </asp:TableCell>
            </asp:TableRow>
            
            <asp:TableRow Width="400">
                <asp:TableCell Width="200">
                    <asp:Label ID="Label4" runat="server" Text="User Name:"  Width="200" ></asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="200">
                    <asp:TextBox ID="UserNameTextbox" runat="server"  Width="500"></asp:TextBox>
                </asp:TableCell>

            </asp:TableRow>

            <asp:TableRow>
                <asp:TableCell Width="200">
                    <asp:Label ID="Label2" runat="server" Text="Password:" Width="200"></asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="200">
                    <asp:TextBox ID="PasswordTextbox" runat="server"  Width="500" TextMode="Password"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell Width="200">
                    <asp:Button ID="AuthenticateButton" runat="server" OnClick="AuthenticateButton_Click"  Width="150" Text="Authenticate" />
                </asp:TableCell>

            </asp:TableRow>
          
            <asp:TableRow>
                
                    <asp:TableCell Width="200">
                        <asp:Label ID="Label3" runat="server" Text="Domain:" Width="200"></asp:Label>
                    </asp:TableCell>

                    <asp:TableCell Width="200">
                        <asp:ListBox runat="server" Width="500" ID="DomainList" AutoPostBack="False"></asp:ListBox>
                    </asp:TableCell>

                    <asp:TableCell Width="200">
                        <asp:Button ID="SelectDomainButton" runat="server" OnClick="SelectDomainButton_Click" Width="150" Text="Select Domain" />
                    </asp:TableCell>
            </asp:TableRow>

            
            <asp:TableRow>
                
                    <asp:TableCell Width="200">
                        <asp:Label ID="Label5" runat="server" Text="Project:" Width="200"></asp:Label>
                    </asp:TableCell>

                    <asp:TableCell Width="200">
                        <asp:ListBox runat="server" Width="500" ID="ProjectList" AutoPostBack="False"></asp:ListBox>
                    </asp:TableCell>

                    <asp:TableCell Width="200">
                        <asp:Button ID="SelectProjectButton" runat="server" OnClick="SelectProjectButton_Click" Width="150" Text="Select Project" />
                    </asp:TableCell>
            </asp:TableRow>


               <asp:TableRow>          
                    <asp:TableCell Width="200">
                    </asp:TableCell> 
                     <asp:TableCell Width="200">
                    </asp:TableCell> 
                    <asp:TableCell Width="200">
                        <asp:Button ID="ConnectButton" runat="server" OnClick="ConnectButton_Click" Width="150" Text="Login" />
                    </asp:TableCell>
            </asp:TableRow>
  
            <asp:TableRow>
                    <asp:TableCell Width="200">
                        <asp:Label ID="Status" runat="server" Text=""></asp:Label>
                    </asp:TableCell>
            </asp:TableRow>

        </asp:table>
            </div>
 
        </div>


        </div>
    </form>
</body>
</html>
