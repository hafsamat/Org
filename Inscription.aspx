<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Inscription.aspx.cs" Inherits="Org.Inscription" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <link href="Styles/Inscription.css" rel="stylesheet" />
    <link rel="preconnect" href="https://fonts.googleapis.com"/>
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link
        href="https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,100;1,200;1,300;1,400;1,500;1,600;1,700;1,800;1,900&display=swap"
        rel="stylesheet">
    <title>Inscription</title>
</head>
<body>
    <form id="form1" runat="server">
        <header>
        <img src="Images/logo.png" class="mg">
        <div class="bar"></div>
    </header>
    <div class="container">
        <div class="cont1">
                <table>
                    <tr>
                        <td>
                            <span>Créer un Compte Au Menara Holding.</span>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <label>Matricule</label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                            
                        </td>
                        <td>
                            <asp:Button ID="Button2" runat="server" Text="Rechercher" OnClick="Button2_Click" />
                        </td>    
                    </tr>
                    <tr>
                        <td>
                            <label>Nom</label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:TextBox ID="TextBox2" runat="server" ReadOnly></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <label>Prénom</label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:TextBox ID="TextBox3" runat="server" ReadOnly></asp:TextBox>
                        </td>
                    </tr>
                </table>
        </div>
        <div class="cont2">
            <div class="line"></div>
        </div>
        <div class="cont3">
                <table>
                    <tr>
                        <td>
                            <label>Société</label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:TextBox ID="TextBox4" runat="server" ReadOnly></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <label>Type de Compte</label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DropDownList ID="DropDownList1" runat="server">
                                <asp:ListItem>Admin</asp:ListItem>
                                <asp:ListItem>Visiteur</asp:ListItem>
                                <asp:ListItem>RH</asp:ListItem>
                                <asp:ListItem>DG</asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <label>Mot de Passe *</label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:TextBox ID="TextBox5" TextMode="Password" runat="server"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Button ID="Button1" runat="server" Text="Enregister" OnClick="Button1_Click1" />
                        </td>
                    </tr>
                </table>
        </div>
    </div>
        
    </form>
</body>
</html>
