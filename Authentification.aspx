<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Authentification.aspx.cs" Inherits="Org.Authentification" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head runat="server">
   <meta charset="UTF-8"/>
  <meta http-equiv="X-UA-Compatible" content="IE=edge"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <link href="Styles/Authentification.css" rel="stylesheet"/>
  <link rel="preconnect" href="https://fonts.googleapis.com"/>
  <link rel="preconnect" href="https://fonts.gstatic.com"/>
  <link href="https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,100;1,200;1,300;1,400;1,500;1,600;1,700;1,800;1,900&display=swap" rel="stylesheet"/>
  <title>Authentification</title>
</head>
<body>
    <form id="form1" runat="server">
    <header>
    <img src="Logo/holdinglogo.png" class="mg"/>
    <div class="bar"></div>
    </header>
  <div class="container">
    <div class="cont1">
      <form>
        <table>
          <tr>
            <td><span>Authentification</span></td>
          </tr>
          <tr>
            <td>
              <label>Bienvenue! Connectez Vous avec votre matricule.</label>
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
          </tr>
          <tr>
            <td>
              <label>Mot de Passe</label>
            </td>
          </tr>
          <tr>
            <td>
              <asp:TextBox ID="TextBox2" TextMode="Password" runat="server"></asp:TextBox>
            </td>
          </tr>
          <tr>
            <td>
              <asp:Button ID="Button1" runat="server" Text="Se Connecter" OnClick="Button1_Click" />
            </td>
          </tr>
        </table>
      </form>
    </div>
    <div class="cont2">
    <div class="line"></div>
    </div>
    <div class="cont3">
      <img src="Images/mol_pc.png"/>
    </div>
  </div>
    </form>
</body>
</html>
