<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="OrgPage.aspx.cs" Inherits="Org.OrgPage" %>

<!DOCTYPE html/>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link href='https://fonts.googleapis.com/css?family=Poppins' rel='stylesheet' />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css" />
    <link href="Styles/StyleOrg.css" rel="stylesheet" />
    <script src="Scripts/orgchart.js"></script>
    
    <title>Organigramme</title>
</head>
<body>
    <script src="Scripts/data.js"></script>
    <script>
        Name.sc = "MH";
        Name.st = "Menara Holding";
        var e = document.getElementById("Text1");
        Name.dr = e.options[e.selectedIndex].text;

        function MH() {
            document.getElementById("nav").style.backgroundColor = "#2c2727";
            document.getElementById("bar").style.backgroundColor = "#9a806a";
            document.getElementById("img").src = "logo/holdinglogo.png";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#ffffff";
            }
        }

        function G() {
            document.getElementById("nav").style.backgroundColor = "#2c2727";
            document.getElementById("bar").style.backgroundColor = "#9a806a";
            document.getElementById("img").src = "logo/holdinglogo.png";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#ffffff";
            }
            Name.sc = "G";
            Name.st = "Global";
        }

        function MP() {
            document.getElementById("nav").style.backgroundColor = "#2d88c0";
            document.getElementById("bar").style.backgroundColor = "#67696B";
            document.getElementById("img").src = "logo/préfalogo.png";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#000000";
            }
            Name.sc = "MP";
            Name.st = "Menara Prefa";
        }

        function CT() {
            document.getElementById("nav").style.backgroundColor = "#C5C9CA";
            document.getElementById("bar").style.backgroundColor = "#14978F";
            document.getElementById("img").src = "logo/ctmlogo.svg";
            Name.sc = "CT";
            Name.st = "Carriere Transport Menara";
        }

        function TC() {
            document.getElementById("nav").style.backgroundColor = "#D76A3E";
            document.getElementById("bar").style.backgroundColor = "#BEB5AF";
            document.getElementById("img").src = "logo/tcgmlogo.svg";
            Name.sc = "TC";
            Name.st = "TCGM";
        }

        function ML() {
            document.getElementById("nav").style.backgroundColor = "#F39132";
            document.getElementById("bar").style.backgroundColor = "#FFFFFF";
            document.getElementById("img").src = "logo/logistiquelogo.png";
            Name.sc = "ML";
            Name.st = "Menara Logistique";
        }

        function MTL() {
            document.getElementById("nav").style.backgroundColor = "#ffffff";
            document.getElementById("bar").style.backgroundColor = "#F4812A";
            document.getElementById("img").src = "logo/mtllogo.png";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#000000";
            }
            Name.sc = "MTL";
            Name.st = "Menara Transport et Logistique";
        }

        function IT() {
            document.getElementById("nav").style.backgroundColor = "#88D2F7";
            document.getElementById("bar").style.backgroundColor = "#E1C3AC";
            document.getElementById("img").src = "logo/tensiftlogo.svg";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#000000";
            }
            Name.sc = "IT";
            Name.st = "Immobiliere Tensift";
        }

        function FZ() {
            document.getElementById("nav").style.backgroundColor = "#2D88c0";
            document.getElementById("bar").style.backgroundColor = "#263d84";
            document.getElementById("img").src = "logo/zahidlogo.svg";
            Name.sc = "FZ";
            Name.st = "Fondation Zahid";
        }

        function AM() {
            document.getElementById("nav").style.backgroundColor = "#ffffff";
            document.getElementById("bar").style.backgroundColor = "#2d88c0";
            document.getElementById("img").src = "logo/menaralogo.svg";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#000000";
            }
            Name.sc = "AM";
            Name.st = "Academie Menara";
        }

        function MG() {
            document.getElementById("nav").style.backgroundColor = "#cccccc";
            document.getElementById("bar").style.backgroundColor = "#000000";
            document.getElementById("bar").style.color = "#000000";
            document.getElementById("img").src = "logo/mgplogo.svg";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#000000";
            }
            Name.sc = "MG";
            Name.st = "Menara Grand Prix";
        }

        function A() {
            document.getElementById("nav").style.backgroundColor = "#FFFFFF";
            document.getElementById("bar").style.backgroundColor = "#f39132";
            for (let i = 0; i < document.getElementsByClassName("dropbtn").length; i++) {
                document.getElementsByClassName("dropbtn")[i].style.color = "#000000";
            }
            document.getElementById("img").src = "logo/akdlogo.png";
            Name.sc = "A";
            Name.st = "Aakar";
        }
    </script>
    <form id="form1" runat="server">
        <nav>
            <div class="container" id="nav" runat="server">
                <div class="cont1">
                    <img id="img" runat="server" src="logo/holdinglogo.png">
                </div>
                <div class="cont2">
                    <div class="mincont2" id="btn_aj" runat="server">
                        <a id="ajout" runat="server" href="Inscription.aspx">
                            <label class="dropbtn">Ajouter</label></a>
                    </div>
                    <div class="mincont1">
                        <div class="dropdown">
                            <asp:ScriptManager ID="ScriptManager2" runat="server"></asp:ScriptManager>
                            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                <ContentTemplate>
                                    <button class="dropbtn" type="button">
                                        Exporter
                                        <i class="fa fa-caret-down"></i>
                                    </button>
                                </ContentTemplate>
                            </asp:UpdatePanel>
                            <div class="dropdown-content">
                                <asp:UpdatePanel ID="UpdatePanel3" runat="server">
                                    <ContentTemplate>
                                        <asp:Button ID="elem" type="button" runat="server" Text="Exporter PDF" />
                                    </ContentTemplate>
                                </asp:UpdatePanel>
                                <asp:Button ID="Btn_excel" runat="server" Text="Exporter Excel" OnClick="Btn_excel_Click" />
                            </div>
                        </div>
                    </div>
                    <div class="mincont6" id="btn_soc" runat="server">
                        <div class="dropdown">
                            <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                <ContentTemplate>
                                    <button type="button" class="dropbtn" runat="server">
                                        Societé
                                        <i class="fa fa-caret-down"></i>
                                    </button>
                                </ContentTemplate>
                            </asp:UpdatePanel>
                            <div class="dropdown-content">
                                <asp:Button ID="G" CssClass="bt" runat="server" Text="Global" OnClick="G_Click" />
                                <asp:Button ID="MH" CssClass="bt" runat="server" Text="Ménara Holding" OnClick="MH_Click" />
                                <asp:Button ID="MP" CssClass="bt" runat="server" Text="Ménara Préfa" OnClick="MP_Click" />
                                <asp:Button ID="CT" CssClass="bt" runat="server" Text="CTM" OnClick="CT_Click" />
                                <asp:Button ID="TC" CssClass="bt" runat="server" Text="TCGM" OnClick="TC_Click" />
                                <asp:Button ID="ML" CssClass="bt" runat="server" Text="Ménara Logistique" OnClick="ML_Click" />
                                <asp:Button ID="MT" CssClass="bt" runat="server" Text="MTL" OnClick="MT_Click" />
                                <asp:Button ID="IT" CssClass="bt" runat="server" Text="Immobilière Tensift" OnClick="IT_Click" />
                                <asp:Button ID="FZ" CssClass="bt" runat="server" Text="Fondation Zahid" OnClick="FZ_Click" />
                                <asp:Button ID="AM" CssClass="bt" runat="server" Text="Académie Ménara" OnClick="AM_Click" />
                                <asp:Button ID="MG" CssClass="bt" runat="server" Text="MGP" OnClick="MG_Click" />
                                <asp:Button ID="A" CssClass="bt" runat="server" Text="Aakar" OnClick="A_Click" />
                            </div>
                        </div>
                    </div>
                    <div class="mincont3" id="btn_grad" runat="server">
                        <table>
                            <tr>
                                <td>
                                    <asp:DropDownList ID="Text3" runat="server"></asp:DropDownList>
                                </td>
                                <td>
                                    <asp:Button ID="Button3" OnClick="FillG" CssClass="dropbtn" runat="server" Text="Generer" />
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div class="mincont4" id="btn_reg" runat="server">
                        <table>
                            <tr>
                                <td>
                                    <asp:DropDownList ID="Text2" runat="server"></asp:DropDownList>
                                </td>
                                <td>
                                    <asp:Button ID="Button2" OnClick="FillREG" CssClass="dropbtn" runat="server" Text="Generer" />
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div class="mincont5" id="btn_dir" runat="server">
                        <table>
                            <tr>
                                <td>
                                    <asp:DropDownList ID="Text1" runat="server"></asp:DropDownList>
                                </td>
                                <td>
                                    <asp:Button ID="Button1" OnClick="Fill" CssClass="dropbtn" runat="server" Text="Generer" />
                                </td>
                            </tr>
                        </table>
                    </div>
                </div>
            </div>
            <div id="bar" class="bar" runat="server"></div>
        </nav>
        <div id="tree"></div>
    </form>
</body>
</html>
