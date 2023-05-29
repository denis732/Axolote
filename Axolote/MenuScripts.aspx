<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MenuScripts.aspx.cs" Inherits="Axolote.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Générer un fichier JSON à partir d'un fichier Excel</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Button id="btnProduireJSON"
            Text="Submit"
            OnClick="btnProduireJSON_Click"
            runat="server" />
        </div>
      <asp:label id="Message" runat="server"/>
    </form>
</body>
</html>
