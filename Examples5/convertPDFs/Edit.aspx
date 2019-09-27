<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Edit.aspx.cs" Inherits="convertPDFs_Edit" %>

<%@ Register assembly="Aceoffix, Version=5.0.0.1, Culture=neutral, PublicKeyToken=e6a26169e940f541" namespace="Aceoffix" tagprefix="ace" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
            function Save() {
                document.getElementById("AceoffixCtrl1").WebSave();
            }
    </script>
</head>
<body>
    <a href="index.aspx">Return to the index page</a>
    <form id="form1" runat="server">
    <div  style="width:auto; height:900px;">
      <ace:AceoffixCtrl ID="AceoffixCtrl1" runat="server">
      </ace:AceoffixCtrl>
    </div>
    </form>
</body>
</html>
