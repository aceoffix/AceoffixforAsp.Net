<%@ Page Language="C#" AutoEventWireup="true" CodeFile="saveworddata.aspx.cs" Inherits="word_saveworddata" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Extracts data from the Word document</title>
    <style type="text/css">
        .TypeValue{ color:#0000FF; text-decoration:underline;}
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <p><img src="../image/word.png" alt="Word Document" style="float:left;" /> 
        The following data was just typed by you and was submitted by AceoffixCtrl when you clicked the save button. This demo shows how to get the data inputted by user from Word document. 
        Typically, you can save the following data to database.</p>
        <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
    </div>
    </form>
</body>
</html>
