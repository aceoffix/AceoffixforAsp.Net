<%@ Page Language="C#" AutoEventWireup="true" CodeFile="viewhtml.aspx.cs" Inherits="word_viewhtml" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="width:810px; margin:0 auto;">
    <p>The following HTML page is submitted by Aceoffix. Back to the <a href="../default.aspx">Example Center</a>.</p>
    <iframe width="800" height="600" src="<%=m_htmlURL%>"></iframe>
    </div>
    </form>
</body>
</html>
