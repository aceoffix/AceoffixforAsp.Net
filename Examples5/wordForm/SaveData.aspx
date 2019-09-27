<%@ Page Language="C#" AutoEventWireup="true" CodeFile="SaveData.aspx.cs" Inherits="wordForm_SaveData" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <div class="errMainArea" id="error163"><div class="errTopArea" style="TEXT-ALIGN:left"></div>
			<div class="errTxtArea" style="HEIGHT:150px; TEXT-ALIGN:left">
				<b class="txt_title">
					<div style="color:#FF0000;">Please refill the "Days":</div>
					<ul>
					<%=ErrorMsg%>
					</ul>
					
				</b>
				
			</div>
			<div class="errBtmArea"><input type="button" class="btnFn" value="Close" onClick="window.opener=null;window.close();"></div>
		</div>
</body>
</html>
