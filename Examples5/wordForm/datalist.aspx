<%@ Page Language="C#" AutoEventWireup="true" CodeFile="datalist.aspx.cs" Inherits="wordForm_datalist" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" method="post">
		</form>

		<table width='98%' class="zz-talbe"
			style="margin-top: 20px; width: 460px;">
			<thead>
				<tr>
					<th width='20%' height='26' valign="middle">
					Database field
					</th>
					<th width='80%' height='23' valign="middle">
						Database values
					</th>
				</tr>
			</thead>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE">
					Theme
				</td>
				<td width='80%' height='23' valign="middle">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docSubject%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE">
					Name
				</td>
				<td width='80%' height='23' valign="middle">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docName%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE">
					Department
				</td>
				<td width='80%' height='23' valign="middle">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docDept%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE">
					Reason
				</td>
				<td width='80%' height='23' valign="middle">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docCause%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE">
					Days
				</td>
				<td width='80%' height='23' valign="middle">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docDays%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE">
					Date
				</td>
				<td width='80%' height='23' valign="middle">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docDate%></td>
			</tr>


		</table>
</body>
</html>
