<%@ Page Language="C#" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="convertPDFs_index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
        checkit = true;
        function selectall() {
            if (checkit) {
                var obj = document.all.check;
                for (var i = 0; i < obj.length; i++) {
                    obj[i].checked = true;
                    checkit = false;
                }
            } else {
                var obj = document.all.check;
                for (var i = 0; i < obj.length; i++) {
                    obj[i].checked = false;
                    checkit = true;
                }

            }

        }

        var count = 0;
        var strLength = 0;
        var str = new Array();
        function getID() {
            var ids = "";
            //Get the value from the checkbox in loop
            var check = document.getElementsByName("check");
            for (var i = 0; i < check.length; i++) {
                if (check[i].checked) {
                    ids += check[i].value + ",";
                }
            }

            if (ids == "" || ids == null || ids == "on,") {
                alert("Select at least one document for converting!");
                return;

            }
            // Remove the last ","
            ids = ids.substring(0, ids.length - 1);
            var i = ids.indexOf("on");
            if (i == 0 || i > 0) {
                ids = ids.replace("on,", "");
            }
            str = ids.split(",");
            if (count > 0) {
                alert("Pleaes refresh the current page and reselect documents for converting! ");
                return;
            }
            //Auto-refresh for the first time
            document.getElementById("iframe1").src = "Convert.aspx?id=" + str[count];
            strLength = str.length;
        }

        function convert() {
		document.getElementById("btnConvert").disabled=true;
            count = count + 1;
            if (count < strLength) {
                document.getElementById("iframe1").src = "Convert.aspx?id=" + str[count];
            } else {
                alert("Converting complete！");
                document.getElementById("aDiv").style.display = "";
		document.getElementById("btnConvert").disabled=false;
            }
        }
		
    </script>

</head>

 <body>
         <div style="margin:100px" align="center">
          <form id="form1"  method="post" >
          <h2>Documents Index</h2>
          <table  id="table1" style="border-spacing:20px;border:1px;background-color: darkgray; width: 600px;" cellpadding="3" cellspacing="1">       
               <tr>
               <td><input name="check" type="checkbox" onclick="selectall()"  /></td>
               <td>Index Number</td>
               <td>Documents</td>
               <td>Edit</td>
             </tr>
              <tr>
               <td><input  name="check"  type="checkbox" value="1"/></td>
               <td>01</td>
               <td>Aceoffix Introduction.docx</td>
               <td><a href="Edit.aspx?id=1">Edit</a></td>
             </tr>
              <tr>
               <td><input  name="check"  type="checkbox" value="2"/></td>
               <td>02</td>
               <td>Using Aceoffix trial version.docx</td>
              <td><a href="Edit.aspx?id=2">Edit</a></td>
             </tr>
              <tr>
               <td><input name="check"  type="checkbox" value="3"/></td> 
               <td>03</td>
               <td>How to Integrate Aceoffix.docx</td>
                <td><a href="Edit.aspx?id=3">Edit</a></td>
              </tr>
               <tr>
               <td><input name="check"  type="checkbox" value="4"/></td>
               <td>04</td>
               <td>How to Support all Browsers.docx</td>
               <td><a href="Edit.aspx?id=4">Edit</a></td>
              </tr>
           </table></br></br>
		   <input type="button" value="Convert Word documents to PDFs in batch" onclick="getID()" id="btnConvert" />
          </form>
      </div> 
	 
	 <div id="aDiv" style="display: none; color: Red; font-size: 15px;text-align:center;"  >
            <span>Converting complete. Please view the .pdf files from <%=url %></span>
     </div>
	 
	  <div style="width: 1px; height: 1px; overflow: hidden;">
        <iframe id="iframe1" name="iframe1" src=""></iframe>
    </div>
</body>
</html>
