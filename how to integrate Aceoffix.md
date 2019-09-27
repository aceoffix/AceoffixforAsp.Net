# How to integrate Aceoffix into your web project

1. Run the setup-server.exe to install the Aceoffix Server Components on the web server.

![截图1](https://github.com/cupandcup/Edit-Word-Online-By-AceOffix/blob/master/Examples5/image/integrate_1.png?raw=true)
             
2. Add the references of Aceoffix to your web project or web site.
  
![截图2](https://github.com/cupandcup/Edit-Word-Online-By-AceOffix/blob/master/Examples5/image/integrate_2.png?raw=true)

![截图3](https://github.com/cupandcup/Edit-Word-Online-By-AceOffix/blob/master/Examples5/image/integrate_3.png?raw=true)

3. Copy the Install\aceoffix-runtime folder to the root folder of your website.
     
![截图4](https://github.com/cupandcup/Edit-Word-Online-By-AceOffix/blob/master/Examples5/image/integrate_4.png?raw=true)

4. Add the following code to the parent page of which page you want to edit Office document.
Write the following code in the `<head>` tag. 

        <script type="text/javascript" src="aceoffix-runtime/js/jquery.min.js"></script>
        
        <script type="text/javascript" src="aceoffix-runtime/js/aceoffix.js" id="ace_js_main"></script>

    Note: The paths of jquery.min.js  and aceoffix.js are relative to the root of your website.
Write the following link to pop-up Aceoffix to edit Office document. We assume that the page which contains Aceoffix control is word/editword.aspx.

        <a href="javascript:AceBrowser.openWindowModeless('word/editword.aspx', 'width=1200px;height=800px;');" > Edit Word document online</a>
        
    Add the following reference to word/editword.aspx.
    
        <%@ Register assembly="Aceoffix, Version=5.0.0.1, Culture=neutral, PublicKeyToken=e6a26169e940f541" namespace="Aceoffix" tagprefix="ace" %>
        
    Then write the following code where you want to show Aceoffix in editword.aspx.
    
        <div style="width:auto; height:600px;">
           <ace:AceoffixCtrl ID="AceoffixCtrl1" runat="server" BorderStyle="BorderThin" 
              Menubar="True" Theme="Office2010" >
           </ace:AceoffixCtrl>
        </div>
    
    Then write the following server code in word/editword.aspx.cs.
    
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("../doc/editword.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");
    
    Note: If your url does not endwith “.aspx” in your web project, you must set the properties as following code:
    
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server";
        AceoffixCtrl1.SaveFilePage = "savefile";
        
    Otherwise, you will get the errors:

![截图5](https://github.com/cupandcup/Edit-Word-Online-By-AceOffix/blob/master/Examples5/image/integrate_5.png?raw=true)

    Or

![截图6](https://github.com/cupandcup/Edit-Word-Online-By-AceOffix/blob/master/Examples5/image/integrate_6.png?raw=true)

    Add a new web page called savefile.aspx if your user want to save his/her document.
    
        protected void Page_Load(object sender, EventArgs e)
        {
                // This page is used to receive the file saved by AceoffixCtrl.
                Aceoffix.FileRequest freq = new Aceoffix.FileRequest();
                freq.SaveToFile(Server.MapPath("../doc/") + freq.FileName);
                freq.Close();
        }
5. Congratulations. Your user can open, edit, save the Word documents in your web project online now. You have finished the integration of Aceoffix. Please refer to the examples5 to learn how to use more features of Aceoffix.

![截图7](https://github.com/cupandcup/Edit-Word-Online-By-AceOffix/blob/master/Examples5/image/integrate_7.png?raw=true)
