# How to support all the browsers

Recently, Chrome and Firefox browsers claimed that they do not support NPAPI plugins. Aceoffix V4 and the earlier versions cannot run in the latest Chrome and Firefox browsers. So we develop a new feature to solve this problem in Aceoffix V5.

How to support all the browsers in Aceoffix V5? Take the following steps:

1. Add the following JavaScript references to your web page.

        <script type="text/javascript" src="aceoffix-runtime/js/jquery.min.js"></script>
		
        <script type="text/javascript" src="aceoffix-runtime/js/aceoffix.js" id="ace_js_main"></script>

    **Note: The paths of jquery.min.js  and aceoffix.js are relative to the root of your website.**
    
    ![](https://github.com/aceoffix/AceoffixforAsp.Net/blob/master/Examples5/image/support_1.png?raw=true)
	
2. Change the old link to new link.

    Refer to Examples5 /Default.aspx to learn how to change the old link to new link.

    The old link:

    ![](https://github.com/aceoffix/AceoffixforAsp.Net/blob/master/Examples5/image/support_2.png?raw=true)

    For example: The old link is 
	
	    <a href=”word/editword.aspx”>Edit Word document</a>
	
    The new link:
    
    ![](https://github.com/aceoffix/AceoffixforAsp.Net/blob/master/Examples5/image/support_3.png?raw=true)
    
    The new link is 
	
	    <a href=”javascript:AceBrowser.openWindowModeless('word/editword.aspx', 'width=1200px;height=800px;');”>Edit Word document</a>`

3. The web page with Aceoffix will pop up when you click the new link in Chrome or Firefox browsers.  This new link can also work in other browsers (e.g. IE,  Opera, Edge).

![](https://github.com/aceoffix/AceoffixforAsp.Net/blob/master/Examples5/image/support_4.png?raw=true)
