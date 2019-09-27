<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="wordForm_Default" %>

 
<!DOCTYPE HTML>
<!--[if lt IE 7]><html class="no-js lt-ie9 lt-ie8 lt-ie7"><![endif]-->
<!--[if IE 7]><html class="no-js lt-ie9 lt-ie8"><![endif]-->
<!--[if IE 8]><html class="no-js lt-ie9"><![endif]-->
<!--[if gt IE 8]><!--><html class="no-js"><!--<![endif]--><head>

<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

<!-- Mobile Specific Metas -->
<meta name="viewport" content="width=device-width, initial-scale = 1.0, user-scalable = no">
<meta http-equiv="X-UA-Compatible" content="IE=Edge">

<title>Acesoft Corporation -  A demo of Absence Request</title>

<!-- Theme styles -->
<link href="../css/acestyle.css" rel="stylesheet">

<link rel='stylesheet' id='contact-form-7-css'  href='../css/styles.css?ver=4.2.2' type='text/css' media='all' />
<!-- Aceoffix -->
<script type="text/javascript" src="../aceoffix-runtime/js/jquery.min.js"></script>
<script type="text/javascript" src="../aceoffix-runtime/js/aceoffix.js" id="ace_js_main"></script>
</head>

<body>
<!--[if lt IE 7]>
	<p class="chromeframe">You are using an <strong>outdated</strong> browser. Please <a href="http://browsehappy.com/">upgrade your browser</a> or <a href="http://www.google.com/chromeframe/?redirect=true">activate Google Chrome Frame</a> to improve your experience.</p>
	<![endif]-->

<div id="banner-wrap" class="clearfix">
	<div id="top-nav" class="container clearfix">
        <div class="grid_10 center alpha">
        	Web Development Platform for Microsoft Office
        </div>
        <div class="grid_2 right omega">
        	<a href="http://www.aceoffix.com/index.php/contact-us/"><img src="../image/contactw.png" alt="contact">&nbsp;&nbsp;Contact Us</a>
        </div>
	</div>
    <div id="nav-wrap" class="clearfix">
        <div class="container">
            <div id="logo" class="grid_3 alpha"><a href="http://www.aceoffix.com"><img src="../image/aceoffixlogo.png" height="48" width="229" alt="Aceoffix" class="scale-with-grid" /></a></div>
            <div class="grid_9 omega">
                <nav class="nav">
                    <ul class="nav-list"><li class="page_item page-item-2 current_page_item menu-item-8"><a href="http://www.aceoffix.com/">Home</a></li>
<li class="page_item menu-item-244"><a href="http://www.aceoffix.com/index.php/overview/">Products</a>
</li>
<li class="page_item menu-item-224"><a href="http://www.aceoffix.com/index.php/download/">Download</a>
</li>
<li class="page_item menu-item-236"><a href="http://www.aceoffix.com/index.php/recent-update/">Support</a>
</li>
<li class="page_item menu-item-205"><a href="http://www.aceoffix.com/index.php/purchase/">Purchase</a>
</li>
<li class="page_item menu-item-17"><a href="http://www.aceoffix.com/index.php/about-us/">About Us</a></li>
</ul>                </nav>
            </div>
        </div>
    </div>


    <div id="title-wrap" class="clearfix">
        <div class="container left">
        Current Page:  Example Center
            <h1>Example Center - Aceoffix V5</h1>
            
        </div>
    </div>
    		
    </div>
    <br/>
    <div class="container left">
        <div class="left"><span>Current Page: <a href="../Default.aspx" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>A demo of Absence Request</em></strong></span></div>
        <div class="right">
            
        </div>
        <div class="clear"></div>
    </div>

	<div id="inner-wrap" class="clearfix">
	<div class="container">
    	<div class="inner-content">

<!--************** Start of AceoffixCtrl JavaScript ************************-->
 <script type="text/javascript">
     function onColor(td) {
         td.style.backgroundColor = '#D7FFEE';
     }
     function offColor(td) {
         td.style.backgroundColor = '';
     }
     function SetLinkUrl(svrpage, fileid) {
         AceBrowser.openWindowModeless(svrpage + '?ID=' + fileid, 'width=1200px;height=800px;');
     }
     function openDataList(svrpage, fileid) {
         window.open(
				svrpage + '?ID=' + fileid, "", "fullscreen=0,toolbar=0,location=1,directories=0,status=0,menubar=0,scrollbars=1,resizable=0,width="
				+ 500 + ",height=" + 320 + " ,top=200,left=100", true);

     }
</script>
		
		<form name="form1" id="form1" action="DocList.aspx" method="post">
			<input id="FileSubject" name="FileSubject" type="hidden" />
			<input id="TemplateName" name="TemplateName" type="hidden" />
		</form>
		<div id="content">

			<div style="margin: 25px 0; font-size: 20px; font-weight: bold; text-align:center;" >
				Use template to generate document
			</div>
			<div id="textcontent">
				<p>
					<strong>Description: </strong>
				</p>
				<div class="flow1">
					This demo shows how to fill Word template to generate a formatted document; how to get data from Word document and store them to database on server; 
                    and how to control the editable regions in Word document by specifying the editable regions and how to enable user to select data from drop-down box instead of editing the content in DataRegion directly.
				</div>	
			</div>
			<br/><br/>
			<div class="flow1" style="text-align:center">
				<table class="zz-talbe" border="1" cellpadding="3" cellspacing="0" style="width: 80%;margin:auto" >
					<caption style="font-size: 20px;">
						List of Absence Request
					</caption><br/><br/>
					<thead style="border: 1px solid #999999;">
						<tr>
							<th width="300" style="border: 1px solid #999999;">
								Theme
							</th>
							<th width="80" style="border: 1px solid #999999;">
								Date
							</th>
							<th width="200" style="border: 1px solid #999999;"> 
								Operation
							</th>
						</tr>
					</thead>
                     <tbody>
                    <asp:Literal id="LiteralGrid" runat="server"></asp:Literal>
                </tbody>
				</table>
			</div>
		</div>
    <!--************** End of AceoffixCtrl JavaScript ************************-->	
</div></div></div>
    
<div id="quicklink-wrap" class="clearfix">
	<div class="container">
    	<h1>Quick Link</h1>
       <div class="footer-column grid_3 alpha">
       		<h3>Product</h3>
			<a href="http://www.aceoffix.com/index.php/overview/">Overview</a>
<a href="http://www.aceoffix.com/index.php/functions-features/">Functions &#038; Features</a>
<a href="http://www.aceoffix.com/index.php/trial-license/">Trial &#038; License</a>
<a href="http://www.aceoffix.com/index.php/purchase/">Purchase</a>
       </div>
       <div class="footer-column grid_3 mobile-omega2">
			<h3>Download</h3>
			<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for ASP.NET</a>
<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for JAVA</a>
<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for PHP</a>
       </div>
       <div class="footer-column grid_3 mobile-alpha2">
			<h3>Support</h3>
			<a href="http://www.aceoffix.com/index.php/technical-faq/">Technical FAQ</a>
<a href="http://www.aceoffix.com/index.php/videos/">Videos</a>
<a href="http://www.aceoffix.com/index.php/recent-update/">Recent Update</a>
       </div>
       <div class="footer-column grid_3 omega">
			<h3>Follow Us</h3>
			<a href="https://www.facebook.com/Aceoffix-152929314763836/">Facebook</a>
<a href="https://twitter.com/Aceoffix">Twitter</a>
       </div>
    </div>
</div>

<div id="footer-wrap" class="clearfix">
    <div class="container center">
		<a href="http://www.aceoffix.com/">Home</a>
<a href="http://www.aceoffix.com/index.php/acesoft-customer-privacy-policy/">Privacy Policy</a>
<a href="http://www.aceoffix.com/index.php/acesoft-terms-of-use/">Terms of Use</a>
<a href="http://www.aceoffix.com/index.php/contact-us/">Contact Us</a>
<a href="http://www.aceoffix.com/index.php/sitemap/">Sitemap</a>
		<br>
		&copy; 2018 Acesoft Corporation. All rights reserved.
    </div>
</div>
</body>
</html>
