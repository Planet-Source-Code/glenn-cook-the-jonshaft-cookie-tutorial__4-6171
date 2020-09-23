<div align="center">

## The JonShaft Cookie Tutorial


</div>

### Description

This Cookie tutorial is designed for anyone interested in learning how to control a cookie with ASP.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Glenn Cook](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/glenn-cook.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/glenn-cook-the-jonshaft-cookie-tutorial__4-6171/archive/master.zip)





### Source Code


<p><strong><font color="#000000" face="Verdana"><small><small>Important Update:</small></small></font></strong><br>
<font face="Verdana"><small><small>ASP Trencher, Bryant L. from Baarns
Consulting, sent me an e-mail recently telling me that&nbsp; IIS4.0 machines
using ASP2.0 do not need the leading &quot;.&quot; for defining the domain
variable for the cookie, and it works with IE and Netscape!&nbsp; I tested this
out on a local machine of my own and sure enough, it worked perfectly.&nbsp; I
haven't updated the cookie-code you'll see below to reflect that because not
everyone is using that exact configuration, but as ASP and IIS improve so will
these little headaches.&nbsp; Thanks, Bryant.<br>
</small></small></font><font color="#000000" face="Verdana"><br>
</font><font face="Verdana"><strong><font color="#000000">How it Works:</font><br>
<font color="#008000">Green = Server-side ASP code</font><font color="#000000"><br>
</font><font color="#800080">Purple= HTML Code</font><font color="#000000"><br>
Black= Visible HTML Text<br>
</font><font color="#ff0000">Red= My Comments</font></strong></font></p>
<table border="1" cellPadding="4" cellSpacing="0" width="100%">
 <tbody>
  <tr>
   <td vAlign="top" width="58%">
    <p><font face="Courier New" color="forestgreen" size="+0"><small>&lt;%
    If Request.Cookies(&quot;JonShaft&quot;).HasKeys Then %&gt;</small></font><font face="Courier New"><small><font color="limegreen" face><br>
    </font></small></font></p>
    <p><font face="Courier New"><small><font color="#800080">&lt;HTML&gt;<br>
    &lt;HEAD&gt;<br>
    &lt;TITLE&gt;</font>A Jon Shaft Cookie. It's cheesy!<font color="#800080">&lt;/TITLE&gt;</font><br>
    <font color="#800080">&lt;/HEAD&gt;<br>
    </font></small></font></p>
    <p><font color="#008000" face="Courier New"><small><br>
    </small></font>&nbsp;</p>
   </td>
   <td vAlign="top" width="42%"><small><font color="#ff0000" face="Arial">'
    This&nbsp; is my &quot; if then&quot; where I find out if the user
    already has the JonShaft Cookie on their system.&nbsp; The HasKeys
    attribute is real handy for checking cookies which have multiple values
    associated with them- those values are referred to as Keys by ASP.&nbsp;
    This cookie says, if they've got the cookie, execute the next statement.</font></small>
    <p>&nbsp;</p>
   </td>
  </tr>
  <tr>
   <td vAlign="top" width="58%">
    <p><font face="Courier New"><small>Welcome Back, <font color="#008000">&lt;%Response.Write
    (Request.Cookies(&quot;JonShaft&quot;)(&quot;FirstName&quot;))%&gt;</font><font color="#800080">&amp;#32</font></small><br>
    <small><font color="#008000">&lt;%Response.Write(Request.Cookies(&quot;JonShaft&quot;)(&quot;LastName&quot;))%&gt;</font><font color="#000000">!</font></small></font></p>
    <p><font face="Courier New"><small><font color="#ff0000"><em><small><img src="http://www.aspalliance.com/glenncook/images/cookieboy.jpg" width="128" height="119"></small></em></font></small></font></p>
   </td>
   <td vAlign="top" width="42%"><font color="#ff0000" face="Arial"><small>'This&nbsp;
    line basically says, &quot;OK, they've got the cookie, let's Request the
    cookie's keys/info and write them to the page.&quot;&nbsp; The Response
    Object allows me to spit information to the user, the Request Object
    allows me to extract it from the user.&nbsp; Basically what we've done
    is said, &quot;Check for cookie(Request), extract cookie(Request), write
    cookie to page(Response)&quot;</small><big><br>
    </big><small>**The &amp;#32 tells HTML to enter a space**</small></font><font color="#008000"><br>
    </font>
    <p>&nbsp;</p>
   </td>
  </tr>
  <tr>
   <td vAlign="top" width="58%">&nbsp;</td>
   <td vAlign="top" width="42%">
    <p>&nbsp;</p>
   </td>
  </tr>
  <tr>
   <td vAlign="top" width="58%"><font color="#008000" face="Courier New"><small>&lt;%<br>
    Else If &quot;BadMutha&quot; = Request(&quot;ActionType&quot;) Then</small><br>
    <small>TheFirstName=Request(&quot;FirstName&quot;)</small><br>
    <small>TheLastName=Request(&quot;LastName&quot;)</small></font></td>
   <td vAlign="top" width="42%"><font color="#ff0000" face="Arial"><small><small><font size="+0" color="#ff0000" face="Arial">'The
    &quot;Else in the first line says,&quot;Ok, the &quot;If-Then&quot;
    wasn't true....But there's more ahead!&quot;<br>
    </font></small><font size="+0">'</font>This section is for the form
    input and it creates the cookie.&nbsp; You see, this single page of code
    serves three functions: It's for people who've been here before, people
    who haven't, and it makes a cookie for the people based on their form
    input.&nbsp; You'll notice that just after the FORM METHOD html I have
    some ASP code which actually asks for it's own name so it can post to
    itself!</small><big><br>
    </big><small>The &quot;If-Then&quot; statement checks to see if the user
    sent a form with the name &quot;ActionType&quot; which has the value
    equal to &quot;BadMutha&quot;!</small><big><br>
    </big><small>It also makes two variables based on the user input to
    stick into the JonShaft cookie.&nbsp; I call the variables TheFirstName
    and TheLastName appropriately.</small></font>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
   </td>
  </tr>
  <tr>
   <td vAlign="top" width="58%"><font face="Courier New"><font color="#008000"><small>Response.Cookies(&quot;JonShaft&quot;)(&quot;FirstName&quot;)
    = TheFirstName</small><br>
    <small>Response.Cookies(&quot;JonShaft&quot;)(&quot;LastName&quot;) =
    TheLastName</small><br>
    <small>Response.Cookies(&quot;JonShaft&quot;).Expires = #September 3,</small>
    <small>2001#<br>
    Response.Cookies(&quot;JonShaft&quot;).Domain=&nbsp; &amp;_ &quot;.www.activeserverpages.com&quot;</small><br>
    <small>Response.Cookies(&quot;JonShaft&quot;).Path = &quot;/glenncook&quot;<br>
    Response.Write &quot;Thanks for your submission, &quot;<br>
    Response.Write(Request(&quot;FirstName&quot;))%&gt;</small></font><small><font color="#000000">!</font></small></font></td>
   <td vAlign="top" width="42%"><small><font color="#ff0000" face="Arial">'The
    Response Object is your cookie writing friend! This code actually writes
    the cookie to the client's system.&nbsp; You'll notice that I make the
    The FirstName key equal to the &quot;TheFirstName&quot; variable which I
    extracted from the Request(&quot;FirstName&quot;) Querystring from the
    input form.(Whoooo! That was a mouthful!)</font></small>
    <p><small><font color="#ff0000" face="Arial">Then I tell the cookie when
    to expire, the domain that it should be sent to, and the path within the
    domain.&nbsp; But the secret recipe is that little period in the
    domain=&quot;.www.activeserverpages.com&quot;&nbsp; Actually without
    that little period, no cookie!&nbsp; Charles Carol helped me on this
    little issue which drove me nuts. &nbsp; MAKE SURE THE DOT IS THERE!
    Also make sure the path is EXACTLY as I wrote it.</font></small></p>
    <p>&nbsp;</p>
   </td>
  </tr>
  <tr>
   <td vAlign="top" width="58%"><font color="#008000" face="Courier New"><small>&lt;%Else%&gt;</small></font></td>
   <td vAlign="top" width="42%"><small><font color="#ff0000" face="Arial">'The
    &quot;Else&quot; code here basically says,&quot; Ok, they don't have the
    cookie and they didn't send any form information, send them the
    following code!</font></small>
    <p>&nbsp;</p>
   </td>
  </tr>
  <tr>
   <td vAlign="top" width="58%"><font face="Courier New"><small><font color="#800080">&lt;FORM
    METHOD=POST ACTION=&quot;</font><font color="#008000">&lt;%=Request.ServerVariables(&quot;SCRIPT_NAME&quot;)%&gt;</font><font color="#800080">&quot;&gt;</font><br>
    <font color="#800080">&lt;input type=&quot;hidden&quot; name=&quot;ActionType&quot;
    value=&quot;BadMutha&quot;&gt;</font><br>
    <br>
    You must be new around here. Gimme your name?!<font color="#800080">&lt;p&gt;</font><br>
    FIRST NAME:<font color="#800080">&lt;input type=&quot;text&quot;
    name=&quot;FirstName&quot; size=&quot;15&quot;&gt;&lt;br&gt;</font><br>
    LAST NAME:<font color="#800080">&amp;#32&lt;input type=&quot;text&quot;
    name=&quot;LastName&quot; size=&quot;15&quot;&gt;&lt;p&gt;<br>
    </font><br>
    <font color="#800080">&lt;input type=&quot;reset&quot; value=&quot;Clear
    Form&quot;&gt;<br>
    &lt;input TYPE=&quot;submit&quot; VALUE=&quot;Submit Info!&quot;&gt;</font><br>
    <font color="#008040">&lt;%End If%&gt;</font><br>
    <font color="#008000">&lt;%End If%&gt;</font><br>
    <font color="#800080">&lt;/BODY&gt;<br>
    &lt;/HTML&gt;<br>
    </font></small></font></td>
   <td vAlign="top" width="42%"><small><font color="#ff0000" face="Arial">'This
    section is your HTML input form for the new visitor!&nbsp; You'll notice
    that I stuck a hidden input box in there.&nbsp; Well basically&nbsp;
    that's so I can get &quot;Bad Mutha&quot; as the Action Type but is very
    effective for passing stuff that the user doesn't need to see.</font></small>
    <p><small><font color="#ff0000" face="Arial">I also extract the name of
    this asp page -which I mentioned above- using the
    Request.ServerVariables object.&nbsp; Remember: If you need some
    information in ASP pages, just &quot;Request&quot; it.</font></small></p>
    <p><small><font color="#ff0000" face="Arial">Finally, don't forget to
    End your If!</font></small></p>
   </td>
  </tr>
  <tr>
   <td vAlign="top" width="58%"><font color="#ff0000" face="Verdana"><strong>Some
    tips and suggestions!</strong></font>
    <ul>
     <li><font color="#ff0000" face="Verdana"><small>One of the most useful
      things I've seen cookies used for is with Human Resource type
      applications where the basic user information is stored as a cookie.&nbsp;
      That way everytime they go to access a form to make changes they
      don't have to re-type the form input information.&nbsp; Remember, DO
      NOT store sensitive information in a cookie.</small></font></li>
    </ul>
    <ul>
     <li><font color="#ff0000" face="Verdana"><small>If you are rewriting
      any cookie information back to an existing cookie you need to update
      all of the cooky's information (e.g. &quot;path&quot;,
      &quot;domain&quot;, &quot;expiration date&quot; etc.)</small></font></li>
    </ul>
    <p>&nbsp;</p>
   </td>
   <td vAlign="top" width="42%">&nbsp;
    <ul>
     <li><small><font color="#ff0000" face="Verdana">Although a cookie
      might be useful for &quot;extra&quot; authentication, it should
      never be used for secure authentication purposes. Check out <a href="http://www.activeserverpages.com/flicks/">Kevin
      Flicks</a> site for some great info on authentication methods and
      security considerations.</font></small></li>
    </ul>
    <ul>
     <li><font color="#ff0000" face="Verdana"><small>IMPORTANT!</small><br>
      <small>Internet Explorer likes the domain info like &quot;.www.domain.com&quot;
      where Netscape likes it like &quot;.domain.com&quot;.&nbsp; (My
      cookie is written for Explorer.)</small></font></li>
    </ul>
    <ul>
     <li><font color="#ff0000" face="Verdana"><small>Cookies are finicky.&nbsp;
      Any little mistake will have you pulling your hair out for hours.</small></font></li>
    </ul>
    <p>&nbsp;</p>
   </td>
  </tr>
 </tbody>
</table>
<p></p>

