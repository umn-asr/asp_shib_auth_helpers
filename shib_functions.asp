<%

function shib_is_logged_in()
  shib_is_logged_in = (Request.ServerVariables("HTTP_REMOTEUSER") <> "")
end function

function shib_internet_id()
  shib_internet_id = shib_get_attribute("internet_id")
end function

function shib_get_attribute(which_one)
  select case which_one
    case "internet_id"
      ' From goggins@umn.edu, returns "goggins"
      a = split(Request.ServerVariables("HTTP_REMOTEUSER"), "@")
      shib_get_attribute = a(0)
    case "tld"
      a = split(Request.ServerVariables("HTTP_REMOTEUSER"), "@")
      shib_get_attribute = a(1)
    case else
      Err.Raise 8 ' 8 is a user-defined error
      Err.Description = "shib_get_attribute does not support a " & which_one & " mode"
    end select
end function

function shib_login_and_redirect_url()
  shib_login_and_redirect_url = "https://" & _
    request.servervariables("HTTP_HOST") & _
    "/Shibboleth.sso/Login?target=" & _
    server.urlencode(shib_get_current_url())
end function

function shib_get_current_url()
  dest_url = request.servervariables("HTTP_HOST") & _
    request.servervariables("SCRIPT_NAME")
  first = true
  for each key in request.querystring
    if first then
      dest_url = dest_url & "?"
      first = false
    else
      dest_url = dest_url & "&"
    end if
    dest_url = dest_url & key & "=" & request.servervariables(key)
  next
  shib_get_current_url = "https://" & dest_url
end function

function shib_logout_url()
  shib_logout_url = "https://" & _
    request.servervariables("HTTP_HOST") & _
    "/Shibboleth.sso/Logout?" & _
    "return=" & _
    server.urlencode(shib_get_current_url())
end function

' If you except the web-server to protect a script, you can place this 
' at the top of the script to simply cause the script to exist if the user is not logged in
function shib_auth_required()
  if not shib_is_logged_in() then
    msg="Sorry, an unexpected error has occurred (Shibboleth authentication credentials are not available)." & Chr(10) & Chr(13) & _
    "Please contact the administer of this page if this error persists."
    Err.Raise 8 ' 8 is a user-defined error
    Err.Description = msg 
    response.end
  end if
end function

%>
