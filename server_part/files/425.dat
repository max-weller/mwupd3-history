
reference("System.IO")
reference("System.Net")
IO = luanet.System.IO
Net = luanet.System.Net

function geturl(url)
  local rq = Net.WebRequest.Create(url)
  rq.Method = "GET"
  local resp = rq:GetResponse()
  local str = IO.StreamReader(resp:GetResponseStream())
  local res = str:ReadToEnd()
  str:Close()
  resp:Close()
  return res
end

function posturl(url, postData, method)
  local rq = Net.WebRequest.Create(url)
  rq.Method = method or "POST"
  local rqstr = IO.StreamWriter(rq:GetRequestStream())
  rqstr:Write(postData)
  rqstr:Close()
  local resp = rq:GetResponse()
  local str = IO.StreamReader(resp:GetResponseStream())
  local res = str:ReadToEnd()
  str:Close()
  resp:Close()
  return res
end

