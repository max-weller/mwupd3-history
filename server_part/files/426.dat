require("win32apihelper")

--
-- debug interaktiv
function debuginteraktiv()
  luanet.load_assembly("System.Windows.Forms")
  luanet.load_assembly("System.Drawing")
  
  local WinForms = luanet.System.Windows.Forms
  
  local info = debug.getinfo(debuginteraktiv, "S")
  trace("Initialized debuginteraktiv", "omitting file "..info.short_src, "ini")
  local editor = ide:getActiveTab()
  local lastfile = nil
  local isBreak = false
  local stackDepth = 0
  local function traceme (event, line)
    local func = debug.getinfo(2)
    local s = func.short_src
    if s == info.short_src or s:sub(0, 1) == "[" then return end
    if event == "call" then
      stackDepth = stackDepth + 1
      printline(stackDepth + 30, func.short_src .. ":" .. func.what .. ": " .. (func.name or "<unk>") .. " (" .. func.namewhat .. ")", type(func.nups))
      printline(13, "sd", stackDepth)
      
    elseif event == "return" then
      printline(stackDepth + 30, "--", "")
      stackDepth = stackDepth - 1
      printline(13, "sd", stackDepth)
      
    elseif event == "line" then
      if s ~= lastfile then
        ide:NavigateFile(s)
        luanet.System.Windows.Forms.Application.DoEvents()
        luanet.System.Threading.Thread.Sleep(100)
        lastfile = s
        editor = ide:getActiveTab()
      end
      printline(11, "", "file: " .. s)
      printline(12, "line: ", line)
      
      editor:highlightExecutingLine(line - 1)
      luanet.System.Threading.Thread.Sleep(30)
      luanet.System.Windows.Forms.Application.DoEvents()
      if is_pause_button() then --  Pause key
        while is_pause_button() do
          luanet.System.Threading.Thread.Sleep(100)
        end
        isBreak = true
      end
      while isBreak do -- 121 = F10   120 = F9    119 = F8
        if win32apihelper.getAsyncKeyState(121) ~= 0 then isBreak = false end
        if win32apihelper.getAsyncKeyState(120) ~= 0 then break end
        if win32apihelper.getAsyncKeyState(119) ~= 0 then trace_locals() end
        if is_pause_button() then error("Debug break") end
        
        luanet.System.Threading.Thread.Sleep(100)
        luanet.System.Windows.Forms.Application.DoEvents()
      end
    end
  end
  function is_pause_button()
    return win32apihelper.getAsyncKeyState(19) ~= 0 or win32apihelper.getAsyncKeyState(44) ~= 0
  end
  function trace_locals()
    local i = 1
    local name, value = debug.getlocal(3, i)
    while name ~= nil do
      trace("Locals: " .. name, type(value) .. ": " .. tostring(value))
      i = i + 1
      name, value = debug.getlocal(3, i)
    end
  end
  stop = function() isBreak = true end
  
  debug.sethook(traceme, "lcr")
end


