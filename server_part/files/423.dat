
-->>> MW Lua Helper Lib

--> aliases
sprintf = string.format
reference = luanet.load_assembly


--> references
reference("System.Windows.Forms")
reference("System.Drawing")
reference("ScriptIDE.Core")

WinForms = luanet.System.Windows.Forms
Drawing = luanet.System.Drawing
ResourceLoader = luanet.ScriptIDE.Core.ResourceLoader

--> messageBox
function msgbox(text, title)
  WinForms.MessageBox.Show(text, title or "Alert", WinForms.MessageBoxButtons.OK,
                           WinForms.MessageBoxIcon.Information)
end
function alert(text, title)
  WinForms.MessageBox.Show(text, title or "Alert", WinForms.MessageBoxButtons.OK,
                           WinForms.MessageBoxIcon.Warning)
end

--> Lua: missing funcs
function enum(o)
   local e = o:GetEnumerator()
   return function()
      if e:MoveNext() then return e.Current end
   end
end
function explode(d,p)   -- explode(seperator, string)
  local t, ll
  t={}
  ll=0
  if(#p == 1) then return {p} end
    while true do
      l=string.find(p,d,ll,true) -- find the next d in the string
      if l~=nil then -- if "not not" found then..
        table.insert(t, string.sub(p,ll,l-1)) -- Save it in our array.
        ll=l+1 -- save just after where we found it for searching next time.
      else
        table.insert(t, string.sub(p,ll)) -- Save what's left in our array.
        break -- Break at end, as it should be, according to the lua manual.
      end
    end
  return t
end
-- Find all matches and return the last one.
function string.lastIndexOf(haystack, needle, plain)
  if plain == nil then plain = true end
  local i = -1
  local j, answer
  repeat
    answer = i
    i, j = string.find(haystack, needle, i + 1, plain)
  until (i == nil)
  if answer == -1 then return nil else return answer end
end


--
--> D U M P V A R : debugausgabe rekursiv
local recursionBlocker = {}
function dumpvar(desc, var)
  local out = {}   recursionBlocker = {}
  if not var then var = desc desc = "(unnamed)" end
  table.insert(out, sprintf("%-25s", desc))
  dumpvarint(out, var, "")
  return table.concat(out)
end
function dumpvarint(out, var, indent)
  if #out>1000 then table.insert(out, "--- snip ---\r\n") return end
  
  --[[
  if recursionBlocker[var] then table.insert(out, "recursionBlocker\r\n") return end
  recursionBlocker[var] = true   
  --]]
  
  indent = indent .. "+ "
  local typename = type(var)
  table.insert(out, sprintf("%11s", "("..typename..") "))
  
  if typename == "string" then
    table.insert(out, sprintf("%q\r\n", var))
  elseif typename == "table" then
    table.insert(out, ">>>\r\n")
    for k,v in pairs(var) do
      table.insert(out, sprintf("%-25s", indent .. tostring(k)))
      dumpvarint(out, v, indent)
    end
  elseif typename == "userdata" then
    table.insert(out, "--\r\n")
  else
    table.insert(out, tostring(var).."\r\n")
  end
end


--
-->>> Windows Forms lib
function Form(opt)
  local frm = WinForms.Form()
  setProperties(frm, opt)
  for i,v in ipairs(opt) do frm.Controls:Add(v) end
  return frm
end
function Menu(menus)
  local strip = WinForms.MenuStrip()
  for i,v in ipairs(menus) do
    trace ("MainMenuitem", v.Text)
    local menu = WinForms.ToolStripMenuItem(v.Text)
    for ii,vv in ipairs(v) do
      trace("--> Item","--> "..vv[1])
      local mi
      if vv[1] == "-" then
        mi = WinForms.ToolStripSeparator()
      else
        mi = WinForms.ToolStripMenuItem(vv[1])
      end
      menu.DropDownItems:Add(mi)
      if #vv > 1 then mi.Click:Add(vv[2]) end
    end
    strip.Items:Add(menu)
  end
  return strip
end
function Toolbar(tb)
  local strip = WinForms.ToolStrip()
  for i,v in ipairs(tb) do
    trace ("ToolStripItem", v[0])
    local ti = WinForms.ToolStripItem(v[0])
    if v[1] ~= nil then ti.Image = ResourceLoader.GetImageCached(v[1]) end
    if v.onclick ~= nil then ti.Click:Add(v.onclick) end
    
    strip.Items:Add(ti)
  end
  return strip
end
function SplitContainer(opt)
  local spl = WinForms.SplitContainer()
  if opt[1] ~= nil then opt[1].Controls:Add(spl) spl.Dock = WinForms.DockStyle.Fill end
  return spl, spl.Panel1, spl.Panel2
end
function Button(txt, onclick)
  local btn = Control("Button", x, y, xx, yy)
  btn.Text = txt
  btn.Click:Add(onclick)
  return btn
end
function Control(type, opt)
  local ctrl = WinForms[type]()
  if opt ~= nil then setProperties(ctrl, opt) end
  return ctrl
end
function SetBounds(x, y, xx, yy, ctrl)
  ctrl.Bounds = Drawing.Rectangle(x, y, xx, yy)
  return ctrl
end
function FindByID(parent, id)
  return parent.Controls:Find(id, true)[0]
end
function setProperties(ctrl, opt)
  for k,v in pairs(opt) do 
    if type(k) == "string" then
      if type(v) == "table" and type(v[1]) == "string" and v[1]:sub(1,1) == "#" then
        ctrl[k] = Drawing.ColorTranslator.FromHtml(v[1])
      else
        ctrl[k] = v
      end
    end
  end
end

local EnumToObject = luanet.get_method_bysig(luanet.System.Enum, "ToObject", "System.Type", "System.Int32")
local WinFormsAnchorStylesType = WinForms.AnchorStyles.Top:GetType()
AnchorTop, AnchorLeft, AnchorRight, AnchorBottom = 1, 4, 8, 2
function Anchor(flags)
  return EnumToObject(WinFormsAnchorStylesType, flags)
end


--
-->>> F l o w  L a y o u t

FlowLayout = {}
FlowLayout.__index = FlowLayout

--> constructors
function FlowLayout.create(opt)
  local panel = WinForms.Panel()
  setProperties(panel, opt or {})
  return FlowLayout.fromContainer(panel)
end
function FlowLayout.fromContainer(container)
  local inst = {}
  setmetatable(inst, FlowLayout)
  inst.container = container
  inst.offsetX = 10   inst.currentX = 10  inst.currentY = 10
  inst.lastRowHeight = 0
  inst.paddingX = 5   inst.paddingY = 5
  inst.defaultLabelWidth = 90  inst.defaultTextWidth = 120
  return inst
end

--> members

function FlowLayout:Do(calls)
  trace("do",type(calls))
  for method,params in pairs(calls) do
    self[method](self, unpack(params))
  end
  return self
end

function FlowLayout:Init(func)
  func(self)
  return self
end

function FlowLayout:SetCursor(x,y) self.offsetX=x self.currentX=x self.currentY=y end
function FlowLayout:Reset(x,y) self.container.Controls:Clear() self:SetCursor(x,y) end

function FlowLayout:SetInputDefault(labelWidth, textWidth, labelOpt, textOpt)
  if labelWidth ~= nil then self.defaultLabelWidth = labelWidth end
  if textWidth ~= nil then self.defaultTextWidth = textWidth end
  if labelOpt ~= nil then self.defaultLabelOpt = labelOpt end
  if textOpt ~= nil then self.defaultTextOpt = textOpt end
end

function FlowLayout:AddInput(name, label, text, labelWidth, textWidth, labelOpt, textOpt)
  trace("FlowLayout:AddInput", label.." "..name)
  local l, t = WinForms.Label(), WinForms.TextBox()
  l.Width = labelWidth or self.defaultLabelWidth
  t.Width = textWidth or self.defaultTextWidth
  l.Height = t.Height
  l.Name = "iLabel_" .. name     l.Text = label
  t.Name = "iText_"  .. name     t.Text = text or ""
  if labelOpt ~= nil then setProperties(l, labelOpt) else setProperties(l, self.defaultLabelOpt) end
  if textOpt ~= nil then setProperties(t, textOpt) else setProperties(t, self.defaultTextOpt) end
  self:AddControl(l)
  self:AddControl(t)
  return l,t
end

function FlowLayout:AddControl(ctrl, opt)
  if opt ~= nil then setProperties(ctrl, opt) end
  self.container.Controls:Add(ctrl)
  ctrl.Left = self.currentX  ctrl.Top = self.currentY
  if ctrl.Height > self.lastRowHeight then self.lastRowHeight = ctrl.Height end
  self.currentX = self.currentX + ctrl.Width + self.paddingX
end

FlowLayout.Add = FlowLayout.AddControl

function FlowLayout:Break()
  self.currentX = self.offsetX
  self.currentY = self.currentY + self.lastRowHeight + self.paddingY
  self.lastRowHeight = 0
end

function FlowLayout:ByID(id)
  return self.container.Controls:Find(id, true)[0]
end
