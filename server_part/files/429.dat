require("helper")
require("db.msaccess")
require("igrid")

function LadeBeispielDaten()
  local con = db.msaccess.connect([[C:\Users\mw\Dropbox\netProjekte1\UU\SVS\DB\es-adr-test2.mdb]])

  local header, content = con:getall([[
    SELECT
      id,name,vorname,plzneu,ort,ausland
    FROM
      [UGB-ADR-TAB]
    WHERE
      id < 2000
  ]])
  
  return header, content
end

function Filter()
  local such = FindById(f, "iText_such").Text
  for i = 0, G.Rows.Count-1 do
    if string.find(G.Cells(i, 1).Value, such, 1, true) then G.Rows(i).Visible = true else G.Rows(i).Visible = false end
  end
end

local f = Form { Text = "FlowLayoutTest", Width = 500 }

local spl, p1, p2 = SplitContainer { f; Orientation = WinForms.Orientation.Horizontal }

local L = FlowLayout.fromContainer ( p1 )
  L:SetInputDefault( 100, 200, { BackColor = {"#333333"}, ForeColor = {"#FFFFFF"} }, {} )

local TB = FlowLayout.create { Dock = WinForms.DockStyle.Top }
  p2.Controls:Add(TB.container)
  TB:SetInputDefault( 100, 200, { BackColor = {"#333333"}, ForeColor = {"#FFFFFF"} }, {} )
  TB:AddInput ( "such", "suche:", "" )
  TB:Add(Button("OK", Filter))

local G = iGrid ({ Dock = WinForms.DockStyle.Fill }, LadeBeispielDaten() )
  p2.Controls:Add(G)
  G.CurRowChanged:Add(function()
    L:Reset(10, 10)
    for i = 0, G.Cols.Count-1 do
      L:AddInput ( i, G.Cols:get_Item(i).Text, G.CurRow.Cells:get_Item(i).Value )        L:Break()
    end
    
    L:Add(Button("OK", function() f:Close() end))
  end)

f:Show()

