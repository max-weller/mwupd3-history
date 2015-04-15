
-- include the iGrid assembly
reference("TenTec.Windows.iGridLib.iGrid.dll")

iGLib = luanet.TenTec.Windows.iGridLib

function iGrid(opt, headers, content)
  local ig = iGLib.iGrid()
  setProperties(ig, opt)
  if headers ~= nil then
    for i,v in ipairs(headers) do
      ig.Cols:Add(v)
    end
  end
  if content ~= nil then
    ig.Rows.Count = #content
    for i,v in ipairs(content) do
      for j,vv in ipairs(v) do
        ig.Cells:get_Item(i - 1, j - 1).Value = vv
      end
    end
  end
  ig.KeyUp:Add(function(sender,e)
    if e.KeyCode == WinForms.Keys.X and e.Control then
      clipset(iGGetSelRowsAsString(ig))
      iGDeleteSelRows(ig)
    end
    if e.KeyCode == WinForms.Keys.C and e.Control then
      clipset(iGGetSelRowsAsString(ig))
    end
    if e.KeyCode == WinForms.Keys.V and e.Control then
      iGInsertRowsFromString(ig, clipget())
    end
  end)
  return ig
end

function iGGetSelRowsAsString(ig)
  local out = {}
  local rows = ig.SelectedRows
  for i = 0, rows.Count - 1 do
    local c = {}
    for j = 0, ig.Cols.Count - 1 do
      table.insert(c, rows:get_Item(i).Cells:get_Item(j).Value)
    end
    table.insert(out, table.concat(c, "\t"))
  end
  return table.concat(out, "\n")
end

function iGDeleteSelRows(ig)
  ig:BeginUpdate()
  local rows = ig.SelectedRows
  trace("selRows", rows.Count)
  for i = rows.Count - 1, 0, -1 do
    trace("i",i)
    local row = rows:get_Item(i)
    ig.Rows:RemoveAt(row.Index)
  end
  ig:EndUpdate()
end

function iGInsertRowsFromString(ig, str, pos)
  if pos == nil then
    if ig.CurRow == nil then pos = 0 else pos = ig.CurRow.Index end
  end
  iGDeselectAll(ig)
  local lines = explode("\n", str)
  for offset,line in ipairs(lines) do
    local parts = explode("\t", line)
    local row = ig.Rows:Insert(pos + offset - 1)
    row.Selected = true
    for col, v in ipairs(parts) do
      row.Cells:get_Item(col - 1).Value = v
    end
  end
  
  ig.CurRow = ig.Rows:get_Item(pos)
end

function iGDeselectAll(ig)
  for i = 0, ig.Rows.Count - 1 do
    ig.Rows:get_Item(i).Selected = false
  end
end
