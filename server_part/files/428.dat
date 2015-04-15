require("helper")
reference("System.Data")

DataOdbc = luanet.System.Data.Odbc


local DbMsaccess = {}
DbMsaccess.__index = DbMsaccess

function DbMsaccess.connect(filename)
  local db = {}
  setmetatable(db, DbMsaccess)
  --local connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s"
  local connectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=%s;Uid=Admin;Pwd=;"
  --local connectionString = "Driver={Microsoft Access Driver (*.mdb)};DBQ=%s"
  trace("connectionString", string.format(connectionString, filename), "dump")
  local con = DataOdbc.OdbcConnection(string.format(connectionString, filename))
  con:Open()
  db.connection = con
  return db
end

function DbMsaccess:close()
  self.connection:Close()
end

function DbMsaccess:getschema(what)
  local dat_tab = self.connection:GetSchema(what, nil)
  local schema = {}
  for row in enum(dat_tab.Rows) do
    local d = {}
    for cols in enum(dat_tab.Columns) do
      d[cols.Caption] = row:get_Item(cols)
    end
    table.insert(schema, d)
  end
  return schema
end


function DbMsaccess:gettables()
  local dat_tab = self.connection:GetSchema("Tables", nil)
  local tableList = {}
  for row in enum(dat_tab.Rows) do
    if row:get_Item("TABLE_TYPE") == "TABLE" then
      table.insert(tableList, row:get_Item("TABLE_NAME"))
    end
  end
  return tableList
end

function DbMsaccess:getsingle(sql)
  local cmd = DataOdbc.OdbcCommand(sql)
  cmd.Connection = self.connection
  return cmd:ExecuteScalar()
end

local function unnull(reader, i)
  if reader:IsDBNull(i) then
    return nil
  else
    return reader:GetValue(i)
  end
end

function DbMsaccess:getall(sql)
  local cmd = DataOdbc.OdbcCommand(sql)
  cmd.Connection = self.connection
  local reader = cmd:ExecuteReader()
  if reader.HasRows then
    local header, content = {}, {}
    for i = 0, reader.FieldCount - 1 do table.insert(header, reader:GetName(i)) end
    repeat
      local row = {}
      for i = 0, reader.FieldCount - 1 do table.insert(row, unnull(reader, i)) end
      table.insert(content, row)
    until reader:Read() == false
    return header, content
  else
    return {}, {}
  end
end



db = db or {}
db.msaccess = DbMsaccess

