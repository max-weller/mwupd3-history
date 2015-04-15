<?php
  #require_once "trace.inc";
  require_once "define.inc";
  
  class db {
  
  static $print_performed_queries = false;
  
  static $inifile;
  static $settings;
  
  static $link;
  static $query;
  static $lastresult;
  
  static $is_connected = false;
  
  static $named_queries = array(
    "GetPageByDidAndURL"     => "SELECT * FROM unc_page WHERE did = %d AND url = '%s'",
    "GetPageById"      => "SELECT * FROM unc_page WHERE pid = %d",
    "GetPageURLById"   => "SELECT CONCAT('http://',domain,url) FROM unc_page INNER JOIN unc_domain on unc_page.did=unc_domain.did WHERE pid = %d",
    "GetPageURLByOldCmsNodeId"   => "SELECT CONCAT('http://',domain,url) FROM unc_page INNER JOIN unc_domain on unc_page.did=unc_domain.did WHERE old_cms_nodeid = %d",
    "GetRedirectURL"   => "SELECT to_url, to_pid, is_external FROM unc_redirect WHERE did = %d AND from_url = '%s'",
    "GetUserByName"    => "SELECT * FROM unc_user WHERE id = %d",
    "GetUrlByOldNodeID"=> "SELECT url FROM unc_page WHERE old_cms_nodeid = %d",
    "CreateNewPage"    => "INSERT INTO unc_page (did,url,title_heading,title_browser,published_to,created,created_by,lastmodified,lastmodified_by,page_type,old_cms_nodeid,content_main,rel_pid)
    VALUES(%d,'%s','%s','%s','2038-01-19 00:00:00',NOW(),'%s',NOW(),'%s','%s',%d,'%s',%d)",
    "CreateNewSeminar" => "INSERT INTO unc_seminar (kuerzel,semtitel,semtext1,semtext2,semtext3,created,created_by,lastmodified,lastmodified_by)
    VALUES('%s','%s','%s','%s','%s',NOW(),'%s',NOW(),'%s')",
    "CreateNewSemTermin" => "INSERT INTO unc_semtermin (sid,old_cms_nodeid,nummer,location,street,city,
     start_dat,finish_dat,published_from,published_to,created,created_by,lastmodified,lastmodified_by)
    VALUES(%d,%d,'%s','%s','%s','%s',  '%s','%s','0000-00-00 00:00:00','2038-01-19 00:00:00' ,   NOW(),'%s',NOW(),'%s')",
    "DeletePage"       => "DELETE FROM unc_page WHERE pid = %d",
  );
  
  static function read_ini() {# var_dump($_COOKIE);
    if ($_COOKIE["lab_enable_dbdebug"])
        DB::$print_performed_queries = true;
    DB::$inifile = CONFIG_FILE;
    DB::$settings = parse_ini_file(DB::$inifile,true);
    DB::$settings["Salt"]= DB::$settings["Login"]["Salt"];
    return DB::$settings;
  }
  
  static function set_config($host, $user, $pass, $db) {
    DB::$settings = array("DB"=>array("Host"=>$host,"User"=>$user,"Pass"=>$pass,"Database"=>$db));
  }
  
  static function connect($alt_db="") {
    if (DB::$is_connected == true) return;
    if (!DB::$settings) DB::read_ini();
    $sect = $alt_db==null?"DB":"alt_db_$alt_db";
    DB::$link = @mysql_connect(DB::$settings[$sect]['Host'], DB::$settings[$sect]['User'], DB::$settings[$sect]['Pass']);
    if (!DB::$link) { trigger_error("Database Error occured. Please try again later.",E_USER_ERROR); }
    
    if (!@mysql_select_db(DB::$settings[$sect]['Database'])) { trigger_error("Database Error occured. Please try again later.",E_USER_ERROR); }
    
    mysql_set_charset('utf8');
    mysql_query("SET character_set_results = 'utf8', character_set_client = 'utf8', character_set_connection = 'utf8', character_set_database = 'utf8', character_set_server = 'utf8'");
    
    DB::$is_connected = true;
  }
  
  static function disconnect() {
    mysql_close();
    DB::$is_connected = false;
  }
  
  
  static function sql() {     // SQL sicher zusammenbauen
    DB::sql_int(func_get_args());
  }
  
  static function sql_int($params) {
    $max = count($params);
    if ($max < 2) {
      DB::$query = $params[0];
    } else {
      $q = $params[0];
      if ($q{0} == "@") $q = DB::$named_queries[substr($q,1)];
      for($i=1; $i<$max; $i++) $args[] = mysql_real_escape_string($params[$i], DB::$link);
      DB::$query = vsprintf($q, $args);
    }
  }
  
  static function execute() { DB::doQuery();
  }
  static function doQuery($sql = null) {
    if($sql) DB::$query = $sql;
    //trace(__CLASS__."::".__FUNCTION__, DB::$query);
    
    DB::printPerformedQuery(DB::$query);
    DB::$lastresult = mysql_query(DB::$query);
    if (!DB::$lastresult) { DB::dieWithError(); return false; }
    
    return DB::$lastresult;
  }
  
  static function printPerformedQuery($sql) {
    if (DB::$print_performed_queries) echo "<pre style='border:1px solid #448; background: #eef; padding: 10px'>".htmlspecialchars($sql)."</pre>";
    
  }
  
  static function nextRow() {
    $row = @mysql_fetch_assoc(DB::$lastresult);
    if (!$row) @mysql_free_result(DB::$lastresult);
    return $row;
  }
  
  static function numRows() {
    
    return mysql_num_rows(DB::$lastresult);
  }
  
  static function affectedRows() {
    
    return mysql_affected_rows();
  }
  
  static function insertId() {
    return mysql_insert_id();
  }
  
  static function dieWithError() {
    #if(!$GLOBALS["DONT_TRACE"])trace(__CLASS__."::".__FUNCTION__, "SQL-ERROR", mysql_error());
    trigger_error("Database Error occured. Please try again later.<br>".DB::$query."<br>".mysql_error(),E_USER_ERROR); return false;
    
  }
  
  static function printError() {
    printf("Es ist ein Datenbankfehler aufgetreten!<br>[%s] %s", mysql_errno(), mysql_error());
  }
  
  
  static function get() {    //ehemals mysql_getone
    if (func_num_args()>0) {
      $q = func_get_arg(0);
      $max = func_num_args();
      for($i=1; $i<$max; $i++) $args[] = mysql_real_escape_string(func_get_arg($i), DB::$link);
      DB::$query = vsprintf($q, $args);
    }
    //if (!$GLOBALS["DONT_TRACE"]) trace(__CLASS__."::".__FUNCTION__, DB::$query);
    DB::printPerformedQuery(DB::$query);
    $result = mysql_query(DB::$query);
    if (!$result) { DB::dieWithError(); return false; }
    
    $row = @mysql_fetch_assoc($result);
    @mysql_free_result($result);
    
    return $row;
  }
  
  static function getsingle() {
    if (func_num_args()>0) {
      $q = func_get_arg(0);
      $max = func_num_args();
      for($i=1; $i<$max; $i++) $args[] = mysql_real_escape_string(func_get_arg($i), DB::$link);
      DB::$query = vsprintf($q, $args);
    }
    #if (!$GLOBALS["DONT_TRACE"]) trace(__CLASS__."::".__FUNCTION__, DB::$query);
    DB::printPerformedQuery(DB::$query);
    $result = mysql_query(DB::$query);
    if (!$result) { DB::dieWithError(); return false; }
    
    $row = @mysql_fetch_row($result);
    @mysql_free_result($result);
    
    return $row[0]; // return only first value!
  }
  
  static function getlist() {   // ehemals mysql_getall
    if (func_num_args()>0) {
      $q = func_get_arg(0);
      $max = func_num_args();
      for($i=1; $i<$max; $i++) $args[] = mysql_real_escape_string(func_get_arg($i), DB::$link);
      DB::$query = vsprintf($q, $args);
    }
    #if (!$GLOBALS["DONT_TRACE"]) trace(__CLASS__."::".__FUNCTION__, DB::$query);
    DB::printPerformedQuery(DB::$query);
    $result = mysql_query(DB::$query);
    if (!$result) { DB::dieWithError(); return false; }
    
    $items = array();
    while ($row = @mysql_fetch_assoc($result)) {
      $items[] = $row;
    }
    @mysql_free_result($result);
    
    return $items;
  }
  
  static function escape($str) {
    return mysql_real_escape_string($str, DB::$link);
  }
  
  static function sql_getlink ($user_a, $user_b, $morewhere = "") {
    DB::$query = "SELECT * FROM twiki_userlink WHERE ((user_a = '$user_a' AND user_b = '$user_b') OR (user_a = '$user_b' AND user_b = '$user_a')) $morewhere";
    
    return DB::$query;
  }
  
  static function sql_insertinto($tablename, $data, $ignore="") {
    $sqlNAMES = implode(', ', array_keys($data));
    $sqlVALUES = implode(', ', array_values($data));
    
    DB::$query = "INSERT $ignore INTO `$tablename` ( $sqlNAMES ) VALUES ( $sqlVALUES );";
    
    return DB::$query;
  }
  
  static function sql_update($tablename, $where, $data, $apo = "'") {
    
    $sqlparts = array();
    foreach($data as $k=>$v) {
      $sqlparts[] = "`$k` = $apo$v$apo";
    }
    
    DB::$query  = "UPDATE `$tablename` SET " . implode(", ", $sqlparts) . " WHERE $where;"; 
    
    return DB::$query;
  }
  
  static function sql_update2($tablename, $where, $sqlparts) {
    DB::$query  = "UPDATE `$tablename` SET " . implode(", ", $sqlparts) . " WHERE $where;"; 
    
    return DB::$query;
  }
  
  } //endClass
  
  
  /*    WRAPPER FOR OLD DB-FUNCTIONS   */
  function mysql_getone() {
    DB::sql_int(func_get_args());
    return DB::get();
  }
  function mysql_exec() {
    DB::sql_int(func_get_args());
    return DB::doQuery();
  }
  
  function mysql_getsingle() {
    DB::sql_int(func_get_args());
    return DB::getsingle();
  }
  
  function mysql_getall() {
    DB::sql_int(func_get_args());
    return DB::getlist();
  }
  
  function mysql_insertinto($tablename, $data, $ignore="") {
    DB::sql_insertinto($tablename, $data, $ignore);
    return DB::doQuery();
  }
  
  function mysql_update($tablename, $where, $data, $apo = "'") {
  //trace("mysqlupdate!!!!!!!!!!!!!!!!!!!!!");
    DB::sql_update($tablename, $where, $data, $apo);
    return DB::doQuery();
  }
  
  function mysql_update2($tablename, $where, $sqlparts) {
    DB::sql_update2($tablename, $where, $sqlparts);
    return DB::doQuery();
  }
?>
