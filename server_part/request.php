<?php
  //header("Content-Type: text/plain");
  require_once 'define.inc';
  require_once 'trace.inc';
//  require_once 'usermanagement.inc';
  require_once 'mysql.inc';
//  require_once 'funcs.inc';
  DB::set_config($SITE_CONFIG['MYSQL_HOST'], $SITE_CONFIG['MYSQL_USER'], $SITE_CONFIG['MYSQL_PASSWORD'], $SITE_CONFIG['MYSQL_DATABASE']);

  DB::connect();
  
  $folder = SRV_ROOT."/vhosts/mwupd3.teamwiki.net/files/";
  
//  trace_r(getallheaders  (),'allheader');
  trace_r(file_get_contents("php://input"),"php://stdin");
  trace_r($_COOKIE);
  trace_r($_POST,"Post");
  trace_r($_GET,"Get");
  trace_r($_SESSION,"Session");
  
  if ($_GET["c"] == "get_updater_version") {
    header("Content-Type: text/plain");
    readfile(WIKI_DIR."/2/webspace/download/mwupd3/version.txt");
    exit;
  }
  
  //--> upload_file
  if ($_GET["c"] == "upload_file") {
    header("X-MWupd3-Ver: MWUPD TeamWiki Software Update Service Version 3.0 ".getLoginUser()."X");
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    if (loginOK() == false || !is_permission("MWupd3")) sendErrorResult("Bitte einloggen");
    
    
    $app_id = intval($_POST["appid"]);
    $appInfo = mysql_getone("SELECT * FROM wupd_apps WHERE ID = $app_id");
    if (!$appInfo) sendErrorResult("Ungültige AppID");
    
    $fileName = onlyAZ09($_POST["fileName"]);
    $fileInfo = mysql_getone("SELECT FileID,UploadCount FROM wupd_files WHERE AppID = $app_id AND FileName = '$fileName'");
    if (!$fileInfo) {
      $sql = array("AppID" => $app_id, "FileName" => "'$fileName'",
                   "EinDat" => "NOW()", "EinDIZ" => "'".getLoginUser()."'","AendDIZ" => "'".getLoginUser()."'");
      mysql_insertinto("wupd_files", $sql);
      $fileInfo = mysql_getone("SELECT FileID,UploadCount FROM wupd_files WHERE AppID = $app_id AND FileName = '$fileName'");
    }
    
    if (is_uploaded_file($_FILES['fileContent']['tmp_name'])) {
      $targetFilespec = $folder.$fileInfo["FileID"].".dat";
      $ok = move_uploaded_file($_FILES['fileContent']['tmp_name'], $targetFilespec);
      if (file_exists($folder.$fileInfo["FileID"].".dat")) {
        $sql2 = array("AendDIZ" => "'".getLoginUser()."'", "AendDat" => "NOW()", "UploadCount" => $fileInfo['UploadCount']+1);
        if (isset($_POST["localSource"])) $sql2["LocalSource"] = "'".mysql_escape_string($_POST["localSource"])."'";
        if (isset($_POST["localTarget"])) $sql2["LocalTarget"] = "'".mysql_escape_string($_POST["localTarget"])."'";
        if (isset($_POST["flags"])) $sql2["Flags"] = "'".mysql_escape_string($_POST["flags"])."'";
        mysql_update("wupd_files", " FileID = $fileInfo[FileID] ", $sql2, "");
        header("X-MWupd3-Result: OK, $fileInfo[FileID], ".date("Y-m-d H:i:s").", ".getLoginUser());
        echo "\r\n\r\nOK\r\n$fileInfo[FileID]";
      } else {
        sendErrorResult("File not Uploaded");
      }
    } else {
      sendErrorResult("is_uploaded_file returned false");
    }
    
  }
  
  
  //--> remove_file
  if ($_GET["c"] == "remove_file") {
    header("X-MWupd3-Ver: MWUPD TeamWiki Software Update Service Version 3.0 ".getLoginUser()."X");
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    if (loginOK() == false || !is_permission("MWupd3")) sendErrorResult("Bitte einloggen");
    
    
    $app_id = intval($_POST["appid"]);
    $file_id = intval($_POST["fileid"]);
    
    $ok2=false;
    $ok1=mysql_query("DELETE FROM wupd_files WHERE AppID = $app_id AND FileID = $file_id");
    if (mysql_affected_rows() == 1) {
      $ok2=unlink($folder.$file_id.".dat");
      echo "\r\n\r\nOK - ";var_dump($ok1,$ok2);
    } else {
      echo "ERROR\r\n\r\n";
    }
  }
  
  
  //--> create_app
  if ($_GET["c"] == "create_app") {
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    if (loginOK() == false || !is_permission("MWupd3")) die("ERR: Bitte einloggen\r\n\r\n");
    
    $appname = strtolower(onlyAZ09($_POST["app_name"]));
    if (!$appname) die("ERR: Ungültiger app_name\r\n\r\n");
    
    $appInfo = mysql_getone("SELECT * FROM wupd_apps WHERE AppName = '$appname'");
    if ($appInfo) die("Der Anwendungsname $appname existiert schon!\r\n\r\n$appInfo[ID]");
    
    $sql = array("AppName" => "'".mysql_escape_string($appname)."'",
                 "AppTitle" => "'".mysql_escape_string($appname)."'",
                 "EinDat" => "NOW()", "EinDIZ" => "'".getLoginUser()."'",
                 "AendDat" => "NOW()", "AendDIZ" => "'".getLoginUser()."'");
    mysql_insertinto("wupd_apps", $sql);
    echo "\r\n\r\n". mysql_insert_id();
  }
  
  
  //--> update_app
  if ($_GET["c"] == "update_app") {
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    if (loginOK() == false || !is_permission("MWupd3")) die("Bitte einloggen\r\n\r\n");
    
    $appname = onlyAZ09($_GET["app_name"]);
    $sqlParts= array("AppTitle" => $_POST["apptitle"], "Desc" => $_POST["desc"],
                     "BigIconURL" => $_POST["bigicon"], "SmallIconURL" => $_POST["smallicon"],
                     "CategoryID" => intval($_POST["catid"]),
                     "DefaultInstallDir" => mysql_escape_string($_POST["DefaultInstallDir"]),
                     "HomepageURL" => mysql_escape_string($_POST["HomepageURL"]),
                     "VersionInfoURL" => mysql_escape_string($_POST["VersionInfoURL"]));
    
    $ok=mysql_update("wupd_apps", " ID = ".intval($_POST["appid"]), $sqlParts);
    if (!$ok) {
      echo "ERR: ".mysql_error()."\r\n\r\n";
    } else {
      echo "\r\n\r\nOK";
    }
  }
  
  //--> save_version_info
  if ($_GET["c"] == "save_version_info") {
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    if (loginOK() == false || !is_permission("MWupd3")) die("Bitte einloggen\r\n\r\n");
    
    $ok=file_put_contents("files/ver_".intval($_POST["appid"]).".txt", $_POST["content"]);
    if (!$ok) {
      echo "\r\n\r\n";
    } else {
      echo "\r\n\r\nOK";
    }
  }
  
  
  //--> get_version_info
  if ($_GET["c"] == "get_version_info") {
    echo '<html><head><title>Versionsinformationen</title>';
    echo '<style> html,body,td {font: 9pt "Courier New",monospace;margin:0;padding:0;} body {padding: 5px; } p.footer{margin:20px 0 0;padding: 10px;font:status-bar;background:#ccc;}</style>';
    echo '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head><body>';
    if (isset($_GET["nourl"])) die("<p class='footer'>Zu dieser Anwendung sind keine Versionshinweise hinterlegt.</p></body></html>");
    
    $info=mysql_getone("SELECT AppTitle FROM wupd_apps WHERE ID = ".intval($_GET["appid"]));
    if(!$info) die("<p class='footer'>Anwendungs-ID nicht gefunden.</p></body></html>");
    
    $cont=file_get_contents("files/ver_".intval($_GET["appid"]).".txt");
    
    echo nl2br($cont);
    echo "\r\n\r\n<p class='footer'>Versionsinformationen zu $info[AppTitle]<br />powered by Teamwiki.net</p></body></html>";

  }
  
  //--> list_apps
  if ($_GET["c"] == "list_apps") {
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n\r\n\r\n";
    
    $where = "";
    if ($_GET["catid"]) $where = " WHERE CategoryID = ".intval($_GET["catid"]);
    $apps = mysql_getall("SELECT * FROM wupd_apps $where ");
    foreach ($apps as $d) {
      foreach ($d as $v) echo "$v\t";
      echo "\r\n";
    }
    
  }
  
  
  //--> list_cats_and_apps
  if ($_GET["c"] == "list_cats_and_apps") {
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    
    $fileRes = mysql_query("SELECT * FROM wupd_categories");
    for ($i=0; $i < mysql_num_fields($fileRes); $i++)
      $out.= mysql_field_name($fileRes,$i)."\t";
    $out.="\r\n";
    while ($data=mysql_fetch_row($fileRes)) {
      foreach($data as $d) $out.= "$d\t";
      $out .= "\r\n";
    }
    mysql_free_result($fileRes);
    
    $out .= "\r\n";
    
    $fileRes = mysql_query("SELECT * FROM wupd_apps");
    for ($i=0; $i < mysql_num_fields($fileRes); $i++)
      $out.= mysql_field_name($fileRes,$i)."\t";
    $out.="\r\n";
    while ($data=mysql_fetch_row($fileRes)) {
      foreach($data as $d) $out.= "$d\t";
      $out .= "\r\n";
    }
    mysql_free_result($fileRes);
    echo "\r\n\r\n$out";
  }
  
  
  //--> get_true_app_info
  if ($_GET["c"] == "get_true_app_info") {
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    $app_id = intval($_GET["appid"]);
    $appInfo = mysql_getone("SELECT * FROM wupd_apps WHERE ID = $app_id");
    
    
    $out="";
    $out.= "\r\n\r\n";
    foreach($appInfo as $k=>$v) {
      $out.=str_pad($k,18," ",STR_PAD_LEFT).": ".$v."\r\n";
    }
    $out.=str_pad("RootURL",18," ",STR_PAD_LEFT).": "."http://mwupd3.teamwiki.net/files/"."\r\n";
    $out.="\r\n";
    
    $fileRes = mysql_query("SELECT * FROM wupd_files WHERE AppID = $app_id");
    for ($i=0; $i < mysql_num_fields($fileRes); $i++)
      $out.= mysql_field_name($fileRes,$i)."\t";
    $out.="\r\n";
    
    
    while ($data=mysql_fetch_row($fileRes)) {
      foreach($data as $d) $out.= "$d\t";
      $out .= "\r\n";
    }
    mysql_free_result($fileRes);
    
    echo $out;
  }
  
  
  //--> get_app_info
  if ($_GET["c"] == "get_app_info") {
    echo "MWUPD TeamWiki Software Update Service Version 3.0\r\n";
    $app_id = intval($_GET["appid"]);
    $appInfo = mysql_getone("SELECT * FROM wupd_apps WHERE ID = $app_id");
    
    
    $out="";
    $out.= "\r\n\r\n";
    foreach($appInfo as $k=>$v) {
      $out.=str_pad($k,18," ",STR_PAD_LEFT).": ".$v."\r\n";
    }
    $out.=str_pad("RootURL",18," ",STR_PAD_LEFT).": "."http://mwupd3.teamwiki.net/files/"."\r\n";
    $out.="\r\n";
    
    //DefaultApp ausgeben
    $fileRes = mysql_query("SELECT * FROM wupd_files WHERE AppID = 0");
    for ($i=0; $i < mysql_num_fields($fileRes); $i++)
      $out.= mysql_field_name($fileRes,$i)."\t";
    $out.="\r\n";
    
    while ($data=mysql_fetch_row($fileRes)) {
      foreach($data as $d) $out.= "$d\t";
      $out .= "\r\n";
    }
    mysql_free_result($fileRes);
    
    //Eigentliche Anwendung ausgeben
    $fileRes = mysql_query("SELECT * FROM wupd_files WHERE AppID = $app_id");
    while ($data=mysql_fetch_row($fileRes)) {
      foreach($data as $d) $out.= "$d\t";
      $out .= "\r\n";
    }
    mysql_free_result($fileRes);
    
    echo $out;
  }
  
  
  
  function onlyAZ09($str) {
    return preg_replace("/[^a-z0-9A-Z_.-]/", "", $str);
  }
  
  function sendErrorResult($err) {
    header("X-MWupd3-Result: ERROR, $err");
    header("X-MWupd3-Error: $err");
    die("$err\r\n\r\n");
  }
  
?>
