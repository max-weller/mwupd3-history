<?php
  //header("Content-Type: text/plain");
  require_once 'server_part/define.inc';
  require_once 'server_part/trace.inc';
//  require_once 'usermanagement.inc';
  require_once 'server_part/mysql.inc';
//  require_once 'funcs.inc';
  DB::set_config($SITE_CONFIG['MYSQL_HOST'], $SITE_CONFIG['MYSQL_USER'], $SITE_CONFIG['MYSQL_PASSWORD'], $SITE_CONFIG['MYSQL_DATABASE']);

  DB::connect();
  
  $folder = "./server_part/files/";
  

$files= mysql_getall("SELECT * FROM wupd_files");
foreach($files as $f)  {
$app = mysql_getone("SELECT * FROM wupd_apps WHERE id = ".$f["AppID"]);
$cat = mysql_getone("SELECT * FROM wupd_categories WHERE Catid = ".$app["CategoryID"]);
$c = preg_replace("/[^a-zA-Z0-9]+/","-",$cat["Caption"]);
echo "mkdir -p apps/".$c."/".$app["AppName"]."/\n";
echo "cp $folder$f[FileID].dat apps/".$c."/".$app["AppName"]."/".$f["FileName"];
echo "\n";

}



