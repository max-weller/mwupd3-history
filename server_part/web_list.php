<?php
  //header("Content-Type: text/plain");
  require_once 'define.inc';
  require_once 'trace.inc';
  require_once 'usermanagement.inc';
  require_once 'mysql.inc';
  require_once 'funcs.inc';
  DB::connect();
  
  function onlyAZ09($str) {
    return preg_replace("/[^a-z0-9A-Z_.-]/", "", $str);
  }
  
  function sendErrorResult($err) {
    header("X-MWupd3-Result: ERROR, $err");
    header("X-MWupd3-Error: $err");
    die("$err\r\n\r\n");
  }
  $folder = ROOT."/webspace/mwupd3.teamwiki.net/files/";
  
  
?><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>MWupd3 Application List</title>
<style>
body,td { font: 10pt sans-serif; }
</style>
</head>

<body>
  
  
  <?php
  //--> list_cats_and_apps
    
    $cats = mysql_getall("SELECT * FROM wupd_categories");
    $apps = mysql_getall("SELECT * FROM wupd_apps");
    
    foreach($cats as $cat) {
      echo "<img src='http://mw.teamwiki.net/docs/img/win-icons/SHELL32_4-16.png'> <b>$cat[Caption]</b><br>";
      foreach($apps as $app) {
        if ($app["CategoryID"]==$cat["CatID"]) {
          echo "&nbsp; &nbsp; ";
          if ($app[SmallIconURL]=="") $app[SmallIconURL]="http://mw.teamwiki.net/docs/img/win-icons/rundll32_100-16.png";
          echo "<img src='$app[SmallIconURL]' /> ";
          echo "$app[AppTitle]<br>";
          
        }
      }
    }
    
  
  
?>
</body>
</html>