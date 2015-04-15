<?php
require_once "define.inc";
if($_GET["dl"]) {
//http://mw.teamwiki.net/docs/img/mwupd3-download.png
$data=parse_ini_file(WIKI_DIR."/2/webspace/download/mwupd3/version.txt");
$url=$data["DownloadURL"];
header("Location: $url");

//echo "<a href='$url'><img src='http://mw.teamwiki.net/docs/img/mwupd3-download.png' title='Neueste Version herunterladen' alt='MWupd3 - Jetzt downloaden' border='0' /></a>";
} else {
readfile(WIKI_DIR."/2/webspace/download/mwupd3/version.txt");
}
?>