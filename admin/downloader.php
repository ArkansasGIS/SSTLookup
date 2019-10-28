<?php
$filename = $_GET['filename'];
$target_dir = ini_get('upload_tmp_dir');
header('Content-Type: application/force-download');
header("Content-disposition: attachment; filename=$filename");
header("Content-Transfer-Encoding: binary");
header("Cache-Control: no-cache, must-revalidate"); //The page must not be stored in the cache 
header("Expires: Mon, 01 Jan 2000 00:00:00 GMT");

  if ($fp = fopen($target_dir ."\\".$filename, 'rb')) {
   while(!feof($fp) and (connection_status()==0)) {
     echo(fread($fp, 1024*8));
     flush();
   }
   fclose($fp);
  }
  unlink($target_dir ."\\".$filename);

?>