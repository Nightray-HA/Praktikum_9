<?php
	//konfigurasi database
	$host 		= "localhost";
	$user 		= "root";
	$password 	= "";
	$database 	= "db_siswa";

	// Perintah php untuk akses ke database
	$koneksi = mysqli_connect($host, $user, $password, $database);
	if(!$koneksi){
		die("Koneksi gagal : " . msqli_connect_error());
	}
 
?>