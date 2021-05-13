<?php  
	include ('koneksi.php');
	require '../../reportexcel/vendor/autoload.php';//open library
	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

	$spreadsheet = new Spreadsheet();
	$sheet = $spreadsheet->getActiveSheet();
	$sheet->setCellValue('A1',	'No');
	$sheet->setCellValue('B1',	'Jenis Pendaftaran');
	$sheet->setCellValue('C1',	'Tanggal Masuk Sekolah');
	$sheet->setCellValue('D1',	'NIS');
	$sheet->setCellValue('E1',	'Nomer Peserta');
	$sheet->setCellValue('F1',	'PAUD');
	$sheet->setCellValue('G1',	'TK');
	$sheet->setCellValue('H1',	'SKHUN');
	$sheet->setCellValue('I1',	'IJAZAH');
	$sheet->setCellValue('J1',	'Hobi');
	$sheet->setCellValue('K1',	'Cita - cita');
	$sheet->setCellValue('L1',	'Nama Lengkap');
	$sheet->setCellValue('M1',	'Jenis Kelamin');
	$sheet->setCellValue('N1',	'NISN');
	$sheet->setCellValue('O1',	'NIK');
	$sheet->setCellValue('P1',	'Tempat Lahir');
	$sheet->setCellValue('Q1',	'Tanggal Lahir');
	$sheet->setCellValue('R1',	'Agama');
	$sheet->setCellValue('S1',	'Berkebutuhan Khusus');
	$sheet->setCellValue('T1',	'Alamat');
	$sheet->setCellValue('U1',	'RT');
	$sheet->setCellValue('V1',	'RW');
	$sheet->setCellValue('W1',	'Nama Dusun');
	$sheet->setCellValue('X1',	'Nama Desa/Kelurahan');
	$sheet->setCellValue('Y1',	'Kecamatan');
	$sheet->setCellValue('Z1',	'Kode Pos');
	$sheet->setCellValue('AA1',	'Tinggal');
	$sheet->setCellValue('AB1',	'Transportasi');
	$sheet->setCellValue('AC1',	'No HP');
	$sheet->setCellValue('AD1',	'No Telp');
	$sheet->setCellValue('AE1',	'Email');
	$sheet->setCellValue('AF1',	'Penerima KIP');
	$sheet->setCellValue('AG1',	'NO KIP');
	$sheet->setCellValue('AH1',	'Kewarganegaraan');

	$query = mysqli_query($koneksi, "SELECT * FROM pendaftaran"); // select data dari database
	$i = 2;
	$no = 1;
	while($row = mysqli_fetch_array($query)){
		$sheet->setCellValue('A'.$i, $no++);
		$sheet->setCellValue('B'.$i, $row['JENIS_PENDAFTARAN']);
		$sheet->setCellValue('C'.$i, $row['TANGGAL_MASUK']);
		$sheet->setCellValue('D'.$i, $row['NIS']);
		$sheet->setCellValue('E'.$i, $row['NOMOR_PESERTA']);
		$sheet->setCellValue('F'.$i, $row['PAUD']);
		$sheet->setCellValue('G'.$i, $row['TK']);
		$sheet->setCellValue('H'.$i, $row['NO_SKHUN']);
		$sheet->setCellValue('I'.$i, $row['NO_IJAZAH']);
		$sheet->setCellValue('J'.$i, $row['HOBI']);
		$sheet->setCellValue('K'.$i, $row['CITA_CITA']);
		$sheet->setCellValue('L'.$i, $row['NAMA']);
		$sheet->setCellValue('M'.$i, $row['JENIS_KELAMIN']);
		$sheet->setCellValue('N'.$i, $row['NISN']);
		$sheet->setCellValue('O'.$i, $row['NIK']);
		$sheet->setCellValue('P'.$i, $row['TEMPAT_LAHIR']);
		$sheet->setCellValue('Q'.$i, $row['TANGGAL_LAHIR']);
		$sheet->setCellValue('R'.$i, $row['AGAMA']);
		$sheet->setCellValue('S'.$i, $row['BERKEBUTUHAN_KHUSUS']);
		$sheet->setCellValue('T'.$i, $row['ALAMAT']);
		$sheet->setCellValue('U'.$i, $row['RT']);
		$sheet->setCellValue('V'.$i, $row['RW']);
		$sheet->setCellValue('W'.$i, $row['DUSUN']);
		$sheet->setCellValue('X'.$i, $row['KELURAHAN']);
		$sheet->setCellValue('Y'.$i, $row['KECAMATAN']);
		$sheet->setCellValue('Z'.$i, $row['KODE_POS']);
		$sheet->setCellValue('AA'.$i, $row['TEMPAT_TINGGAL']);
		$sheet->setCellValue('AB'.$i, $row['TRANSPORTASI']);
		$sheet->setCellValue('AC'.$i, $row['NO_HP']);
		$sheet->setCellValue('AD'.$i, $row['NO_TELP']);
		$sheet->setCellValue('AE'.$i, $row['EMAIL']);
		$sheet->setCellValue('AF'.$i, $row['pPENERIMA_KPS']);
		$sheet->setCellValue('AG'.$i, $row['NO_KPS']);
		$sheet->setCellValue('AH'.$i, $row['KEWARGANEGARAAN']);
		$i++;
	}

	$styleArray = [
		'borders' => [
			'allBorders' => [
				'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
			],
		],
	];
	$i = $i - 1;
	$sheet->getStyle('A1:AK'.$i)->applyFromArray($styleArray);

	$writer = new Xlsx($spreadsheet);
	$writer->save('Report Data Siswa.xlsx'); //menyimpan file dengan nama Report Data Siswa.xlsx
	header("location:formpendaftaran.php");
	alert("Tersimpan")
?>