<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="pt-br">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>PHP Excel Reader</title>

</head>
<body>

<h1>PHP Excel Reader</h1>
<p>
	Componente em PHP para leitura de arquivos excel. 
	Para mais detalhes veja o código fonte.<br />
	Ou se preferir faça o 
	<a href="http://files.edysegura.com/labs/php-excel-reader.zip">downlaod</a> deste 
	<a href="http://files.edysegura.com/labs/php-excel-reader.zip">exemplo</a>.
</p>

<form action="<?php echo $_SERVER['PHP_SELF']; ?>" method="post" enctype="multipart/form-data">
	<fieldset>
		<legend>Importar Arquivos do Excel:</legend>
		<label>
			Arquivo do Excel:<br />
			<input type="file" name="upfile" />
		</label>
		<input type="submit" value="Importar" />
	</fieldset>
</form>

<p>&nbsp;</p>

<?php
require 'Excel/reader.php';

$data = new Spreadsheet_Excel_Reader();
$data->setOutputEncoding('UTF-8');

if(!empty($_FILES['upfile']) && $_FILES['upfile']['type'] == "application/vnd.ms-excel") {
	$data->read($_FILES['upfile']['tmp_name']);
}
else {
	$data->read('test.xls');
}

$totalplanilha = 0;
$policial = array();
$matriculaatual = '';

echo "<table border=\"1\">";
	for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
		echo "<tr>";
			for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
				$celldata = utf8_encode((!empty($data->sheets[0]['cells'][$i][$j])) ? $data->sheets[0]['cells'][$i][$j] : "&nbsp;");
				echo "<td>$celldata</td>";

				if ($j == 4){
					$matriculaatual = trim($celldata);
				}

				if ($j == 6){

					if (isset ($policial[$matriculaatual])){
						$policial[$matriculaatual] += $celldata ;
					}else{
						$policial[$matriculaatual] = $celldata;
					}

					$totalplanilha += $celldata;
				}
			}
		echo "</tr>";
	}


$totalagrupado = 0;
foreach ($policial as $k => $v) {
	echo "\$policial[$k] => $k.\n.\n.$v <p>";
	$totalagrupado += $v;
}


/*
$meuArray = array();

$meuArray['12345'] = 0;
$meuArray['67890'] = 1;
$meuArray['12345'] += 10;
$meuArray['12345'] += 10;

//print_r($meuArray);

foreach ($meuArray as $k => $v) {


	echo "\$a[$k] => $k.\n.\n.$v";

}
*/

echo $totalplanilha;
echo "<p>";
echo "total agrupado " + $totalagrupado;

echo "</table>";


?>

</body>
</html>