<meta charset="UTF-8">
<style>
* {
	font-size:12px
}
@media print {
    footer {page-break-after: always;}
}
@media print{@page {size: landscape}}
@media print
{    
    .no-print, .no-print *
    {
        display: none !important;
    }
}
   table { page-break-inside:auto;
	border: 1px solid black;
border-collapse: collapse;
}
   tr    { page-break-inside:avoid; page-break-after:auto }
</style>
<div class="no-print"><form action="" method="post" enctype="multipart/form-data">
    Select XLS to upload:
    <input type="file" name="fileToUpload" id="fileToUpload">
    <input type="submit" value="Upload XLS" name="submit">
</form>

<?php
require_once 'reader.php';

$target_dir = "uploads/";
$target_file = $target_dir . basename($_FILES["fileToUpload"]["name"]);
$uploadOk = 1;
$imageFileType = strtolower(pathinfo($target_file,PATHINFO_EXTENSION));
// Check if image file is a actual image or fake image
if(isset($_POST["submit"])) {
// Check if file already exists
if (file_exists($target_file)) {
    echo "Sorry, file already exists.";
    $uploadOk = 0;
}
// Check file size
if ($_FILES["fileToUpload"]["size"] > 500000) {
    echo "Sorry, your file is too large.";
    $uploadOk = 0;
}
// Check if $uploadOk is set to 0 by an error
if ($uploadOk == 0) {
    echo "Sorry, your file was not uploaded.";
// if everything is ok, try to upload file
} else {
    if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], "temp.xls")) {
        echo "The file ". basename( $_FILES["fileToUpload"]["name"]). " has been uploaded.";
    } else {
        echo "Sorry, there was an error uploading your file.";
    }
}
echo "</div>"; //No Print

$data = new Spreadsheet_Excel_Reader();
//$data->setOutputEncoding('CPa25a');
$data->read("temp.xls");
error_reporting(E_ALL ^ E_NOTICE);
$newpage =1;
$footer=0;
$summa=0;
$u=0;
$n=0;
/*echo "<table width=100% border=1>
<tr>
<th>'Orgnr<th>Avdeling<th>Gruppnamn<th>Regno<th>Fabrikat<th>From datum<th>Til datum<th>Belopp
<tr>
";*/
$orgnr2="";
for ($i = 2; $i <= $data->sheets[0]['numRows']; $i++) {
	if ($data->sheets[0]['cells'][$i][1] <> $orgnr2){
		$orgnr2 =  $data->sheets[0]['cells'][$i][1];
                if ($n == 1){
		echo "
		<tr><td colspan=6><td><b>Summa<td><b>$summa
		</table><footer></footer>
		";
                }
		$summa=0;
                $u++;
		$newpage = 1;
		$footer=0;
		echo "<table width=100% border=1>
		    <tr>
		        <th>Orgnr<th>Avdeling<th>Gruppnamn<th>Regno<th>Fabrikat<th>From datum<th>Til datum<th>Belopp
		    </tr>
		";
		$newpage = 0;
	}
	$orgnr =  $data->sheets[0]['cells'][$i][1];
	$avdeling =  utf8_encode($data->sheets[0]['cells'][$i][2]);
	$gruppnamn =  utf8_encode($data->sheets[0]['cells'][$i][5]);
	$regno =  $data->sheets[0]['cells'][$i][6];
	$fabrikat =  utf8_encode($data->sheets[0]['cells'][$i][8]);
	$from_datum =  $data->sheets[0]['cells'][$i][10];
	$UNIX_DATE = ($from_datum - 25569) * 86400;
	$from_datum = gmdate("Y-m-d", $UNIX_DATE);
	$till_datum =  $data->sheets[0]['cells'][$i][11];
	$UNIX_DATE = ($till_datum - 25569) * 86400;
	$till_datum = gmdate("Y-m-d", $UNIX_DATE);

	$belopp =  $data->sheets[0]['cells'][$i][22];
//	$belopp = round($belopp);
	$summa += $belopp;
	echo "
		<tr>
			<td width=100><nobr>$orgnr</nobr>
			<td width=450>$avdeling
			<td><nobr>$gruppnamn
			<td width=100>$regno
			<td width=200>$fabrikat
			<td width=100>$from_datum
			<td width=100>$till_datum
			<td width=100>$belopp
		</tr>
";
$n =1;
}
                echo "
                <tr><td colspan=6><td><b>Summa<td><b>$summa
                </table><footer></footer>
                ";

}
unlink("temp.xls");
?>
