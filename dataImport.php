<?php
error_reporting(E_ALL ^ E_NOTICE);
require_once 'excel_reader2.php';
require_once('config.php');
$file = "example.xls";
$tableName = "tablename";
$con = mysql_connect(DB_HOST, DB_USER, DB_PASSWORD);
if (!$con)
{
    die('Could not connect: ' . mysql_error());
}
mysql_select_db(DB_DATABASE, $con);




function importToNewTable($file,$tableName){
    $data = new Spreadsheet_Excel_Reader($file);
    $colcount = $data->colcount()."<br />";
    $rowcount = $data->rowcount()."<br />";

    $longTexts = '';

    for($rowno=2;$rowno<$rowcount;$rowno++){
        for($colno=1;$colno<$colcount;$colno++){
            if (strlen($data->value($rowno,$colno))>250){
                $longTexts = $longTexts."{$colno},";
            }
        }
    }

    $longTexts = array_values(array_unique(explode(',',substr($longTexts,0,-1))));

    mysql_query("DROP TABLE `{$tableName}`");
    $sql = "CREATE TABLE `{$tableName}` (`line_id` INT(10) NOT NULL AUTO_INCREMENT";
    for($colno=1;$colno<=$colcount;$colno++){
        $cell = $data->value(1,$colno);
        $colsArray[$colno] = $cell;
        if(in_array($colno,$longTexts)){
            $sql = $sql.", `{$cell}` LONGTEXT NOT NULL";
        }
        else{
            $sql = $sql.", `{$cell}` VARCHAR(255) NOT NULL";
        }
    }
    $sql = $sql.", PRIMARY KEY (`line_id`))";
    mysql_query($sql);

    

    $actualInserts=0;
    for($rowno=2;$rowno<=$rowcount;$rowno++){
        
        $sql = "INSERT INTO `{$tableName}` (";
        for($no=1;$no<=count($colsArray);$no++){
            $sql = $sql."`{$colsArray[$no]}`,";
        }
        $sql = substr($sql,0,-1).") VALUES (";

        for($colno=1;$colno<=$colcount;$colno++){
            $sql = $sql."'".mysql_real_escape_string($data->value($rowno,$colno))."',";
        }
        $sql = substr($sql,0,-1).")";
        if(!mysql_query($sql)){
            echo mysql_error();
        }
        else{
            $actualInserts++;
        }
    }
    echo "<br />".$actualInserts." rows inserted";
}



function importToExistingTable($file,$tableName){
    $data = new Spreadsheet_Excel_Reader($file);
    $colcount = $data->colcount();
    $rowcount = $data->rowcount();

    $sql = mysql_query("SELECT * FROM `{$tableName}`");
    for($no=0;$no<=$colcount;$no++){
        $tableColsArray[$no] = mysql_field_name($sql,$no);
    }
    echo "<br />";
    print_r($tableColsArray);
    echo "<br />";
    for($colno=1;$colno<=$colcount;$colno++){
        $cell = $data->value(1,$colno);
        if($tableColsArray[$colno]!=$cell){
            die("Column $colno does not match its counterpart in the database");
        }
        else{
            echo "spreadsheet part ($cell) perfectly matches db part ($tableColsArray[$colno])<br />";
        }
    }
    
    
    $actualInserts=0;
    for($rowno=2;$rowno<=$rowcount;$rowno++){
        
        $sql = "INSERT INTO `{$tableName}` (";
        for($no=1;$no<count($tableColsArray);$no++){
            $sql = $sql."`{$tableColsArray[$no]}`,";
        }
        $sql = substr($sql,0,-1).") VALUES (";

        for($colno=1;$colno<=$colcount;$colno++){
            $sql = $sql."'".mysql_real_escape_string($data->value($rowno,$colno))."',";
        }
        $sql = substr($sql,0,-1).")";
        if(!mysql_query($sql)){
            echo mysql_error();
        }
        else{
            $actualInserts++;
        }
    }
    echo "<br />".$actualInserts." rows inserted";

}


importToExistingTable($file,$tableName);
?>
