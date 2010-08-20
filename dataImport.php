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


    //This works out which columns contain cells with more than 250 characters in them
    //so that we can make the field type LONGTEXT instead of VARCHAR

    for($rowno=2;$rowno<$rowcount;$rowno++){
        for($colno=1;$colno<$colcount;$colno++){
            if (strlen($data->value($rowno,$colno))>250){
                $longTexts = $longTexts."{$colno},";
            }
        }
    }
    //now we put these into an array for use later:
    $longTexts = array_values(array_unique(explode(',',substr($longTexts,0,-1))));

    //First we drop any table with the same name, then we build the create table statement by going through the column heads in the
    //sheet again:
    mysql_query("DROP TABLE `{$tableName}`");
    $sql = "CREATE TABLE `{$tableName}` (`line_id` INT(10) NOT NULL AUTO_INCREMENT";
    for($colno=1;$colno<=$colcount;$colno++){
        $cell = $data->value(1,$colno);
        $colsArray[$colno] = $cell;
        //This checks to see if this particular column has been previously identified as one that should be LONGTEXT rather than VARCHAR:
        if(in_array($colno,$longTexts)){
            $sql = $sql.", `{$cell}` LONGTEXT NOT NULL";
        }
        else{
            $sql = $sql.", `{$cell}` VARCHAR(255) NOT NULL";
        }
    }
    $sql = $sql.", PRIMARY KEY (`line_id`))";

    //Then we execute the create table:
    if(!mysql_query($sql)){
        echo "Could not create the table: ".mysql_error();
    }

    
    //This goes through the whole sheet building the insert statement from the data:
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

        //This executes the insert and counts the successful inserts:
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

    //This gets the field names in the database so that they can be confirmed as matching those in the spreadsheet:
    $sql = mysql_query("SELECT * FROM `{$tableName}`");
    for($no=0;$no<=$colcount;$no++){
        $tableColsArray[$no] = mysql_field_name($sql,$no);
    }

    //This goes through the spreadsheet and compares the columns with those in the database:
    for($colno=1;$colno<=$colcount;$colno++){
        $cell = $data->value(1,$colno);
        if($tableColsArray[$colno]!=$cell){
            die("Column $colno does not match its counterpart in the database");
        }
        else{
            echo "spreadsheet part ($cell) perfectly matches db part ($tableColsArray[$colno])<br />";
        }
    }
    
    //This goes through the whole sheet building the insert statement from the data:
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
        //This executes the insert and counts the successful inserts:
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
