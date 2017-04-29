<?php
// Include Composer autoloader if not already done.
include 'vendor/autoload.php';
include 'class/Converter.php';

/**
 * File:  convert.php
 * Author: Sharan Girdhani
 */
error_reporting(0);

echo "\nLoading given template\n";
$obj_nt = Converter::loadBasicTemplate();
$obj_wt = Converter::loadBasicTemplate();

// Parse pdf file and build necessary objects.
$parser = new \Smalot\PdfParser\Parser();


// Taking all the files presented in the given folder
echo "\nTaking the input files: \n";
$files = scandir('InputFiles/');

$i = 1;
foreach($files as $input_file) 
{
    // if the file is not pdf then ignoring it
    if(!strpos($input_file, ".pdf"))
        continue;
    
    echo "\nProcessing file ".$input_file."\n\n\n";
    // Processing the pdf file
    $pdf = $parser->parseFile('InputFiles/'.$input_file);

    // Retrieve all pages from the pdf file.
    $pages  = $pdf->getPages();
    
    $i++;
    // Loop over each page to extract text and add to the required xlsx file.
    $data_exp = [];
    foreach ($pages as $page) 
    {
        $data = $page->getText();
        $data_exp = array_merge($data_exp, explode('|', $data));
    }
    $isSpecial = false;
    //echo $data_exp[9][0]; die;
    if($data_exp[9][0] != 'P')
    {
        $temp = array_slice($data_exp, 0, 9);
        $temp[9] = "Special Case";
        $data_exp = array_merge($temp, array_slice($data_exp, 9));
        
        $temp = array_slice($data_exp, 0, 74);
        $temp = array_merge($temp, array_slice($data_exp, 77, 64));
        $temp = array_merge($temp, [0 => "Special Case"]);
        $temp = array_merge($temp, array_slice($data_exp, 74, 3));
        $temp = array_merge($temp, array_slice($data_exp, 145, 6));
        $temp = array_merge($temp, array_slice($data_exp, 141, 3));
        $data_exp = array_merge($temp, array_slice($data_exp, 151, 3));
        $isSpecial = true;
    }
    // echo "<pre>";
    // print_r($data_exp);
    // die;

    foreach($data_exp as &$str)
    {
        $str = str_replace(",", "", $str);
    }
    // Adding the data to the output file
    $obj_nt = Converter::addNonTobaccoData($data_exp, $obj_nt, $i, $isSpecial);
    $obj_wt = Converter::addTobaccoData($data_exp, $obj_wt, $i, $isSpecial);
}

$filename= "Output/output_non_tobacco.xlsx";

$objWriter = PHPExcel_IOFactory::createWriter($obj_nt, 'Excel2007');
$objWriter->save($filename);


$filename= "Output/output_with_tobacco.xlsx";

$objWriter = PHPExcel_IOFactory::createWriter($obj_wt, 'Excel2007');
$objWriter->save($filename);
 
echo "Process Completed\n";

echo "\nPlease check the output.xlsx file placed in the output folder.\n\nThank You\n";
?>

