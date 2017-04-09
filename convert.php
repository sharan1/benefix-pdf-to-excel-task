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
$obj = Converter::loadBasicTemplate();

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
    
    // Loop over each page to extract text and add to the required xlsx file.
    foreach ($pages as $page) 
    {
        $i++;
        $data = $page->getText();

        $data_exp = explode('|', $data);
        $str = $data_exp[1];
        $str = str_replace("Valid for Effective Dates: ","",$str);

        // Adding the data to the output file
        $obj = Converter::addData($data_exp, $obj, $i);
    }
}

$filename= "Output/output.xlsx";

$objWriter = PHPExcel_IOFactory::createWriter($obj, 'Excel2007');
$objWriter->save($filename);
 
echo "Process Completed\n";

echo "\nPlease check the output.xlsx file placed in the output folder.\n\nThank You\n";
?>

