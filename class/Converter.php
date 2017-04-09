<?php

include 'vendor/autoload.php';

/**
 * Class:  Converter
 * Author: Sharan Girdhani
 */
class Converter
{
    // Creating a basic template
    public static function loadBasicTemplate()
    {
        $obj = PHPExcel_IOFactory::load("template.xlsx");
        $obj->setActiveSheetIndex(0);
        $objSheet = $obj->getActiveSheet();

        $objSheet->getStyle('A1:AZ1')->getFont()->setBold(true)->setSize(14);
        $objSheet->getColumnDimension('A')->setWidth(25);
        $objSheet->getColumnDimension('B')->setWidth(25);
        $objSheet->getColumnDimension('C')->setWidth(50);
        $objSheet->getColumnDimension('D')->setWidth(30);
        $objSheet->getColumnDimension('E')->setWidth(30);
        $objSheet->getColumnDimension('F')->setWidth(20);
        $objSheet->getColumnDimension('G')->setWidth(20);
        $objSheet->getColumnDimension('H')->setWidth(20);
        $objSheet->getColumnDimension('I')->setWidth(20);
        $objSheet->getColumnDimension('J')->setWidth(20);
        $objSheet->getColumnDimension('K')->setWidth(20);
        $objSheet->getColumnDimension('L')->setWidth(20);
        $objSheet->getColumnDimension('M')->setWidth(20);
        $objSheet->getColumnDimension('N')->setWidth(20);
        $objSheet->getColumnDimension('O')->setWidth(20);
        $objSheet->getColumnDimension('P')->setWidth(20);
        $objSheet->getColumnDimension('Q')->setWidth(20);
        $objSheet->getColumnDimension('R')->setWidth(20);
        $objSheet->getColumnDimension('S')->setWidth(20);
        $objSheet->getColumnDimension('T')->setWidth(20);
        $objSheet->getColumnDimension('U')->setWidth(20);
        $objSheet->getColumnDimension('V')->setWidth(20);
        $objSheet->getColumnDimension('W')->setWidth(20);
        $objSheet->getColumnDimension('X')->setWidth(20);
        $objSheet->getColumnDimension('Y')->setWidth(20);
        $objSheet->getColumnDimension('Z')->setWidth(20);
        $objSheet->getColumnDimension('AA')->setWidth(20);
        $objSheet->getColumnDimension('AB')->setWidth(20);
        $objSheet->getColumnDimension('AC')->setWidth(20);
        $objSheet->getColumnDimension('AD')->setWidth(20);
        $objSheet->getColumnDimension('AE')->setWidth(20);
        $objSheet->getColumnDimension('AF')->setWidth(20);
        $objSheet->getColumnDimension('AG')->setWidth(20);
        $objSheet->getColumnDimension('AH')->setWidth(20);
        $objSheet->getColumnDimension('AI')->setWidth(20);
        $objSheet->getColumnDimension('AJ')->setWidth(20);
        $objSheet->getColumnDimension('AK')->setWidth(20);
        $objSheet->getColumnDimension('AL')->setWidth(20);
        $objSheet->getColumnDimension('AM')->setWidth(20);
        $objSheet->getColumnDimension('AN')->setWidth(20);
        $objSheet->getColumnDimension('AO')->setWidth(20);
        $objSheet->getColumnDimension('AP')->setWidth(20);
        $objSheet->getColumnDimension('AQ')->setWidth(20);
        $objSheet->getColumnDimension('AR')->setWidth(20);
        $objSheet->getColumnDimension('AS')->setWidth(20);
        $objSheet->getColumnDimension('AT')->setWidth(20);
        $objSheet->getColumnDimension('AU')->setWidth(20);
        $objSheet->getColumnDimension('AV')->setWidth(20);
        $objSheet->getColumnDimension('AW')->setWidth(20);
        $objSheet->getColumnDimension('AX')->setWidth(20);
        $objSheet->getColumnDimension('AY')->setWidth(20);
        $objSheet->getColumnDimension('AZ')->setWidth(20);

        return $obj;
    }

    // Adding Data to the file
    public static function addData($data_exp, $obj, $i)
    {
        $str = $data_exp[1];
        $str = str_replace("Valid for Effective Dates: ","",$str);
        $arr = explode("-", $str);

        $start_date = trim($arr[0]);
        $end_date = trim($arr[1]);

        $obj->getActiveSheet()->setCellValue('A'.$i, $start_date);
        $obj->getActiveSheet()->setCellValue('B'.$i, $end_date);
        $obj->getActiveSheet()->setCellValue('C'.$i, $data_exp[8]);
        $obj->getActiveSheet()->setCellValue('D'.$i, $data_exp[109]);
        $obj->getActiveSheet()->setCellValue('E'.$i, $data_exp[3]);
        $obj->getActiveSheet()->setCellValue('F'.$i, $data_exp[16]);
        $obj->getActiveSheet()->setCellValue('G'.$i, $data_exp[16]);
        $obj->getActiveSheet()->setCellValue('H'.$i, $data_exp[22]);
        $obj->getActiveSheet()->setCellValue('I'.$i, $data_exp[28]);
        $obj->getActiveSheet()->setCellValue('J'.$i, $data_exp[34]);
        $obj->getActiveSheet()->setCellValue('K'.$i, $data_exp[40]);
        $obj->getActiveSheet()->setCellValue('L'.$i, $data_exp[46]);
        $obj->getActiveSheet()->setCellValue('M'.$i, $data_exp[52]);
        $obj->getActiveSheet()->setCellValue('N'.$i, $data_exp[58]);
        $obj->getActiveSheet()->setCellValue('O'.$i, $data_exp[64]);
        $obj->getActiveSheet()->setCellValue('P'.$i, $data_exp[70]);
        $obj->getActiveSheet()->setCellValue('Q'.$i, $data_exp[76]);
        $obj->getActiveSheet()->setCellValue('R'.$i, $data_exp[82]);
        $obj->getActiveSheet()->setCellValue('S'.$i, $data_exp[88]);
        $obj->getActiveSheet()->setCellValue('T'.$i, $data_exp[94]);
        $obj->getActiveSheet()->setCellValue('U'.$i, $data_exp[100]);
        $obj->getActiveSheet()->setCellValue('V'.$i, $data_exp[18]);
        $obj->getActiveSheet()->setCellValue('W'.$i, $data_exp[24]);
        $obj->getActiveSheet()->setCellValue('X'.$i, $data_exp[30]);
        $obj->getActiveSheet()->setCellValue('Y'.$i, $data_exp[36]);
        $obj->getActiveSheet()->setCellValue('Z'.$i, $data_exp[42]);
        $obj->getActiveSheet()->setCellValue('AA'.$i, $data_exp[48]);
        $obj->getActiveSheet()->setCellValue('AB'.$i, $data_exp[54]);
        $obj->getActiveSheet()->setCellValue('AC'.$i, $data_exp[60]);
        $obj->getActiveSheet()->setCellValue('AD'.$i, $data_exp[66]);
        $obj->getActiveSheet()->setCellValue('AE'.$i, $data_exp[72]);
        $obj->getActiveSheet()->setCellValue('AF'.$i, $data_exp[78]);
        $obj->getActiveSheet()->setCellValue('AG'.$i, $data_exp[84]);
        $obj->getActiveSheet()->setCellValue('AH'.$i, $data_exp[90]);
        $obj->getActiveSheet()->setCellValue('AI'.$i, $data_exp[96]);
        $obj->getActiveSheet()->setCellValue('AJ'.$i, $data_exp[102]);
        $obj->getActiveSheet()->setCellValue('AK'.$i, $data_exp[20]);
        $obj->getActiveSheet()->setCellValue('AL'.$i, $data_exp[26]);
        $obj->getActiveSheet()->setCellValue('AM'.$i, $data_exp[32]);
        $obj->getActiveSheet()->setCellValue('AN'.$i, $data_exp[38]);
        $obj->getActiveSheet()->setCellValue('AO'.$i, $data_exp[44]);
        $obj->getActiveSheet()->setCellValue('AP'.$i, $data_exp[50]);
        $obj->getActiveSheet()->setCellValue('AQ'.$i, $data_exp[56]);
        $obj->getActiveSheet()->setCellValue('AR'.$i, $data_exp[62]);
        $obj->getActiveSheet()->setCellValue('AS'.$i, $data_exp[68]);
        $obj->getActiveSheet()->setCellValue('AT'.$i, $data_exp[74]);
        $obj->getActiveSheet()->setCellValue('AU'.$i, $data_exp[80]);
        $obj->getActiveSheet()->setCellValue('AV'.$i, $data_exp[86]);
        $obj->getActiveSheet()->setCellValue('AW'.$i, $data_exp[92]);
        $obj->getActiveSheet()->setCellValue('AX'.$i, $data_exp[98]);
        $obj->getActiveSheet()->setCellValue('AY'.$i, $data_exp[104]);
        $obj->getActiveSheet()->setCellValue('AZ'.$i, $data_exp[104]);

        return $obj;
    }
}
