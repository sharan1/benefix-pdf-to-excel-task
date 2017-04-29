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
    public static function addNonTobaccoData($data_exp, $obj, $i, $isSpecial)
    {
        $str = $data_exp[5];
        $str = str_replace("Rates effective from ","",$str);
        $arr = explode("through", $str);
        $start_date = trim($arr[0]);
        $end_date = trim($arr[1]);

        $str_product = $data_exp[6];
        $str_product = str_replace("Keystone ","",$str_product);

        $obj->getActiveSheet()->setCellValue('A'.$i, $start_date." 00:00:00 UTC");
        $obj->getActiveSheet()->setCellValue('B'.$i, $end_date." 00:00:00 UTC");
        $obj->getActiveSheet()->setCellValue('C'.$i, $str_product);
        $obj->getActiveSheet()->setCellValue('D'.$i, "PA");
        $obj->getActiveSheet()->setCellValue('E'.$i, "8");
        $obj->getActiveSheet()->setCellValue('F'.$i, (double)ltrim($data_exp[15], '$'));
        $obj->getActiveSheet()->setCellValue('G'.$i, (double)ltrim($data_exp[15], '$'));
        $obj->getActiveSheet()->setCellValue('H'.$i, (double)ltrim($data_exp[18], '$'));
        $obj->getActiveSheet()->setCellValue('I'.$i, (double)ltrim($data_exp[21], '$'));
        $obj->getActiveSheet()->setCellValue('J'.$i, (double)ltrim($data_exp[24], '$'));
        $obj->getActiveSheet()->setCellValue('K'.$i, (double)ltrim($data_exp[27], '$'));
        $obj->getActiveSheet()->setCellValue('L'.$i, (double)ltrim($data_exp[30], '$'));
        $obj->getActiveSheet()->setCellValue('M'.$i, (double)ltrim($data_exp[33], '$'));
        $obj->getActiveSheet()->setCellValue('N'.$i, (double)ltrim($data_exp[36], '$'));
        $obj->getActiveSheet()->setCellValue('O'.$i, (double)ltrim($data_exp[39], '$'));
        $obj->getActiveSheet()->setCellValue('P'.$i, (double)ltrim($data_exp[42], '$'));
        $obj->getActiveSheet()->setCellValue('Q'.$i, (double)ltrim($data_exp[45], '$'));
        $obj->getActiveSheet()->setCellValue('R'.$i, (double)ltrim($data_exp[48], '$'));
        $obj->getActiveSheet()->setCellValue('S'.$i, (double)ltrim($data_exp[51], '$'));
        $obj->getActiveSheet()->setCellValue('T'.$i, (double)ltrim($data_exp[54], '$'));
        $obj->getActiveSheet()->setCellValue('U'.$i, (double)ltrim($data_exp[57], '$'));
        $obj->getActiveSheet()->setCellValue('V'.$i, (double)ltrim($data_exp[60], '$'));
        $obj->getActiveSheet()->setCellValue('W'.$i, (double)ltrim($data_exp[63], '$'));
        $obj->getActiveSheet()->setCellValue('X'.$i, (double)ltrim($data_exp[66], '$'));
        $obj->getActiveSheet()->setCellValue('Y'.$i, (double)ltrim($data_exp[69], '$'));
        $obj->getActiveSheet()->setCellValue('Z'.$i, (double)ltrim($data_exp[72], '$'));
        $obj->getActiveSheet()->setCellValue('AA'.$i, (double)ltrim($data_exp[140], '$'));
        $obj->getActiveSheet()->setCellValue('AB'.$i, (double)ltrim($data_exp[143], '$'));
        $obj->getActiveSheet()->setCellValue('AC'.$i, (double)ltrim($data_exp[146], '$'));
        $obj->getActiveSheet()->setCellValue('AD'.$i, (double)ltrim($data_exp[79], '$'));
        $obj->getActiveSheet()->setCellValue('AE'.$i, (double)ltrim($data_exp[82], '$'));
        $obj->getActiveSheet()->setCellValue('AF'.$i, (double)ltrim($data_exp[85], '$'));
        $obj->getActiveSheet()->setCellValue('AG'.$i, (double)ltrim($data_exp[88], '$'));
        $obj->getActiveSheet()->setCellValue('AH'.$i, (double)ltrim($data_exp[91], '$'));
        $obj->getActiveSheet()->setCellValue('AI'.$i, (double)ltrim($data_exp[94], '$'));
        $obj->getActiveSheet()->setCellValue('AJ'.$i, (double)ltrim($data_exp[97], '$'));
        $obj->getActiveSheet()->setCellValue('AK'.$i, (double)ltrim($data_exp[100], '$'));
        $obj->getActiveSheet()->setCellValue('AL'.$i, (double)ltrim($data_exp[103], '$'));
        $obj->getActiveSheet()->setCellValue('AM'.$i, (double)ltrim($data_exp[106], '$'));
        $obj->getActiveSheet()->setCellValue('AN'.$i, (double)ltrim($data_exp[109], '$'));
        $obj->getActiveSheet()->setCellValue('AO'.$i, (double)ltrim($data_exp[112], '$'));
        $obj->getActiveSheet()->setCellValue('AP'.$i, (double)ltrim($data_exp[115], '$'));
        $obj->getActiveSheet()->setCellValue('AQ'.$i, (double)ltrim($data_exp[118], '$'));
        $obj->getActiveSheet()->setCellValue('AR'.$i, (double)ltrim($data_exp[121], '$'));
        $obj->getActiveSheet()->setCellValue('AS'.$i, (double)ltrim($data_exp[124], '$'));
        $obj->getActiveSheet()->setCellValue('AT'.$i, (double)ltrim($data_exp[127], '$'));
        $obj->getActiveSheet()->setCellValue('AU'.$i, (double)ltrim($data_exp[130], '$'));
        $obj->getActiveSheet()->setCellValue('AV'.$i, (double)ltrim($data_exp[133], '$'));
        $obj->getActiveSheet()->setCellValue('AW'.$i, (double)ltrim($data_exp[136], '$'));
        $obj->getActiveSheet()->setCellValue('AX'.$i, (double)ltrim($data_exp[149], '$'));
        $obj->getActiveSheet()->setCellValue('AY'.$i, (double)ltrim($data_exp[152], '$'));
        $obj->getActiveSheet()->setCellValue('AZ'.$i, (double)ltrim($data_exp[152], '$'));

        return $obj;
    }

    // Adding Data to the file
    public static function addTobaccoData($data_exp, $obj, $i, $isSpecial)
    {
        $str = $data_exp[5];
        $str = str_replace("Rates effective from ","",$str);
        $arr = explode("through", $str);
        $start_date = trim($arr[0]);
        $end_date = trim($arr[1]);

        $str_product = $data_exp[6];
        $str_product = str_replace("Keystone ","",$str_product);

        $obj->getActiveSheet()->setCellValue('A'.$i, $start_date." 00:00:00 UTC");
        $obj->getActiveSheet()->setCellValue('B'.$i, $end_date." 00:00:00 UTC");
        $obj->getActiveSheet()->setCellValue('C'.$i, $str_product);
        $obj->getActiveSheet()->setCellValue('D'.$i, "PA");
        $obj->getActiveSheet()->setCellValue('E'.$i, "8");
        $obj->getActiveSheet()->setCellValue('F'.$i, (double)ltrim($data_exp[16], '$'));
        $obj->getActiveSheet()->setCellValue('G'.$i, (double)ltrim($data_exp[16], '$'));
        $obj->getActiveSheet()->setCellValue('H'.$i, (double)ltrim($data_exp[19], '$'));
        $obj->getActiveSheet()->setCellValue('I'.$i, (double)ltrim($data_exp[22], '$'));
        $obj->getActiveSheet()->setCellValue('J'.$i, (double)ltrim($data_exp[25], '$'));
        $obj->getActiveSheet()->setCellValue('K'.$i, (double)ltrim($data_exp[28], '$'));
        $obj->getActiveSheet()->setCellValue('L'.$i, (double)ltrim($data_exp[31], '$'));
        $obj->getActiveSheet()->setCellValue('M'.$i, (double)ltrim($data_exp[34], '$'));
        $obj->getActiveSheet()->setCellValue('N'.$i, (double)ltrim($data_exp[37], '$'));
        $obj->getActiveSheet()->setCellValue('O'.$i, (double)ltrim($data_exp[40], '$'));
        $obj->getActiveSheet()->setCellValue('P'.$i, (double)ltrim($data_exp[43], '$'));
        $obj->getActiveSheet()->setCellValue('Q'.$i, (double)ltrim($data_exp[46], '$'));
        $obj->getActiveSheet()->setCellValue('R'.$i, (double)ltrim($data_exp[49], '$'));
        $obj->getActiveSheet()->setCellValue('S'.$i, (double)ltrim($data_exp[52], '$'));
        $obj->getActiveSheet()->setCellValue('T'.$i, (double)ltrim($data_exp[55], '$'));
        $obj->getActiveSheet()->setCellValue('U'.$i, (double)ltrim($data_exp[58], '$'));
        $obj->getActiveSheet()->setCellValue('V'.$i, (double)ltrim($data_exp[61], '$'));
        $obj->getActiveSheet()->setCellValue('W'.$i, (double)ltrim($data_exp[64], '$'));
        $obj->getActiveSheet()->setCellValue('X'.$i, (double)ltrim($data_exp[67], '$'));
        $obj->getActiveSheet()->setCellValue('Y'.$i, (double)ltrim($data_exp[70], '$'));
        $obj->getActiveSheet()->setCellValue('Z'.$i, (double)ltrim($data_exp[73], '$'));
        $obj->getActiveSheet()->setCellValue('AA'.$i, (double)ltrim($data_exp[141], '$'));
        $obj->getActiveSheet()->setCellValue('AB'.$i, (double)ltrim($data_exp[144], '$'));
        $obj->getActiveSheet()->setCellValue('AC'.$i, (double)ltrim($data_exp[147], '$'));
        $obj->getActiveSheet()->setCellValue('AD'.$i, (double)ltrim($data_exp[80], '$'));
        $obj->getActiveSheet()->setCellValue('AE'.$i, (double)ltrim($data_exp[83], '$'));
        $obj->getActiveSheet()->setCellValue('AF'.$i, (double)ltrim($data_exp[86], '$'));
        $obj->getActiveSheet()->setCellValue('AG'.$i, (double)ltrim($data_exp[89], '$'));
        $obj->getActiveSheet()->setCellValue('AH'.$i, (double)ltrim($data_exp[92], '$'));
        $obj->getActiveSheet()->setCellValue('AI'.$i, (double)ltrim($data_exp[95], '$'));
        $obj->getActiveSheet()->setCellValue('AJ'.$i, (double)ltrim($data_exp[98], '$'));
        $obj->getActiveSheet()->setCellValue('AK'.$i, (double)ltrim($data_exp[101], '$'));
        $obj->getActiveSheet()->setCellValue('AL'.$i, (double)ltrim($data_exp[104], '$'));
        $obj->getActiveSheet()->setCellValue('AM'.$i, (double)ltrim($data_exp[107], '$'));
        $obj->getActiveSheet()->setCellValue('AN'.$i, (double)ltrim($data_exp[110], '$'));
        $obj->getActiveSheet()->setCellValue('AO'.$i, (double)ltrim($data_exp[113], '$'));
        $obj->getActiveSheet()->setCellValue('AP'.$i, (double)ltrim($data_exp[116], '$'));
        $obj->getActiveSheet()->setCellValue('AQ'.$i, (double)ltrim($data_exp[119], '$'));
        $obj->getActiveSheet()->setCellValue('AR'.$i, (double)ltrim($data_exp[122], '$'));
        $obj->getActiveSheet()->setCellValue('AS'.$i, (double)ltrim($data_exp[125], '$'));
        $obj->getActiveSheet()->setCellValue('AT'.$i, (double)ltrim($data_exp[128], '$'));
        $obj->getActiveSheet()->setCellValue('AU'.$i, (double)ltrim($data_exp[131], '$'));
        $obj->getActiveSheet()->setCellValue('AV'.$i, (double)ltrim($data_exp[134], '$'));
        $obj->getActiveSheet()->setCellValue('AW'.$i, (double)ltrim($data_exp[137], '$'));
        $obj->getActiveSheet()->setCellValue('AX'.$i, (double)ltrim($data_exp[150], '$'));
        $obj->getActiveSheet()->setCellValue('AY'.$i, (double)ltrim($data_exp[153], '$'));
        $obj->getActiveSheet()->setCellValue('AZ'.$i, (double)ltrim($data_exp[153], '$'));

        return $obj;
    }
}
