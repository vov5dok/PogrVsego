<?php
/**
 * Ключевые показатели деятельности Регионов
 * 
 */ 
require_once('Utils.php');
require_once('../Classes/PHPExcel.php');

class KPDnew {
    
private $year, $m, $reg, $pryear;
private $date, $fdate, $arReg,$arPok, $all_data;

function KPDnew($type, $year, $mes, $reg) {
    
    $this->type = $type; //нарастание (0 - за месяц или 1 - с начала года)
    if ($type==1) //нарастанием с начала года
    {
        $this->fdate = $year."-01-01"; //начало года  
        $this->date = Utils::LastDay($mes, $year);
        $this->year = $year;
        $this->m = $mes;
        $this->pryear = (int)$year-1;
        $this->fdatepr = $this->pryear."-01-01"; //начало прошлого года  
        $this->datepr = Utils::LastDay($mes, $this->pryear);
        $this->kolM=(int)$mes; //количество месяцев 
    }
 else {   //за 1 месяц     
        $this->date = Utils::LastDay($mes, $year);
        $this->fdate = Utils::FirstDay($this->date); //1 день месяца
        $this->year = $year;
        $this->m = $mes;
        $this->pryear = (int)$year-1;       
        $this->datepr = Utils::LastDay($mes, $this->pryear);
        $this->fdatepr = Utils::FirstDay($this->datepr); //1 день месяца
        $this->kolM=1; //количество месяцев         
    }
        
    $this->arReg = array(
        1=>"Пензенский",
        2=>"Волго-Камский",
        3=>"Самарский",
        4=>"Башкирский"
    );
    $this->reg = $reg;
    $this->nreg = $this->arReg[$this->reg];        
       
    $this->arPok = array(    
    0=>"Погрузка грузов",
    1=>"Выполнение расписания движения пассажирских поездов (по прибытию на станции посадки/высадки)",
    2=>"Возмещение выпадающих доходов (убытков) от пригородных перевозок",    
    3=>"Количество пострадавших со смертельным исходом",    
    4=>"Уровень безопасности движения",
    //5=>"Интенсивность отказов",
    6=>"Удельное количество отказов 1,2 категории",
    7=>"Удельное количество технологических нарушений",
    8=> "Итого показателей по региону",
    9=>"Итого не выполнено показателей",
    10=>"Процент невыполнения"
    );
    
    $this->arCol = array(
    1=>"№п/п",
    2=>"Наименование показателя",
    3=>"Ед.изм.",
    4=>"Отчет ".$this->pryear." г.",    
    5=>"План ".$this->year." г.",    
    6=>"Выполнение ".$this->year." г.", 
    7=>"% выполнения плана",
    8=> "% к прошлому году",
    9=>"Оценка",
    10=>"Ответственность за невыполнение показателя",
    11=>"Ответственность за ввод показателя"
    );      
    
  $this->GetData(6, $this->year, $this->m);//факт  
  $this->GetData(4, $this->pryear, $this->m);//пр.год  
  
  $this->GetPlan();  //план    
 }
 
  // $Key1 - id столбца  (arCol),  $Key2 - id показателя  (arPok)
  function setValue($Key1, $Key2, $Value) {
    $val = Utils::Delim($Value);//заменяем , на .
    $this->all_data[$Key1][$Key2] = $val;
  } 
   
  function getValue($Key1, $Key2) {
    $Value = isset($this->all_data[$Key1][$Key2]) ? $this->all_data[$Key1][$Key2] : 0;
    return  $Value;  
  }
 
 /*проверяем подписанность Показателей КПД, введенных в УПВД (НЗ-Тер)
  * в БД хранятся варианты неподписанные без ФИО и даты
  * и подписанные с ФИО ответственного и датой подписи
  * выбираем максимальную подпись по показателю
  * $id - ID показателя, $m - месяц
  */
 function GetPodp($id, $m, $y) {   
   $sql = "select coalesce(max(PODP),0) as podp  from PLAN.KPD  where ID = ".$id." and year(date) = ".$y." and month(date) = ".$m; 
   $result = Database::select($sql);        
   foreach($result as $str) {
     $podp = $str['PODP'];      
   } 
    return $podp;
 }
 
//факт
 function GetData($type, $y, $m) {  

  foreach ($this->arPok  as $idPok =>$nPok)
  {
      if ($idPok<=7) //основные показатели (1-7 в справке)
      {
          if ($idPok==0) //погрузка грузов 
          {
            if ($this->type==1) { //накопление
            $sql = "select sum(POGRVES) as VALUE from MURATOV.GO10  
            where RYEAR = ".$y." and RMONTH <= ".$m." and GR10_1=99 and nod = ".$this->reg." group by nod";
            $result = Database::select($sql);
            foreach ($result as $str) {              
               $value = $str['VALUE'];   
            }
           } else { //за месяц
            $sql = "select coalesce(sum(POGRVES),0) as VALUE from MURATOV.GO10  
            where RYEAR = ".$y." and RMONTH = ".$m." and GR10_1=99 and nod = ".$this->reg;
            $result = Database::select($sql);
            foreach ($result as $str) {              
               $value = $str['VALUE'];                                         
            }   
           }
           $value = Utils::Delim($value)/1000;
           $value = number_format(round($value, 3), 3, '.', '');
           $this->setValue($type, $idPok, $value); 
          }   
          
          //**********************добавлено 03/2018*********************************************************
          //отражаются данные, введенные НАРАСТАНИЕМ на отчетный месяц  через УПВД
          //т.к. невозможно посчитать нарастание по введенным месяцам
          //******************************************************************************************************
          else if ($idPok==4 || $idPok==6 || $idPok==7) //уровень безопасности или интенсивность отказов 
          {
             $value = 0;
            if ($this->type==1) { //накопление
             $sql = "select coalesce(F".$this->reg.",0) as VALUE from PLAN.KPD
                where ID = ".$idPok." and Year(DATE)= ".$y." and month(Date)=".$m." and podp = ".self::GetPodp($idPok, $m, $y);
            $result = Database::select($sql);
            foreach ($result as $str) {              
               $value = $str['VALUE'];   
            }
           } else { //за месяц
             $sql = "select coalesce(F".$this->reg.",0) as VALUE from PLAN.KPD
                where ID = ".$idPok." and Year(DATE)= ".$y." and month(Date)=".$m." and podp = ".self::GetPodp($idPok, $m, $y);
            $result = Database::select($sql);
            foreach ($result as $str) {              
               $value = $str['VALUE'];                                         
            }   
           }
           $value = Utils::Delim($value);
           $value = number_format(round($value, 3), 3, '.', '');
           $this->setValue($type, $idPok, $value); 
          }
          //*****************************************************************************************************************
          else { //остальные КПД
              $val = 0;
              if ($this->type==1) { //накопление
              //ищем значения КПД по максимальной подписанности за каждый месяц и складываем
              for ($mm=1; $mm<=$m; $mm++)
              {
                $sql = "select coalesce(sum(F".$this->reg."),0) as VALUE from PLAN.KPD
                where ID = ".$idPok." and Year(DATE)= ".$y." and month(Date)=".$mm." and podp = ".self::GetPodp($idPok, $mm, $y);
                $result = Database::select($sql);
                foreach ($result as $str) {              
                   $value = $str['VALUE'];  
                   $val = $val+Utils::Delim($value);
                }
              }
            } else { //за месяц
                $sql = "select coalesce(sum(F".$this->reg."),0) as VALUE from PLAN.KPD
                where ID = ".$idPok." and Year(DATE)= ".$y." and month(Date)=".$m." and podp = ".self::GetPodp($idPok, $m,$y);
                $result = Database::select($sql);
                foreach ($result as $str) {              
                   $value = $str['VALUE'];  
                   $val = Utils::Delim($value);
            }              
         }
         //факт всех показателей - сумма планов за месяц????????              
        $valF = $val;
        if ($idPok==1) //выполнение расписания
        {
            $valF = $val/$this->kolM;
           // $valF = number_format(round($valF, 1), 1, '.', ''); //до скольки округлять?????????????
        }
        if (($idPok==5)||($idPok==4)) //интенсивность, уровень безопасности - занесены в УПВД нарастанием!!!
        {
            $valF = $val;
        }
        $this->setValue($type, $idPok, $valF);              
     }
  }
  } 
 }
 
 //план
 function GetPlan() {
           
     foreach ($this->arPok  as $idPok =>$nPok)
    {
      if ($idPok<=7) //основные показатели 
      {          
          if ($idPok==0) //погрузка грузов 
          {  
            $plan = 0;
            if ($this->type==1) { //накопление
              $sql = "Select month(Date) as MN, TONN as VALUE from PLAN.P_POGR_MAIN_N
                where Year(DATE)=".$this->year." and month(Date)<= ".$this->m."
                and ROD_GR = '99' and KOD_OTDEL = '0".$this->reg."'";
              $result = Database::select($sql);
              foreach($result as $str) {
                   $month = $str['MN'];
                   $kol = Utils::NumDayinMonth1($month, $this->year);  //кол-во дней в месяце                               
                   $val = $str['VALUE']*$kol;
                   $plan += $val; 
                 }
             }
            else { //за месяц
                $sql = "Select month(Date) as MN, TONN as VALUE from PLAN.P_POGR_MAIN_N
                where Year(DATE)=".$this->year." and month(Date)= ".$this->m."
                and ROD_GR = '99' and KOD_OTDEL = '0".$this->reg."'";
              $result = Database::select($sql);
              foreach($result as $str) {
                   $month = $str['MN'];
                   $kol = Utils::NumDayinMonth1($month, $this->year);  //кол-во дней в месяце                               
                   $val = $str['VALUE']*$kol;
                   $plan = $val; 
                 }
              }
              $plan = Utils::Delim($plan)/1000;
              $plan = number_format(round($plan, 3), 3, '.', '');
              $this->setValue(5, $idPok, $plan);
            }          
            else {
              $valP = 0;
              if ($this->type==1) { //накопление                
              //ищем значения КПД по максимальной подписанности за каждый месяц и складываем
              for ($mm=1; $mm<=$this->m; $mm++)
              {    
                $sql = "select month(Date) as MN, coalesce(sum(PL".$this->reg."),0) as VALUE  from PLAN.KPD
                   where ID = ".$idPok." and Year(DATE)= ".$this->year." and month(Date)=".$mm." and podp = ".self::GetPodp($idPok, $mm,$this->year)." group by date";
               // echo $sql."<br>";
                $result = Database::select($sql);
                foreach ($result as $str) {              
                   $value = $str['VALUE'];  
                   $valP = $valP+Utils::Delim($value);
                }
              }
              } else { //за месяц
                $sql = "select month(Date) as MN, coalesce(sum(PL".$this->reg."),0) as VALUE  from PLAN.KPD
                   where ID = ".$idPok." and Year(DATE)= ".$this->year." and month(Date)=".$this->m." and podp = ".self::GetPodp($idPok, $this->m,$this->year)." group by date";
               // echo $sql."<br>";
                $result = Database::select($sql);
                foreach ($result as $str) {              
                   $value = $str['VALUE'];  
                   $valP = Utils::Delim($value);
                }  
              }
              //возмещение убытков ($idPok==2) - сумма планов за месяц????????
              //по остальным - нули
              $plan = $valP;
              
              if (($idPok==1)||($idPok==4)||($idPok==5)) //выполнение расписания, уровень, интенсивность
              {                  
                  $plan = $valP/$this->kolM;
                  //$plan = number_format(round($plan, 1), 1, '.', ''); //до скольки округлять?????????????
              }
              $this->setValue(5, $idPok, $plan);
           }
      
      }
   }    
  }
 
 //вычисляем % Type1 / Type2
 function CalcPerc($IdPok, $Type1, $Type2) {
    
    $pok1 = $this->getValue($Type1, $IdPok); 
    $pok2 = $this->getValue($Type2, $IdPok);  
    $perc = $pok2!=0 ? $pok1*100/$pok2 :0;
    //********округление до 0,1
    $perc = number_format(round($perc, 1), 1, '.', '');
    return $perc;  
  }        
  
  // Получаем Excel-файл    
  function GetExcelTable() {
   
    $template = "template/KPDNzTerNew.xlsx"; //файл-шаблон
    $titR = substr($this->nreg, 0, -2)."ого региона Куйбышевской железной дороги"  ;
    $title1 = iconv('windows-1251', 'UTF-8', $titR);
    if ($this->type==1) //нарастание
    {
        if ($this->m==1) $mmm= " месяц "; else if (($this->m==2)||($this->m==3)||($this->m==4)) $mmm= " месяца "; else $mmm= " месяцев ";
        $title2 = iconv('windows-1251', 'UTF-8', "за ".$this->m.$mmm.$this->year." года");
        $fname = "КПД_".$this->reg."РЕГ_".$this->m."мес".$this->year.".xlsx";//имя файла для сохранения    
    }
    else { //за месяц
        $title2 = iconv('windows-1251', 'UTF-8', "за ".Utils::GetNMonth($this->m)." ".$this->year." года");
        $fname = "КПД_".$this->reg."РЕГ_".str_pad($this->m, 2, '0', STR_PAD_LEFT).".".$this->year.".xlsx";//имя файла для сохранения        
    }
    
    $xls = new PHPExcel();
    $cacheMethod = PHPExcel_CachedObjectStorageFactory:: cache_to_phpTemp;
    $cacheSettings = array( ' memoryCacheSize ' => '1024MB');
    PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
    if (!PHPExcel_Settings::setCacheStorageMethod($cacheMethod,$cacheSettings))
       die('CACHEING ERROR');
    $xls = PHPExcel_IOFactory::load($template);   
    $xls->setActiveSheetIndex(0);      
    $sheet = $xls->getActiveSheet();
    $sheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);   
    
    //регион    
    $sheet->setCellValue('A2', $title1);
    //количество месяцев в заголовок отчета    
    $sheet->setCellValue('A3', $title2);
    //даты в шапки    
    $pry = iconv('windows-1251', 'UTF-8', $this->arCol[4]);
    $sheet->setCellValue('D4', $pry);
    $plan = iconv('windows-1251', 'UTF-8', $this->arCol[5]);
    $sheet->setCellValue('E4', $plan);
    $fact = iconv('windows-1251', 'UTF-8', $this->arCol[6]);
    $sheet->setCellValue('F4', $fact);
        
    $row=7; //с какой строки начинаем вывод данных - конец шапки      
   $nevyp = 0;
   $ColorRed = new PHPExcel_Style_Color();
   $ColorRed->setRGB('FF0000');  
   $ColorGreen = new PHPExcel_Style_Color();
   $ColorGreen->setRGB('9ACD32');  
   $ColorWhite = new PHPExcel_Style_Color();
   $ColorWhite->setRGB('FFFFFF');  
   $ColorBlue = new PHPExcel_Style_Color();
   $ColorBlue->setRGB('0000CD');  
   
    
   foreach ($this->arPok  as $idPok =>$nPok)
   {
      if ($idPok<=7) //основные показатели 
      {
      //    if ($idPok==1) //погрузка грузов 
      //    {
         //подписанный или нет
          //текущий год
          $podpT = self::GetPodp($idPok, $this->m, $this->year);
          //предыдущий
          $podpL = self::GetPodp($idPok, $this->m, $this->pryear);
          
            $r = $row+$idPok;
            $rr = (string)$r;
            if ($idPok == 6 || $idPok == 7)
            {
                $rr = $rr - 1;
            }
            $pry = $this->getValue(4,$idPok);  //прошлый год
            $pry = Utils::Delim($pry);           
          //  $pry = $pry!=0 ? $pry : '';
            $sheet->setCellValue('D'.$rr, $pry);  
            if ($podpL==0) $sheet->getStyle('D'.$rr)->getFont()->setColor($ColorBlue);
            
            $pl = $this->getValue(5,$idPok);  //план
            $pl = Utils::Delim($pl);
            if ($idPok>=3) $pl = $pl!=0 ? $pl : '';  
            $sheet->setCellValue('E'.$rr, $pl);
            if ($podpT==0) $sheet->getStyle('E'.$rr)->getFont()->setColor($ColorBlue);
            
            $fact = $this->getValue(6,$idPok);  //факт
            $fact = Utils::Delim($fact);
          //  $fact = $fact!=0 ? $fact : '';  
            $sheet->setCellValue('F'.$rr, $fact);
            if ($podpT==0) $sheet->getStyle('F'.$rr)->getFont()->setColor($ColorBlue);
            
           /**
           * вычисляем % iType1/iType2
           * 20.02.19 для первых трёх в заливке зел добавил  || ($pl==0)&&($fact==0)          
           */
           if (($idPok==0)||($idPok==1)||($idPok==2))
           {
            if (($pl==0)&&($fact!=0) || ($pl==0)&&($fact==0)) {
                //заливка ячеек зеленым
                    $sheet->getStyle('I'.$rr)->getFont()->setColor($ColorGreen); 
                    $sheet->getStyle('I'.$rr)->getFill()->getStartColor()->applyFromArray(array('rgb' => '9ACD32'));               
            } else  {
               if (self::CalcPerc($idPok, 6, 5)<100)
                {
                   //заливка ячеек красным
                    $sheet->getStyle('I'.$rr)->getFont()->setColor($ColorRed);                                             
                    $sheet->getStyle('I'.$rr)->getFill()->getStartColor()->applyFromArray(array('rgb' => 'FF0000'));
                    $nevyp =$nevyp +1;
                } else {
                    //заливка ячеек зеленым
                    $sheet->getStyle('I'.$rr)->getFont()->setColor($ColorGreen); 
                    $sheet->getStyle('I'.$rr)->getFill()->getStartColor()->applyFromArray(array('rgb' => '9ACD32'));
                }  
            }         
                
          } else { 
            if (($idPok==3)||($idPok==4)||($idPok==6)||($idPok==7))
              {
                 if (($pry==0)&&($fact!=0)) {                
                  //заливка ячеек красным
                  $sheet->getStyle('I'.$rr)->getFont()->setColor($ColorRed);                                             
                  $sheet->getStyle('I'.$rr)->getFill()->getStartColor()->applyFromArray(array('rgb' => 'FF0000'));
                  $nevyp =$nevyp +1;
              } else
                 {
                  if (self::CalcPerc($idPok, 6, 4)>=100)
                  {
                     //заливка ячеек красным
                      $sheet->getStyle('I'.$rr)->getFont()->setColor($ColorRed);                                             
                      $sheet->getStyle('I'.$rr)->getFill()->getStartColor()->applyFromArray(array('rgb' => 'FF0000'));
                      $nevyp =$nevyp +1;
                  } else {
                      //заливка ячеек зеленым
                      $sheet->getStyle('I'.$rr)->getFont()->setColor($ColorGreen); 
                      $sheet->getStyle('I'.$rr)->getFill()->getStartColor()->applyFromArray(array('rgb' => '9ACD32'));
                  }  
                }
              }     
            }                           
        }          
      }
      
      $nevyp = $nevyp!=0 ? $nevyp : '';  
      $sheet->setCellValue('I15', $nevyp);
      //$sheet->setSaved(true);
      
      if ($this->type==0) //за месяц, приписка о введенных показателях РБ
      {
          $p4 = iconv('windows-1251', 'UTF-8', 'Уровень безопасности движения*');
          $sheet->setCellValue('B11', $p4);
          $p5 = iconv('windows-1251', 'UTF-8', 'Удельное количество отказов 1,2 категории*');
          $sheet->setCellValue('B12', $p5);
          $p6 = iconv('windows-1251', 'UTF-8', 'Удельное количество технологических нарушений*');
          $sheet->setCellValue('B13', $p6);
          $info = iconv('windows-1251', 'UTF-8', '* показатели введены нарастанием с начала года');
          $sheet->setCellValue('A18', $info);
      }
      
   ob_end_clean();

   $objWriter = PHPExcel_IOFactory::createWriter($xls, 'Excel2007');
   header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');      
   header('Content-Disposition: attachment;filename="'.$fname.'"');
   header('Cache-Control: max-age=0');     
   $objWriter->setPreCalculateFormulas(false);
   $objWriter->save('php://output');
   exit;    
   }   
    
}

?>