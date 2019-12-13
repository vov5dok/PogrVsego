<?php
// ini_set('error_reporting', E_ALL);
// ini_set('display_errors', 1);
// ini_set('display_startup_errors', 1);

/**
 * Погрузка и начисленная выручка всего по дороге с расчетом через ЦФТО
 * @author Иванов В. Ю.
 * @copyright 2019
 */

require_once('../Classes/PHPExcel.php');

class PogrVsego
{
    private $arrGruz;
    private $date;
    private $dateEpl;
    private $y;
    private $m;
    private $d;
    private $pokazPlanDoroga;
    private $pokazPlanCfto;
    private $pokazFactDoroga;
    private $pokazFactCfto;
    private $titleSpravka;

    // Инициализируем свойства класса
    public function __construct($date) {

        $this->pokazPlanDoroga = array(1, 2, 4, 6, 7, 8, 9); // Показатели плана на Дороге по порядку справок в excel-шаблоне
        $this->pokazPlanCfto = array(17, 10, 5, 11, 12, 13, 14); // Показатели плана ЦФТО по порядку справок в excel-шаблоне

        $this->pokazFactDoroga = array(0, 1, 2, 4, 5, 6, 3, 7); // Показатели плана на Дороге по порядку справок в excel-шаблоне
        $this->pokazFactCfto = array(8, 9, 10, 12, 13, 14, 1, 15); // Показатели плана ЦФТО по порядку справок в excel-шаблоне

        // Наименования справок строго по порядку листов в excel-шаблоне
        $this->titleSpravka = array(
            "Погрузка и начисленная выручка всего ",
            "Погрузка и начисленная выручка всего - внутрироссийское сообщение ",
            "Погрузка и начисленная выручка всего - международное сообщение ",            
            "Погрузка и начисленная выручка всего - международное сообщение через погранпереходы в 3-и страны ",
            "Погрузка и начисленная выручка всего - международное сообщение через погранпереходы в страны СНГ (кроме Белорусии) ",
            "Погрузка и начисленная выручка всего - международное сообщение в Белорусию ",            
            "Погрузка и начисленная выручка всего - международное сообщение через порты ",
        );
        // Наименования грузов строго по excel-шаблону
        $this->arrGruz = array(
            1 => 'Каменный уголь',
            2 => 'Кокс',
            3 => 'Нефть и нефтепродукты',
            4 => 'Торф и торфяная продукция',
            5 => 'Сланцы и горючие',
            6 => 'Флюсы',
            7 => 'Руда железная и марганцевая',
            8 => 'Руда цветная и серное сырье',
            9 => 'Черные металлы',
            10 => 'Машины и оборудование',
            11 => 'Металлические конструкции',
            12 => 'Метизы',
            13 => 'Лом черных металлов',
            14 => 'Сельскохозяйственные машины',
            15 => 'Автомобили',
            16 => 'Цв. металлы, изделия из них и лом цветных металлов',
            17 => 'Химические и минеральные удобрения',
            18 => 'Химикаты и сода',
            19 => 'Строительные грузы',
            20 => 'Промышленное сырье и формовочные материалы',
            21 => 'Гранулированные шлаки',
            22 => 'Огнеупоры',
            23 => 'Цемент',
            24 => 'Лесные грузы',
            25 => 'Сахар',
            26 => 'Мясо и масло животное',
            27 => 'Рыба',
            28 => 'Картофель, овощи и фрукты',
            29 => 'Соль поваренная',
            30 => 'Остальные продовольственные грузы',
            31 => 'Промышленные товары народного потребления',
            32 => 'Хлопок',
            33 => 'Сахарная свекла и семена',
            34 => 'Зерно',
            35 => 'Продукты перемола',
            36 => 'Комбикорма',
            37 => 'Живность',
            38 => 'Жмыхи',
            39 => 'Бумага',
            40 => 'Перевалка грузов с водного на ж/д транспорт',
            41 => 'Импортные грузы без учета транзита',
            42 => 'Грузы в контейнерах',
            43 => 'Остальные и сборные грузы',
            44 => 'Транзит в импортных грузах',
            51 => 'Порожние вагоны',
            52 => 'Нерабочий парк',
            53 => 'Переадресовка',
            54 => 'Всего по отправлению',
            55 => 'Всего по прибытию',
        );

        $this->date = $date;
        $this->dateEpl = explode(".", $date);
        $this->d = $this->dateEpl[0];
        $this->m = $this->dateEpl[1];
        $this->y = $this->dateEpl[2];
    }

    // Метод, возвращающий коды грузов Плана/Факта
    // Параметр $typeData должен принимать либо 'PLAN', либо 'FACT'
    private function getKodGryz($typeData)
    {
        $result = array();
        $strKodGr = implode(",", array_keys($this->arrGruz)); // Строка кодов грузов из шаблона Excel
        $sql = "select ID_GRUZ_FACT, ID_GRUZ_PLAN from MURATOV.NSI_GRUZ_FP where ID_TEMPLATE in ($strKodGr) order by ID_TEMPLATE";
        Database::connect();
        $result = Database::select($sql);        
        Database::disconnect();
        switch ($typeData) {
            case 'FACT':
                for ($i = 0; $i < count($result); $i++) {
                    $arrKodGruz[$i] = $result[$i]['ID_GRUZ_FACT'];
                }
                break;

            case 'PLAN':
                for ($i = 0; $i < count($result); $i++) {
                    $arrKodGruz[$i] = $result[$i]['ID_GRUZ_PLAN'];
                }
                break;            
        }
        return $arrKodGruz;
    }

    // Метод, возвращающий данные Плана по дороге и ЦФТО
    public function getDataPlan($pokazDoroga, $pokazCfto)
    {
        $sqlDatePlan = $this->y . "-" . $this->m . "-01"; // Дата для sql-запроса
        $strKGrArr = $this->getKodGryz('PLAN'); // Коды грузов для плана     
        $strKGr = implode(",", array_diff($strKGrArr, array(''))); // Коды грузов в строку

        $sql = "select ID_POKAZ, POGRUZKA, NACHISLENO, GRUZ from MURATOV.PLAN_MONEY where DT = '$sqlDatePlan' and GRUZ in ($strKGr) and (ID_POKAZ = $pokazDoroga or ID_POKAZ = $pokazCfto) order by GRUZ, ID_POKAZ";

        Database::connect();
        $resultData = Database::select($sql);
        Database::disconnect();

        // Формирование корректного массива с данными под шаблон Excel
        $k = 0;
        for ($i = 0; $i < count($resultData); $i++) {
            if ($resultData[$i]['POGRUZKA'] == -1) {
                $resultData[$i]['POGRUZKA'] = 0;
                if ($resultData[$k]['POGRUZKA'] == -1) {
                    $resultData[$k]['POGRUZKA'] = 0;
                }
            }
            
            $arrData[$k] = array(
                'GRUZ' => $resultData[$i]['GRUZ'],
                'POGRUZKA' => $resultData[$i]['POGRUZKA'] + $resultData[$i = $i + 1]['POGRUZKA'],
                'NACHISLENO' => $resultData[$i - 1]['NACHISLENO'] + $resultData[$i]['NACHISLENO'],
            );

            $k++;
        }

        for ($i = 0; $i < count($strKGrArr); $i++) {
            $valGruz = $strKGrArr[$i];
            $keySearchArray = $this->array_column($valGruz, $arrData);
            $arrDataNew[$i] = array(
                'GRUZ' => $valGruz,
                'POGRUZKA' => $arrData[$keySearchArray]['POGRUZKA'],
                'NACHISLENO' => $arrData[$keySearchArray]['NACHISLENO'],
            );
        }
        return $arrDataNew;
    }

    // Метод, возвращающий данные Факта по дороге и ЦФТО (Полностью аналогичен getDataPlan() только построен под показатели Факта)
    public function getDataFact($pokazDoroga, $pokazCfto)
    {
        $sqlDateFact = $this->y . "-" . $this->m . "-" . $this->d;
        $strKGrArr = $this->getKodGryz('FACT');        
        $strKGr = implode(",", array_diff($strKGrArr, array('')));

        $sql = "select ID_POKAZ, POGRUZKA, NACHISLENO, GRUZ from MURATOV.FACT_MONEY where DT = '$sqlDateFact' and GRUZ in ($strKGr) and (ID_POKAZ = $pokazDoroga or ID_POKAZ = $pokazCfto)  order by GRUZ, ID_POKAZ";

        Database::connect();
        $resultData = Database::select($sql);
        Database::disconnect();

        $k = 0;
        for ($i = 0; $i < count($resultData); $i++) {
            if ($resultData[$i]['POGRUZKA'] == -1) {
                $resultData[$i]['POGRUZKA'] = 0;
                if ($resultData[$k]['POGRUZKA'] == -1) {
                    $resultData[$k]['POGRUZKA'] = 0;
                }
            }

            $arrData[$k] = array(
                'GRUZ' => $resultData[$i]['GRUZ'],
                'POGRUZKA' => $resultData[$i]['POGRUZKA'] + $resultData[$i = $i + 1]['POGRUZKA'],
                'NACHISLENO' => $resultData[$i - 1]['NACHISLENO'] + $resultData[$i]['NACHISLENO'],
            );

            $k++;
        }

        for ($i = 0; $i < count($strKGrArr); $i++) {
            $valGruz = $strKGrArr[$i];
            $keySearchArray = $this->array_column($valGruz, $arrData);
            if (!isset($arrData[$keySearchArray])) {
                continue;
            }
            $arrDataNew[$i] = array(
                'GRUZ' => $valGruz,
                'POGRUZKA' => $arrData[$keySearchArray]['POGRUZKA'],
                'NACHISLENO' => $arrData[$keySearchArray]['NACHISLENO'],
            );
        }

        return $arrDataNew;
    }

    // Формирование excel-таблицы
    public function GetExcelTable()
        {
            $templateExcel = "template/PogrVsego.xlsx"; // Расположение шаблона        
            $arHeadStyle = array( //Задаем цвет текста (красный) для ячеек с отрицательным значением
                'font'  => array(
                    'color' => array('rgb' => 'FF0000'),
            ));
            $fName = "PogrVsego_{$this->date}.xls"; // Название файла для сохранения
            $cell = 9;// С какой ячейки начинать ввод данных
            $countDayM = Utils::NumDayinMonth1($this->m, $this->y); // Количество дней в месяце
            
            $xls = new PHPExcel();
            
            $xls = PHPExcel_IOFactory::load($templateExcel); // Получаем объект документа Excel

            //Прописываем названия отчетов в каждом листе
            for ($i = 0; $i < count($this->titleSpravka); $i++) {
                $xls->setActiveSheetIndex($i);
                $sheetActive = $xls->getActiveSheet(); // Получаем активный лист
                $sheetActive->setCellValue("C1", $this->titleSpravka[$i] . "за");
                $sheetActive->mergeCells("C1:N1"); // Объединение ячеек для названия отчета
                $sheetActive->getStyle('C1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // Выравнивание текста в ячейке по центру

                $sheetActive->setCellValue("C2", $this->getMonthNumb($this->m));
                $sheetActive->mergeCells("C2:N2"); // Объединение ячеек для названия отчета
                $sheetActive->getStyle('C2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // Выравнивание текста в ячейке по центру

                $sheetActive->setCellValue("C3", "с 01." . $this->m . "." . $this->y . " по " . $this->d . "." . $this->m . "." . $this->y);
                $sheetActive->mergeCells("C3:N3"); // Объединение ячеек для названия отчета
                $sheetActive->getStyle('C3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // Выравнивание текста в ячейке по центру
            }            

            // Вводим значения в ячейки (PLAN)
            for ($i = 0; $i < 7; $i++) {
                $arrDataPlan[$i] = array($this->getDataPlan($this->pokazPlanDoroga[$i], $this->pokazPlanCfto[$i]));

                $xls->setActiveSheetIndex($i); // Устанавливаем индекс активного листа
                $sheetActive = $xls->getActiveSheet(); // Получаем активный лист

                for ($j = 0; $j < 56; $j++) {
                    $gruz = $arrDataPlan[$i][0][$j]['GRUZ'];
                    $pogruzka = $arrDataPlan[$i][0][$j]['POGRUZKA'] / $countDayM * $this->d;
                    $nachisleno = $arrDataPlan[$i][0][$j]['NACHISLENO'] / $countDayM * $this->d;

                    if ($nachisleno < 0 && $nachisleno >= -1) {
                        $nachisleno = 0;
                    }

                    $rowTableGruz = array_search($gruz, $this->getKodGryz('PLAN'));                    

                    if ($gruz == 50) {
                        if ($i == 3 || $i == 4 || $i == 5 || $i == 6) {
                            continue;
                        }  

                        $rowTableGruz = 48;
                        $pogruzka = '';
                    }

                    $cellPogrPlan = "C" . ($rowTableGruz + $cell);
                    $cellNachisPlan = "G" . ($rowTableGruz + $cell);                    

                    $sheetActive->setCellValue($cellPogrPlan, $pogruzka); // Вставляем значение Плана Погрузки
                    $sheetActive->setCellValue($cellNachisPlan, $nachisleno); // Вставляем значение Плана Начисления

                    if ($j == 1) {
                        $sheetActive->setCellValue("C9", "777");
                    }
                }
            }

            // Вводим значения в ячейки (FACT)
            for ($i = 0; $i < 8; $i++) {
                $arrData[$i] = array($this->getDataFact($this->pokazFactDoroga[$i], $this->pokazFactCfto[$i]));

                if ($i == 4) {
                    $arrData[$i+3] = array($this->getDataFact($this->pokazFactDoroga[$i+3], $this->pokazFactCfto[$i+3]));
                    for ($key = 0; $key < count($arrData[4][0]); $key++) {
                        if (!isset($arrData[4][0][$key])) {
                            continue;
                        }
                        $arrData[4][0][$key]['POGRUZKA'] += $arrData[7][0][$key]['POGRUZKA'];
                        $arrData[4][0][$key]['NACHISLENO'] += $arrData[7][0][$key]['NACHISLENO'];
                    }
                }

                if ($i == 7) {
                    unset($arrData[7]);
                    continue;
                }

                for ($j = 0; $j < 56; $j++) {

                    if (!isset($arrData[$i][0][$j]) || $i == 7) {
                        continue;
                    }

                    $gruz = $arrData[$i][0][$j]['GRUZ'];
                    $pogruzka = $arrData[$i][0][$j]['POGRUZKA'];
                    $nachisleno = $arrData[$i][0][$j]['NACHISLENO'];

                    if ($nachisleno < 0 && $nachisleno >= -1) {
                        $nachisleno = 0;
                    }

                    $rowTableGruz = array_search($gruz, $this->getKodGryz('FACT'));

                    $xls->setActiveSheetIndex($i); // Устанавливаем индекс активного листа
                    $sheetActive = $xls->getActiveSheet(); // Получаем активный лист

                    if ($gruz == 39) {
                        if ($i == 3 || $i == 4 || $i == 5 || $i == 6) {
                            continue;
                        }
                        $rowTableGruz = 48;
                        $pogruzka = '';
                    }
                    $cellPogrFact = "D" . ($rowTableGruz + $cell);
                    $cellNachisFact = "H" . ($rowTableGruz + $cell);

                    $sheetActive->setCellValue($cellPogrFact, $pogruzka); // Вставляем значение Плана Погрузки
                    $sheetActive->setCellValue($cellNachisFact, $nachisleno); // Вставляем значение Плана Начисления

                }
            }

            // Добавлен цикл с 12.12.2019, для дублирования ввода плановых показателей груза Каменный уголь
            for ($list = 0; $list < 7; $list++) {
                $xls->setActiveSheetIndex($list); // Устанавливаем индекс активного листа
                $sheetActive = $xls->getActiveSheet(); // Получаем активный лист
                $sheetActive->setCellValue("C9", $arrDataPlan[$list][0][0]['POGRUZKA'] / $countDayM * $this->d);
                $sheetActive->setCellValue("G9", $arrDataPlan[$list][0][0]['NACHISLENO'] / $countDayM * $this->d);
            }

            // Установка красного шрифта для отрицательных значений в столбцах E, I, M 
            for ($list = 6; $list > -1; $list--) {
                $xls->setActiveSheetIndex($list);
                $sheetActive = $xls->getActiveSheet();

                for ($i = 7; $i < 59; $i++) {
                    $cellE = 'E' . $i;
                    $cellI = 'I' . $i;
                    $cellM = 'M' . $i;
                    $arrCell = array($cellE, $cellI, $cellM);
                    for ($j = 0; $j < 3; $j++) {
                        $cellAct = $xls->getActiveSheet()->getCell($arrCell[$j])->getCalculatedValue();
                        if ($cellAct < 0) {            
                            $sheetActive->getStyle($arrCell[$j])->applyFromArray($arHeadStyle);
                        }
                    }            
                }
            }           

            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' . $fName . '"');
            header('Cache-Control: max-age=0');
            $objWriter = PHPExcel_IOFactory::createWriter($xls, 'Excel5');
            $objWriter->save('php://output');
            return $bool;
        }

    private function getMonthNumb($m)
    {
        $arrMonth = array(1 => "январь", 2 => "февраль", 3 => "март", 4 => "апрель", 5 => "май", 6 => "июнь", 7 => "июль", 8 => "август", 9 => "сентябрь", 10 => "октябрь", 11 => "ноябрь", 12 => "декабрь");
        return $arrMonth[$m];
    }

    private function array_column($id, $array) {
        foreach ($array as $key => $val) {
            if ($val['GRUZ'] == $id) {
                return $key;
            }
        }
        return null;
    }
}
