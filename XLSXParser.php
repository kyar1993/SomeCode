<?php

require_once 'Classes/PHPExcel/IOFactory.php';

class XLSXParser
{
	private $objPHPExcel;

    /**
     * objPHPExcel->getActiveSheet()
     */
    private $aSheet;

    /**
     * Путь к файлу
     * @var string
     */
    private $file;

    /**
     * Конечная строка
     * @var int
     */
    private $endRow;

    /**
     * Конечная колонка
     * @var string
     */
    private $endCell;

	/**
	* Возвращать формат json?
	*/
    private $json;

	/**
	* Статус
	*/
    private $status;

    /**
     * XLSXParser constructor.
     * @param array $params [
     	filename Путь к файлу
        endrow Конечная строка
        endcolumn Конечная колонка
        returnjson формат возвращаемый true = json, false = array
        status блок статуса
     ]
     * 
     */
    function __construct(array $params)
    {
    	if ($params['filename'] == '') {
    		throw new Exception('Не указан файл');
    	}

        $this->file = $params['filename'];
        $this->endRow = $params['endrow'] ?? 20000;
        $this->endCell = $params['endcolumn'] ?? 'O';
        $this->json = $params['returnjson'] ?? 'false';
        $this->status = $params['status'] ?? '';
    }

    /**
     * Возвращает данные преобразованные из файла .xlsx
     * @return array|string
     * @throws Exception
     */
    public function getArray()
    {
        $type = PHPExcel_IOFactory::identify($this->file);
        $objReader = PHPExcel_IOFactory::createReader($type);
        $objReader->setReadFilter(new ReadFilter(1, $this->endRow));
        $objReader->setReadDataOnly(true);
        $this->objPHPExcel = $objReader->load($this->file);
        $this->aSheet = $this->objPHPExcel->getActiveSheet();

        // получаем 1 строку (заголовки)
        $headerRow = [];
        $rows = [];
        $resArray = ['success' => true, 'msg' => ''];

        $headersRus = [
            'A' => 'Площадка',
            'B' => 'Визиты',
            'C' => 'Показы рекламы',
            'D' => 'ROI',
            'E' => 'Клики',
            'F' => 'Просмотры',
            'G' => 'CTR',
            'H' => 'Отказы',
            'I' => 'Статус',
            'J' => 'J нужно добавить описание заголовка в парсер',
            'K' => 'K нужно добавить описание заголовка в парсер',
            'L' => 'L нужно добавить описание заголовка в парсер',
            'M' => 'M нужно добавить описание заголовка в парсер',
            'N' => 'N нужно добавить описание заголовка в парсер',
            'O' => 'O нужно добавить описание заголовка в парсер'
        ];

        foreach ($this->aSheet->getRowIterator(1, 1) as $fRow) {
            foreach ($fRow->getCellIterator('A', $this->endCell) as $key => $fCell) {
                $headerRow[$key] = $fCell->getCalculatedValue();
                $resArray['data']['headers'][] = [
                    'key' => $headersRus[$key],
                    'value' => $fCell->getCalculatedValue(),
                    'description' => ''
                ];
            }
        }

        $obj = [];

        // перебираем строки
        foreach ($this->aSheet->getRowIterator(1, $this->endRow) as $row) {

            // индекс текущей строки
            $rowIndex = $row->getRowIndex();

            if ($rowIndex == 1) {
            	continue;
            }

            // значения ячеек строки
            $item = [];

                        // заголовок текущей строки (значение колонки А)
            $cellA = $this->aSheet->getCellByColumnAndRow(0, $rowIndex)->getCalculatedValue();

            // уровень текущей строки
            $currentRowLevel = $this->aSheet->getRowDimension($rowIndex)->getOutlineLevel();

            // уровень следующей строки
            $nextRowLevel = $this->aSheet->getRowDimension($rowIndex + 1)->getOutlineLevel();

            // уровень предыдущей строки
            $prevRowLevel = $this->aSheet->getRowDimension($rowIndex - 1)->getOutlineLevel();

            // перебираем ячейки
            foreach ($row->getCellIterator('A', $this->endCell) as $key => $cell) {
                $cellVal = $cell->getCalculatedValue();

				//статус меняем на кастомный из параметров
                if ($currentRowLevel > 0 && $key == 'O') {

                    if ($this->status != '') {
                        $this->status['value'] = array_rand(['true', 'false'], 1);
                        $item[$headerRow[$key]]['value'] = $this->status;
                    } else {
                        $item[$headerRow[$key]]['value'] = $cellVal;
                    }

                } else {
                	if (is_numeric($cellVal)) {
                    	$item[$headerRow[$key]]['value'] = round($cellVal, 2);
	                } else {
	                    $item[$headerRow[$key]]['value'] = $cellVal;
	                }
            	}
            }

            unset($this->objPHPExcel);

            // 0 уровень
            if ($currentRowLevel == 0) {

                if ($nextRowLevel > $currentRowLevel) {
                    $obj['row'] = $item;
                } else {
                    $resArray['data']['spoilers'][]['row'] = $item;
                }

                // 1 уровень
            } elseif ($currentRowLevel == 1) {

                if ($prevRowLevel < $currentRowLevel) {
                    $obj['shown'] = false;
                }

                if ($nextRowLevel == $currentRowLevel || $nextRowLevel > $currentRowLevel) {
                    $obj['inner']['spoilers'][] = ['row' => $item];
                } elseif ($nextRowLevel < $currentRowLevel) {
                    $obj['inner']['spoilers'][] = ['row' => $item];
                    $resArray['data']['spoilers'][] = $obj;

                    $obj = [];
                } else {
                    echo 'error at level №1';
                }

                // 2 уровень
            } elseif ($currentRowLevel == 2) {

                if ($prevRowLevel < $currentRowLevel) {
                    $obj['inner']['spoilers'][0]['shown'] = false;
                }

                if ($nextRowLevel == $currentRowLevel || $nextRowLevel > $currentRowLevel) {
                    $obj['inner']['spoilers'][0]['inner']['spoilers'][] = ['row' => $item];
                } elseif ($nextRowLevel < $currentRowLevel) {
                    $obj['inner']['spoilers'][0]['inner']['spoilers'][] = ['row' => $item];
                    $resArray['data']['spoilers'][] = $obj;

                    $obj = [];
                } else {
                    echo 'error at level №2';
                }

                // 3 уровень
            } elseif ($currentRowLevel == 3) {

                if ($prevRowLevel < $currentRowLevel) {
                    $obj['inner']['spoilers'][0]['inner']['spoilers'][0]['shown'] = false;
                }

                if ($nextRowLevel == $currentRowLevel) {
                    $obj['inner']['spoilers'][0]['inner']['spoilers'][0]['inner']['spoilers'][] = ['row' => $item];
                } elseif ($nextRowLevel < $currentRowLevel) {
                    $obj['inner']['spoilers'][0]['inner']['spoilers'][0]['inner']['spoilers'][] = ['row' => $item];
                    $resArray['data']['spoilers'][] = $obj;

                    $obj = [];
                } else {
                    echo 'error at level №3';
                }


                // уровень не описан
            } else {
                throw new Exception('Данный уровень не описан!!!');
            }

        }

        //echo '<br><br><br> !!! The end of parsing !!! <br><br><br>';

        if ($this->json == 'true') {
            return json_encode($resArray, JSON_UNESCAPED_UNICODE);
        } else {
            return $resArray;
        }
    }

    /**
     * Значение заголовка (колонки А) родительского элемента(с уровнем вложенности 0)
     * @param int $rowIndex номер текущей строки
     * @return array
     */
    private function getParent(int $rowIndex): array
    {
        $rowLevel = $this->aSheet->getRowDimension($rowIndex)->getOutlineLevel();
        $res[] = $this->aSheet->getCellByColumnAndRow(0, $rowIndex)->getCalculatedValue();

        for ($i = $rowIndex - 1; $i != 0; $i--) {
            $curLevel = $this->aSheet->getRowDimension($i)->getOutlineLevel();

            if ($curLevel < $rowLevel) {
                $res[] = $this->aSheet->getCellByColumnAndRow(0, $i)->getCalculatedValue();
                $rowLevel = $curLevel;
            } else {
                continue;
            }

            if ($curLevel == 0) {
                break;
            }
        }

        return array_reverse($res);
    }

    private function change_key($key, $new_key, &$arr, $rewrite = true)
    {
        if (!array_key_exists($new_key, $arr) || $rewrite) {
            $arr[$new_key] = $arr[$key];
            unset($arr[$key]);
            return true;
        }
        return false;
    }
}

class ReadFilter implements PHPExcel_Reader_IReadFilter
{
    private $startRow = 1;

    private $endRow = 0;

    /**
     * ReadFilter constructor.
     * @param int $startRow
     * @param int $endRow
     */
    public function __construct(int $startRow, int $endRow)
    {
        $this->startRow = $startRow;
        $this->endRow = $endRow;
    }

    public function readCell($column, $row, $worksheetName = '')
    {
        if ($row > 0 && $row <= $this->endRow) {
            return true;
        }

        return false;
    }
}