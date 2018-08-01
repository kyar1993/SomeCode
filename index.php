<?
include_once("Classes/PHPExcel.php");
include_once("XLSXParser.php");

set_time_limit(18000);

$status = [
	"type" => "switcher",
        "value" => 'true',
        "data" => [
	        "id"=> 8,
	        "type" => "campaign",
	        "source" => "Yandex",
	        "url" => "/dist/jsondata/vue/report/switch.json"
	    ]
];

$par = [
	'filename' => "test.xlsx",
	'endrow' => 18100,
	'endcolumn' => 'O',
	'returnjson' => 'true',
	'status' => $status
];

$parser = new XLSXParser($par);

header('Content-type: application/json; charset=utf-8');

$key = 'demo_excel_' . time();
// Пробуем извлечь $data из кэша.
// $data = $cache->get($key);

// if ($data === false) {
    // $data нет в кэше, получаем
    $data = $parser->getArray();

    // Сохраняем значение $data в кэше. Данные можно получить в следующий раз.
//     $cache->set($key, $data);
// }

//json
echo($data);

//array
// echo '<pre>';
// var_dump($data);
// echo '</pre>';


?>
<?/* </body>
</html>*/?>