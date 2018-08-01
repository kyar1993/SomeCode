<?
require_once 'vendor/autoload.php';
require_once("XLSXToJSONParser.php");

header('Content-type: application/json; charset=utf-8');
set_time_limit(18000);

$status = [
	"type" => "switcher",
        "value" => 'true',// в парсере рандомит на 110
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

$parser = new XLSXToJSONParser($par);

$key = 'demo_excel_' . time();
// Пробуем извлечь $data из кэша.
// $data = $cache->get($key);

// if ($data === false) {
    // $data нет в кэше, получаем
    $data = $parser->getArray();

    // Сохраняем значение $data в кэше. Данные можно получить в следующий раз.
//     $cache->set($key, $data);
// }

//to output json
echo($data);

//to output array
// echo '<pre>';
// var_dump($data);
// echo '</pre>';
?>