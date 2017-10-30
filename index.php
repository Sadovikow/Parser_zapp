<?
    if($_POST['file'])
    {
      function getXLS($xls){
          //include_once 'Classes/PHPExcel/IOFactory.php'
        include_once 'PHPExcel.php';
        $objPHPExcel = PHPExcel_IOFactory::load($xls);
        $objPHPExcel->setActiveSheetIndex(0);
        $aSheet = $objPHPExcel->getActiveSheet();

        //этот массив будет содержать массивы содержащие в себе значения ячеек каждой строки
        $array = array();
        //получим итератор строки и пройдемся по нему циклом
        foreach($aSheet->getRowIterator() as $row){
          //получим итератор ячеек текущей строки
          $cellIterator = $row->getCellIterator();
          //пройдемся циклом по ячейкам строки
          //этот массив будет содержать значения каждой отдельной строки
          $item = array();
          foreach($cellIterator as $cell){
            //заносим значения ячеек одной строки в отдельный массив
              //array_push($item, iconv('utf-8', 'cp1251', $cell->getCalculatedValue()));
              array_push($item, iconv('utf-8', 'utf-8', $cell->getCalculatedValue()));
          }
          //заносим массив со значениями ячеек отдельной строки в "общий массв строк"
          array_push($array, $item);
        }
        return $array;
      }


    $xlsData = getXLS($_POST['file']); //извлеаем данные из XLS
    }

    $filename = $_POST['file'];
    if($filename == '' || !$filename) {
        $filename = 'table.xlsx';
    }
?>

<form action="index.php?action=parse" method="post" id="subaru_parsing">
    <h2>Парсинг сайта zzap.ru</h2>
    <?if($_GET['action'] == 'parse'):?>
    <h4>В таблице <?=count($xlsData);?> позиций</h4>
    <?endif;?>
    <div class="option">
        <label>Время запроса одной записи:</label> <input type="text" value="2000" name="time" />
    </div>
    <div class="option">
        <label>Название файла:</label> <input type="text" value="<?=$filename?>" name="file" id="parsingfile" /> <i class="checking"><input type="button" id="checkfile" value="Проверить файл" /></i>
    </div>
    <div class="option">
        <input type="submit" value="Парсить!" />
    </div>

</form>
<?
if($_GET['action'] == 'parse' && $_POST['time'] != ''):

    $xlsData = array_diff($xlsData, array(''));

    foreach($xlsData as $key=>$article): // Перебераем массив артикулов
         // преобразуем массив в URL-кодированную строку
        $vars = http_build_query($paramsArray);
        // создаем параметры контекста
        $options = array(
            'http' => array(
                        'method'  => 'POST',  // метод передачи данных
                        'header'  => 'Content-type: application/x-www-form-urlencoded',  // заголовок
                        'content' => $vars,  // переменные
                    )
        );
        $context  = stream_context_create($options);  // создаём контекст потока
        if($result = file_get_contents('https://www.zzap.ru/public/search.aspx?rawdata='.$article[1].'&class_man=SUBARU'))
        {
            //$result = ; //отправляем запрос

    ?>
        <div id="parse_<?=$key?>" style="display: none;">
    <?
        var_dump($result); // Достаём всю таблицу со стороннего сайта
    ?>
    </div>
    <script>
        function loader() {
            $('#loader').html('Подождите, идёт загрузка данных...');
        }

        function parser() { // Наш парсер, достаём нужные данные и выводим их
            /* Номера колонок выводимой инфы */
                var article_column = 4;
                var name_column = 5;
                var refresh_column = 6;
                var instock_column = 7;
                var price_column = 8;
                var shop_column = 9;
            /* /Номера колонок выводимой инфы */

            var res = $('#parse_<?=$key?> #ctl00_BodyPlace_SearchGridView_DXDataRow3').html();
            var article = $('#parse_<?=$key?> #ctl00_BodyPlace_SearchGridView_DXDataRow3 td:nth-child('+article_column+') span').html();
            var name = $('#parse_<?=$key?> #ctl00_BodyPlace_SearchGridView_DXDataRow3 td:nth-child('+name_column+')').text();
            var refresh = $('#parse_<?=$key?> #ctl00_BodyPlace_SearchGridView_DXDataRow3 td:nth-child('+refresh_column+')').text();
            var instock = $('#parse_<?=$key?> #ctl00_BodyPlace_SearchGridView_DXDataRow3 td:nth-child('+instock_column+') span').text();
            var price = $('#parse_<?=$key?> #ctl00_BodyPlace_SearchGridView_DXDataRow3 td:nth-child('+price_column+') ').text();
            var shop = $('#parse_<?=$key?> #ctl00_BodyPlace_SearchGridView_DXDataRow3 td:nth-child('+shop_column+') a').html();
            var position = $('#parse_<?=$key?> a:contains("ООО СЕРВИС-СУБАРУ")').parents('tr.dxgvDataRow_ZzapAqua').attr('id');
            var position = parseInt(position.replace(/\D+/g,""));
            // Так как в классах содержится номер строки, берём позицию от туда, (#ctl00_BodyPlace_SearchGridView_DXDataRow3) <- Вот класс, Его позиция ровна 3-2=1
            position = parseInt(position) - 2;
            // но нужно учитывать что ПЕРВАЯ строка записана как 3яя, поэтому надо отнять 2
            if(shop == 'ООО СЕРВИС-СУБАРУ') shop += '<br><b>Вы на первом месте!</b>';
            /* Выводим инфу в таблицу и удаляем временную информацию */
            $('#parse_result tbody').append('<tr><td><a href="https://www.zzap.ru/public/search.aspx?rawdata=<?=$article[1]?>&class_man=SUBARU">'+article+'</a></td><td>'+name+'</td><td>'+refresh+'</td><td>'+instock+'</td><td>'+price+'</td><td>'+shop+'</td><td>'+position+'</td></tr>');
            $('#parse_<?=$key?>').remove();
            $('#loader').html('Результат:');
            <?if(count($xlsData)-1 == $key):?>
                $('#resultat').html('');
            <?endif;?>
        }

        setTimeout(parser, <?echo $_POST['time']*count($xlsData); ?>); // Тут желательно поставить обработчик, Который будет понимать, что результат пришел
        loader();
    </script>
    <?
        }
    ?>
    <?endforeach;?>
<?endif;?>
<div id="resultat" >
    Загрузка . . .
</div>
<span id="loader">Результат:</span>
<table id="parse_result" border="1px">
    <thead>
        <th>Артикул</th>
        <th>Название</th>
        <th>Обновление</th>
        <th>Наличие</th>
        <th>Цена</th>
        <th>Магазин лидер</th>
        <th>Ваша позиция</th>
    </thead>
    <tbody>

    </tbody>
</table>

<script>
    $(document).ready(function() {

        function checkFile(name) { // Проверка на наличие файла
            $.ajax({
            url: name,
            type:'HEAD',
            error:
                function(){
                    $('i.checking input').val('Файл не найден');
                },
            success:
                function(){
                    $('i.checking').html('✓');
                }
            });
        }

        $('#checkfile').on('click', function() {
            var file = $('#parsingfile').val();
            checkFile(file);
        });

    });
</script>

<style>

    table#parse_result {
        background: #fff;
    }

    table#parse_result {
        border: 2px solid #aaa;
        width: 100%;
    }

    table#parse_result thead {
        font-weight: 600;
        font-size: 16px;
    }

    table#parse_result tr td {
        padding: 10px;
    }

    table#parse_result th {
        padding: 20px 10px;
    }

    #subaru_parsing label {
        float: left;
        width: 150px;
    }

    #subaru_parsing .option {
        margin: 10px 0;
        padding: 5px;
    }

    #subaru_parsing input[type="checkbox"] {
        -webkit-appearance: checkbox !important;
    }
</style>