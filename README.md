# stat.gov.kz

Парсер для stat.gov.kz - реестр юридических лиц Казахстана.

В качестве основы был взят готовый парсер (http://analyst.kz/index.php/category-list/business-intelligense/19-%D0%BF%D0%B0%D1%80%D1%81%D0%B5%D1%80-%D0%B4%D0%BB%D1%8F-stat-gov-kz-%D1%80%D0%B5%D0%B5%D1%81%D1%82%D1%80-%D1%8E%D1%80%D0%B8%D0%B4%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8%D1%85-%D0%BB%D0%B8%D1%86).
Код был немного переписан для удобства.

Источник данных: http://stat.gov.kz/juridical/list

Вкратце как работает парсер:

1. С сайта stat.gov.kz скачиваются 16 архивов с БИН-ами по областям и городам республиканского значения.
2. Распаковываем архивы -> получаем 16 Excel файлов.
3. Открываем по очереди Excel файлы и из каждого листа копируем данные в общий csv-файл.
