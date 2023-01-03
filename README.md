# Начало

## Введение

Office Open XML (OOXML, DOCX, XLSX, PPTX, проект ISO (International Organization for Standardization / IEC International Electrotechnical Commission IS 29500:2008) — серия форматов файлов для хранения электронных документов пакетов офисных приложений — в частности, Microsoft Office.

Формат представляет собой zip-архив, содержащий текст в виде XML, графику и другие данные, которые ранее хранились в двоичных форматах DOC, XLS и т. д. 

## Работа приложения

Основная идея формирования отчета docx — это замена текста в word’е исходном файле с текстовыми метками, используются ключевые текстовые метки, которые будут храниться внутри тега (w:r) Run в его теге (w:t) Text, а также вставка табличных данных по меткам в таблицах.

Для таблиц устанавливается уникальный идентификатор таблицы Id для каждой таблицы в ее первой строке, чтобы осуществлять поиск меток для вывода табличных данных в ячейках таблицы. Имеется возможность вывода вложенных таблиц, работа с хедером и футером документа docx.

Более подробно с принципом работы можно ознакомиться в презентации в файле Description/АЭБ_OpenXml.pptx. Имеется серверная и клиентская части приложения. Со стороны клиента формируется и заполняется объект класса WinWordRenderModel.

Для отображения используется серверная часть. Для примеров используются юнит-тесты, нужно их запустить, исходные файлы docx и результаты работы будут выведены в папке bin\Debug\net5.0\TestWinWord.
