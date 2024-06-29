﻿\# Enterprise\_analysis\_Excelforms

Для студентов экономического направления в аграрном ВУЗе в рамках учебной детятельности ставятся задачи для вычисления и анализа состояния предприятия. В процесс решения входит расчёт формул, для демонстрации состояния предприятия и прогнозирование результатов деятельности в будущем. Данные для расчета формул берут из форм бухгалтерской отчетности, имеющей одинаковую структуру. Решение задачи подразумевает поиск необходимых значений среди других не используемых при расчете основных формул данных.

Разработан инструмент для расчета экономических показателей предприятия, используемый специалистами для оценки финансового состояния фирмы. Для работы с таблицами используются формы Excel. Программа рассчитывает показатели: рентабельность, норма прибыли, фондовооруженность и коэффициент оборачиваемости. Для дополнительного удобства и проверки есть функция просмотра таблицы внутри программы

\## Алгоритм работы

1. В программу интегрированы формулы расчета. Для выполнения алгоритма загружаются Excel формы отчетности в порядке очередности: №2, №5 апк, №6, №1.
1. Из массива данных выбираются нужные ячейки с числами и высчитывается выбранный показатель.
1. Расчеты делаются в очередном порядке их расположения, и программа запоминает уже загруженные данные. Для считывания одинаковых значений используемых при расчете повторное открытие таблицы не требуется.

\## Как пользоваться

![enter image description here](https://sun9-46.userapi.com/impg/MRMWzlYkw11occHhb45SLH1ouWSOFt75UJ0Jtg/2cju2zW5MrM.jpg?size=1044x525&quality=95&sign=afcb6fb7f4c32356c4d06f2d071f6279&type=album)

Для расчета показателей необходимо скачать формы бухгалтерской отчетности, а именно №1,2,6,5 апк (для примера приложены к проекту в папке "формы").

Расчеты производятся по порядку, чтобы программа запомнила вводимые данные:

\##### Рентабельность -> Норма Прибыли -> Фондовооруженность -> Коэффициент оборачиваемости.

1. Загрузите форму №2, нажав на кнопку \*\*Рентабельность\*\*. Программа найдет нужные данные и по имеющейся формуле рассчитает показатель, выведя результат возле кнопки показателя

![enter image description here](https://sun9-3.userapi.com/impg/8Hc2zkw4Xv4dupA-AtrGGYgg4VnFGuAGDn0s6A/Fxogypbb1lY.jpg?size=1046x520&quality=96&sign=4b7124a94f40b66cdaf47a8f5b6b32b1&type=album)


1. Повторите действия в соответствие с описанной последовательностью ![enter image description here](https://sun9-7.userapi.com/impg/j509Ctx7H8xrGXLoE--pTbOk3oHDgNkehMmyDg/OFXqcUE-o6k.jpg?size=1042x522&quality=96&sign=aae07187f059477c2a32cee321d7a4b8&type=album)



1. Приложение имеет функцию сохранения результата в текстовый файл. Нажмите кнопку \*\*Сохранить\*\* после проведения расчетов, введите имя файла и выберете путь сохранения

![enter image description here](https://sun9-51.userapi.com/impg/ocVhVGuAY\_VZxFyAiguBn-t-3Lb4JlrYBaiMaw/KX\_kmJHmzQI.jpg?size=1011x515&quality=96&sign=c2914020a4ed31d0613a33ca524354fa&type=album)

1. Кнопка \*\*Загрузить\*\* выводит таблицу Excel для предварительного просмотра

![enter image description here](https://sun9-69.userapi.com/impg/HHIUx-PkMkIHJa7h3V1j0OsTDet4T\_bSB8W8lw/1zlTTNCL25A.jpg?size=1046x521&quality=96&sign=c48161c6ff3c9f2493fa2c620726ba6c&type=album)

\## Стек технологии

- С++
- Visual Studio
- WinForm

\### [ElizavetaRychko] https://github.com/ElizavetaRychko)



