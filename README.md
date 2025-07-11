# DHCPScan
Скрипт предназначен для мониторинга DHCP-серверов с целью выявления новых аренд IP-адресов, а также для просмотра зарезервированных IP-адресов.
Он реализован в двух режимах:

CLI (командная строка) — автоматический режим для задач типа Task Scheduler;

GUI (графический интерфейс) — визуальный режим с интерактивной формой и удобным управлением.


CLI-режим (командная строка) <br>
Используется для автоматизации (например, через Task Scheduler). <br>
Функции:
- Поиск новых IP-аренд по сравнению со вчерашними;
- Фильтрация по VLAN;
- Логирование в CSV;
- Отправка отчёта на почту;
- Очистка старых логов.

GUI-режим (графический интерфейс) <br>
Интерактивная форма для ручного использования. <br>
Функции:
- Выбор DHCP-сервера и VLAN;
- Поиск новых арендаторов IP;
- Просмотр зарезервированных IP-адресов;
- Удобный вывод результатов в списке.

## Скриншоты (Графический интерфейс)

### Главное окно  
<img src="screens/sc1.png">

Для поиска необходимо выбрать DHCP-сервер

<img src="screens/sc2.png">

Нажать кнопку "SEARCH"

<img src="screens/sc3.png">

Также можно произвести поиск по конкретному VLAN, выбрав его из выпадающего списка.

<img src="screens/sc4.png">
<img src="screens/sc6.png">

В случае отсутствия новых аренд отобразится соответствующее уведомление.

<img src="screens/sc5.png">

Также можно отобразить все зарезервированные IP-адреса, нажав на кнопку "RESERVED"

<img src="screens/sc7.png">
<img src="screens/sc8.png">

## Скриншоты (Командная строка)

Выполнение в командной строке 

<img src="screens/sc9.png">

Отчет на почту 

<img src="screens/sc10.png">
