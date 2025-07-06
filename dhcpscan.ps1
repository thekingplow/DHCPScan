#v1.1 (01.08.2025)
#Developed by Danilovich M.D.










# --- CLI-режим при переданных аргументах ---
if ($args.Count -ge 1) {
    # Карта серверов
    $serverMap = @{
        "1" = "dhcp-01.domain.local"     #Указать свои dhcp серверы
        "2" = "dhcp-02.domain.local"
    }

    $arg1 = $args[0]
    $arg2 = if ($args.Count -ge 2) { $args[1] } else { $null }

    if (-not $serverMap.ContainsKey($arg1)) {
        Write-Host "Ошибка: неверный идентификатор сервера. Используйте 1 или 2." -ForegroundColor Red
        exit 1
    }

    $server = $serverMap[$arg1]
    $vlanFilter = $arg2  # может быть null

    # Сбор scopes и формирование карты
    $scopesMap = @{}
    try {
        $scopes = Get-DhcpServerv4Scope -ComputerName $server
        foreach ($scope in $scopes) {
            $thirdOctet = ($scope.ScopeId.ToString() -split '\.')[2]
            $scopesMap[$thirdOctet] = $scope.ScopeId
        }
    } catch {
        Write-Host "Ошибка при получении VLAN с $server : $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    # Получаем аренды
    $leases = @()

    if ($vlanFilter -and $scopesMap.ContainsKey($vlanFilter)) {
        $scopeId = $scopesMap[$vlanFilter]
        $leases = Get-DhcpServerv4Lease -ComputerName $server -ScopeId $scopeId
    } else {
        foreach ($scopeId in $scopesMap.Values) {
            $leases += Get-DhcpServerv4Lease -ComputerName $server -ScopeId $scopeId
        }
    }

    # Подготовка логов
    $logDir = "C:\Logs\Dhcp"
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null


        # Удаление логов старше 10 дней
    Get-ChildItem -Path $logDir -Filter "leases-$server-*.csv" | Where-Object {
        $_.LastWriteTime -lt (Get-Date).AddDays(-10)
    } | ForEach-Object {
        try {
            Remove-Item $_.FullName -Force -ErrorAction Stop
        } catch {
            # Ошибки подавляем или логируем при необходимости
        }
    }


    $today = Get-Date -Format "yyyy-MM-dd"
    $yesterday = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")

    $todayFile = Join-Path $logDir "leases-$server-$today.csv"
    $yesterdayFile = Join-Path $logDir "leases-$server-$yesterday.csv"

    $leases | Export-Csv -Path $todayFile -NoTypeInformation -Force

    if (Test-Path $yesterdayFile) {
        $yesterdayLeases = Import-Csv $yesterdayFile
        $yesterdayIPs = $yesterdayLeases.IPAddress

        $newLeases = $leases | Where-Object {
            $_.IPAddress -and ($_.IPAddress -notin $yesterdayIPs)
        }

if ($newLeases.Count -eq 0) {
    Write-Host "Нет новых арендаторов на $server за сегодня." -ForegroundColor Yellow
} else {
    Write-Host "Новые арендаторы на $server :`n" -ForegroundColor Green
    $body = @"
<html>
<head>
<style>
  table { border-collapse: collapse; }
  th, td { border: 1px solid #ddd; padding: 8px; font-family: Arial;  width: 25%; }
  th { background-color: #f2f2f2; }
</style>
</head>
<body>
<h3>Новые арендаторы за сегодня на сервере $server</h3>
<table>
  <tr>
    <th>IP-адрес</th>
    <th>MAC</th>
    <th>Имя</th>
    <th>Аренда до</th>
  </tr>
"@

    foreach ($lease in $newLeases) {
        $hostname   = if ($lease.HostName) { $lease.HostName } else { "(без имени)" }
        $expiryTime = if ($lease.LeaseExpiryTime) { $lease.LeaseExpiryTime.ToString("g") } else { "(нет даты)" }
        $mac        = if ($lease.ClientId) { $lease.ClientId } else { "(нет MAC)" }

        Write-Host "IP: $($lease.IPAddress) | MAC: $mac | Name: $hostname | До: $expiryTime"

        $body += "<tr><td>$($lease.IPAddress)</td><td>$mac</td><td>$hostname</td><td>$expiryTime</td></tr>`n"
    }

    $body += "</table></body></html>"

    # Параметры SMTP
    $smtpServer = "mail.smtpserver.by"
    $fromAddress = "from@mail.by"
    $toAddress = "to@mail.by"
    $subject = "Новые DHCP аренды на $server"

    try {
        Send-MailMessage -SmtpServer $smtpServer `
                         -From $fromAddress `
                         -To $toAddress `
                         -Subject $subject `
                         -Body $body `
                         -BodyAsHtml `
                         -Encoding ([System.Text.Encoding]::UTF8)
        Write-Host "Отчёт успешно отправлен на $toAddress" -ForegroundColor Cyan
    } catch {
        Write-Host "Ошибка отправки письма: $($_.Exception.Message)" -ForegroundColor Red
                }
            }

        }
    exit 0  # Прерываем скрипт — GUI не запускается
}

























Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'ru-RU'
#[System.Threading.Thread]::CurrentThread.CurrentCulture = 'ru-RU'


# Словарь для связи текста в ComboBox с реальными ScopeId
$scopesMap = @{}




function Get-DhcpData {
    param (
        [string]$Server,
        [System.Collections.Hashtable]$ScopesMap,
        [string]$SelectedScopeText,
        [ScriptBlock]$Fetcher
    )

    $results = @()

    if ($SelectedScopeText -ne "Все VLAN" -and $ScopesMap.ContainsKey($SelectedScopeText)) {
        $selectedScopeId = $ScopesMap[$SelectedScopeText]
        $results = & $Fetcher -Server $Server -ScopeId $selectedScopeId
    } else {
        $scopes = Get-DhcpServerv4Scope -ComputerName $Server
        foreach ($scope in $scopes) {
        $results += & $Fetcher -Server $Server -ScopeId $scope.ScopeId
        }
    }
    return $results
}



function Get-AllDhcpLeases {
    param (
        [string]$Server,
        [System.Collections.Hashtable]$ScopesMap,
        [string]$SelectedScopeText
    )
    return Get-DhcpData -Server $Server -ScopesMap $ScopesMap -SelectedScopeText $SelectedScopeText -Fetcher {
        param ($Server, $ScopeId)
        Get-DhcpServerv4Lease -ComputerName $Server -ScopeId $ScopeId
    }
}



function Get-AllDhcpReservations {
    param (
        [string]$Server,
        [System.Collections.Hashtable]$ScopesMap,
        [string]$SelectedScopeText
    )
    return Get-DhcpData -Server $Server -ScopesMap $ScopesMap -SelectedScopeText $SelectedScopeText -Fetcher {
        param ($Server, $ScopeId)
        Get-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId
    }
}















# === Создание формы ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "IP search"
$form.Size = New-Object System.Drawing.Size(800,700)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle # Установка фиксированного размера формы
$form.MaximizeBox = $false # Отключение кнопки максимизации
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold) # Устанавливаем стиль и размер шрифта для всех элементов формы


$scriptPath = $PSScriptRoot

# Установка иконки
$iconPath = Join-Path -Path $scriptPath -ChildPath "images\ip.ico" # Укажите путь к вашей иконке
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)


# Загружаем изображение из файла (замените путь на свой)
$imagePath = Join-Path -Path $scriptPath -ChildPath "images\bg.jpg"
$image = [System.Drawing.Image]::FromFile($imagePath)

# Устанавливаем изображение как фон формы
$form.BackgroundImage = $image
$form.BackgroundImageLayout = "Stretch"  # Растягиваем изображение на всю форму


# метка номера версии
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "v1.1 (01.08.2025)"
$labelVersion.Location = New-Object System.Drawing.Point(0, 0)
$labelVersion.Font = New-Object System.Drawing.Font("Arial", 7.5, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelVersion.AutoSize = $true  # Автоматический размер под текст
$labelVersion.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelVersion.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelVersion)


# Создание заголовка
$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = "IP SEARCH IN DHCP"
$labelTitle.Location = New-Object System.Drawing.Point(215, 50)
$labelTitle.Font = New-Object System.Drawing.Font("Arial", 26, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelTitle.AutoSize = $true  # Автоматический размер под текст
$labelTitle.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelTitle.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelTitle)


# === Метка DHCP сервера ===
$labelServer = New-Object System.Windows.Forms.Label
$labelServer.Text = "DHCP-server:"
$labelServer.Location = New-Object System.Drawing.Point(130, 165)
$labelServer.Size = New-Object System.Drawing.Size(200, 20)
$labelServer.AutoSize = $true  # Автоматический размер под текст
$labelServer.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelServer.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelServer)



# === ComboBox выбора сервера ===
$comboBoxServers = New-Object System.Windows.Forms.ComboBox
$comboBoxServers.Location = New-Object System.Drawing.Point(250, 162)
$comboBoxServers.Width = 140
$comboBoxServers.DropDownStyle = 'DropDownList'
$comboBoxServers.Items.AddRange(@("dhcp-01.domain.local", "dhcp-02.domain.local"))        #Указать свои dhcp серверы
$comboBoxServers.SelectedIndex = -1

$comboBoxServers.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$comboBoxServers.ForeColor = [System.Drawing.Color]::Black
$comboBoxServers.BackColor = [System.Drawing.Color]::WhiteSmoke

$comboBoxServers.Cursor = [System.Windows.Forms.Cursors]::Hand

$form.Controls.Add($comboBoxServers)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($comboBoxServers, "Выберите один из доступных DHCP серверов")



# === Метка VLAN ===
$labelVlan = New-Object System.Windows.Forms.Label
$labelVlan.Text = "VLAN"
$labelVlan.Location = New-Object System.Drawing.Point(130, 230)
$labelVlan.Size = New-Object System.Drawing.Size(100, 20)
$labelVlan.AutoSize = $true  # Автоматический размер под текст
$labelVlan.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelVlan.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelVlan)



# === ComboBox VLAN ===
$comboBoxScopes = New-Object System.Windows.Forms.ComboBox
$comboBoxScopes.Location = New-Object System.Drawing.Point(200, 227)
$comboBoxScopes.Width = 190
$comboBoxScopes.DropDownStyle = 'DropDownList'
$comboBoxScopes.Items.Add("Все VLAN")
$comboBoxScopes.SelectedIndex = 0

$comboBoxScopes.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$comboBoxScopes.ForeColor = [System.Drawing.Color]::Black
$comboBoxScopes.BackColor = [System.Drawing.Color]::WhiteSmoke

$comboBoxScopes.Cursor = [System.Windows.Forms.Cursors]::Hand

$form.Controls.Add($comboBoxScopes)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($comboBoxScopes, "Выберите Vlan")




# === Обработчик изменения сервера ===
$comboBoxServers.Add_SelectedIndexChanged({
    $comboBoxScopes.Items.Clear()
    $comboBoxScopes.Items.Add("Все VLAN")
    $scopesMap.Clear()  # Очищаем карту

    $comboBoxScopes.SelectedIndex = 0

    $server = $comboBoxServers.SelectedItem
    try {
        $scopes = Get-DhcpServerv4Scope -ComputerName $server
        foreach ($scope in $scopes) {
            $ipString = $scope.ScopeId.ToString()
            $thirdOctet = ($ipString -split '\.')[2]
            $display = "$($scope.Name) [$thirdOctet]"
            $comboBoxScopes.Items.Add($display)
            $scopesMap[$display] = $scope.ScopeId
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при получении VLAN с $server : $($_.Exception.Message)", "Ошибка", "OK", "Error")
    }
})










# === Кнопка поиска ===
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = "SEARCH"
$buttonSearch.Location = New-Object System.Drawing.Point(500, 150)
$buttonSearch.Width = 130
$buttonSearch.Height = 45      # Устанавливаем высоту кнопки
$buttonSearch.BackColor = [System.Drawing.Color]::Silver
$buttonSearch.ForeColor = [System.Drawing.Color]::SeaGreen
$buttonSearch.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

# Установка курсора при наведении
$buttonSearch.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonSearch.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonSearch.FlatAppearance.BorderColor = [System.Drawing.Color]::SeaGreen
$buttonSearch.FlatAppearance.BorderSize = 2

# Установка размера и стиля шрифта
$buttonSearch.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$magnifierPath = Join-Path -Path $scriptPath -ChildPath "images\magnifier.png"

# Загрузка и установка иконки для кнопки
$magnifier = [System.Drawing.Image]::FromFile($magnifierPath)
$buttonSearch.Image = $magnifier

# Устанавливаем выравнивание иконки
$buttonSearch.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonSearch.Padding = New-Object System.Windows.Forms.Padding(5, 0, 5, 0)


$form.Controls.Add($buttonSearch)







# === Кнопка Зарезервированные IP ===
$buttonReserved = New-Object System.Windows.Forms.Button
$buttonReserved.Text = "RESERVED"
$buttonReserved.Location = New-Object System.Drawing.Point(490, 215)
$buttonReserved.Width = 150
$buttonReserved.Height = 45      # Устанавливаем высоту кнопки
$buttonReserved.BackColor = [System.Drawing.Color]::Silver
$buttonReserved.ForeColor = [System.Drawing.Color]::Brown
$buttonReserved.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

# Установка курсора при наведении
$buttonReserved.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonReserved.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonReserved.FlatAppearance.BorderColor = [System.Drawing.Color]::RosyBrown
$buttonReserved.FlatAppearance.BorderSize = 2
 
# Установка размера и стиля шрифта
$buttonReserved.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$pinPath = Join-Path -Path $scriptPath -ChildPath "images\pin.png"

# Загрузка и установка иконки для кнопки
$pin = [System.Drawing.Image]::FromFile($pinPath)
$buttonReserved.Image = $pin

# Устанавливаем выравнивание иконки
$buttonReserved.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonReserved.Padding = New-Object System.Windows.Forms.Padding(5, 0, 5, 0)

$form.Controls.Add($buttonReserved)





# === Список результатов ===
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(100, 320)
$listBox.Size = New-Object System.Drawing.Size(580,300)
$listBox.Font = New-Object System.Drawing.Font("Calibri", 11, [System.Drawing.FontStyle]::Regular)
$listBox.ForeColor = [System.Drawing.Color]::White      # Цвет текста
$listBox.BackColor = [System.Drawing.Color]::DimGray  # Цвет фона
$form.Controls.Add($listBox)

# === Создание ToolTip для ListBox ===
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay = 100
$toolTip.ShowAlways = $true

# Храним индекс, чтобы не обновлять tooltip постоянно
$global:lastIndex = -1

# Обработка наведения мыши на элементы ListBox
$listBox.Add_MouseMove({
    $p = $listBox.PointToClient([System.Windows.Forms.Cursor]::Position)
    $index = $listBox.IndexFromPoint($p)

    if ($index -ge 0 -and $index -ne $global:lastIndex) {
        $text = $listBox.Items[$index].ToString()
        $toolTip.SetToolTip($listBox, $text)
        $global:lastIndex = $index
    }
})


















# === Логика поиска ===
$buttonSearch.Add_Click({
    $null = $listBox.Items.Clear()
    $server = $comboBoxServers.SelectedItem
    $selectedScopeText = $comboBoxScopes.SelectedItem

    # Принудительная очистка памяти
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    try {
        # Получаем все аренды DHCP
        $allLeases = Get-AllDhcpLeases -server $server -scopesMap $scopesMap -selectedScopeText $selectedScopeText

        # Подготовка директорий и путей
        $logDir = "C:\Logs\Dhcp"
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null

        $today      = Get-Date -Format "yyyy-MM-dd"
        $yesterday  = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")

        $todayFile     = Join-Path $logDir "leases-$server-$today.csv"
        $yesterdayFile = Join-Path $logDir "leases-$server-$yesterday.csv"

        # Сохраняем текущие аренды
        $allLeases | Export-Csv -Path $todayFile -NoTypeInformation -Force


    # Удаление старых логов (старше 2 дней)
        Get-ChildItem -Path $logDir -Filter "leases-$server-*.csv" | Where-Object {
            $_.LastWriteTime -lt (Get-Date).AddDays(-10)
                } | ForEach-Object {
                try {
                    Remove-Item $_.FullName -Force -ErrorAction Stop
                } catch {
                # Ошибки подавляем или можно логировать в файл, если нужно
            }
        }


        if (Test-Path $yesterdayFile) {
            $yesterdayLeases = Import-Csv $yesterdayFile
            $yesterdayIPs = $yesterdayLeases.IPAddress

            # Определяем новые IP, которых не было вчера
            $newLeases = $allLeases | Where-Object {
                $_.IPAddress -and ($_.IPAddress -notin $yesterdayIPs)
            }

            if ($newLeases.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Сегодня нет новых арендаторов на сервере $server в VLAN '$selectedScopeText'.", "Информация", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            } else {
                $listBox.Items.Add("Новые арендаторы за сегодня:")
            foreach ($lease in $newLeases) {
            $hostname   = if ($lease.HostName) { $lease.HostName } else { "(без имени)" }
            $expiryTime = if ($lease.LeaseExpiryTime) { $lease.LeaseExpiryTime.ToString("g") } else { "(нет даты)" }
            $mac        = if ($lease.ClientId) { $lease.ClientId } else { "(нет MAC)" }
    
            $listBox.Items.Add("IP: $($lease.IPAddress) | MAC: $mac | Name: $hostname | До: $expiryTime")
}

            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Вчерашний файл не найден:`n$yesterdayFile`nСравнение невозможно.", "Предупреждение", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите DHCP-сервер.", "Внимание", "OK", "Warning")
    }
})







$buttonReserved.Add_Click({
    $listBox.Items.Clear()
    $server = $comboBoxServers.SelectedItem
    $selectedScopeText = $comboBoxScopes.SelectedItem

    if (-not $server) {
        [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите DHCP-сервер.", "Внимание", "OK", "Warning")
        return
    }

    try {
        $reservations = Get-AllDhcpReservations -server $server -scopesMap $scopesMap -selectedScopeText $selectedScopeText

        if ($reservations.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Нет зарезервированных IP-адресов на $server.", "Информация", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            foreach ($r in $reservations) {
                $name = if ($r.Name) { $r.Name } else { "(без имени)" }
                $listBox.Items.Add("IP: $($r.IPAddress) | MAC: $($r.ClientId) | Name: $name")
            }
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при получении резервов с $server : $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})







[void]$form.ShowDialog()
