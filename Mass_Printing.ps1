#Dmitry Ornatsky (kda2495), 2022-2024, version 1.5
Write-Host "Выберите файлы для печати
(для выделения всех файлов в папке по порядку выделите первый файл в списке и нажмите Ctrl+A)"
#Открытие окна выбора файлов:
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Multiselect = $true
$OpenFileDialog.Filter = "Все файлы (*.pdf,*.doc,*.docx,*.xls,*.xlsx,*.ppt,*.pptx)|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx"
$OpenFileDialog.ShowDialog() | Out-Null
$FilesToPrint = $OpenFileDialog.FileNames | Sort-Object
if (!($FilesToPrint)) {
	break
}
#Подсчет количества файлов, выбранных для печати:
$total_number_of_files = (Get-ChildItem $FilesToPrint | Measure-Object).Count
Write-Host "Выбрано файлов: $total_number_of_files"
#Проверка запущенных процессов, которые могут помешать работе:
$checkprocess = 'winword,powerpnt,AcroRd32,Acrobat' -Split ','
do {
	$array = @()
	foreach ($item in $checkprocess) {
		$proc = Get-Process | Where-Object ProcessName -Match "$($item)" | Select -Unique
		$array += New-Object psobject -Property @{ 'Description' = $proc }
	}
if ($array.Description.Description -eq $null) {
	break
}
if ($array.Description.Description -ne $null) {
	$wshell = New-Object -ComObject Wscript.Shell
	$output = $wshell.Popup("Данные приложения должны быть закрыты перед запуском печати:`n`n" + ($array.Description.Description -Join "`n") + "`n`nПожалуйста, закройте вышеуказанные приложения и нажмите Повтор для начала печати.",0,"Закрытие приложений",5+48)
}
} until (($array.Description.Description -eq $null) -and ($output -eq 4) -or ($output -eq 2))

if ($output -eq 2) {
	exit
}
#Ввод количества копий:
$copies = Read-Host "Введите количество копий"
if (!($copies)) {
	break
}
#Количество секунд между печатью очередного файла:
$seconds = 3.6
#Подсчет времени печати выбранных файлов:
$time_to_print = (New-TimeSpan -Seconds ($total_number_of_files * $copies * $seconds)).ToString("hh\:mm\:ss")
Write-Host "Время печати выбранных файлов: $time_to_print"
#Печать файлов:
for ($copies_default = 1; $copies_default -le $copies; $copies_default++) {
	$i = 1
	foreach ($file in Get-ChildItem $FilesToPrint) {
		Write-Host "Копия $($copies_default): $i/$total_number_of_files.Печать файла $($file.name)"
		Start-Process -FilePath $file -Verb Print -WindowStyle Minimized
		Start-Sleep $seconds
		$i++
	}
}
#Путь к файлам для дальнейшего перемещения:
$directory = $OpenFileDialog.FileName | Get-ChildItem | Split-Path -Parent
#Вычисление текущей даты для папки Распечатано:
$date = Get-Date -Format dd.MM.yyyy
#Запрос на перемещение файлов:
$wshell = New-Object -ComObject Wscript.Shell
$output = $wshell.Popup("Переместить распечатанные файлы в папку Распечатано?",0,"Перемещение файлов",4+32)
switch ($output) {
	'6' {
		#Перемещение файлов в папку Распечатано:
		$printed = "$directory/Распечатано_$($date)"
		$testedpath = $printed
		if(!(test-path $testedpath)) {
			New-Item -ItemType Directory -Force -Path $testedpath
		}
		Write-Host "Перемещаем файлы..."
		Set-Location $printed
		$destination_folder = Get-Location
		$i = 1
		foreach ($file in Get-ChildItem $FilesToPrint){
			Write-Host "$i/$total_number_of_files.Файл $($file.name) перемещен в $($destination_folder.Path)"
			Move-Item -Path $file.fullname -Destination $destination_folder
			$i++
		}
		#Открытие папки с перемещенными файлами:
		Invoke-Item .
		#Закрытие программ:
		foreach ($item in $checkprocess) {
			Get-Process | Where-Object ProcessName -Match "$($item)" | % { $_.CloseMainWindow() | Out-Null } | Stop-Process -Force
		}
	}
	'7' {
		#Закрытие программ:
		foreach ($item in $checkprocess) {
			Get-Process | Where-Object ProcessName -Match "$($item)" | % { $_.CloseMainWindow() | Out-Null } | Stop-Process -Force
		}
	}
}
Write-Host -NoNewLine "Печать завершена, нажмите любую клавишу для продолжения..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')