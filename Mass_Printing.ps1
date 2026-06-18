# Настройка консоли:
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class ConsoleFont {
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public struct CONSOLE_FONT_INFO_EX {
		public uint cbSize;
		public uint nFont;
		public short dwFontSizeX;
		public short dwFontSizeY;
		public int FontFamily;
		public int FontWeight;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
		public string FaceName;
	}
	[DllImport("kernel32.dll", SetLastError = true)]
	public static extern bool SetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref CONSOLE_FONT_INFO_EX lpConsoleCurrentFontEx);
	[DllImport("kernel32.dll", SetLastError = true)]
	public static extern IntPtr GetStdHandle(int nStdHandle);
	public static void SetFont(string fontName, short fontSize = 12) {
		IntPtr hConsole = GetStdHandle(-11); // STD_OUTPUT_HANDLE
		CONSOLE_FONT_INFO_EX fontInfo = new CONSOLE_FONT_INFO_EX();
		fontInfo.cbSize = (uint)Marshal.SizeOf(fontInfo);
		fontInfo.FaceName = fontName;
		fontInfo.dwFontSizeY = fontSize;
		SetCurrentConsoleFontEx(hConsole, false, ref fontInfo);
	}
}
"@

[ConsoleFont]::SetFont("Consolas", 16)
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

# Функция разделителя:
function Separator {
	Write-Host "================================================" -ForegroundColor Green
}

Separator
Write-Host "Mass_Printing 1.6.4"
Separator

# Проверка запущенных процессов:
$CheckProcess = 'winword', 'powerpnt', 'excel', 'AcroRd32', 'Acrobat'
$Wshell = New-Object -ComObject Wscript.Shell

do {
	$Proc = Get-Process -Name $CheckProcess -ErrorAction SilentlyContinue
	
	if (!$Proc) {
		break
	}

	$RunningApps = $Proc | ForEach-Object { if ($_.Description) { $_.Description } else { $_.ProcessName } } | Select-Object -Unique	
	$Output = $Wshell.Popup("Данные приложения должны быть закрыты перед запуском печати:`n`n" + ($RunningApps -join "`n") + "`n`nПожалуйста, закройте их и нажмите 'Повтор' для продолжения.", 0, "Закрытие приложений", 5 + 48)

	if ($Output -eq 2) { 
		exit
	}
} while ($true)

# Выбор файлов для печати:
Write-Host "Выберите файлы для печати (для выделения всех файлов нажмите Ctrl + A):"

Add-Type -AssemblyName System.Windows.Forms | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Multiselect = $true
$OpenFileDialog.Filter = "Документы (*.pdf,*.doc,*.docx,*.xls,*.xlsx,*.ppt,*.pptx)|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx"

# Если файлы не выбраны:
if ($OpenFileDialog.ShowDialog() -ne 'OK') {
	Separator
	Write-Host "Ошибка: Файлы не выбраны." -ForegroundColor DarkRed
	Separator
	exit
}

# Подсчет количества выбранных файлов:
$FilesToPrint = Get-Item $OpenFileDialog.FileNames | Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(10, '0') }) }
$FilesTotal = $FilesToPrint.Count
Separator
Write-Host "Выбрано файлов: $FilesTotal"
Separator

# Запрос количества копий:
do {
	$CopiesInput = Read-Host "Введите количество копий"
	Separator
	if ($CopiesInput -match '^\d+$' -and [int]$CopiesInput -gt 0) {
		break
	}
	Write-Host "Ошибка: Неверный ввод." -ForegroundColor DarkRed
	Separator
} while ($true)
$Copies = [int]$CopiesInput

# Задержка между печатью файлов (в секундах):
$Seconds = 3.6

# Расчёт примерного времени печати:
$PrintTime = [TimeSpan]::FromSeconds($FilesTotal * $Copies * $Seconds).ToString("hh\:mm\:ss")
Write-Host "Примерное время печати: $PrintTime"
Separator

# Печать файлов:
$FailedFiles = [System.Collections.Generic.List[string]]::new()

for ($CopiesDefault = 1; $CopiesDefault -le $Copies; $CopiesDefault++) {
	$i = 1
	foreach ($file in $FilesToPrint) {
		Write-Host "Копия $($CopiesDefault): $i/$FilesTotal. Печать файла $($file.Name)"

		try {
			Start-Process -FilePath $file.FullName -Verb Print -WindowStyle Minimized -ErrorAction Stop
		} catch {
			Separator
			Write-Host "Ошибка: Не удалось отправить на печать: $($file.Name): `n$_" -ForegroundColor DarkRed
			$FailedFiles.Add($file.Name)
		}

		Start-Sleep -Seconds $Seconds
		$i++
	}
	Separator
}

if ($FailedFiles.Count -gt 0) {
	Write-Host "Не удалось напечатать следующие файлы:" -ForegroundColor DarkRed
	$FailedFiles | ForEach-Object { Write-Host "$_" -ForegroundColor DarkRed }
	Separator
}

# Закрытие программ перед перемещением:
Write-Host "Завершение фоновых процессов..."
Separator
$Proc = Get-Process -Name $CheckProcess -ErrorAction SilentlyContinue
if ($Proc) {
	foreach ($item in $Proc) {
		$null = $item.CloseMainWindow()
	}
	Start-Sleep -Seconds 1
	Get-Process -Name $CheckProcess -ErrorAction SilentlyContinue | Stop-Process -Force
}

# Перемещение файлов:
$Directory = Split-Path -Parent $OpenFileDialog.FileName
$Date = Get-Date -Format "dd.MM.yyyy"

# Запрос на перемещение файлов:
$Output = $Wshell.Popup("Переместить распечатанные файлы в папку Распечатано?", 0, "Перемещение файлов", 4 + 32)

if ($Output -eq 6) { 
	$Printed = Join-Path $Directory "Распечатано_$Date"
	$TestedPath = $Printed
	
	if (!(Test-Path $TestedPath)) {
		New-Item -ItemType Directory -Force -Path $TestedPath | Out-Null
	}
	
	Write-Host "Перемещаем файлы..."
	Separator
	
	$DestinationFolder = $Printed
	$i = 1
	foreach ($file in $FilesToPrint) {
		if (Test-Path $file.FullName) {
			Write-Host "$i/$FilesTotal. Файл $($file.name) перемещен в $DestinationFolder"
			Move-Item -Path $file.FullName -Destination $DestinationFolder -Force
		}
		$i++
	}
	Separator
	
	# Открытие папки с перемещенными файлами:
	Invoke-Item $DestinationFolder
}
