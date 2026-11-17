# Версия скрипта:
$ScriptVersion = "2.0.5"

# Задержка между печатью файлов в миллисекундах:
$PrintDelayMs = 3600

# Функция настройки консоли:
function Initialize-Console {
	# Установка шрифта консоли вне Windows Terminal:
	if (-not $env:WT_SESSION) {
		# Компиляция типа при первом запуске в текущей сессии:
		if (-not ('ConsoleFont' -as [type])) {
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
		IntPtr hConsole = GetStdHandle(-11);
		CONSOLE_FONT_INFO_EX fontInfo = new CONSOLE_FONT_INFO_EX();
		fontInfo.cbSize = (uint)Marshal.SizeOf(fontInfo);
		fontInfo.FaceName = fontName;
		fontInfo.dwFontSizeY = fontSize;
		SetCurrentConsoleFontEx(hConsole, false, ref fontInfo);
	}
}
"@
		}
		
		[ConsoleFont]::SetFont("Consolas", 16)
	}
	
	# Настройка размера буфера экрана:
	try {
		$RawUI = $Host.UI.RawUI
		$WindowSize = $RawUI.WindowSize
		$BufferSize = $RawUI.BufferSize
		
		# Настройки ширины и высоты буфера экрана:
		$BufferWidth = [Math]::Max([Math]::Max($BufferSize.Width, 90), $WindowSize.Width)
		$BufferHeight = [Math]::Max([Math]::Max($BufferSize.Height, 9001), $WindowSize.Height)
		
		if ($BufferSize.Width -ne $BufferWidth -or $BufferSize.Height -ne $BufferHeight) {
			$RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($BufferWidth, $BufferHeight)
		}
		
		# Расширение окна до 90 символов:
		$WindowWidth = [Math]::Min([Math]::Max($WindowSize.Width, 90), $RawUI.MaxPhysicalWindowSize.Width)
		if ($WindowWidth -gt $WindowSize.Width) {
			$RawUI.WindowSize = New-Object System.Management.Automation.Host.Size($WindowWidth, $WindowSize.Height)
		}
	} catch { }
}

# Функция разделителя:
function Separator {
	Write-Host "=====================================================" -ForegroundColor Green
}

# Функция запроса положительного числа:
function Read-PositiveInt {
	param ([string]$Prompt)
	$Number = 0
	while ($true) {
		$InputText = Read-Host $Prompt
		Separator
		if ([int]::TryParse($InputText, [System.Globalization.NumberStyles]::None, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$Number) -and $Number -gt 0) {
			return $Number
		}
		Write-Host "Ошибка: Неверный ввод." -ForegroundColor DarkRed
		Separator
	}
}

# Функция выбора файлов для печати:
function Select-PrintFiles {
	Add-Type -AssemblyName System.Windows.Forms
	$TopForm = New-Object System.Windows.Forms.Form
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	try {
		$TopForm.TopMost = $true
		$OpenFileDialog.Multiselect = $true
		$OpenFileDialog.Filter = "Документы (*.pdf,*.doc,*.docx,*.xls,*.xlsx,*.ppt,*.pptx,*.txt)|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.txt"
		if ($OpenFileDialog.ShowDialog($TopForm) -eq 'OK') {
			return $OpenFileDialog.FileNames
		}
	} finally {
		$OpenFileDialog.Dispose()
		$TopForm.Dispose()
	}
}

# Настройка консоли на Windows:
Initialize-Console

# Включение UTF-8 по умолчанию:
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Баннер с версией скрипта:
Separator
Write-Host "Mass_Printing $ScriptVersion"
Separator

# Выбор файлов для печати:
Write-Host "Выберите файлы для печати:"
$SelectedPaths = @(Select-PrintFiles)

# Проверка выбора файлов:
if ($SelectedPaths.Count -eq 0) {
	Separator
	Write-Host "Ошибка: Файлы не выбраны." -ForegroundColor DarkRed
	Separator
	Write-Host "Нажмите любую клавишу для выхода..."
	Separator
	exit
}

# Формирование списка файлов с естественной сортировкой:
$FilesToPrint = @(Get-Item -LiteralPath $SelectedPaths | Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(10, '0') }) })
$FilesTotal = $FilesToPrint.Count
Separator
Write-Host "Выбрано файлов: $FilesTotal"
Separator

# Запрос количества копий:
$Copies = Read-PositiveInt -Prompt "Введите количество копий"

# Печать файлов:
$FailedFiles = [System.Collections.Generic.HashSet[string]]::new()

for ($Copy = 1; $Copy -le $Copies; $Copy++) {
	Write-Host "Печать файлов:"
	for ($i = 0; $i -lt $FilesTotal; $i++) {
		$File = $FilesToPrint[$i]
		Write-Host "Копия ${Copy}: $($i + 1)/$FilesTotal. Печать файла: $($File.Name)"
		
		try {
			Start-Process -FilePath $File.FullName -Verb Print -WindowStyle Minimized -ErrorAction Stop
		} catch {
			Separator
			Write-Host "Ошибка: Не удалось отправить на печать: $($File.Name): `n$_" -ForegroundColor DarkRed
			[void]$FailedFiles.Add($File.FullName)
			continue
		}
		
		Start-Sleep -Milliseconds $PrintDelayMs
	}
	Separator
}

# Список файлов с ошибками печати:
if ($FailedFiles.Count -gt 0) {
	Write-Host "Не удалось напечатать следующие файлы:" -ForegroundColor DarkRed
	foreach ($File in $FilesToPrint) {
		if ($FailedFiles.Contains($File.FullName)) { Write-Host $File.Name -ForegroundColor DarkRed }
	}
	Separator
}

# Формирование списка файлов для перемещения:
$PrintedFiles = @($FilesToPrint | Where-Object { -not $FailedFiles.Contains($_.FullName) })

if ($PrintedFiles.Count -gt 0) {
	# Запрос на перемещение файлов:
	$WShell = New-Object -ComObject Wscript.Shell
	try {
		$Answer = $WShell.Popup("Переместить распечатанные файлы в папку Распечатано?", 0, "Перемещение файлов", 4 + 32 + 4096)
	} finally {
		[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($WShell)
	}
	
	# Перемещение файлов:
	if ($Answer -eq 6) {
		$PrintedFolder = Join-Path -Path $FilesToPrint[0].DirectoryName -ChildPath "Распечатано_$(Get-Date -Format 'dd.MM.yyyy')"
		New-Item -ItemType Directory -Force -Path $PrintedFolder | Out-Null
		
		Write-Host "Перемещение файлов:"
		
		for ($i = 0; $i -lt $PrintedFiles.Count; $i++) {
			$File = $PrintedFiles[$i]
			if (Test-Path -LiteralPath $File.FullName) {
				try {
					Move-Item -LiteralPath $File.FullName -Destination $PrintedFolder -Force -ErrorAction Stop
					Write-Host "$($i + 1)/$($PrintedFiles.Count). Файл $($File.Name) перемещен в $PrintedFolder."
				} catch {
					Write-Host "$($i + 1)/$($PrintedFiles.Count). Ошибка: Файл $($File.Name) занят программой и не перемещен." -ForegroundColor DarkRed
				}
			}
		}
		Separator
	}
}

Write-Host "Печать завершена. Нажмите любую клавишу для выхода..."
Separator