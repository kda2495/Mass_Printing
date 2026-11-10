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
	
	# Расширение окна до 90 символов, если позволяет экран:
	$WindowWidth = [Math]::Min([Math]::Max($WindowSize.Width, 90), $RawUI.MaxPhysicalWindowSize.Width)
	if ($WindowWidth -gt $WindowSize.Width) {
		$RawUI.WindowSize = New-Object System.Management.Automation.Host.Size($WindowWidth, $WindowSize.Height)
	}
} catch {
}

# Включение UTF-8 по умолчанию:
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Функция разделителя:
function Separator {
	Write-Host "=====================================================" -ForegroundColor Green
}

# Версия скрипта:
Separator
Write-Host "Mass_Printing 2.0.5"
Separator

# Выбор файлов для печати:
Write-Host "Выберите файлы для печати:"
$WShell = New-Object -ComObject Wscript.Shell
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$TopForm = New-Object System.Windows.Forms.Form
$TopForm.TopMost = $true
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Multiselect = $true
$OpenFileDialog.Filter = "Документы (*.pdf,*.doc,*.docx,*.xls,*.xlsx,*.ppt,*.pptx,*.txt)|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.txt"

# Если файлы не выбраны:
if ($OpenFileDialog.ShowDialog($TopForm) -ne 'OK') {
	Separator
	Write-Host "Ошибка: Файлы не выбраны." -ForegroundColor DarkRed
	Separator
	Write-Host "Нажмите любую клавишу для выхода..."
	Separator
	
	# Очистка памяти:
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WShell) | Out-Null
	$TopForm.Dispose()
	exit
}

$TopForm.Dispose()

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
} until ($CopiesInput -match '^\d+$' -and [int]$CopiesInput -gt 0)

$Copies = [int]$CopiesInput

# Задержка между печатью файлов (в секундах):
$Seconds = 3.6

# Печать файлов:
$FailedFiles = [System.Collections.Generic.List[string]]::new()

for ($CopiesDefault = 1; $CopiesDefault -le $Copies; $CopiesDefault++) {
	Write-Host "Печать файлов:"
	for ($i = 0; $i -lt $FilesToPrint.Count; $i++) {
		$file = $FilesToPrint[$i]
		Write-Host "Копия $($CopiesDefault): $($i + 1)/$FilesTotal. Печать файла: $($file.Name)"
		
		try {
			Start-Process -FilePath $file.FullName -Verb Print -WindowStyle Minimized -ErrorAction Stop
		} catch {
			Separator
			Write-Host "Ошибка: Не удалось отправить на печать: $($file.Name): `n$_" -ForegroundColor DarkRed
			$FailedFiles.Add($file.Name)
		}
		
		Start-Sleep -Seconds $Seconds
	}
	Separator
}

if ($FailedFiles.Count -gt 0) {
	Write-Host "Не удалось напечатать следующие файлы:" -ForegroundColor DarkRed
	$FailedFiles | ForEach-Object { Write-Host "$_" -ForegroundColor DarkRed }
	Separator
}

# Перемещение файлов:
$Directory = Split-Path -Parent $OpenFileDialog.FileName
$Date = Get-Date -Format "dd.MM.yyyy"

# Запрос на перемещение файлов:
$Output = $WShell.Popup("Переместить распечатанные файлы в папку Распечатано?", 0, "Перемещение файлов", 4 + 32 + 4096)

if ($Output -eq 6) { 
	$Printed = Join-Path $Directory "Распечатано_$Date"
	
	if (!(Test-Path $Printed)) {
		New-Item -ItemType Directory -Force -Path $Printed | Out-Null
	}
	
	Write-Host "Перемещение файлов:"
	
	for ($i = 0; $i -lt $FilesToPrint.Count; $i++) {
		$file = $FilesToPrint[$i]
		if (Test-Path $file.FullName) {
			try {
				Move-Item -Path $file.FullName -Destination $Printed -Force -ErrorAction Stop
				Write-Host "$($i + 1)/$FilesTotal. Файл $($file.Name) перемещен в $Printed."
			} catch {
				Write-Host "$($i + 1)/$FilesTotal. Ошибка: Файл $($file.Name) занят программой и не перемещен." -ForegroundColor DarkRed
			}
		}
	}
	Separator
}
# Очистка памяти:
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WShell) | Out-Null

Write-Host "Печать завершена. Нажмите любую клавишу для выхода..."
Separator