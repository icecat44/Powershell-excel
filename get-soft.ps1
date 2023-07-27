# Составление таблицы Excel со списком установленных программ.
# Необходимо, что бы на компе был установлен Microsoft Excel. 
#
#
# Получаем список установленных программ
$soft = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher)
# Получаем имя компьютера
$computername = (Get-WmiObject win32_computersystem).DNSHostName
# Запускаем Excel в фоне
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
# Открываем документ
$EWB = $excel.Workbooks.Open("\\192.168.17.7\Share\po\po-orpo.xlsx")
# Проверка на наличие дубликата, 
# если в документе уже есть лист с текущем значением имени 
# компьтера - закрываем документ и выходим.
$Cheets = ($EWB.worksheets | Select-Object Name)
if ($Cheets -match $computername) 
{
   $EWB.Close()
   $excel.Quit()
   exit
}
# Добавляем новый лист и переименовываем его
$EWB.worksheets.add()
$NewSheet = $EWB.Worksheets.Item(1)
$NewSheet.Name = "$computername"
# Строим таблицу и заполняем лист значениями списка устрановленных программ
$NewSheet.Cells.Item(1,1) = 'Name'
$NewSheet.Cells.Item(1,2) = 'Version'
$NewSheet.Cells.Item(1,3) = 'Publisher'
$NewSheet.Range("A1:C1").font.size = 18
$NewSheet.Range("A1:C1").font.bold = $true
$NewSheet.Range("A1:C1").font.ColorIndex = 2
$NewSheet.Range("A1:C1").interior.colorindex = 1
$NewSheet.Columns.Item(2).NumberFormat= "@"
$NewSheet.Columns.Item(2).HorizontalAlignment = -4108
$NewSheet.Columns.Item(2).VerticalAlignment = -4108
$j = 2
foreach($arr in $soft)
{
   $NewSheet.Cells.Item($j,1) = $arr.DisplayName
   $NewSheet.Cells.Item($j,2) = $arr.DisplayVersion
   $NewSheet.Cells.Item($j,3) = $arr.DisplayName
   $j++
}
$NewSheet.Columns.AutoFit()
# Сохраняем и выходим.
$EWB.Save()
$EWB.Close()
$excel.Quit()
