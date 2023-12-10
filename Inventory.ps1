#get serial number of the computer
$serialnumber = (Get-WmiObject -Class win32_bios).SerialNumber
#get manufacturer of the computer
$manufacturer = (Get-CimInstance -ClassName win32_ComputerSystem).Manufacturer
#get model of the computer
$model = (Get-CimInstance -ClassName win32_ComputerSystem).Model
#get hostname of the computer
$hostname = (Get-CimInstance -ClassName win32_ComputerSystem).Name
#ask user for his login
$username = Read-Host "enter your login"
#ask user for his first name
$name = Read-Host "Enter your firstname"
#ask user for his lastname
$surname = Read-Host "Enter your lastname"
#ask user for the department where they work
$department = Read-Host "Enter the department where you work"

#message for the user
Write-Host "You have successfully added a new device!"

#assign excel to a variable
$ExcelObj = New-Object -ComObject Excel.Application

#run excel
$ExcelObj.visible=$true

#run inventory file
$ExcelWorkBook = $ExcelObj.Workbooks.Open("C:\inventory\example.xlsx")

#open inventory sheet
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Arkusz1")


#find empty cell and assign data to it
while ($number -le 10000)
{
    if ([string]::IsNullOrEmpty($ExcelWorkSheet.Range("B$number").Text) -eq $false) {
        $number++
    } else {
        $ExcelWorkSheet.Range("A$number") = "$number"
        $ExcelWorkSheet.Range("B$number") = "$model"
        $ExcelWorkSheet.Range("C$number") = "$hostname"
        $ExcelWorkSheet.Range("D$number") = "$serialnumber"
        $ExcelWorkSheet.Range("E$number") = "$username".ToUpper()
        $ExcelWorkSheet.Range("F$number") = "$name".ToUpper()
        $ExcelWorkSheet.Range("G$number") = "$surname".ToUpper()
        $ExcelWorkSheet.Range("H$number") = "$department".ToUpper()
        $ExcelWorkSheet.Range("I$number") = "$location".ToUpper()
        $ExcelWorkSheet.Range("J$number") = "$comment".ToUpper()
        $ExcelWorkBook.Save()
        $ExcelWorkBook.close($true)
        $ExcelObj.Quit()
        break
    }
}