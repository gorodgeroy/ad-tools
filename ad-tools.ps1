import-module ActiveDirectory
Add-Type -assembly System.Windows.Forms

# Параметры формы
function Form-Create {
    $main_form = New-Object System.Windows.Forms.Form
    $main_form.Text = 'gorodgeroy AD TOOLS'
    $main_form.Width = 640
    $main_form.Height = 480
    $main_form.AutoSize = $true

    $LabelMode = New-Object System.Windows.Forms.Label
    $LabelMode.Text = 'Выберите действие'
    $LabelMode.Location = New-Object System.Drawing.Point(10,20)
    $LabelMode.AutoSize = $true
    $main_form.Controls.Add($LabelMode)

    $ModeSwitch = New-Object System.Windows.Forms.ComboBox
    $ModeSwitch.DataSource = @('Cписок групп юзера','Список членов группы', 'Членство в группах', 'Поиск по шаблону', 'Ящики с FullAccess')
    $ModeSwitch.Location = New-Object System.Drawing.Point(10,60)
    $ModeSwitch.AutoSize = $true
    $ModeSwitch.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $main_form.Controls.Add($ModeSwitch)

    $LabelExclude = New-Object System.Windows.Forms.Label
    $LabelExclude.Text = 'Исключить шаблон (для поиска по шаблону)'
    $LabelExclude.Location = New-Object System.Drawing.Point(10,100)
    $LabelExclude.AutoSize = $true
    $main_form.Controls.Add($LabelExclude)

    $TextExclude = New-Object System.Windows.Forms.TextBox
    $TextExclude.Location = New-Object System.Drawing.Point(10,120)
    $TextExclude.Text = ''
    $TextExclude.AutoSize = $true
    $main_form.Controls.Add($TextExclude)

    $LabelInput = New-Object System.Windows.Forms.Label
    $LabelInput.Text = 'Имя пользователя или группы'
    $LabelInput.Location = New-Object System.Drawing.Point(200,20)
    $LabelInput.AutoSize = $true
    $main_form.Controls.Add($LabelInput)

    $TextboxInput = New-Object System.Windows.Forms.TextBox
    $TextboxInput.Location = New-Object System.Drawing.Point(200,60)
    $TextboxInput.AutoSize = $true
    $TextboxInput.Text = ''
    $main_form.Controls.Add($TextboxInput)

    $ButtonCompare = New-Object System.Windows.Forms.Button
    $ButtonCompare.Location = New-Object System.Drawing.Point(310,60)
    $ButtonCompare.Text = 'Сравнить'
    $ButtonCompare.AutoSize = $true
    $ButtonCompare.Add_Click({(Get-Compare)})
    $main_form.Controls.Add($ButtonCompare)
    
    $LabelInput1 = New-Object System.Windows.Forms.Label
    $LabelInput1.Text = 'Имя пользователя или группы для сравнения'
    $LabelInput1.Location = New-Object System.Drawing.Point(400,20)
    $LabelInput1.AutoSize = $true
    $main_form.Controls.Add($LabelInput1)
    
    $TextboxInput1 = New-Object System.Windows.Forms.TextBox
    $TextboxInput1.Location = New-Object System.Drawing.Point(400,60)
    $TextboxInput1.AutoSize = $true
    $TextboxInput1.Text = ''
    $main_form.Controls.Add($TextboxInput1)

    $LabelOutput = New-Object System.Windows.Forms.Label
    $LabelOutput.Text = 'Результат'
    $LabelOutput.Location = New-Object System.Drawing.Point(660,20)
    $LabelOutput.AutoSize = $true
    $main_form.Controls.Add($LabelOutput)

    $TextboxOutput = New-Object System.Windows.Forms.RichTextBox
    $TextboxOutput.Location = New-Object System.Drawing.Point(660,40)
    $TextboxOutput.Size = New-Object System.Drawing.Size(400,300)
    $TextboxOutput.Multiline = $true
    $TextboxOutput.ReadOnly = $true
    $main_form.Controls.Add($TextboxOutput)

    $ButtonExport = New-Object System.Windows.Forms.Button
    $ButtonExport.Location = New-Object System.Drawing.Point(460,360)
    $ButtonExport.AutoSize = $true
    $ButtonExport.Text='Экспорт в файл'
    $ButtonExport.Add_Click({ExportToFile})
    #$main_form.Controls.Add($ButtonExport)

    $ButtonCopy = New-Object System.Windows.Forms.Button
    $ButtonCopy.Location = New-Object System.Drawing.Point(460,350)
    $ButtonCopy.Text = 'Копировать в буфер обмена'
    $ButtonCopy.AutoSize = $true
    #$main_form.Controls.Add($ButtonCopy)

    $LabelConsole = New-Object System.Windows.Forms.Label
    $LabelConsole.Text = 'Консоль'
    $LabelConsole.Location = New-Object System.Drawing.Point(10,210)
    $LabelConsole.AutoSize = $true
    $main_form.Controls.Add($LabelConsole)
    
    $ListConsole = New-Object System.Windows.Forms.ListBox
    $ListConsole.Location  = New-Object System.Drawing.Point(10,230)
    $ListConsole.Size = New-Object System.Drawing.Size(200,200)
    $ListConsole.Items.Add('Однако, здравствуйте');
    $main_form.Controls.add($ListConsole)

    $GetData = New-Object System.Windows.Forms.Button
    $GetData.Text = 'Получить данные'
    $GetData.Location = New-Object System.Drawing.Point(10,160)
    $GetData.AutoSize = $true
    $GetData.Add_Click({(Get-Data)})
    $main_form.Controls.Add($GetData)
    $main_form.ShowDialog()
}

# Обработка нажатия кнопки "Сравнить"

function Get-Compare {
    #TODO сравнить объекты, на которые ссылаются $TextboxInput и $TextboxInput1
}


#Обработка нажатия кнопки "Получить данные"
Function Get-Data
{
    $ListOutput.Clear
    if ($ModeSwitch.SelectedItem -eq 'Cписок групп юзера') {
        GetUserGroups
    } 
    elseif ($ModeSwitch.SelectedItem -eq 'Список членов группы') {
        GetGroupMembers
    } 
    elseif ($ModeSwitch.SelectedItem -eq 'Поиск по шаблону') {
        SearchByExample
    }
    elseif ($ModeSwitch.SelectedItem -eq 'Ящики с FullAccess') {
        WhoHasFullAccess
    }
    elseif ($ModeSwitch.SelectedItem -eq 'Членство в группах') {
        GetGroupMembership
    } 
}


#режим Поиск по шаблону
Function SearchByExample
{
    $TextboxOutput.Text = ''
    Get-ADGroup -Filter 'name -like $TextboxInput.Text'| select name | where name -notlike $TextExclude.Text | ForEach-Object {
        $TextboxOutput.AppendText($_.name)
        $TextboxOutput.AppendText("`n")    
    }
    
}

#режим Cписок групп юзера
Function GetUserGroups
{
    $TextboxOutput.Text = ''
    if ($TextboxInput.Text -eq '') {
        $ListConsole.Items.Add('Имя юзера-то введи, например')
        }
    Else {
        $Result = UserMembership $TextboxInput.Text | Sort Memberof
        $Result = $Result | Foreach {$_ -replace "\\#", "#"}
        $Result = $Result | Foreach {$_ -replace "@{Memberof=", ""}
        $Result = $Result | Foreach {$_ -replace "}", ""} | ForEach-Object {
            $TextboxOutput.AppendText($_)
            $TextboxOutput.AppendText("`n")
        }

    }

}

Function Get-UserMembership
{
    Param($UserAccount)
    Process
    {
        Try {
            $Groups = (Get-ADUser -Identity $UserAccount -Properties MemberOf |
              Select-Object MemberOf).MemberOf
        }
        Catch {
            $ListConsole.Items.Add('Небось юзернейм неверный')
            Return $Nothing
        }
        $GroupItems = @()
        ForEach ($Group in $Groups)
        {
          $var = $group.split(",")
          $var1 = $var[0]
          $ADGroup = $var1.Substring(3)  
          $GrpItems = New-Object -TypeName PSObject -Property @{
          Memberof = $ADGroup}
          $GroupItems += $GrpItems
        }
        Return $GroupItems | Sort memberOf
    }
}

#режим Членство в группах
Function GetGroupMembership
{
    $TextboxOutput.Text = ''
    if ($TextboxInput.Text -eq '') {
        $ListConsole.Items.Add('Имя группы-то введи, например')
        }
    Else {
        $Result = Get-GroupMembership $TextboxInput.Text | Sort Memberof
        $Result = $Result | Foreach {$_ -replace "\\#", "#"}
        $Result = $Result | Foreach {$_ -replace "@{Memberof=", ""}
        $Result = $Result | Foreach {$_ -replace "}", ""} | ForEach-Object {
            $TextboxOutput.AppendText($_)
            $TextboxOutput.AppendText("`n")
        }

    }

}

Function Get-GroupMembership
{
    Param($UserGroup)
    Process
    {
        Try {
            $Groups = (Get-ADGroup -Identity $UserGroup -Properties MemberOf |
              Select-Object MemberOf).MemberOf
        }
        Catch {
            $ListConsole.Items.Add('Небось имя группы неверное')
            Return $Nothing
        }
        $GroupItems = @()
        ForEach ($Group in $Groups)
        {
          $var = $group.split(",")
          $var1 = $var[0]
          $ADGroup = $var1.Substring(3)  
          $GrpItems = New-Object -TypeName PSObject -Property @{
          Memberof = $ADGroup}
          $GroupItems += $GrpItems
        }
        Return $GroupItems | Sort memberOf
    }
}

#режим Список членов группы

Function GetGroupMembers
{
    $TextboxOutput.Text=''
    if ($TextboxInput.Text -eq '') {
        $ListConsole.Items.Add('Имя группы-то введи, например')
        }
    Else {
    $GroupName = $TextboxInput.Text
    try {
        $GroupMembers = Get-ADGroupMember $GroupName | Select-Object -Property name
    }
    Catch {
        $ListConsole.Items.Add('Небось имя группы неверное')
        Return $nothing
    }
        
        $GroupMembers = $GroupMembers | Foreach {$_ -replace "@{name=", ""}
        $GroupMembers = $GroupMembers | Foreach {$_ -replace "}", ""}
        ForEach ($GroupMember in $GroupMembers) {
            $TextboxOutput.AppendText($GroupMember)
            $TextboxOutput.AppendText("`n")
        }
    }
}

#режим Ящики с FullAcess
Function WhoHasFullAccess
{
    $username = $TextboxInput.Text
    "" > D:\POWERSHELL\EXCHAUDIT\mailbox-permission.txt
    
    $TextboxOutput.Text = ''
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://srv-exch2/powershell" -Authentication Kerberos
Import-PSSession $Session -AllowClobber
    $allboxes = get-mailbox -resultsize unlimited 
    $allboxes | Get-MailboxPermission | 
    where-object {($_.Deny -eq $False) -and($_.isinherited -eq $false) -and ($_.User -notlike "*SELF*") -and ($_.User -like "*$username")} | 
    sort-object  identity |format-table User, AccessRights -GroupBy Identity -AutoSize > D:\POWERSHELL\EXCHAUDIT\mailbox-permission.txt
    $TextboxOutput.AppendText("Данные выгружены по пути D:\POWERSHELL\EXCHAUDIT\mailbox-permission.txt")
    $TextboxOutput.AppendText("`n")  
    Remove-PSSession $Session
}

# Экспорт в файл
Function ExportToFile
{
    $ExportFileName = $TextboxInput.Text
    $TextboxOutput.Text > .\$ExportFileName.txt
}

# Рисуем форму

Form-Create