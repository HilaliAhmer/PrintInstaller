function Show-ComboBox {
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory=$true)] $Items,
        [Parameter(Mandatory=$false)] [switch] $ReturnIndex,
        [Parameter(Mandatory=$true)] [string] $FormTitle,
        [Parameter(Mandatory=$false)] [string] $ButtonText = "OK"
    )
    begin {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    }
    process {

        $LabelComboSize= New-Object System.Drawing.Size
        $LabelComboSize.Height = 20
        $LabelComboSize.Width = 280

        $LabelComboPosition = New-Object -TypeName System.Drawing.Point
        $LabelComboPosition.X = 10
        $LabelComboPosition.Y = 10

        $LabelCombo = New-Object -TypeName System.Windows.Forms.Label
        $LabelCombo.Location = $LabelComboPosition
        $LabelCombo.Size = $LabelComboSize
        $LabelCombo.Name = "labelComboName"
        $LabelCombo.Text = 'Yazıcı Driver Listesi:'
        $LabelCombo.TabIndex = 0

        $ComboBoxSize = New-Object System.Drawing.Size
        $ComboBoxSize.Height = 20
        $ComboBoxSize.Width = 260

        $ComboBoxPosition = New-Object -TypeName System.Drawing.Point
        $ComboBoxPosition.X = 10
        $ComboBoxPosition.Y = 30

        $ComboBox = New-Object -TypeName System.Windows.Forms.ComboBox
        $ComboBox.Location = $ComboBoxPosition
        $ComboBox.DataBindings.DefaultDataSourceUpdateMode = 0
        $ComboBox.FormattingEnabled = $true
        $ComboBox.Name = "comboBox1"
        $ComboBox.TabIndex = 0
        $ComboBox.Size = $ComboBoxSize

        $LabelTextBoxSize= New-Object System.Drawing.Size
        $LabelTextBoxSize.Height = 20
        $LabelTextBoxSize.Width = 280

        $LabelTextBoxPosition = New-Object -TypeName System.Drawing.Point
        $LabelTextBoxPosition.X = 10
        $LabelTextBoxPosition.Y = 60

        $LabelTextBox = New-Object -TypeName System.Windows.Forms.Label
        $LabelTextBox.Location = $LabelTextBoxPosition
        $LabelTextBox.Size = $LabelTextBoxSize
        $LabelTextBox.Name = "LabelTextBoxName"
        $LabelTextBox.Text = 'A4 Printer Ip Adres:'
        $LabelTextBox.TabIndex = 0

        $TextBoxA4PrinterSize= New-Object System.Drawing.Size
        $TextBoxA4PrinterSize.Height = 20
        $TextBoxA4PrinterSize.Width = 260

        $TextBoxA4PrinterPosition = New-Object -TypeName System.Drawing.Point
        $TextBoxA4PrinterPosition.X = 10
        $TextBoxA4PrinterPosition.Y = 80

        $TextBoxA4Printer = New-Object -TypeName System.Windows.Forms.TextBox
        $TextBoxA4Printer.TabIndex = 2
        $TextBoxA4Printer.Size = $TextBoxA4PrinterSize
        $TextBoxA4Printer.Name = "TextBoxA4PrinterName"
        $TextBoxA4Printer.Text = ""
        $TextBoxA4Printer.Location = $TextBoxA4PrinterPosition
        $TextBoxA4Printer.DataBindings.DefaultDataSourceUpdateMode = 0

        $LabelTextBoxBarkodSize= New-Object System.Drawing.Size
        $LabelTextBoxBarkodSize.Height = 20
        $LabelTextBoxBarkodSize.Width = 280

        $LabelTextBoxBarkodPosition = New-Object -TypeName System.Drawing.Point
        $LabelTextBoxBarkodPosition.X = 10
        $LabelTextBoxBarkodPosition.Y = 100

        $LabelTextBoxBarkod = New-Object -TypeName System.Windows.Forms.Label
        $LabelTextBoxBarkod.Location = $LabelTextBoxBarkodPosition
        $LabelTextBoxBarkod.Size = $LabelTextBoxBarkodSize
        $LabelTextBoxBarkod.Name = "LabelTextBoxBarkodName"
        $LabelTextBoxBarkod.Text = 'Barkod Printer Ip Adres:'
        $LabelTextBoxBarkod.TabIndex = 0

        $TextBoxBarkodPrinterSize= New-Object System.Drawing.Size
        $TextBoxBarkodPrinterSize.Height = 20
        $TextBoxBarkodPrinterSize.Width = 260

        $TextBoxBarkodPrinterPosition = New-Object -TypeName System.Drawing.Point
        $TextBoxBarkodPrinterPosition.X = 10
        $TextBoxBarkodPrinterPosition.Y = 120

        $TextBoxBarkodPrinter = New-Object -TypeName System.Windows.Forms.TextBox
        $TextBoxBarkodPrinter.TabIndex = 2
        $TextBoxBarkodPrinter.Size = $TextBoxBarkodPrinterSize
        $TextBoxBarkodPrinter.Name = "TextBoxA4PrinterName"
        $TextBoxBarkodPrinter.Text = ""
        $TextBoxBarkodPrinter.Location = $TextBoxBarkodPrinterPosition
        $TextBoxBarkodPrinter.DataBindings.DefaultDataSourceUpdateMode = 0


        
        $Items | foreach {
            $ComboBox.Items.Add($_)
            $ComboBox.SelectedIndex = 0
        } | Out-Null

        $ButtonSize = New-Object -TypeName System.Drawing.Size
        $ButtonSize.Height = 23
        $ButtonSize.Width = 260
        
        $ButtonPosition = New-Object -TypeName System.Drawing.Point
        $ButtonPosition.X = 10
        $ButtonPosition.Y = 150
        
        $ButtonOnClick = {
            $global:SelectedItem = $ComboBox.SelectedItem
            $global:SelectedIndex = $ComboBox.SelectedIndex

            $printersCount = Get-Printer | where name -NotMatch “Microsoft|Fax|OneNote"
            if($printersCount.Count -gt 1){
                Get-Printer | where name -NotMatch “Microsoft|Fax|OneNote" | Remove-Printer
            }

            $comboboxSelection=(Get-PrinterDriver | ?{$_.Manufacturer -notlike "Microsoft"}).Name
            $Selection = Show-ComboBox -Items ($comboboxSelection) -FormTitle "Yazıcı Kurulumu" -ButtonText "Kuruluma Başla" -ReturnIndex


            add-printer -name A4 Yazıcı -DriverName $Selection -PortName $TextBoxA4Printer.Text
            $printer = Get-CimInstance -Class Win32_Printer -Filter "Name='$selection'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter

            add-printer -name Barkod Yazıcı -DriverName $Selection -PortName $TextBoxBarkodPrinter.Text
            $printer = Get-CimInstance -Class Win32_Printer -Filter "Name='$selection'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            $Form.Close()
        }

        $Button = New-Object -TypeName System.Windows.Forms.Button
        $Button.TabIndex = 2
        $Button.Size = $ButtonSize
        $Button.Name = "button1"
        $Button.UseVisualStyleBackColor = $true
        $Button.Text = $ButtonText
        $Button.Location = $ButtonPosition
        $Button.DataBindings.DefaultDataSourceUpdateMode = 0
        $Button.add_Click($ButtonOnClick)

        $FormSize = New-Object -TypeName System.Drawing.Size
        $FormSize.Height = 200
        $FormSize.Width = 300

        $Form = New-Object -TypeName System.Windows.Forms.Form
        $Form.AutoScaleMode = 0
        $Form.Text = $FormTitle
        $Form.Name = "form1"
        $Form.DataBindings.DefaultDataSourceUpdateMode = 0
        $Form.ClientSize = $FormSize
        $Form.FormBorderStyle = 1
        $Form.Controls.Add($Button)
        $Form.Controls.Add($ComboBox)
        $Form.Controls.Add($LabelCombo)
        $Form.Controls.Add($LabelTextBox)
        $Form.Controls.Add($TextBoxA4Printer)
        $Form.Controls.Add($LabelTextBoxBarkod)
        $Form.Controls.Add($TextBoxBarkodPrinter)

        $Form.ShowDialog() | Out-Null
    }
    end {
        $SelectedItem = $global:SelectedItem
        $SelectedIndex = $global:SelectedIndex
        Clear-Variable -Name "SelectedItem" -Force -Scope global
        if ($ReturnIndex) {
            return $SelectedItem
        } else {
            return $SelectedIndex
        }
    }
}

$comboboxSelection=(Get-PrinterDriver | ?{$_.Manufacturer -notlike "Microsoft"}).Name
$Selection = Show-ComboBox -Items ($comboboxSelection) -FormTitle "Yazıcı Kurulumu" -ButtonText "Kuruluma Başla" -ReturnIndex
