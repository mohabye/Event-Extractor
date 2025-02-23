Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Event Viewer Filter Tool"  
$form.Size = New-Object System.Drawing.Size(600, 500)  
$form.StartPosition = "CenterScreen"  

$labelStart = New-Object System.Windows.Forms.Label
$labelStart.Text = "Start Date/Time:"
$labelStart.Location = New-Object System.Drawing.Point(10, 20)
$labelStart.AutoSize = $true
$form.Controls.Add($labelStart)

$dateTimePickerStart = New-Object System.Windows.Forms.DateTimePicker
$dateTimePickerStart.Location = New-Object System.Drawing.Point(150, 20)
$dateTimePickerStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dateTimePickerStart.CustomFormat = "MM/dd/yyyy HH:mm:ss"  
$form.Controls.Add($dateTimePickerStart)

$labelEnd = New-Object System.Windows.Forms.Label
$labelEnd.Text = "End Date/Time:"
$labelEnd.Location = New-Object System.Drawing.Point(10, 60)
$labelEnd.AutoSize = $true
$form.Controls.Add($labelEnd)

$dateTimePickerEnd = New-Object System.Windows.Forms.DateTimePicker
$dateTimePickerEnd.Location = New-Object System.Drawing.Point(150, 60)
$dateTimePickerEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dateTimePickerEnd.CustomFormat = "MM/dd/yyyy HH:mm:ss"  
$form.Controls.Add($dateTimePickerEnd)

$labelEventID = New-Object System.Windows.Forms.Label
$labelEventID.Text = "Event IDs (comma-separated):"
$labelEventID.Location = New-Object System.Drawing.Point(10, 100)
$labelEventID.AutoSize = $true
$form.Controls.Add($labelEventID)

$textBoxEventID = New-Object System.Windows.Forms.TextBox
$textBoxEventID.Location = New-Object System.Drawing.Point(150, 100)
$textBoxEventID.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textBoxEventID)

$checkboxAllEventIDs = New-Object System.Windows.Forms.CheckBox
$checkboxAllEventIDs.Text = "All Event IDs"
$checkboxAllEventIDs.Location = New-Object System.Drawing.Point(360, 100)
$checkboxAllEventIDs.AutoSize = $true
$checkboxAllEventIDs.Checked = $true  
$checkboxAllEventIDs.Add_CheckedChanged({
    $textBoxEventID.Enabled = -not $checkboxAllEventIDs.Checked
})
$form.Controls.Add($checkboxAllEventIDs)

$labelEventLevel = New-Object System.Windows.Forms.Label
$labelEventLevel.Text = "Event Level:"
$labelEventLevel.Location = New-Object System.Drawing.Point(10, 140)
$labelEventLevel.AutoSize = $true
$form.Controls.Add($labelEventLevel)

$comboBoxEventLevel = New-Object System.Windows.Forms.ComboBox
$comboBoxEventLevel.Location = New-Object System.Drawing.Point(150, 140)
$comboBoxEventLevel.Size = New-Object System.Drawing.Size(200, 20)
$comboBoxEventLevel.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxEventLevel.Items.AddRange(@("Information", "Warning", "Error", "Critical", "All"))  
$comboBoxEventLevel.SelectedIndex = 4  
$form.Controls.Add($comboBoxEventLevel)

$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Text = "Source:"
$labelSource.Location = New-Object System.Drawing.Point(10, 180)
$labelSource.AutoSize = $true
$form.Controls.Add($labelSource)

$comboBoxSource = New-Object System.Windows.Forms.ComboBox
$comboBoxSource.Location = New-Object System.Drawing.Point(150, 180)
$comboBoxSource.Size = New-Object System.Drawing.Size(200, 20)
$comboBoxSource.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxSource.Items.AddRange(@("Application", "System", "Security", "Setup", "All")) 
$comboBoxSource.SelectedIndex = 4  
$form.Controls.Add($comboBoxSource)

$labelKeyword = New-Object System.Windows.Forms.Label
$labelKeyword.Text = "Keyword (optional):"
$labelKeyword.Location = New-Object System.Drawing.Point(10, 220)
$labelKeyword.AutoSize = $true
$form.Controls.Add($labelKeyword)

$textBoxKeyword = New-Object System.Windows.Forms.TextBox
$textBoxKeyword.Location = New-Object System.Drawing.Point(150, 220)
$textBoxKeyword.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textBoxKeyword)

$labelTimeZone = New-Object System.Windows.Forms.Label
$labelTimeZone.Text = "Time Zone:"
$labelTimeZone.Location = New-Object System.Drawing.Point(10, 260)
$labelTimeZone.AutoSize = $true
$form.Controls.Add($labelTimeZone)

$comboBoxTimeZone = New-Object System.Windows.Forms.ComboBox
$comboBoxTimeZone.Location = New-Object System.Drawing.Point(150, 260)
$comboBoxTimeZone.Size = New-Object System.Drawing.Size(200, 20)
$comboBoxTimeZone.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$timeZones = [System.TimeZoneInfo]::GetSystemTimeZones() | 
    Sort-Object { $_.DisplayName } |  
    ForEach-Object { $_.DisplayName }  
$comboBoxTimeZone.Items.AddRange($timeZones)
$comboBoxTimeZone.Items.Add("Local Time")
$comboBoxTimeZone.SelectedItem = [System.TimeZoneInfo]::Local.DisplayName  
$form.Controls.Add($comboBoxTimeZone)

$labelOutputPath = New-Object System.Windows.Forms.Label
$labelOutputPath.Text = "Output File Path:"
$labelOutputPath.Location = New-Object System.Drawing.Point(10, 300)
$labelOutputPath.AutoSize = $true
$form.Controls.Add($labelOutputPath)

$textBoxOutputPath = New-Object System.Windows.Forms.TextBox
$textBoxOutputPath.Location = New-Object System.Drawing.Point(150, 300)
$textBoxOutputPath.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($textBoxOutputPath)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Browse"
$buttonBrowse.Location = New-Object System.Drawing.Point(460, 300)
$buttonBrowse.Size = New-Object System.Drawing.Size(75, 20)
$buttonBrowse.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Output File"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxOutputPath.Text = $saveFileDialog.FileName
    }
})
$form.Controls.Add($buttonBrowse)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export to CSV"
$buttonExport.Location = New-Object System.Drawing.Point(150, 340)
$buttonExport.Size = New-Object System.Drawing.Size(100, 30)
$buttonExport.BackColor = [System.Drawing.Color]::LightGray  
$buttonExport.Add_Click({
    $buttonExport.BackColor = [System.Drawing.Color]::LightBlue  
    
    $startTime = $dateTimePickerStart.Value
    $endTime = $dateTimePickerEnd.Value
    $eventIDs = $textBoxEventID.Text
    $eventLevel = $comboBoxEventLevel.SelectedItem
    $source = $comboBoxSource.SelectedItem
    $keyword = $textBoxKeyword.Text.Trim()
    $timeZone = $comboBoxTimeZone.SelectedItem
    $outputPath = $textBoxOutputPath.Text

    if ([string]::IsNullOrEmpty($outputPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output file path.", "Error")
        $buttonExport.BackColor = [System.Drawing.Color]::LightGray
        return
    }

    if ($startTime -gt $endTime) {
        [System.Windows.Forms.MessageBox]::Show("Start date/time must be earlier than end date/time.", "Error")
        $buttonExport.BackColor = [System.Drawing.Color]::LightGray
        return
    }

    if ($timeZone -ne "Local Time") {
        $tzInfo = [System.TimeZoneInfo]::GetSystemTimeZones() | Where-Object { $_.DisplayName -eq $timeZone }
        if ($tzInfo) {
            $startTime = [System.TimeZoneInfo]::ConvertTime($startTime, $tzInfo)
            $endTime = [System.TimeZoneInfo]::ConvertTime($endTime, $tzInfo)
        }
    }

    $filter = @{
        StartTime = $startTime
        EndTime = $endTime
    }

    if ($source -ne "All") {
        $filter.Add("LogName", $source)
    } else {
        $filter.Add("LogName", @("Application", "System", "Security", "Setup"))
    }

    if (-not $checkboxAllEventIDs.Checked -and $eventIDs) {
        $eventIdArray = $eventIDs -split ',' | ForEach-Object { [int]($_.Trim()) }
        $filter.Add("ID", $eventIdArray)
    }

    if ($eventLevel -ne "All") {
        $levelMap = @{
            "Information" = 4
            "Warning" = 3
            "Error" = 2
            "Critical" = 1
        }
        $filter.Add("Level", $levelMap[$eventLevel])
    }

    try {
        Write-Host "Starting event retrieval with filter: $filter"
        $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop

        Write-Host "Found $($events.Count) events before keyword filtering."

        if ($keyword) {
            $events = $events | Where-Object {
                ($_.Message -match [regex]::Escape($keyword)) -or
                ($_.Id.ToString() -eq $keyword)
            }
            Write-Host "Found $($events.Count) events after keyword filtering."
        }

        if (-not $events) {
            [System.Windows.Forms.MessageBox]::Show("No events found matching the criteria.", "No Results")
            $buttonExport.BackColor = [System.Drawing.Color]::LightGray
            return
        }

        $events | Select-Object @{
            Name = "TimeCreated"; 
            Expression = { $_.TimeCreated }
        }, Id, LevelDisplayName, ProviderName, Message | 
        Export-Csv -Path $outputPath -NoTypeInformation

        [System.Windows.Forms.MessageBox]::Show("Exported to $outputPath", "Export Complete")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error")
    } finally {
        $buttonExport.BackColor = [System.Drawing.Color]::LightGray  # Reset color after operation
    }
})
$form.Controls.Add($buttonExport)

$form.Add_Shown({ $form.Activate() })
[void] $form.ShowDialog()