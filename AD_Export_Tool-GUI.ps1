# Load necessary .NET assemblies for WPF
Add-Type -AssemblyName PresentationFramework

# Define the XAML for the WPF UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD Export Tool" Height="900" Width="1200">
    <Grid>
        <GroupBox Header="Live Logs" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="1000" Height="150">
            <TextBox x:Name="txtLiveLogs" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" IsReadOnly="True" TextWrapping="Wrap" Margin="5"/>
        </GroupBox>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,10,0" VerticalAlignment="Top">
            <Ellipse x:Name="ellipseConnectionStatus" Fill="Transparent" Stroke="Black" StrokeThickness="1" Width="20" Height="20"/>
            <Button x:Name="btnTestConnection" Content="Test Connection" Margin="10,0,0,0" Width="120" Height="30"/>
        </StackPanel>

        <GroupBox Header="Domain" HorizontalAlignment="Left" Margin="10,170,0,0" VerticalAlignment="Top" Width="200" Height="200">
            <StackPanel>
                <RadioButton x:Name="rbDomain1" Content="yourdomain.com" Margin="5"/>
                <RadioButton x:Name="rbDomain2" Content="corp.yourdomain.com" Margin="5"/>
                <RadioButton x:Name="rbDomain3" Content="test.local" Margin="5"/>
                <RadioButton x:Name="rbDomain4" Content="sec.yourdomain.net" Margin="5"/>
                <RadioButton x:Name="rbDomain5" Content="imbs.yourdomain.net" Margin="5"/>
                <RadioButton x:Name="rbDomain6" Content="sub.dmz.local" Margin="5"/>
                <RadioButton x:Name="rbDomain7" Content="cloud.yourdomain.com" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <GroupBox HorizontalAlignment="Left" Margin="10,380,0,0" VerticalAlignment="Top" Width="200" Height="Auto">
            <Grid>
                <GroupBox Header="Enable" HorizontalAlignment="Left" Margin="0,0,0,0" VerticalAlignment="Top" Width="180" Height="180">
                    <StackPanel>
                        <CheckBox x:Name="cbUserExport" Content="User" Margin="5" Background="LightGreen" VerticalAlignment="Top" HorizontalAlignment="Left" IsChecked="True"/>
                        <CheckBox x:Name="cbGroupExport" Content="Group" Margin="5" Background="LightGreen" VerticalAlignment="Center" HorizontalAlignment="Left" IsChecked="True"/>
                        <CheckBox x:Name="cbOUExport" Content="OU" Margin="5" Background="LightGreen" VerticalAlignment="Bottom" HorizontalAlignment="Left" IsChecked="True"/>
                    </StackPanel>
                </GroupBox>

                <GroupBox Header="Export Format" HorizontalAlignment="Left" Margin="0,190,0,0" VerticalAlignment="Top" Width="180" Height="100">
                    <StackPanel>
                        <RadioButton x:Name="rbCSVFormat" Content="CSV" Margin="5" VerticalAlignment="Center" HorizontalAlignment="Left" IsChecked="True"/>
                        <RadioButton x:Name="rbXMLFormat" Content="XML" Margin="5" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                    </StackPanel>
                </GroupBox>
            </Grid>
        </GroupBox>

        <GroupBox HorizontalAlignment="Left" Margin="220,170,10,10" VerticalAlignment="Top" Width="Auto" Height="Auto">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <GroupBox x:Name="groupBoxUserAttributes" Header="User Attributes" Grid.Row="0" Margin="5,5,5,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>
                        <StackPanel Orientation="Horizontal" Margin="5" Grid.Row="0">
                            <CheckBox x:Name="cbAllUserAttributes" Content="All Attributes" Margin="5" Background="LightBlue" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                            <CheckBox x:Name="cbIncludeMemberships" Content="Include Memberships" Margin="5" Background="Orange" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                        </StackPanel>
                        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Grid.Row="1">
                            <UniformGrid x:Name="uniformGridUserAttributes" Columns="4" Margin="10">
                                </UniformGrid>
                        </ScrollViewer>
                    </Grid>
                </GroupBox>

                <GroupBox x:Name="groupBoxGroupAttributes" Header="Group Attributes" Grid.Row="1" Margin="5,5,5,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>
                        <StackPanel Orientation="Horizontal" Margin="5" Grid.Row="0">
                            <CheckBox x:Name="cbAllGroupAttributes" Content="All Attributes" Margin="5" Background="LightBlue" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                        </StackPanel>
                        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Grid.Row="1">
                            <UniformGrid x:Name="uniformGridGroupAttributes" Columns="4" Margin="10">
                                </UniformGrid>
                        </ScrollViewer>
                    </Grid>
                </GroupBox>

                <Grid Grid.Row="2" Margin="5,5,5,5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <ProgressBar x:Name="progressBar" Grid.Column="0" Height="20" VerticalAlignment="Center" Minimum="0" Maximum="100" Value="0" />
                    <Button x:Name="btnExport" Grid.Column="1" Content="Export" Width="100" Height="30" HorizontalAlignment="Right" Margin="10,0,0,0"/>
                </Grid>
            </Grid>
        </GroupBox>
    </Grid>
</Window>
"@

# Load the XAML into a WPF window
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Variables for Domain Radio Buttons
$rbDomain1 = $window.FindName("rbDomain1")
$rbDomain2 = $window.FindName("rbDomain2")
$rbDomain3 = $window.FindName("rbDomain3")
$rbDomain4 = $window.FindName("rbDomain4")
$rbDomain5 = $window.FindName("rbDomain5")
$rbDomain6 = $window.FindName("rbDomain6")
$rbDomain7 = $window.FindName("rbDomain7")


# Variables for other controls
$txtLiveLogs = $window.FindName("txtLiveLogs")
$cbUserExport = $window.FindName("cbUserExport")
$cbGroupExport = $window.FindName("cbGroupExport")
$cbOUExport = $window.FindName("cbOUExport")
$rbCSVFormat = $window.FindName("rbCSVFormat")
$rbXMLFormat = $window.FindName("rbXMLFormat")
$cbAllUserAttributes = $window.FindName("cbAllUserAttributes")
$cbIncludeMemberships = $window.FindName("cbIncludeMemberships")
$cbAllGroupAttributes = $window.FindName("cbAllGroupAttributes")
$progressBar = $window.FindName("progressBar")
$btnExport = $window.FindName("btnExport")
$btnTestConnection = $window.FindName("btnTestConnection")
$ellipseConnectionStatus = $window.FindName("ellipseConnectionStatus")
$uniformGridUserAttributes = $window.FindName("uniformGridUserAttributes")
$uniformGridGroupAttributes = $window.FindName("uniformGridGroupAttributes")
$groupBoxUserAttributes = $window.FindName("groupBoxUserAttributes")
$groupBoxGroupAttributes = $window.FindName("groupBoxGroupAttributes")

# Define Functions
Function Update-Log {
    param (
        [string]$message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $window.Dispatcher.Invoke([action]{
        $txtLiveLogs.AppendText("$timestamp - $message`r`n")
        $txtLiveLogs.ScrollToEnd()
    })
}

Function Show-MessageBox {
    param (
        [string]$message
    )
    $window.Dispatcher.Invoke([action]{
        [System.Windows.MessageBox]::Show($message, "AD Export Tool", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
    })
}

Function Populate-UserAttributes {
    # Populate user attributes checkboxes
    $attributes = @(
        "accountExpires", "c", "carLicense", "cn", "co", "comment", "company", "countryCode",
        "department", "departmentNumber", "description", "displayName",
        "distinguishedName", "employeeID", "employeeNumber", "employeeType", "extensionAttribute1",
        "extensionAttribute5", "extensionAttribute7", "extensionAttribute15",
        "givenName", "homePhone", "homePostalAddress", "initials", "l", "mail", "mailNickname",
        "manager", "middleName", "mobile", "msDS-PrincipalName", "o", "ou", "postalAddress", "postalCode",
        "postOfficeBox", "preferredLanguage", "proxyAddresses", "sAMAccountName", "secretary",
        "seeAlso", "sn", "st", "street", "streetAddress", "telephoneNumber", "userPrincipalName"
    )
    $attributes | ForEach-Object {
        $window.Dispatcher.Invoke([action]{
            $checkBox = New-Object Windows.Controls.CheckBox
            $checkBox.Content = $_
            $checkBox.Margin = "5,5,5,5"
            $uniformGridUserAttributes.Children.Add($checkBox)
        })
    }
}

Function Populate-GroupAttributes {
    # Populate group attributes checkboxes
    $attributes = @(
        "name", "sAMAccountName", "groupType", "groupScope", "description", "distinguishedName",
        "info", "notes", "proxyAddresses", "managedBy", "mail", "mailNickname"
    )
    $attributes | ForEach-Object {
        $window.Dispatcher.Invoke([action]{
            $checkBox = New-Object Windows.Controls.CheckBox
            $checkBox.Content = $_
            $checkBox.Margin = "5,5,5,5"
            $uniformGridGroupAttributes.Children.Add($checkBox)
        })
    }
}

# Function to Get the Selected Domain
Function Get-SelectedDomain {
    if ($rbDomain1.IsChecked -eq $true) { return "yourdomain.com" }
    elseif ($rbDomain2.IsChecked -eq $true) { return "corp.yourdomain.com" }
    elseif ($rbDomain3.IsChecked -eq $true) { return "test.local" }
    elseif ($rbDomain4.IsChecked -eq $true) { return "sec.yourdomain.net" }
    elseif ($rbDomain5.IsChecked -eq $true) { return "imbs.yourdomain.net" }
    elseif ($rbDomain6.IsChecked -eq $true) { return "sub.dmz.local" }
    elseif ($rbDomain7.IsChecked -eq $true) { return "cloud.yourdomain.com" }
    else { return $null }
}


# Update the Test-ADConnection function to use the selected domain
Function Test-ADConnection {
    $selectedDomain = Get-SelectedDomain

    if ($null -eq $selectedDomain) {
        Show-MessageBox "Please select a domain before testing the connection."
        return
    }

    $window.Dispatcher.Invoke([action]{
        $ellipseConnectionStatus.Fill = [System.Windows.Media.Brushes]::Yellow
    })

    try {
        # Attempt to connect to the selected domain
        $adDomain = Get-ADDomain -Server $selectedDomain
        if ($adDomain -ne $null) {
            Update-Log "INFO: Successfully connected to the Active Directory domain: $selectedDomain"
            $window.Dispatcher.Invoke([action]{
                $ellipseConnectionStatus.Fill = [System.Windows.Media.Brushes]::Green
            })
        } else {
            throw "Unable to retrieve Active Directory domain information for $selectedDomain."
        }
    }
    catch {
        Update-Log "ERROR: Failed to connect to Active Directory domain $selectedDomain. Error: $_"
        $window.Dispatcher.Invoke([action]{
            $ellipseConnectionStatus.Fill = [System.Windows.Media.Brushes]::Red
        })
    }
}

Function Export-Data {
    try {
        Update-Log "INFO: Starting Export Process..."
        
        $selectedDomain = Get-SelectedDomain
        if ($null -eq $selectedDomain) {
            Show-MessageBox "Please select a domain before exporting."
            return
        }
        
        # Ensure the export directory exists
        $exportDir = "C:\Exports\" + $selectedDomain
        if (-not (Test-Path -Path $exportDir)) {
            New-Item -Path $exportDir -ItemType Directory | Out-Null
        }

        # Initialize export variables
        $userExportData = @()
        $groupExportData = @()
        $ouExportData = @()

        # Export User Data
        if ($cbUserExport.IsChecked -eq $true) {
            $selectedUserAttributes = $uniformGridUserAttributes.Children | Where-Object { $_.IsChecked } | ForEach-Object { $_.Content }

            # Ensure DistinguishedName is included in the selected attributes
            if (-not $selectedUserAttributes -contains "DistinguishedName") {
                $selectedUserAttributes += "DistinguishedName"
            }

            if (-not $selectedUserAttributes) {
                Show-MessageBox "User Export enabled but no Attribute(s) were selected. Please select at least 1 attribute or disable export by unchecking it."
                return
            }

            Update-Log "INFO: User Export selected. Collecting User Data..."

            # Fetch data from Active Directory
            try {
                $userExportData = Get-ADUser -Filter * -Property $selectedUserAttributes | Select-Object $selectedUserAttributes
            } catch {
                Update-Log "ERROR: Error fetching User data with attributes $selectedUserAttributes. Error: $_"
                return
            }

            # Include Memberships if checked
            if ($cbIncludeMemberships.IsChecked -eq $true) {
                Update-Log "INFO: Including Group Memberships in User Export."
                $userExportData | ForEach-Object {
                    $userDN = $_.DistinguishedName
                    Update-Log "DEBUG: Fetching MemberOf for user DN: $userDN"

                    if ($null -ne $userDN) {
                        try {
                            $userMemberOf = Get-ADUser -Identity $userDN -Property MemberOf | Select-Object -ExpandProperty MemberOf
                            $memberOf = if ($userMemberOf) { $userMemberOf -join "," } else { "" }
                            $_ | Add-Member -MemberType NoteProperty -Name "memberOf" -Value $memberOf -Force
                        } catch {
                            Update-Log "ERROR: Failed to retrieve MemberOf for $userDN. Error: $_"
                        }
                    } else {
                        Update-Log "ERROR: User DN is null, skipping MemberOf retrieval."
                    }
                }
            }

            if ($userExportData.Count -gt 0) {
                $totalUsers = $userExportData.Count
                $i = 0
                foreach ($user in $userExportData) {
                    $progress = [math]::Round(($i / $totalUsers) * 100)
                    $window.Dispatcher.Invoke([action]{
                        $progressBar.Value = $progress
                    })
                    Start-Sleep -Milliseconds 10 # Reduced for faster UI updates
                    Update-Log "INFO: Processing user $($i+1) of $totalUsers."
                    $i++
                }
            }

            # Export User Data
            $userFileName = "UserExport"
            $exportUserPath = Join-Path -Path $exportDir -ChildPath "$userFileName.csv"

            if ($rbCSVFormat.IsChecked -eq $true) {
                $userExportData | Export-Csv -Path $exportUserPath -NoTypeInformation -Force
            } elseif ($rbXMLFormat.IsChecked -eq $true) {
                $exportUserPath = Join-Path -Path $exportDir -ChildPath "$userFileName.xml"
                $userExportData | Export-Clixml -Path $exportUserPath
            }

            Update-Log "INFO: User data export complete to $exportUserPath"
        }

        # Export Group Data
        if ($cbGroupExport.IsChecked -eq $true) {
            $selectedGroupAttributes = $uniformGridGroupAttributes.Children | Where-Object { $_.IsChecked } | ForEach-Object { $_.Content }

            if (-not $selectedGroupAttributes) {
                Show-MessageBox "Group Export enabled but no Attribute(s) were selected. Please select at least 1 attribute or disable export by unchecking it."
                return
            }

            Update-Log "INFO: Group Export selected. Collecting Group Data..."

            # Fetch data from Active Directory
            $groupExportData = Get-ADGroup -Filter * -Property $selectedGroupAttributes | Select-Object $selectedGroupAttributes

            # Translate groupType values to legible string
            foreach ($group in $groupExportData) {
                if ($group.groupType) {
                    $groupTypeValue = $group.groupType

                    # Apply bitwise operation to determine group type
                    switch ($groupTypeValue) {
                        {($_ -band 0x80000000) -and ($_ -band 0x00000002)} { $group.groupType = "Global Security Group" }
                        {($_ -band 0x80000000) -and ($_ -band 0x00000004)} { $group.groupType = "Domain Local Security Group" }
                        {($_ -band 0x80000000) -and ($_ -band 0x00000008)} { $group.groupType = "Universal Security Group" }
                        {$_ -eq 0x00000002} { $group.groupType = "Global Distribution Group" }
                        {$_ -eq 0x00000004} { $group.groupType = "Domain Local Distribution Group" }
                        {$_ -eq 0x00000008} { $group.groupType = "Universal Distribution Group" }
                        default { $group.groupType = "Unknown Group Type" }
                    }
                }
            }

            if ($groupExportData.Count -gt 0) {
                $totalGroups = $groupExportData.Count
                $i = 0
                foreach ($group in $groupExportData) {
                    $progress = [math]::Round(($i / $totalGroups) * 100)
                    $window.Dispatcher.Invoke([action]{
                        $progressBar.Value = $progress
                    })
                    Start-Sleep -Milliseconds 10 # Reduced for faster UI updates
                    Update-Log "INFO: Processing group $($i+1) of $totalGroups."
                    $i++
                }
            }

            # Export Group Data
            $groupFileName = "GroupExport"
            $exportGroupPath = Join-Path -Path $exportDir -ChildPath "$groupFileName.csv"

            if ($rbCSVFormat.IsChecked -eq $true) {
                $groupExportData | Export-Csv -Path $exportGroupPath -NoTypeInformation -Force
            } elseif ($rbXMLFormat.IsChecked -eq $true) {
                $exportGroupPath = Join-Path -Path $exportDir -ChildPath "$groupFileName.xml"
                $groupExportData | Export-Clixml -Path $exportGroupPath
            }

            Update-Log "INFO: Group data export complete to $exportGroupPath"
        }

        # Export OU Data
        if ($cbOUExport.IsChecked -eq $true) {
            Update-Log "INFO: OU Export selected. Collecting OU Data..."

            # Fetch OU data from Active Directory
            $ouExportData = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName

            if ($ouExportData.Count -gt 0) {
                $totalOUs = $ouExportData.Count
                $i = 0
                foreach ($ou in $ouExportData) {
                    $progress = [math]::Round(($i / $totalOUs) * 100)
                    $window.Dispatcher.Invoke([action]{
                        $progressBar.Value = $progress
                    })
                    Start-Sleep -Milliseconds 10 # Reduced for faster UI updates
                    Update-Log "INFO: Processing OU $($i+1) of $totalOUs."
                    $i++
                }
            }

            # Export OU Data
            $ouFileName = "OUExport"
            $exportOUPath = Join-Path -Path $exportDir -ChildPath "$ouFileName.csv"

            if ($rbCSVFormat.IsChecked -eq $true) {
                $ouExportData | Export-Csv -Path $exportOUPath -NoTypeInformation -Force
            } elseif ($rbXMLFormat.IsChecked -eq $true) {
                $exportOUPath = Join-Path -Path $exportDir -ChildPath "$ouFileName.xml"
                $ouExportData | Export-Clixml -Path $exportOUPath
            }

            Update-Log "INFO: OU data export complete to $exportOUPath"
        }

        # Final progress bar update
        $window.Dispatcher.Invoke([action]{
            $progressBar.Value = 100
        })
        Update-Log "INFO: Export Process Complete!"
    }
    catch {
        Update-Log "ERROR: Error during export: $_"
    }
}

# Initialize UI
Populate-UserAttributes
Populate-GroupAttributes

# Event Handlers
$btnTestConnection.Add_Click({
    Update-Log "INFO: Testing Connection to the selected domain..."
    Test-ADConnection
})

$cbAllUserAttributes.Add_Checked({
    $uniformGridUserAttributes.Children | ForEach-Object {
        $_.IsChecked = $true
    }
    Update-Log "INFO: All User Attributes selected."
})

$cbAllUserAttributes.Add_Unchecked({
    $uniformGridUserAttributes.Children | ForEach-Object {
        $_.IsChecked = $false
    }
    Update-Log "INFO: All User Attributes unselected."
})

$cbAllGroupAttributes.Add_Checked({
    $uniformGridGroupAttributes.Children | ForEach-Object {
        $_.IsChecked = $true
    }
    Update-Log "INFO: All Group Attributes selected."
})

$cbAllGroupAttributes.Add_Unchecked({
    $uniformGridGroupAttributes.Children | ForEach-Object {
        $_.IsChecked = $false
    }
    Update-Log "INFO: All Group Attributes unselected."
})

$cbUserExport.Add_Checked({
    $window.Dispatcher.Invoke([action]{
        $groupBoxUserAttributes.IsEnabled = $true
    })
    Update-Log "INFO: User Export option enabled."
})

$cbUserExport.Add_Unchecked({
    $window.Dispatcher.Invoke([action]{
        $groupBoxUserAttributes.IsEnabled = $false
        $cbAllUserAttributes.IsChecked = $false
        $cbIncludeMemberships.IsChecked = $false
        $uniformGridUserAttributes.Children | ForEach-Object {
            $_.IsChecked = $false
        }
    })
    Update-Log "INFO: User Export option disabled."
})

$cbGroupExport.Add_Checked({
    $window.Dispatcher.Invoke([action]{
        $groupBoxGroupAttributes.IsEnabled = $true
    })
    Update-Log "INFO: Group Export option enabled."
})

$cbGroupExport.Add_Unchecked({
    $window.Dispatcher.Invoke([action]{
        $groupBoxGroupAttributes.IsEnabled = $false
        $cbAllGroupAttributes.IsChecked = $false
        $uniformGridGroupAttributes.Children | ForEach-Object {
            $_.IsChecked = $false
        }
    })
    Update-Log "INFO: Group Export option disabled."
})

$btnExport.Add_Click({
    Export-Data
})

# Show the window
$window.ShowDialog() | Out-Null
