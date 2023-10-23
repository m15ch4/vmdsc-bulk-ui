# Load necessary assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Import VMware PowerCLI module
Import-Module VMware.PowerCLI -ErrorAction SilentlyContinue

# Create a new DataTable
$newDataTable = New-Object System.Data.DataTable
[void]$newDataTable.Columns.Add('Name', [string])
[void]$newDataTable.Columns.Add('Uuid', [string])
[void]$newDataTable.Columns.Add('Current CPU', [int])
[void]$newDataTable.Columns.Add('Current MemoryMB', [int])
[void]$newDataTable.Columns.Add('Current CoresPerCPU', [int])
[void]$newDataTable.Columns.Add('Desired CPU', [int])
[void]$newDataTable.Columns.Add('Desired MemoryMB', [int])
[void]$newDataTable.Columns.Add('Desired CoresPerCPU', [int])
[void]$newDataTable.Columns.Add('Action', [string])

# XAML code for the UI
[xml]$xaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="VMDSC Bulk UI" Height="880" Width="1200">
    <Grid>
        <Menu HorizontalAlignment="Left" VerticalAlignment="Top">
            <MenuItem Header="_File">
                <MenuItem Name="AboutItem" Header="_About"/>
            </MenuItem>
        </Menu>

        <Label Name="CsvLabel" Content="Select CSV:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,30,0,0"/>
        <TextBox Name="FilePathBox" HorizontalAlignment="Left" Height="25" VerticalAlignment="Top" Width="280" Margin="120,30,0,0" IsReadOnly="True"/>
        <Button Name="BrowseButton" Content="Browse" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Height="25" Margin="410,30,0,0"/>

        <Label Name="vCenterLabel" Content="vCenter hostname:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,70,0,0"/>
        <TextBox Name="vCenterBox" HorizontalAlignment="Left" Height="25" VerticalAlignment="Top" Width="280" Margin="120,70,0,0"/>

        <Label Name="vCenterUserLabel" Content="vCenter User:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,110,0,0"/>
        <TextBox Name="vCenterUserBox" HorizontalAlignment="Left" Height="25" VerticalAlignment="Top" Width="280" Margin="120,110,0,0"/>

        <Label Name="vCenterPassLabel" Content="vCenter password:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,150,0,0"/>
        <PasswordBox Name="vCenterPassBox" HorizontalAlignment="Left" Height="25" VerticalAlignment="Top" Width="280" Margin="120,150,0,0"/>

        <Label Name="VMDSCLabel" Content="VMDSC hostname:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,190,0,0"/>
        <TextBox Name="VMDSCBox" HorizontalAlignment="Left" Height="25" VerticalAlignment="Top" Width="280" Margin="120,190,0,0"/>

        <Button Name="PlanButton" Content="Plan" HorizontalAlignment="Center" VerticalAlignment="Top" Width="100" Height="25" Margin="10,230,0,0"/>

        <DataGrid Name="DataGrid" HorizontalAlignment="Left" Height="500" VerticalAlignment="Top" Width="1160" Margin="10,270,0,0" AutoGenerateColumns="True" CanUserAddRows='False'/>
        
        <Button Name="ApplyButton" IsEnabled="False" Content="Apply" HorizontalAlignment="Center" VerticalAlignment="Top" Width="100" Height="25" Margin="0,780,0,10"/>

        <StatusBar Name="StatusBar" VerticalAlignment="Bottom">
            <TextBlock Name="StatusText" Text="Ready" />
        </StatusBar>
    </Grid>
</Window>
'@

# Create XmlReader and load XAML code
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml.OuterXml))
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Connect controls to variables
$AboutItem = $Window.FindName('AboutItem')
$BrowseButton = $Window.FindName('BrowseButton')
$PlanButton = $Window.FindName('PlanButton')
$FilePathBox = $Window.FindName('FilePathBox')
$vCenterBox = $Window.FindName('vCenterBox')
$DataGrid = $Window.FindName('DataGrid')
$vCenterUserBox = $Window.FindName('vCenterUserBox')
$vCenterPassBox = $Window.FindName('vCenterPassBox')
$VMDSCBox = $Window.FindName('VMDSCBox')
$ApplyButton = $Window.FindName('ApplyButton')
$StatusBar = $Window.FindName('StatusBar')
$StatusText = $Window.FindName('StatusText')


# Browse button event to open file dialog for CSV file selection
$BrowseButton.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    if ($OpenFileDialog.ShowDialog() -eq 'OK') {
        $FilePathBox.Text = $OpenFileDialog.FileName
    }
})

# Bind DataTable to DataGrid
$dataGrid.ItemsSource = $newDataTable.DefaultView

# Create global variable to store vCenter credentials
$credentials = $null

# Run button event to execute operations with vCenter Server
$PlanButton.Add_Click({
    $StatusText.Text = "Planning..."
    $PlanButton.IsEnabled = $false
    $PlanButton.Content = "Planning..."
    # Clear the DataGrid each time the Run button is clicked
    if ($newDataTable.Rows.Count -gt 0) {
        $newDataTable.Rows.Clear()
    }

    # Force the UI to update
    $Window.Dispatcher.InvokeAsync({

        $securePassword = ConvertTo-SecureString -String $vCenterPassBox.Password -AsPlainText -Force
        $credentials = New-Object System.Management.Automation.PSCredential($vCenterUserBox.Text, $securePassword)

        # Connect to vCenter Server
        Connect-VIServer -Server $vCenterBox.Text -Credential $credentials -ErrorAction SilentlyContinue

        # Connect to VMDSC
        Connect-VMDSC -vmdsc $VMDSCBox.Text -username $credentials.UserName -password $credentials.GetNetworkCredential().Password

        if ($?) { # Check if the last command was successful
            # Read VM names from the CSV file and check if they exist
            ##$vmNames = Import-Csv -Path $FilePathBox.Text | Select-Object -ExpandProperty vm_name
            $vmNames = Import-Csv -Path $FilePathBox.Text

            foreach ($vmName in $vmNames) {
                $vm = Get-VM -Name $vmName.vm_name -ErrorAction SilentlyContinue
                if ($vm) {
                    # Create new row in table
                    $newRow = $newDataTable.NewRow()
                    $newRow["Name"] = $vm.Name
                    $newRow["Uuid"] = $vm | ForEach-Object{(Get-View $_.Id).config.uuid}
                    $newRow["Current CPU"] = $vm.NumCpu
                    $newRow["Current MemoryMB"] = $vm.MemoryMB
                    $newRow["Current CoresPerCPU"] = $vm.CoresPerSocket
                    $newRow["Desired CPU"] = $vmName.cpu
                    $newRow["Desired MemoryMB"] = $vmName.mem
                    $newRow["Desired CoresPerCPU"] = $vmName.cores

                    # Check if VMDSC configuration already exists for $vm
                    $vmDSCConfig = Get-VMDSC -uuid $newRow["Uuid"]
                    if (($vmDSCConfig.GetType().Name -ne "String")) {
                        # VMDSC configuration exists
                        # new desired values equal to current values
                        if (($vm.NumCpu -eq $vmName.cpu) -and ($vm.MemoryMB -eq $vmName.mem) -and ($vm.CoresPerSocket -eq $vmName.cores)) {
                            # Remove configuration as no changes to vCenter VM configuration is needed
                            ##Clear-VMDSC -uuid $vmName.uuid
                            $newRow["Action"] = "Remove"
                        } elseif (($vmName.cpu -eq $vmDSCConfig.cpu) -and ($vmName.mem -eq $vmDSCConfig.memsize) -and ($vmName.cores -eq $vmDSCConfig.cores_per_socket)) {
                        # new desired values equal to old desired values   
                            # do not touch VMDSC configuration for this vm :) 
                            $newRow["Action"] = "None (VMDSC config exists)"
                        } else {
                        # new desired values are different from both old desired values and current values
                            # Update configuration
                            ##Set-VMDSC -uuid $temp.Uuid -cpu $temp.DesiredCpu -mem $temp.DesiredMemory -corespersocket $temp.DesiredCores
                            $newRow["Action"] = "Update"        
                        }
                    } elseif (($vmDSCConfig.GetType().Name -eq "String") -and $vmDSCConfig.Contains("Config not found")) {
                        # No VMDSC configuration exists
                        # new desired values equal to current values
                        if (($vm.NumCpu -eq $vmName.cpu) -and ($vm.MemoryMB -eq $vmName.mem) -and ($vm.CoresPerSocket -eq $vmName.cores)) {
                            # do not create VMDSC configuration as no change is needed
                            $newRow["Action"] = "None (vCenter values ok)"
                        } else {
                            # Add new VMDSC configuration
                            ##Add-VMDSC -uuid ... -cpu $vmName.NumCpu -mem $vmName.MemoryMB -corespersocket $vmName.CoresPerSocket
                            $newRow["Action"] = "Create"
                        }
                    } else {
                        # do not create VMDSC configuration as some error occured
                        $newRow["Action"] = "None (error)"
                    }
                    $newDataTable.Rows.Add($newRow)
                }
            }
            $PlanButton.IsEnabled = $true
            $PlanButton.Content = "Plan"
            $ApplyButton.IsEnabled = $true
            $StatusText.Text = "Planning complete."
        } else {
            [System.Windows.MessageBox]::Show('Failed to connect to vCenter. Please check your credentials and vCenter address.', 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
})

$ApplyButton.Add_Click({
    $StatusText.Text = "Applying configuration..."

    $Window.Dispatcher.InvokeAsync({
        $securePassword = ConvertTo-SecureString -String $vCenterPassBox.Password -AsPlainText -Force
        $credentials = New-Object System.Management.Automation.PSCredential($vCenterUserBox.Text, $securePassword)
        # Connect to VMDSC
        Connect-VMDSC -vmdsc $VMDSCBox.Text -username $credentials.UserName -password $credentials.GetNetworkCredential().Password

        foreach ($task in $newDataTable.Rows) {
            if (($task["Action"]).Contains("None")) {
                Write-Host "No action for " $task["Name"]
            } elseif ($task["Action"] -eq "Create") {
                Write-Host "Create action for " $task["Name"]
                Add-VMDSC -uuid $task["Uuid"] -cpu $task["Desired CPU"] -mem $task["Desired MemoryMB"] -corespersocket $task["Desired CoresPerCPU"]
            } elseif ($task["Action"] -eq "Update") {
                Write-Host "Update action for " $task["Name"]
                Set-VMDSC -uuid $task["Uuid"] -cpu $task["Desired CPU"] -mem $task["Desired MemoryMB"] -corespersocket $task["Desired CoresPerCPU"]
            } elseif ($task["Action"] -eq "Remove") {
                Write-Host "Remove action for " $task["Name"]
                Clear-VMDSC -uuid $task["Uuid"]
            }
        }
        $ApplyButton.IsEnabled = $false
        $StatusText.Text = "Done."
    })
})

# Add a Click event handler to the About menu item
$AboutItem.Add_Click({
    [System.Windows.MessageBox]::Show("VMDSC Bulk UI v1.0`nAuthor: Michal Czerwinski (mczerwinski@vmware.com)`nCopyright 2023 @ VMware" , 'About')
})

# Show WPF window
$Window.ShowDialog() | Out-Null
