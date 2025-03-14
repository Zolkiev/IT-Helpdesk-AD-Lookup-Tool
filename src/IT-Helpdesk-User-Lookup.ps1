# Active Directory User Lookup Tool
# This PowerShell script creates a GUI application for IT helpdesk staff to look up user information in Active Directory

# Import required modules
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import the Active Directory module (requires RSAT tools installed)
Import-Module ActiveDirectory


Add-Type -AssemblyName System.Drawing

# Import the Active Directory module (requires RSAT tools installed)
Import-Module ActiveDirectory

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "IT Helpdesk User Lookup Tool"
$form.Size = New-Object System.Drawing.Size(750, 700)  # Increased height from 600 to 700
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create search label
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Location = New-Object System.Drawing.Point(20, 20)
$searchLabel.Size = New-Object System.Drawing.Size(100, 25)
$searchLabel.Text = "Search For:"
$form.Controls.Add($searchLabel)

# Create search textbox
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(130, 20)
$searchBox.Size = New-Object System.Drawing.Size(350, 25)
$searchBox.ShortcutsEnabled = $true  # Ensure shortcuts like Ctrl+A, Ctrl+C, Ctrl+V work

# Enable standard editing features
$searchBox.Add_KeyUp({
    # This empty handler helps ensure proper keyboard handling
})

$form.Controls.Add($searchBox)

# Create search help label
$searchHelpLabel = New-Object System.Windows.Forms.Label
$searchHelpLabel.Location = New-Object System.Drawing.Point(130, 45)
$searchHelpLabel.Size = New-Object System.Drawing.Size(350, 20)
$searchHelpLabel.Text = "Search by name, username, email..."
$searchHelpLabel.ForeColor = [System.Drawing.Color]::Gray
$searchHelpLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($searchHelpLabel)

# Create search button
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(500, 20)
$searchButton.Size = New-Object System.Drawing.Size(100, 30)
$searchButton.Text = "Search"
$searchButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$searchButton.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($searchButton)

# Create clear button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(610, 20)
$clearButton.Size = New-Object System.Drawing.Size(100, 30)
$clearButton.Text = "Clear"
$form.Controls.Add($clearButton)

# Create export button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(500, 55)
$exportButton.Size = New-Object System.Drawing.Size(210, 30)
$exportButton.Text = "Export User Details"
$exportButton.Enabled = $false  # Initially disabled until user details are loaded
$form.Controls.Add($exportButton)

# Create results listview
$resultsList = New-Object System.Windows.Forms.ListView
$resultsList.Location = New-Object System.Drawing.Point(20, 100)
$resultsList.Size = New-Object System.Drawing.Size(690, 200)
$resultsList.View = [System.Windows.Forms.View]::Details
$resultsList.FullRowSelect = $true
$resultsList.GridLines = $true
$resultsList.Columns.Add("Username", 120)
$resultsList.Columns.Add("Full Name", 200)
$resultsList.Columns.Add("Enabled", 80)
$resultsList.Columns.Add("Locked Out", 100)
$resultsList.Columns.Add("Department", 180)
$form.Controls.Add($resultsList)

# Create context menu for user actions
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Add unlock account menu item
$unlockAccountMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$unlockAccountMenuItem.Text = "Unlock Account"
$unlockAccountMenuItem.Enabled = $false
$contextMenu.Items.Add($unlockAccountMenuItem)

# Add reset password menu item
$resetPasswordMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$resetPasswordMenuItem.Text = "Reset Password"
$resetPasswordMenuItem.Enabled = $false
$contextMenu.Items.Add($resetPasswordMenuItem)

# Add enable/disable account menu item
$toggleAccountMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$toggleAccountMenuItem.Text = "Enable/Disable Account"
$toggleAccountMenuItem.Enabled = $false
$contextMenu.Items.Add($toggleAccountMenuItem)

# Separator
$contextMenu.Items.Add("-")

# Add refresh user info menu item
$refreshMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$refreshMenuItem.Text = "Refresh User Info"
$refreshMenuItem.Enabled = $false
$contextMenu.Items.Add($refreshMenuItem)

# Assign context menu to listview
$resultsList.ContextMenuStrip = $contextMenu

# Create details groupbox
$detailsBox = New-Object System.Windows.Forms.GroupBox
$detailsBox.Location = New-Object System.Drawing.Point(20, 320)
$detailsBox.Size = New-Object System.Drawing.Size(690, 320)  # Increased height from 220 to 320
$detailsBox.Text = "User Details"
$form.Controls.Add($detailsBox)

# Create detailsText - a rich text box to display detailed user properties
$detailsText = New-Object System.Windows.Forms.RichTextBox
$detailsText.Location = New-Object System.Drawing.Point(15, 25)
$detailsText.Size = New-Object System.Drawing.Size(660, 280)  # Increased height from 180 to 280
$detailsText.ReadOnly = $true
$detailsText.Font = New-Object System.Drawing.Font("Consolas", 9)
$detailsBox.Controls.Add($detailsText)

# Create export menu strip (for CSV and HTML export options)
$exportMenu = New-Object System.Windows.Forms.ContextMenuStrip
$exportCSVItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exportCSVItem.Text = "Export to CSV"
$exportHTMLItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exportHTMLItem.Text = "Export to HTML"
$exportMenu.Items.Add($exportCSVItem)
$exportMenu.Items.Add($exportHTMLItem)

# Status bar to show operation status
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Function to format date time
function Format-ADDateTime {
    param([DateTime]$DateTime)
    if ($DateTime -eq [DateTime]::MaxValue) {
        return "Never"
    }
    else {
        return $DateTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
}

# Function to parse Active Directory date
function Convert-ADDateTime {
    param([string]$ADDateTime)
    
    if ($ADDateTime -eq $null -or $ADDateTime -eq "") {
        return "Never"
    }
    
    try {
        $datetime = [DateTime]::FromFileTime([Int64]::Parse($ADDateTime))
        return $datetime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        return "Invalid Date"
    }
}

# Function to search users
function Search-ADUsers {
    param(
        [string]$SearchText
    )
    
    $resultsList.Items.Clear()
    $detailsText.Text = ""
    
    if ([string]::IsNullOrWhiteSpace($SearchText)) {
        $statusLabel.Text = "Please enter a search term"
        return
    }
    
    try {
        $statusLabel.Text = "Searching..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Build a filter that searches across multiple properties
        $filter = "(sAMAccountName -like '*$SearchText*') -or " +
                  "(GivenName -like '*$SearchText*') -or " +
                  "(sn -like '*$SearchText*') -or " +
                  "(DisplayName -like '*$SearchText*') -or " +
                  "(mail -like '*$SearchText*') -or " +
                  "(userPrincipalName -like '*$SearchText*')"
        
        # Perform the search
        $users = Get-ADUser -Filter $filter -Properties DisplayName, Department, Enabled, LockedOut | 
                 Select-Object sAMAccountName, DisplayName, Enabled, LockedOut, Department
        
        if ($null -eq $users) {
            $statusLabel.Text = "No users found matching the search criteria"
            return
        }
        
        # Convert to array if single user returned
        if ($users -isnot [System.Array]) {
            $users = @($users)
        }
        
        if ($users.Count -eq 0) {
            $statusLabel.Text = "No users found matching the search criteria"
            return
        }
        
        # Process and display the results
        foreach ($user in $users) {
            # Debug null check for each user
            if ($null -eq $user) {
                continue  # Skip this null user
            }
            
            # Debug null check for sAMAccountName
            if ($null -eq $user.sAMAccountName) {
                $sAMAccountName = "[No Username]"
            } else {
                $sAMAccountName = $user.sAMAccountName
            }
            
            # Create the list item
            $item = New-Object System.Windows.Forms.ListViewItem($sAMAccountName)
            
            # Add display name with null checking
            if ($null -eq $user.DisplayName) {
                $item.SubItems.Add("")
            } else {
                $item.SubItems.Add($user.DisplayName)
            }
            
            # Add Enabled status with null checking and color coding
            if ($null -eq $user.Enabled) {
                $item.SubItems.Add("Unknown")
                $item.ForeColor = [System.Drawing.Color]::Gray
            } else {
                $item.SubItems.Add($user.Enabled.ToString())
                # Color coding for disabled accounts
                if ($user.Enabled -eq $false) {
                    $item.ForeColor = [System.Drawing.Color]::Red
                }
            }
            
            # Add LockedOut status with null checking
            if ($null -eq $user.LockedOut) {
                $item.SubItems.Add("Unknown")
            } else {
                $item.SubItems.Add($user.LockedOut.ToString())
                # Additional color coding for locked accounts
                if ($user.LockedOut -eq $true) {
                    $item.ForeColor = [System.Drawing.Color]::DarkOrange
                }
            }
            
            # Add Department with null checking
            if ($null -eq $user.Department) {
                $item.SubItems.Add("")
            } else {
                $item.SubItems.Add($user.Department)
            }
            
            # Store username in Tag for retrieval when clicked
            $item.Tag = $sAMAccountName
            
            # Add the item to the results list
            $resultsList.Items.Add($item)
        }
        
        $statusLabel.Text = "$($users.Count) user(s) found"
    }
    catch {
        $statusLabel.Text = "Error searching users: $($_.Exception.Message)"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Function to get detailed user information
function Get-UserDetails {
    param([string]$Username)
    
    try {
        $statusLabel.Text = "Loading user details..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Get detailed user information - specify only the properties we need
        $user = Get-ADUser -Identity $Username -Properties DisplayName, GivenName, Surname, 
            mail, Department, Title, Office, telephoneNumber, Enabled, LockedOut, 
            PasswordNeverExpires, pwdLastSet, lastLogonTimestamp, badPwdCount, 
            AccountExpirationDate, MemberOf -ErrorAction Stop
        
        # Format the output
        $sb = New-Object System.Text.StringBuilder
        
        [void]$sb.AppendLine("BASIC INFORMATION")
        [void]$sb.AppendLine("====================")
        [void]$sb.AppendLine("Username: $($user.sAMAccountName)")
        [void]$sb.AppendLine("Display Name: $($user.DisplayName)")
        [void]$sb.AppendLine("First Name: $($user.GivenName)")
        [void]$sb.AppendLine("Last Name: $($user.Surname)")
        [void]$sb.AppendLine("Email: $($user.mail)")
        [void]$sb.AppendLine("Department: $($user.Department)")
        [void]$sb.AppendLine("Title: $($user.Title)")
        [void]$sb.AppendLine("Office: $($user.Office)")
        [void]$sb.AppendLine("Phone: $($user.telephoneNumber)")
        [void]$sb.AppendLine("")
        
        [void]$sb.AppendLine("ACCOUNT STATUS")
        [void]$sb.AppendLine("====================")
        
        # Write the formatted output first
        $detailsText.Text = $sb.ToString()
        
        # Now add the status information with proper coloring
        $detailsText.SelectionFont = New-Object System.Drawing.Font("Consolas", 9)
        $detailsText.AppendText("Account Enabled: ")
        
        # Set appropriate color based on enabled status
        if ($user.Enabled) {
            $detailsText.SelectionColor = [System.Drawing.Color]::Green
        } else {
            $detailsText.SelectionColor = [System.Drawing.Color]::Red
        }
        $detailsText.AppendText("$($user.Enabled)")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black # Reset color
        $detailsText.AppendText("`r`n")
        
        # Add locked status with coloring
        $detailsText.AppendText("Account Locked: ")
        if ($user.LockedOut) {
            $detailsText.SelectionColor = [System.Drawing.Color]::Red
        } else {
            $detailsText.SelectionColor = [System.Drawing.Color]::Green
        }
        $detailsText.AppendText("$($user.LockedOut)")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black # Reset color
        $detailsText.AppendText("`r`n")
        
        # Add password never expires status with coloring
        $detailsText.AppendText("Password Never Expires: ")
        if ($user.PasswordNeverExpires) {
            $detailsText.SelectionColor = [System.Drawing.Color]::Blue
        }
        $detailsText.AppendText("$($user.PasswordNeverExpires)")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black # Reset color
        $detailsText.AppendText("`r`n")
        
        # Add password last set
        $detailsText.AppendText("Password Last Set: $(Convert-ADDateTime $user.pwdLastSet)`r`n")
        # Calculate password expiration
        if (!$user.PasswordNeverExpires -and $user.pwdLastSet -ne 0) {
            try {
                $maxPwdAge = (Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue).MaxPasswordAge.Days
                if ($maxPwdAge -gt 0) {
                    $pwdLastSet = [DateTime]::FromFileTime([Int64]::Parse($user.pwdLastSet))
                    $pwdExpires = $pwdLastSet.AddDays($maxPwdAge)
                    $passwordExpiresText = $pwdExpires.ToString('yyyy-MM-dd HH:mm:ss')
                    $detailsText.AppendText("Password Expires: $passwordExpiresText`r`n")
                    
                    $daysLeft = ($pwdExpires - (Get-Date)).Days
                    
                    # Add days until expiration with color coding based on urgency
                    $detailsText.AppendText("Days Until Expiration: ")
                    
                    # Color code based on how soon the password expires
                    if ($daysLeft -le 5) {
                        $detailsText.SelectionColor = [System.Drawing.Color]::Red
                    } elseif ($daysLeft -le 14) {
                        $detailsText.SelectionColor = [System.Drawing.Color]::Orange
                    } else {
                        $detailsText.SelectionColor = [System.Drawing.Color]::Green
                    }
                    $detailsText.AppendText("$daysLeft")
                    $detailsText.SelectionColor = [System.Drawing.Color]::Black # Reset color
                    $detailsText.AppendText("`r`n")
                    
                } else {
                    $detailsText.AppendText("Password Expiration: Unable to determine (domain policy not available)`r`n")
                }
            } catch {
                $detailsText.AppendText("Password Expiration: Error calculating ($($_.Exception.Message))`r`n")
            }
        } else {
            $detailsText.AppendText("Password Expiration: Not applicable`r`n")
        }
        
        # Add login information section
        $detailsText.AppendText("`r`nLOGIN INFORMATION`r`n")
        $detailsText.AppendText("====================")
        $detailsText.AppendText("`r`n")
        $detailsText.AppendText("Last Logon: $(Convert-ADDateTime $user.lastLogonTimestamp)`r`n")
        $detailsText.AppendText("Bad Logon Count: $($user.badPwdCount)`r`n")
        
        # Add account expiration with proper formatting
        $detailsText.AppendText("Account Expires: ")
        if ($user.AccountExpirationDate) {
            $expiryDate = $user.AccountExpirationDate.ToString('yyyy-MM-dd')
            $detailsText.AppendText("$expiryDate`r`n")
        } else {
            $detailsText.SelectionColor = [System.Drawing.Color]::Green
            $detailsText.AppendText("Never`r`n")
            $detailsText.SelectionColor = [System.Drawing.Color]::Black # Reset color
        }
        
        # Add group membership section
        $detailsText.AppendText("`r`nGROUPS MEMBERSHIP`r`n")
        $detailsText.AppendText("====================")
        $detailsText.AppendText("`r`n")
        try {
            # Get group membership using the MemberOf property instead of Get-ADPrincipalGroupMembership
            $groups = $user.MemberOf | ForEach-Object {
                (Get-ADGroup $_ -ErrorAction SilentlyContinue).Name
            } | Sort-Object
            
            if ($groups -and $groups.Count -gt 0) {
                foreach ($group in $groups) {
                    $detailsText.AppendText("- $group`r`n")
                }
            } else {
                $detailsText.AppendText("No group memberships found`r`n")
            }
        } catch {
            $detailsText.AppendText("Error retrieving group membership: $($_.Exception.Message)`r`n")
        }
        
        # Add a legend explaining the color coding
        $detailsText.AppendText("`r`nCOLOR LEGEND`r`n")
        $detailsText.AppendText("====================")
        $detailsText.AppendText("`r`n")
        
        # Green legend
        $detailsText.AppendText("")
        $detailsText.SelectionColor = [System.Drawing.Color]::Green
        $detailsText.AppendText("Green")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black
        $detailsText.AppendText(": Good Status`r`n")
        
        # Red legend
        $detailsText.AppendText("")
        $detailsText.SelectionColor = [System.Drawing.Color]::Red
        $detailsText.AppendText("Red")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black
        $detailsText.AppendText(": Issue Needing Attention`r`n")
        
        # Orange legend
        $detailsText.AppendText("")
        $detailsText.SelectionColor = [System.Drawing.Color]::Orange
        $detailsText.AppendText("Orange")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black
        $detailsText.AppendText(": Warning`r`n")
        
        # Blue legend
        $detailsText.AppendText("")
        $detailsText.SelectionColor = [System.Drawing.Color]::Blue
        $detailsText.AppendText("Blue")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black
        $detailsText.AppendText(": Special Configuration`r`n")
        
        # Set the status
        $statusLabel.Text = "User details loaded"
    }
    catch {
        $detailsText.Text = ""
        $detailsText.SelectionColor = [System.Drawing.Color]::Red
        $detailsText.AppendText("Error retrieving user details: $($_.Exception.Message)")
        $detailsText.SelectionColor = [System.Drawing.Color]::Black
        $statusLabel.Text = "Error retrieving user details"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Function to export user details to CSV
function Export-UserToCSV {
    param([string]$Username)
    
    try {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $saveDialog.Title = "Save User Details as CSV"
        $saveDialog.FileName = "$Username-Details.csv"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $statusLabel.Text = "Exporting user details to CSV..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            # Get detailed user information
            $user = Get-ADUser -Identity $Username -Properties DisplayName, GivenName, Surname, 
                mail, Department, Title, Office, telephoneNumber, Enabled, LockedOut, 
                PasswordNeverExpires, pwdLastSet, lastLogonTimestamp, badPwdCount, 
                AccountExpirationDate -ErrorAction Stop
            
            # Create custom object with user details
            $userDetails = [PSCustomObject]@{
                Username = $user.sAMAccountName
                DisplayName = $user.DisplayName
                FirstName = $user.GivenName
                LastName = $user.Surname
                Email = $user.mail
                Department = $user.Department
                Title = $user.Title
                Office = $user.Office
                Phone = $user.telephoneNumber
                AccountEnabled = $user.Enabled
                AccountLocked = $user.LockedOut
                PasswordNeverExpires = $user.PasswordNeverExpires
                PasswordLastSet = if ($user.pwdLastSet) { Convert-ADDateTime $user.pwdLastSet } else { "Never" }
                LastLogon = if ($user.lastLogonTimestamp) { Convert-ADDateTime $user.lastLogonTimestamp } else { "Never" }
                BadLogonCount = $user.badPwdCount
                AccountExpires = if ($user.AccountExpirationDate) { $user.AccountExpirationDate.ToString('yyyy-MM-dd') } else { "Never" }
            }
            
            # Export to CSV
            $userDetails | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
            
            # Export groups to a separate section of the CSV
            $groups = $user.MemberOf | ForEach-Object {
                (Get-ADGroup $_ -ErrorAction SilentlyContinue).Name
            } | Sort-Object
            
            if ($groups -and $groups.Count -gt 0) {
                $groupObjects = $groups | ForEach-Object {
                    [PSCustomObject]@{"Group Membership" = $_}
                }
                $groupObjects | Export-Csv -Path "$($saveDialog.FileName.TrimEnd('.csv'))-Groups.csv" -NoTypeInformation
            }
            
            $statusLabel.Text = "User details exported to CSV successfully"
        }
    }
    catch {
        $statusLabel.Text = "Error exporting user details: $($_.Exception.Message)"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Function to export user details to HTML
function Export-UserToHTML {
    param([string]$Username)
    
    try {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
        $saveDialog.Title = "Save User Details as HTML"
        $saveDialog.FileName = "$Username-Details.html"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $statusLabel.Text = "Exporting user details to HTML..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            # Get detailed user information
            $user = Get-ADUser -Identity $Username -Properties DisplayName, GivenName, Surname, 
                mail, Department, Title, Office, telephoneNumber, Enabled, LockedOut, 
                PasswordNeverExpires, pwdLastSet, lastLogonTimestamp, badPwdCount, 
                AccountExpirationDate, MemberOf -ErrorAction Stop
            
            # Calculate password expiration
            $passwordExpires = "Not applicable"
            $daysLeft = "N/A"
            if (!$user.PasswordNeverExpires -and $user.pwdLastSet -ne 0) {
                try {
                    $maxPwdAge = (Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue).MaxPasswordAge.Days
                    if ($maxPwdAge -gt 0) {
                        $pwdLastSet = [DateTime]::FromFileTime([Int64]::Parse($user.pwdLastSet))
                        $pwdExpires = $pwdLastSet.AddDays($maxPwdAge)
                        $passwordExpires = $pwdExpires.ToString('yyyy-MM-dd HH:mm:ss')
                        $daysLeft = ($pwdExpires - (Get-Date)).Days
                    }
                }
                catch {
                    $passwordExpires = "Error calculating"
                }
            }
            
            # Get group membership
            $groups = $user.MemberOf | ForEach-Object {
                (Get-ADGroup $_ -ErrorAction SilentlyContinue).Name
            } | Sort-Object
            
            $groupsList = if ($groups -and $groups.Count -gt 0) {
                $groups | ForEach-Object { "<li>$_</li>" } | Out-String
            } else {
                "<li>No group memberships found</li>"
            }
            
            # Generate HTML content
            $currentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>$($user.DisplayName) - AD User Details</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #0078D4; }
        .container { max-width: 800px; margin: 0 auto; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f2f2f2; }
        .section { margin-bottom: 30px; }
        .footer { font-size: 12px; color: #666; margin-top: 30px; }
        .status-enabled { color: green; font-weight: bold; }
        .status-disabled { color: red; font-weight: bold; }
        .status-warning { color: #FF8C00; font-weight: bold; } /* Dark Orange */
        .status-special { color: blue; font-weight: bold; }
        .groups-list { list-style-type: none; padding-left: 0; }
        .groups-list li { padding: 3px 0; }
        .legend { display: flex; margin-bottom: 10px; }
        .legend-item { margin-right: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Active Directory User Details</h1>
        <div class="section">
            <h2>Basic Information</h2>
            <table>
                <tr><th>Username</th><td>$($user.sAMAccountName)</td></tr>
                <tr><th>Display Name</th><td>$($user.DisplayName)</td></tr>
                <tr><th>First Name</th><td>$($user.GivenName)</td></tr>
                <tr><th>Last Name</th><td>$($user.Surname)</td></tr>
                <tr><th>Email</th><td>$($user.mail)</td></tr>
                <tr><th>Department</th><td>$($user.Department)</td></tr>
                <tr><th>Title</th><td>$($user.Title)</td></tr>
                <tr><th>Office</th><td>$($user.Office)</td></tr>
                <tr><th>Phone</th><td>$($user.telephoneNumber)</td></tr>
            </table>
        </div>
        
        <div class="section">
            <h2>Account Status</h2>
            
            <div class="legend">
                <div class="legend-item"><span class="status-enabled">■</span> Good Status</div>
                <div class="legend-item"><span class="status-disabled">■</span> Issue Needing Attention</div>
                <div class="legend-item"><span class="status-warning">■</span> Warning</div>
                <div class="legend-item"><span class="status-special">■</span> Special Configuration</div>
            </div>
            
            <table>
                <tr>
                    <th>Account Enabled</th>
                    <td class="$(if ($user.Enabled) { 'status-enabled' } else { 'status-disabled' })">$($user.Enabled)</td>
                </tr>
                <tr>
                    <th>Account Locked</th>
                    <td class="$(if ($user.LockedOut) { 'status-disabled' } else { 'status-enabled' })">$($user.LockedOut)</td>
                </tr>
                <tr>
                    <th>Password Never Expires</th>
                    <td class="$(if ($user.PasswordNeverExpires) { 'status-special' } else { '' })">$($user.PasswordNeverExpires)</td>
                </tr>
                <tr><th>Password Last Set</th><td>$(Convert-ADDateTime $user.pwdLastSet)</td></tr>
                <tr><th>Password Expires</th><td>$passwordExpires</td></tr>
                
                <tr>
                    <th>Days Until Expiration</th>
                    <td class="$(if ($daysLeft -le 5) { 'status-disabled' } elseif ($daysLeft -le 14) { 'status-warning' } elseif ($daysLeft -gt 14) { 'status-enabled' } else { '' })">$daysLeft</td>
                </tr>
            </table>
        </div>
        
        <div class="section">
            <h2>Login Information</h2>
            <table>
                <tr><th>Last Logon</th><td>$(Convert-ADDateTime $user.lastLogonTimestamp)</td></tr>
                <tr><th>Bad Logon Count</th><td>$($user.badPwdCount)</td></tr>
                <tr><th>Account Expires</th><td>$(if ($user.AccountExpirationDate) { $user.AccountExpirationDate.ToString('yyyy-MM-dd') } else { 'Never' })</td></tr>
            </table>
        </div>
        
        <div class="section">
            <h2>Group Memberships</h2>
            <ul class="groups-list">
                $groupsList
            </ul>
        </div>
        
        <div class="footer">
            Report generated on $currentDate by IT Helpdesk AD Lookup Tool
        </div>
    </div>
</body>
</html>
"@
            
            # Save HTML file
            $html | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            
            $statusLabel.Text = "User details exported to HTML successfully"
        }
    }
    catch {
        $statusLabel.Text = "Error exporting user details: $($_.Exception.Message)"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Wire up search button event
$searchButton.Add_Click({
    Search-ADUsers -SearchText $searchBox.Text
})

# Wire up clear button event
$clearButton.Add_Click({
    $searchBox.Text = ""
    $resultsList.Items.Clear()
    $detailsText.Text = ""
    $exportButton.Enabled = $false
    $statusLabel.Text = "Ready"
})

# Wire up listview selection changed event
$resultsList.Add_SelectedIndexChanged({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $selectedUsername = $resultsList.SelectedItems[0].Tag
        Get-UserDetails -Username $selectedUsername
    }
})

# Wire up Enter key to search
$searchBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $searchButton.PerformClick()
        $_.SuppressKeyPress = $true
    }
    # Support for Ctrl+A (Select All)
    elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::A -and $_.Control) {
        $searchBox.SelectAll()
        $_.SuppressKeyPress = $true
    }
})

# Wire up export button event
$exportButton.Add_Click({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $exportMenu.Show($exportButton, 0, $exportButton.Height)
    }
})

# Wire up export menu items
$exportCSVItem.Add_Click({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $selectedUsername = $resultsList.SelectedItems[0].Tag
        Export-UserToCSV -Username $selectedUsername
    }
})

$exportHTMLItem.Add_Click({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $selectedUsername = $resultsList.SelectedItems[0].Tag
        Export-UserToHTML -Username $selectedUsername
    }
})

# Create a wrapper for the Get-UserDetails function to enable export button when details are loaded
$originalGetUserDetails = ${function:Get-UserDetails}

# Redefine the function to add export button enabling/disabling
function Get-UserDetails {
    param([string]$Username)
    
    if ([string]::IsNullOrWhiteSpace($Username)) {
        $statusLabel.Text = "Error: No username provided"
        $exportButton.Enabled = $false
        return
    }
    
    # Call the original function implementation
    & $originalGetUserDetails $Username
    
    # Enable the export button when details are loaded successfully
    if ($detailsText.Text -notmatch "Error retrieving user details") {
        $exportButton.Enabled = $true
    } else {
        $exportButton.Enabled = $false
    }
}

# Function to unlock a user account
function Unlock-UserAccount {
    param([string]$Username)
    
    try {
        $statusLabel.Text = "Unlocking account for " + $Username + "..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Attempt to unlock the account
        Unlock-ADAccount -Identity $Username -ErrorAction Stop
        
        # Show success message
        [System.Windows.Forms.MessageBox]::Show(
            "Account for " + $Username + " has been successfully unlocked.", 
            "Unlock Account", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        # Refresh user details to show updated status
        Get-UserDetails -Username $Username
        $statusLabel.Text = "Account unlocked for " + $Username
        
        # Update the selected item in the list view
        if ($resultsList.SelectedItems.Count -gt 0) {
            $selectedIndex = $resultsList.SelectedIndices[0]
            $resultsList.Items[$selectedIndex].SubItems[3].Text = "False" # Update Locked Out column
            $resultsList.Items[$selectedIndex].ForeColor = [System.Drawing.Color]::Black # Reset color
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error unlocking account: $($_.Exception.Message)", 
            "Unlock Account Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $statusLabel.Text = "Error unlocking account"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Function to reset a user's password
function Reset-UserPassword {
    param([string]$Username)
    
    try {
        # Create a form for password reset
        $resetForm = New-Object System.Windows.Forms.Form
        $resetForm.Text = "Reset Password for $Username"
        $resetForm.Size = New-Object System.Drawing.Size(400, 250)
        $resetForm.StartPosition = "CenterParent"
        $resetForm.FormBorderStyle = "FixedDialog"
        $resetForm.MaximizeBox = $false
        $resetForm.MinimizeBox = $false
        $resetForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        
        # Create password label
        $passwordLabel = New-Object System.Windows.Forms.Label
        $passwordLabel.Location = New-Object System.Drawing.Point(20, 20)
        $passwordLabel.Size = New-Object System.Drawing.Size(350, 20)
        $passwordLabel.Text = "Enter new password for " + $Username + ":"
        $resetForm.Controls.Add($passwordLabel)
        
        # Create password textbox
        $passwordBox = New-Object System.Windows.Forms.TextBox
        $passwordBox.Location = New-Object System.Drawing.Point(20, 45)
        $passwordBox.Size = New-Object System.Drawing.Size(350, 25)
        $passwordBox.PasswordChar = '*'
        $resetForm.Controls.Add($passwordBox)
        
        # Create confirm password label
        $confirmLabel = New-Object System.Windows.Forms.Label
        $confirmLabel.Location = New-Object System.Drawing.Point(20, 80)
        $confirmLabel.Size = New-Object System.Drawing.Size(350, 20)
        $confirmLabel.Text = "Confirm new password:"
        $resetForm.Controls.Add($confirmLabel)
        
        # Create confirm password textbox
        $confirmBox = New-Object System.Windows.Forms.TextBox
        $confirmBox.Location = New-Object System.Drawing.Point(20, 105)
        $confirmBox.Size = New-Object System.Drawing.Size(350, 25)
        $confirmBox.PasswordChar = '*'
        $resetForm.Controls.Add($confirmBox)
        
        # Create checkbox for "User must change password at next logon"
        $mustChangeBox = New-Object System.Windows.Forms.CheckBox
        $mustChangeBox.Location = New-Object System.Drawing.Point(20, 140)
        $mustChangeBox.Size = New-Object System.Drawing.Size(350, 20)
        $mustChangeBox.Text = "User must change password at next logon"
        $mustChangeBox.Checked = $true
        $resetForm.Controls.Add($mustChangeBox)
        
        # Create reset button
        $resetButton = New-Object System.Windows.Forms.Button
        $resetButton.Location = New-Object System.Drawing.Point(110, 170)
        $resetButton.Size = New-Object System.Drawing.Size(80, 30)
        $resetButton.Text = "Reset"
        $resetButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $resetForm.AcceptButton = $resetButton
        $resetForm.Controls.Add($resetButton)
        
        # Create cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(200, 170)
        $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $resetForm.CancelButton = $cancelButton
        $resetForm.Controls.Add($cancelButton)
        
        # Show the form and process the result
        $result = $resetForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $newPassword = $passwordBox.Text
            $confirmPassword = $confirmBox.Text
            $mustChange = $mustChangeBox.Checked
            
            # Validate password
            if ([string]::IsNullOrWhiteSpace($newPassword)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Password cannot be empty.", 
                    "Password Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            
            if ($newPassword -ne $confirmPassword) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Passwords do not match.", 
                    "Password Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            
            $statusLabel.Text = "Resetting password for " + $Username + "..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            # Convert plain text password to secure string
            $securePassword = ConvertTo-SecureString -String $newPassword -AsPlainText -Force
            
            # Reset the password
            Set-ADAccountPassword -Identity $Username -NewPassword $securePassword -Reset -ErrorAction Stop
            
            # Set 'User must change password at next logon' if checked
            if ($mustChange) {
                Set-ADUser -Identity $Username -ChangePasswordAtLogon $true -ErrorAction Stop
            }
            
            # Show success message
            [System.Windows.Forms.MessageBox]::Show(
                "Password for " + $Username + " has been successfully reset." + 
                $(if ($mustChange) { "`nUser will be required to change password at next logon." } else { "" }), 
                "Password Reset", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # Refresh user details to show updated status
            Get-UserDetails -Username $Username
            $statusLabel.Text = "Password reset for " + $Username
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error resetting password: $($_.Exception.Message)", 
            "Password Reset Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $statusLabel.Text = "Error resetting password"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Function to enable or disable a user account
function Toggle-UserAccountStatus {
    param(
        [string]$Username,
        [bool]$CurrentStatus
    )
    
    try {
        $action = if ($CurrentStatus) { "disable" } else { "enable" }
        
        # Ask for confirmation
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to " + $action + " the account for " + $Username + "?",
            "Confirm Account $($action.Substring(0,1).ToUpper() + $action.Substring(1))",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $statusLabel.Text = $($action.Substring(0,1).ToUpper() + $action.Substring(1)) + "ing account for " + $Username + "..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            # Enable or disable the account
            if ($CurrentStatus) {
                # Disable account
                Disable-ADAccount -Identity $Username -ErrorAction Stop
            } else {
                # Enable account
                Enable-ADAccount -Identity $Username -ErrorAction Stop
            }
            
            # Show success message
            [System.Windows.Forms.MessageBox]::Show(
                "Account for " + $Username + " has been successfully " + $action + "d.",
                "Account $($action.Substring(0,1).ToUpper() + $action.Substring(1))d",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # Update the selected item in the list view
            if ($resultsList.SelectedItems.Count -gt 0) {
                $selectedIndex = $resultsList.SelectedIndices[0]
                $newStatus = !$CurrentStatus
                $resultsList.Items[$selectedIndex].SubItems[2].Text = $newStatus.ToString() # Update Enabled column
                
                # Update color coding
                if (!$newStatus) {
                    $resultsList.Items[$selectedIndex].ForeColor = [System.Drawing.Color]::Red
                } else {
                    # Reset color unless locked out
                    $lockedStatus = $resultsList.Items[$selectedIndex].SubItems[3].Text
                    if ($lockedStatus -eq "True") {
                        $resultsList.Items[$selectedIndex].ForeColor = [System.Drawing.Color]::DarkOrange
                    } else {
                        $resultsList.Items[$selectedIndex].ForeColor = [System.Drawing.Color]::Black
                    }
                }
            }
            
            # Refresh user details to show updated status
            Get-UserDetails -Username $Username
            $statusLabel.Text = "Account " + $action + "d for " + $Username
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error ${action}ing account: $($_.Exception.Message)",
            "Account $($action.Substring(0,1).ToUpper() + $action.Substring(1)) Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $statusLabel.Text = "Error ${action}ing account"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Function to update context menu item state based on user account status
function UpdateContextMenuState {
    param([Microsoft.ActiveDirectory.Management.ADUser]$User)
    
    # Update menu items based on user properties
    if ($User) {
        # Unlock account - enable only if account is locked
        $unlockAccountMenuItem.Enabled = $User.LockedOut
        
        # Reset password - always enabled for valid user
        $resetPasswordMenuItem.Enabled = $true
        
        # Enable/Disable account - always enabled for valid user, text changes based on current state
        $toggleAccountMenuItem.Enabled = $true
        $toggleAccountMenuItem.Text = if ($User.Enabled) { "Disable Account" } else { "Enable Account" }
        
        # Refresh - always enabled for valid user
        $refreshMenuItem.Enabled = $true
    } else {
        # Disable all menu items if no valid user
        $unlockAccountMenuItem.Enabled = $false
        $resetPasswordMenuItem.Enabled = $false
        $toggleAccountMenuItem.Enabled = $false
        $refreshMenuItem.Enabled = $false
    }
}

# Update the Get-UserDetails function to call UpdateContextMenuState
# Add this line to the Get-UserDetails function before the "return $user" line:
# UpdateContextMenuState -User $user

# Add these event handlers before the "Show the form" line at the end of the script:
$unlockAccountMenuItem.Add_Click({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $selectedUsername = $resultsList.SelectedItems[0].Tag
        Unlock-UserAccount -Username $selectedUsername
    }
})

$resetPasswordMenuItem.Add_Click({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $selectedUsername = $resultsList.SelectedItems[0].Tag
        Reset-UserPassword -Username $selectedUsername
    }
})

$toggleAccountMenuItem.Add_Click({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $selectedUsername = $resultsList.SelectedItems[0].Tag
        $user = Get-ADUser -Identity $selectedUsername -Properties Enabled
        Toggle-UserAccountStatus -Username $selectedUsername -CurrentStatus $user.Enabled
    }
})

$refreshMenuItem.Add_Click({
    if ($resultsList.SelectedItems.Count -gt 0) {
        $selectedUsername = $resultsList.SelectedItems[0].Tag
        Get-UserDetails -Username $selectedUsername
        $statusLabel.Text = "User details refreshed"
    }
})

# Add right-click support for the results list
$resultsList.Add_MouseUp({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        if ($resultsList.SelectedItems.Count -gt 0) {
            # Context menu will show automatically since it's assigned to the ListView
            # But we ensure items are properly enabled/disabled
            $selectedUsername = $resultsList.SelectedItems[0].Tag
            $user = Get-ADUser -Identity $selectedUsername -Properties Enabled, LockedOut
            UpdateContextMenuState -User $user
        }
    }
})

# Show the form
[void]$form.ShowDialog()
