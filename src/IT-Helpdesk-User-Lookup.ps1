# Active Directory User Lookup Tool
# This PowerShell script creates a GUI application for IT helpdesk staff to look up user information in Active Directory

# Import required modules
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import the Active Directory module (requires RSAT tools installed)
Import-Module ActiveDirectory

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "IT Helpdesk User Lookup Tool"
$form.Size = New-Object System.Drawing.Size(750, 600)
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
$form.Controls.Add($searchBox)

# Create search type dropdown
$searchTypeLabel = New-Object System.Windows.Forms.Label
$searchTypeLabel.Location = New-Object System.Drawing.Point(20, 55)
$searchTypeLabel.Size = New-Object System.Drawing.Size(100, 25)
$searchTypeLabel.Text = "Search By:"
$form.Controls.Add($searchTypeLabel)

$searchTypeBox = New-Object System.Windows.Forms.ComboBox
$searchTypeBox.Location = New-Object System.Drawing.Point(130, 55)
$searchTypeBox.Size = New-Object System.Drawing.Size(200, 25)
$searchTypeBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$searchTypeBox.Items.Add("Username (sAMAccountName)")
[void]$searchTypeBox.Items.Add("First Name (givenName)")
[void]$searchTypeBox.Items.Add("Last Name (sn)")
[void]$searchTypeBox.Items.Add("Display Name (displayName)")
[void]$searchTypeBox.Items.Add("Email Address (mail)")
$searchTypeBox.SelectedIndex = 0
$form.Controls.Add($searchTypeBox)

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

# Create details groupbox
$detailsBox = New-Object System.Windows.Forms.GroupBox
$detailsBox.Location = New-Object System.Drawing.Point(20, 320)
$detailsBox.Size = New-Object System.Drawing.Size(690, 220)
$detailsBox.Text = "User Details"
$form.Controls.Add($detailsBox)

# Create detailsText - a rich text box to display detailed user properties
$detailsText = New-Object System.Windows.Forms.RichTextBox
$detailsText.Location = New-Object System.Drawing.Point(15, 25)
$detailsText.Size = New-Object System.Drawing.Size(660, 180)
$detailsText.ReadOnly = $true
$detailsText.Font = New-Object System.Drawing.Font("Consolas", 9)
$detailsBox.Controls.Add($detailsText)

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
        [string]$SearchText,
        [string]$SearchBy
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
        
        # Determine search property based on dropdown selection
        $property = switch ($SearchBy) {
            "Username (sAMAccountName)" { "sAMAccountName" }
            "First Name (givenName)" { "givenName" }
            "Last Name (sn)" { "sn" }
            "Display Name (displayName)" { "displayName" }
            "Email Address (mail)" { "mail" }
        }
        
        # Build the search filter
        $filter = "$property -like '*$SearchText*'"
        
        # Perform the search
        $users = Get-ADUser -Filter $filter -Properties DisplayName, Department, Enabled, LockedOut | 
                 Select-Object sAMAccountName, DisplayName, Enabled, LockedOut, Department
        
        if ($users -eq $null -or $users.Count -eq 0) {
            $statusLabel.Text = "No users found matching the search criteria"
            return
        }
        
        # Process and display the results
        foreach ($user in $users) {
            $item = New-Object System.Windows.Forms.ListViewItem($user.sAMAccountName)
            $item.SubItems.Add($user.DisplayName)
            $item.SubItems.Add($user.Enabled.ToString())
            $item.SubItems.Add($user.LockedOut.ToString())
            $item.SubItems.Add($user.Department)
            $item.Tag = $user.sAMAccountName  # Store username in Tag for retrieval when clicked
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
        [void]$sb.AppendLine("Account Enabled: $($user.Enabled)")
        [void]$sb.AppendLine("Account Locked: $($user.LockedOut)")
        [void]$sb.AppendLine("Password Never Expires: $($user.PasswordNeverExpires)")
        [void]$sb.AppendLine("Password Last Set: $(Convert-ADDateTime $user.pwdLastSet)")
        
        # Calculate password expiration
        if (!$user.PasswordNeverExpires -and $user.pwdLastSet -ne 0) {
            try {
                $maxPwdAge = (Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue).MaxPasswordAge.Days
                if ($maxPwdAge -gt 0) {
                    $pwdLastSet = [DateTime]::FromFileTime([Int64]::Parse($user.pwdLastSet))
                    $pwdExpires = $pwdLastSet.AddDays($maxPwdAge)
                    [void]$sb.AppendLine("Password Expires: $($pwdExpires.ToString('yyyy-MM-dd HH:mm:ss'))")
                    $daysLeft = ($pwdExpires - (Get-Date)).Days
                    [void]$sb.AppendLine("Days Until Expiration: $daysLeft")
                } else {
                    [void]$sb.AppendLine("Password Expiration: Unable to determine (domain policy not available)")
                }
            } catch {
                [void]$sb.AppendLine("Password Expiration: Error calculating ($($_.Exception.Message))")
            }
        } else {
            [void]$sb.AppendLine("Password Expiration: Not applicable")
        }
        
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("LOGIN INFORMATION")
        [void]$sb.AppendLine("====================")
        [void]$sb.AppendLine("Last Logon: $(Convert-ADDateTime $user.lastLogonTimestamp)")
        [void]$sb.AppendLine("Bad Logon Count: $($user.badPwdCount)")
        [void]$sb.AppendLine("Account Expires: $(if ($user.AccountExpirationDate) { $user.AccountExpirationDate.ToString('yyyy-MM-dd') } else { 'Never' })")
        
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("GROUPS MEMBERSHIP")
        [void]$sb.AppendLine("====================")
        try {
            # Get group membership using the MemberOf property instead of Get-ADPrincipalGroupMembership
            $groups = $user.MemberOf | ForEach-Object {
                (Get-ADGroup $_ -ErrorAction SilentlyContinue).Name
            } | Sort-Object
            
            if ($groups -and $groups.Count -gt 0) {
                foreach ($group in $groups) {
                    [void]$sb.AppendLine("- $group")
                }
            } else {
                [void]$sb.AppendLine("No group memberships found")
            }
        } catch {
            [void]$sb.AppendLine("Error retrieving group membership: $($_.Exception.Message)")
        }
        
        # Display the formatted output
        $detailsText.Text = $sb.ToString()
        $statusLabel.Text = "User details loaded"
    }
    catch {
        $detailsText.Text = "Error retrieving user details: $($_.Exception.Message)"
        $statusLabel.Text = "Error retrieving user details"
        Write-Error $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Wire up search button event
$searchButton.Add_Click({
    Search-ADUsers -SearchText $searchBox.Text -SearchBy $searchTypeBox.SelectedItem
})

# Wire up clear button event
$clearButton.Add_Click({
    $searchBox.Text = ""
    $resultsList.Items.Clear()
    $detailsText.Text = ""
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
})

# Show the form
[void]$form.ShowDialog()