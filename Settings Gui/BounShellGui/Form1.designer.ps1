[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
$SettingsForm = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.DataGridView]$grid_Tenants = $null
[System.Windows.Forms.Button]$btn_CancelConfig = $null
[System.Windows.Forms.Button]$Btn_ReloadConfig = $null
[System.Windows.Forms.Button]$Btn_SaveConfig = $null
[System.Windows.Forms.Button]$Btn_Default = $null
[System.Windows.Forms.DataGridViewTextBoxColumn]$Tenant_ID = $null
[System.Windows.Forms.DataGridViewTextBoxColumn]$Tenant_DisplayName = $null
[System.Windows.Forms.DataGridViewTextBoxColumn]$Tenant_Email = $null
[System.Windows.Forms.DataGridViewButtonColumn]$Tenant_Credentials = $null
[System.Windows.Forms.DataGridViewCheckBoxColumn]$Tenant_ModernAuth = $null
[System.Windows.Forms.DataGridViewCheckBoxColumn]$Tenant_Teams = $null
[System.Windows.Forms.DataGridViewCheckBoxColumn]$Tenant_Skype = $null
[System.Windows.Forms.DataGridViewCheckBoxColumn]$Tenant_Exchange = $null
[System.Windows.Forms.CheckBox]$cbx_AutoUpdates = $null
function InitializeComponent
{
[System.Windows.Forms.DataGridViewCellStyle]$dataGridViewCellStyle1 = (New-Object -TypeName System.Windows.Forms.DataGridViewCellStyle)
[System.Windows.Forms.DataGridViewCellStyle]$dataGridViewCellStyle2 = (New-Object -TypeName System.Windows.Forms.DataGridViewCellStyle)
[System.Windows.Forms.DataGridViewCellStyle]$dataGridViewCellStyle3 = (New-Object -TypeName System.Windows.Forms.DataGridViewCellStyle)
$btn_CancelConfig = (New-Object -TypeName System.Windows.Forms.Button)
$Btn_ReloadConfig = (New-Object -TypeName System.Windows.Forms.Button)
$Btn_SaveConfig = (New-Object -TypeName System.Windows.Forms.Button)
$cbx_AutoUpdates = (New-Object -TypeName System.Windows.Forms.CheckBox)
$grid_Tenants = (New-Object -TypeName System.Windows.Forms.DataGridView)
$Btn_Default = (New-Object -TypeName System.Windows.Forms.Button)
$Tenant_ID = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
$Tenant_DisplayName = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
$Tenant_Email = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
$Tenant_Credentials = (New-Object -TypeName System.Windows.Forms.DataGridViewButtonColumn)
$Tenant_ModernAuth = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
$Tenant_Teams = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
$Tenant_Skype = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
$Tenant_Exchange = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
([System.ComponentModel.ISupportInitialize]$grid_Tenants).BeginInit()
$SettingsForm.SuspendLayout()
#
#btn_CancelConfig
#
$btn_CancelConfig.BackColor = [System.Drawing.Color]::White
$btn_CancelConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btn_CancelConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$btn_CancelConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

$btn_CancelConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]821,[System.Int32]368))
$btn_CancelConfig.Name = [System.String]'btn_CancelConfig'
$btn_CancelConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]94,[System.Int32]23))
$btn_CancelConfig.TabIndex = [System.Int32]59
$btn_CancelConfig.Text = [System.String]'Cancel'
$btn_CancelConfig.UseVisualStyleBackColor = $false
$btn_CancelConfig.add_Click($btn_CancelConfig_Click)
#
#Btn_ReloadConfig
#
$Btn_ReloadConfig.BackColor = [System.Drawing.Color]::White
$Btn_ReloadConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Btn_ReloadConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Btn_ReloadConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

$Btn_ReloadConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]589,[System.Int32]368))
$Btn_ReloadConfig.Name = [System.String]'Btn_ReloadConfig'
$Btn_ReloadConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110,[System.Int32]23))
$Btn_ReloadConfig.TabIndex = [System.Int32]58
$Btn_ReloadConfig.Text = [System.String]'Reload Config'
$Btn_ReloadConfig.UseVisualStyleBackColor = $true
$Btn_ReloadConfig.add_Click($Btn_ConfigBrowse_Click)
#
#Btn_SaveConfig
#
$Btn_SaveConfig.BackColor = [System.Drawing.Color]::White
$Btn_SaveConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Btn_SaveConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Btn_SaveConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

$Btn_SaveConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]705,[System.Int32]368))
$Btn_SaveConfig.Name = [System.String]'Btn_SaveConfig'
$Btn_SaveConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110,[System.Int32]23))
$Btn_SaveConfig.TabIndex = [System.Int32]57
$Btn_SaveConfig.Text = [System.String]'Save Config'
$Btn_SaveConfig.UseVisualStyleBackColor = $false
$Btn_SaveConfig.add_Click($Btn_SaveConfig_Click)
#
#cbx_AutoUpdates
#
$cbx_AutoUpdates.AutoSize = $true
$cbx_AutoUpdates.Checked = $true
$cbx_AutoUpdates.CheckState = [System.Windows.Forms.CheckState]::Checked
$cbx_AutoUpdates.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

$cbx_AutoUpdates.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]370))
$cbx_AutoUpdates.Name = [System.String]'cbx_AutoUpdates'
$cbx_AutoUpdates.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]183,[System.Int32]17))
$cbx_AutoUpdates.TabIndex = [System.Int32]75
$cbx_AutoUpdates.Text = [System.String]'Automatically Check For Updates'
$cbx_AutoUpdates.UseVisualStyleBackColor = $true
$cbx_AutoUpdates.add_CheckedChanged($cbx_NoIntLCD_CheckedChanged_1)
#
#grid_Tenants
#
$grid_Tenants.AllowUserToAddRows = $false
$grid_Tenants.AllowUserToDeleteRows = $false
$dataGridViewCellStyle1.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$dataGridViewCellStyle1.BackColor = [System.Drawing.SystemColors]::Control
$dataGridViewCellStyle1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$dataGridViewCellStyle1.ForeColor = [System.Drawing.SystemColors]::WindowText
$dataGridViewCellStyle1.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
$dataGridViewCellStyle1.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
$dataGridViewCellStyle1.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
$grid_Tenants.ColumnHeadersDefaultCellStyle = $dataGridViewCellStyle1
$grid_Tenants.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$grid_Tenants.Columns.AddRange($Tenant_ID,$Tenant_DisplayName,$Tenant_Email,$Tenant_Credentials,$Tenant_ModernAuth,$Tenant_Teams,$Tenant_Skype,$Tenant_Exchange)
$dataGridViewCellStyle2.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$dataGridViewCellStyle2.BackColor = [System.Drawing.SystemColors]::Window
$dataGridViewCellStyle2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$dataGridViewCellStyle2.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

$dataGridViewCellStyle2.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
$dataGridViewCellStyle2.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
$dataGridViewCellStyle2.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
$grid_Tenants.DefaultCellStyle = $dataGridViewCellStyle2
$grid_Tenants.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]12))
$grid_Tenants.Name = [System.String]'grid_Tenants'
$dataGridViewCellStyle3.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$dataGridViewCellStyle3.BackColor = [System.Drawing.SystemColors]::Control
$dataGridViewCellStyle3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$dataGridViewCellStyle3.ForeColor = [System.Drawing.SystemColors]::WindowText
$dataGridViewCellStyle3.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
$dataGridViewCellStyle3.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
$dataGridViewCellStyle3.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
$grid_Tenants.RowHeadersDefaultCellStyle = $dataGridViewCellStyle3
$grid_Tenants.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]903,[System.Int32]336))
$grid_Tenants.TabIndex = [System.Int32]76
$grid_Tenants.add_CellContentClick($grid_Tenants_CellContentClick)
#
#Btn_Default
#
$Btn_Default.BackColor = [System.Drawing.Color]::White
$Btn_Default.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Btn_Default.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Btn_Default.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

$Btn_Default.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]473,[System.Int32]368))
$Btn_Default.Name = [System.String]'Btn_Default'
$Btn_Default.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110,[System.Int32]23))
$Btn_Default.TabIndex = [System.Int32]77
$Btn_Default.Text = [System.String]'Reset to Default'
$Btn_Default.UseVisualStyleBackColor = $true
$Btn_Default.add_Click($Btn_Default_Click)
#
#Tenant_ID
#
$Tenant_ID.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$Tenant_ID.Frozen = $true
$Tenant_ID.HeaderText = [System.String]'ID'
$Tenant_ID.Name = [System.String]'Tenant_ID'
$Tenant_ID.ReadOnly = $true
$Tenant_ID.Width = [System.Int32]43
#
#Tenant_DisplayName
#
$Tenant_DisplayName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$Tenant_DisplayName.Frozen = $true
$Tenant_DisplayName.HeaderText = [System.String]'Display Name'
$Tenant_DisplayName.Name = [System.String]'Tenant_DisplayName'
$Tenant_DisplayName.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$Tenant_DisplayName.Width = [System.Int32]78
#
#Tenant_Email
#
$Tenant_Email.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$Tenant_Email.Frozen = $true
$Tenant_Email.HeaderText = [System.String]'Sign In Address'
$Tenant_Email.Name = [System.String]'Tenant_Email'
$Tenant_Email.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$Tenant_Email.Width = [System.Int32]78
#
#Tenant_Credentials
#
$Tenant_Credentials.Frozen = $true
$Tenant_Credentials.HeaderText = [System.String]'Credentials'
$Tenant_Credentials.Name = [System.String]'Tenant_Credentials'
#
#Tenant_ModernAuth
#
$Tenant_ModernAuth.HeaderText = [System.String]'Uses Modern Auth?'
$Tenant_ModernAuth.Name = [System.String]'Tenant_ModernAuth'
#
#Tenant_Teams
#
$Tenant_Teams.HeaderText = [System.String]'Connect to Teams?'
$Tenant_Teams.Name = [System.String]'Tenant_Teams'
#
#Tenant_Skype
#
$Tenant_Skype.HeaderText = [System.String]'Connect to Skype?'
$Tenant_Skype.Name = [System.String]'Tenant_Skype'
#
#Tenant_Exchange
#
$Tenant_Exchange.HeaderText = [System.String]'Connect to Exchange?'
$Tenant_Exchange.Name = [System.String]'Tenant_Exchange'
#
#SettingsForm
#
$SettingsForm.BackColor = [System.Drawing.Color]::White
$SettingsForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]925,[System.Int32]404))
$SettingsForm.Controls.Add($Btn_Default)
$SettingsForm.Controls.Add($grid_Tenants)
$SettingsForm.Controls.Add($cbx_AutoUpdates)
$SettingsForm.Controls.Add($btn_CancelConfig)
$SettingsForm.Controls.Add($Btn_ReloadConfig)
$SettingsForm.Controls.Add($Btn_SaveConfig)
$SettingsForm.Name = [System.String]'SettingsForm'
$SettingsForm.add_Load($SettingsForm_Load)
([System.ComponentModel.ISupportInitialize]$grid_Tenants).EndInit()
$SettingsForm.ResumeLayout($false)
$SettingsForm.PerformLayout()
Add-Member -InputObject $SettingsForm -Name base -Value $base -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name grid_Tenants -Value $grid_Tenants -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name btn_CancelConfig -Value $btn_CancelConfig -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Btn_ReloadConfig -Value $Btn_ReloadConfig -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Btn_SaveConfig -Value $Btn_SaveConfig -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Btn_Default -Value $Btn_Default -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_ID -Value $Tenant_ID -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_DisplayName -Value $Tenant_DisplayName -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_Email -Value $Tenant_Email -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_Credentials -Value $Tenant_Credentials -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_ModernAuth -Value $Tenant_ModernAuth -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_Teams -Value $Tenant_Teams -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_Skype -Value $Tenant_Skype -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name Tenant_Exchange -Value $Tenant_Exchange -MemberType NoteProperty
Add-Member -InputObject $SettingsForm -Name cbx_AutoUpdates -Value $cbx_AutoUpdates -MemberType NoteProperty
}
. InitializeComponent
