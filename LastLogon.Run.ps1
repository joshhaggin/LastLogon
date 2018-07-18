﻿#------------------------------------------------------------------------
# Source File Information (DO NOT MODIFY)
# Source ID: 16b9e890-0f7b-4418-a68b-a5d7ef5faa24
# Source File: \\kiewitplaza\ktg\Active\KSS\kss_team\Apps\Powershell Tools\Josh Dev stuff\GUI\LastLogon\LastLogon.psproj
#------------------------------------------------------------------------
<#
    .NOTES
    --------------------------------------------------------------------------------
     Code generated by:  SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.149
     Generated on:       3/2/2018 9:03 AM
     Generated by:       Josh.Haggin
    --------------------------------------------------------------------------------
    .DESCRIPTION
        Script generated by PowerShell Studio 2018
#>



#region Source: Startup.pss
#----------------------------------------------
#region Import Assemblies
#----------------------------------------------
[void][Reflection.Assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][Reflection.Assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][Reflection.Assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][Reflection.Assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
#endregion Import Assemblies

#Define a Param block to use custom parameters in the project
#Param ($CustomParameter)

function Main {
<#
    .SYNOPSIS
        The Main function starts the project application.
    
    .PARAMETER Commandline
        $Commandline contains the complete argument string passed to the script packager executable.
    
    .NOTES
        Use this function to initialize your script and to call GUI forms.
		
    .NOTES
        To get the console output in the Packager (Forms Engine) use: 
		$ConsoleOutput (Type: System.Collections.ArrayList)
#>
	Param ([String]$Commandline)
		
	#--------------------------------------------------------------------------
	#TODO: Add initialization script here (Load modules and check requirements)
	
	
	#--------------------------------------------------------------------------
	
	if((Show-MainForm_psf) -eq 'OK')
	{
		
	}
	
	$script:ExitCode = 0 #Set the exit code for the Packager
}







#endregion Source: Startup.pss

#region Source: Globals.ps1
	#--------------------------------------------
	# Declare Global Variables and Functions here
	#--------------------------------------------
	
	
	#Sample function that provides the location of the script
	function Get-ScriptDirectory
	{
	<#
		.SYNOPSIS
			Get-ScriptDirectory returns the proper location of the script.
	
		.OUTPUTS
			System.String
		
		.NOTES
			Returns the correct path within a packaged executable.
	#>
		[OutputType([string])]
		param ()
		if ($null -ne $hostinvocation)
		{
			Split-Path $hostinvocation.MyCommand.path
		}
		else
		{
			Split-Path $script:MyInvocation.MyCommand.Path
		}
	}
	
	#Sample variable that provides the location of the script
	[string]$ScriptDirectory = Get-ScriptDirectory
	
	
	#region Control Helper Functions
	function Update-DataGridView
	{
		<#
		.SYNOPSIS
			This functions helps you load items into a DataGridView.
	
		.DESCRIPTION
			Use this function to dynamically load items into the DataGridView control.
	
		.PARAMETER  DataGridView
			The DataGridView control you want to add items to.
	
		.PARAMETER  Item
			The object or objects you wish to load into the DataGridView's items collection.
		
		.PARAMETER  DataMember
			Sets the name of the list or table in the data source for which the DataGridView is displaying data.
	
		.PARAMETER AutoSizeColumns
		    Resizes DataGridView control's columns after loading the items.
		#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory = $true)]
			[System.Windows.Forms.DataGridView]$DataGridView,
			[ValidateNotNull()]
			[Parameter(Mandatory = $true)]
			$Item,
			[Parameter(Mandatory = $false)]
			[string]$DataMember,
			[System.Windows.Forms.DataGridViewAutoSizeColumnMode]$AutoSizeColumns = 'None'
		)
		$DataGridView.SuspendLayout()
		$DataGridView.DataMember = $DataMember
		
		if ($Item -is [System.Data.DataSet] -and $Item.Tables.Count -gt 0)
		{
			$DataGridView.DataSource = $Item.Tables[0]
		}
		elseif ($Item -is [System.ComponentModel.IListSource]`
			-or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView])
		{
			$DataGridView.DataSource = $Item
		}
		else
		{
			$array = New-Object System.Collections.ArrayList
			
			if ($Item -is [System.Collections.IList])
			{
				$array.AddRange($Item)
			}
			else
			{
				$array.Add($Item)
			}
			$DataGridView.DataSource = $array
		}
		
		if ($AutoSizeColumns -ne 'None')
		{
			$DataGridView.AutoResizeColumns($AutoSizeColumns)
		}
		
		$DataGridView.ResumeLayout()
	}
	
	function ConvertTo-DataTable
	{
		<#
			.SYNOPSIS
				Converts objects into a DataTable.
		
			.DESCRIPTION
				Converts objects into a DataTable, which are used for DataBinding.
		
			.PARAMETER  InputObject
				The input to convert into a DataTable.
		
			.PARAMETER  Table
				The DataTable you wish to load the input into.
		
			.PARAMETER RetainColumns
				This switch tells the function to keep the DataTable's existing columns.
			
			.PARAMETER FilterWMIProperties
				This switch removes WMI properties that start with an underline.
		
			.EXAMPLE
				$DataTable = ConvertTo-DataTable -InputObject (Get-Process)
		#>
		[OutputType([System.Data.DataTable])]
		param (
			[ValidateNotNull()]
			$InputObject,
			[ValidateNotNull()]
			[System.Data.DataTable]$Table,
			[switch]$RetainColumns,
			[switch]$FilterWMIProperties)
		
		if ($null -eq $Table)
		{
			$Table = New-Object System.Data.DataTable
		}
		
		if ($InputObject -is [System.Data.DataTable])
		{
			$Table = $InputObject
		}
		elseif ($InputObject -is [System.Data.DataSet] -and $InputObject.Tables.Count -gt 0)
		{
			$Table = $InputObject.Tables[0]
		}
		else
		{
			if (-not $RetainColumns -or $Table.Columns.Count -eq 0)
			{
				#Clear out the Table Contents
				$Table.Clear()
				
				if ($null -eq $InputObject) { return } #Empty Data
				
				$object = $null
				#find the first non null value
				foreach ($item in $InputObject)
				{
					if ($null -ne $item)
					{
						$object = $item
						break
					}
				}
				
				if ($null -eq $object) { return } #All null then empty
				
				#Get all the properties in order to create the columns
				foreach ($prop in $object.PSObject.Get_Properties())
				{
					if (-not $FilterWMIProperties -or -not $prop.Name.StartsWith('__')) #filter out WMI properties
					{
						#Get the type from the Definition string
						$type = $null
						
						if ($null -ne $prop.Value)
						{
							try { $type = $prop.Value.GetType() }
							catch { Out-Null }
						}
						
						if ($null -ne $type) # -and [System.Type]::GetTypeCode($type) -ne 'Object')
						{
							[void]$table.Columns.Add($prop.Name, $type)
						}
						else #Type info not found
						{
							[void]$table.Columns.Add($prop.Name)
						}
					}
				}
				
				if ($object -is [System.Data.DataRow])
				{
					foreach ($item in $InputObject)
					{
						$Table.Rows.Add($item)
					}
					return @( ,$Table)
				}
			}
			else
			{
				$Table.Rows.Clear()
			}
			
			foreach ($item in $InputObject)
			{
				$row = $table.NewRow()
				
				if ($item)
				{
					foreach ($prop in $item.PSObject.Get_Properties())
					{
						if ($table.Columns.Contains($prop.Name))
						{
							$row.Item($prop.Name) = $prop.Value
						}
					}
				}
				[void]$table.Rows.Add($row)
			}
		}
		
		return @( ,$Table)
	}
	#endregion
	
	function Load-DataGridView
	{
		<#
		.SYNOPSIS
			This functions helps you load items into a DataGridView.
	
		.DESCRIPTION
			Use this function to dynamically load items into the DataGridView control.
	
		.PARAMETER  DataGridView
			The DataGridView control you want to add items to.
	
		.PARAMETER  Item
			The object or objects you wish to load into the DataGridView's items collection.
		
		.PARAMETER  DataMember
			Sets the name of the list or table in the data source for which the DataGridView is displaying data.
	
		#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory = $true)]
			[System.Windows.Forms.DataGridView]$DataGridView,
			[ValidateNotNull()]
			[Parameter(Mandatory = $true)]
			$Item,
			[Parameter(Mandatory = $false)]
			[string]$DataMember
		)
		$DataGridView.SuspendLayout()
		$DataGridView.DataMember = $DataMember
		
		if ($Item -is [System.ComponentModel.IListSource]`
			-or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView])
		{
			$DataGridView.DataSource = $Item
		}
		else
		{
			$array = New-Object System.Collections.ArrayList
			
			if ($Item -is [System.Collections.IList])
			{
				$array.AddRange($Item)
			}
			else
			{
				$array.Add($Item)
			}
			$DataGridView.DataSource = $array
		}
		
		$DataGridView.ResumeLayout()
	}
	#endregion
	
	#-----------------------------
	#  function to test if a variable is a number
	#-----------------------------
	
	function Is-Numeric ($Value)
	{
		return $Value -match "^[\d\.]+$"
	}
	
	function Get-LoggedOnUser
	{
		param ([String[]]$ComputerName = $env:COMPUTERNAME)
		
		$ComputerName | ForEach-Object {
			(quser /SERVER:$_) -replace '\s{2,}', ',' |
			ConvertFrom-CSV |
			Add-Member -MemberType NoteProperty -Name ComputerName -Value $_ -PassThru
		}
	}
	
	function search-AD
	{
		param (
			[string]$ADObject
		)
		
		begin
		{
		}
		
		Process
		{
			if (Is-Numeric $ADObject)
			{
				if ($ADObject.length -gt "6")
				{ $ADObject = $ADObject.TrimStart("0") }
				$object = Get-ADUser -filter { (employeeid -like $ADObject) -and (name -notlike "*admin") } -Properties *
			}
			else
			{
				$object = Get-ADobject -filter { (samaccountname -eq $ADObject) -or (displayname -eq $ADObject) -or (mail -eq $ADObject) } -Properties *
				if ($object -eq $null)
				{
					$object = Get-ADObject -LDAPFilter:"(anr=$ADObject)" -Properties * -SearchBase:"DC=KIEWITPLAZA,DC=COM" -SearchScope:"Subtree"
				}
			}
		}
		
		End
		{ Write-Output $object }
	}
	
	
#endregion Source: Globals.ps1

#region Source: MainForm.psf
function Show-MainForm_psf
{
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form_LastLogon = New-Object 'System.Windows.Forms.Form'
	$buttonSearch = New-Object 'System.Windows.Forms.Button'
	$datagridview_results = New-Object 'System.Windows.Forms.DataGridView'
	$labelSamAccountNameEmailD = New-Object 'System.Windows.Forms.Label'
	$textbox_user = New-Object 'System.Windows.Forms.TextBox'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$form_LastLogon_Load={
		#TODO: Initialize Form Controls here
		
	}
	
	$buttonSearch_Click={
		#TODO: Place custom script here
		
		$DC_array = @()
		$DCs = Get-ADDomainController -Filter *
		
		$account = (search-AD $textbox_user.Text).samaccountname
		Write-Host $account
		foreach ($DC in $DCs)
		{
			try
			{
				$user = get-aduser $account -Server $DC -Properties *
				
				#$myObject = $null
				$myObject = New-Object System.Object
				
				$myObject | Add-Member -Type NoteProperty -Name DC -Value $DC.name
				$myObject | Add-Member -Type NoteProperty -Name LastLogOn -Value $([DateTime]::FromFileTime($user.lastlogon))
				$myObject | Add-Member -Type NoteProperty -Name BadPWCount -Value $user.badPwdCount
				$myObject | Add-Member -Type NoteProperty -Name BadPWTime -Value $([DateTime]::FromFileTime($user.badPasswordTime))
				
				$DC_array = $DC_array + $myObject
			}
			catch
			{ }
			
			$AD_DT = ConvertTo-DataTable -InputObject $DC_array
			Update-DataGridView -DataGridView $datagridview_results -Item $AD_DT
		}
		
		$AD_DT = ConvertTo-DataTable -InputObject $DC_array
		Load-DataGridView -DataGridView $datagridview_results -Item $AD_DT
		
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form_LastLogon.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
		$script:MainForm_datagridview_results = $datagridview_results.SelectedCells
		$script:MainForm_textbox_user = $textbox_user.Text
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$buttonSearch.remove_Click($buttonSearch_Click)
			$form_LastLogon.remove_Load($form_LastLogon_Load)
			$form_LastLogon.remove_Load($Form_StateCorrection_Load)
			$form_LastLogon.remove_Closing($Form_StoreValues_Closing)
			$form_LastLogon.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form_LastLogon.SuspendLayout()
	#
	# form_LastLogon
	#
	$form_LastLogon.Controls.Add($buttonSearch)
	$form_LastLogon.Controls.Add($datagridview_results)
	$form_LastLogon.Controls.Add($labelSamAccountNameEmailD)
	$form_LastLogon.Controls.Add($textbox_user)
	$form_LastLogon.AutoScaleDimensions = '6, 13'
	$form_LastLogon.AutoScaleMode = 'Font'
	$form_LastLogon.ClientSize = '418, 331'
	$form_LastLogon.Name = 'form_LastLogon'
	$form_LastLogon.Text = 'Last Logon'
	$form_LastLogon.add_Load($form_LastLogon_Load)
	#
	# buttonSearch
	#
	$buttonSearch.Location = '338, 27'
	$buttonSearch.Name = 'buttonSearch'
	$buttonSearch.Size = '75, 23'
	$buttonSearch.TabIndex = 3
	$buttonSearch.Text = 'Search'
	$buttonSearch.UseCompatibleTextRendering = $True
	$buttonSearch.UseVisualStyleBackColor = $True
	$buttonSearch.add_Click($buttonSearch_Click)
	#
	# datagridview_results
	#
	$datagridview_results.AllowUserToAddRows = $False
	$datagridview_results.AllowUserToDeleteRows = $False
	$datagridview_results.AllowUserToResizeRows = $False
	$datagridview_results.Anchor = 'Top, Bottom, Left, Right'
	$datagridview_results.AutoSizeColumnsMode = 'Fill'
	$datagridview_results.ColumnHeadersHeightSizeMode = 'AutoSize'
	$datagridview_results.Location = '1, 71'
	$datagridview_results.Name = 'datagridview_results'
	$datagridview_results.ReadOnly = $True
	$datagridview_results.RowHeadersWidthSizeMode = 'AutoSizeToAllHeaders'
	$datagridview_results.Size = '414, 258'
	$datagridview_results.TabIndex = 2
	#
	# labelSamAccountNameEmailD
	#
	$labelSamAccountNameEmailD.AutoSize = $True
	$labelSamAccountNameEmailD.Location = '12, 9'
	$labelSamAccountNameEmailD.Name = 'labelSamAccountNameEmailD'
	$labelSamAccountNameEmailD.Size = '269, 17'
	$labelSamAccountNameEmailD.TabIndex = 1
	$labelSamAccountNameEmailD.Text = 'SamAccountName, Email, DisplayName, or PERNR:'
	$labelSamAccountNameEmailD.UseCompatibleTextRendering = $True
	#
	# textbox_user
	#
	$textbox_user.Location = '12, 29'
	$textbox_user.Name = 'textbox_user'
	$textbox_user.Size = '320, 20'
	$textbox_user.TabIndex = 0
	$form_LastLogon.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form_LastLogon.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form_LastLogon.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form_LastLogon.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$form_LastLogon.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $form_LastLogon.ShowDialog()

}
#endregion Source: MainForm.psf

#Start the application
Main ($CommandLine)
