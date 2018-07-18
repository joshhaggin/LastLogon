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

