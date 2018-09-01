# Office 365 Automation Tool
<#
.SYNOPSIS
        Tool to output user details

.Description
		Use this Tool to get below user details.
		CAS Properties : IMAP Configuration, POP3, ActiveSync, OWA
		Mailbox Properties : Litigation Hold, Legal Hold or In Place Hold, Audit, Mailbox Size and Archive Mailbox Size if enabled
		Mailbox Permissions : Full Access, Send On Behalf, Send As and Calendar Permissions
        Using Tool We can set below properties to end user
		Mailbox Permissions : Add or Remove single permission or for multiple users based on input (Applies for all mailbox permissions available)
		Set Mailbox Size
		

.Notes
        Author          : Rushi Rohit Kotha
        Email Address   : rushirohith.kotha@gmail.com
        Version         : 1.0
        Date Created    : 28/08/2018
        Changes         : 1.0 Script first set up

#>
#Created By : Rushi Rohit Kotha
#Generated Form Function
function GenerateForm{
#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion
#-----------------------------------------------------------------------
#region Generated Form Objects 
$Tool = New-Object System.Windows.Forms.Form
$groupbox1 = New-Object System.Windows.Forms.GroupBox
$groupbox2 = New-Object System.Windows.Forms.GroupBox
$groupbox3 = New-Object System.Windows.Forms.GroupBox
$cbox_MbxPermission = New-Object System.Windows.Forms.ComboBox
$lbl_Connected = New-Object System.Windows.Forms.Label
$lbl_Sync = New-Object System.Windows.Forms.Label
$lbl_RecipientType = New-Object System.Windows.Forms.Label
$lbl_TDI =  New-Object System.Windows.Forms.Label
$labelUserid = New-Object System.Windows.Forms.Label
$txt_Enterid = New-Object System.Windows.Forms.TextBox
$txt_UPN = New-Object System.Windows.Forms.TextBox
$txt_FwdAdd = New-Object System.Windows.Forms.TextBox
$button_Search = New-Object System.Windows.Forms.Button
$lbl_OOF = New-Object System.Windows.Forms.Label
$lbl_LicenseDetails = New-Object System.Windows.Forms.Label
$lbl_Licensed = New-Object System.Windows.Forms.Label
$lbl_Audit = New-Object System.Windows.Forms.Label
$lbl_LHold = New-Object System.Windows.Forms.Label
$lbl_ASync = New-Object System.Windows.Forms.Label
$lbl_Owa = New-Object System.Windows.Forms.Label
$lbl_Pop = New-Object System.Windows.Forms.Label
$lbl_Imap = New-Object System.Windows.Forms.Label
$btn_Search = New-Object System.Windows.Forms.Button
$btn_AddFwdAdd = New-Object System.Windows.Forms.Button
$btn_RemoveFwdAdd = New-Object System.Windows.Forms.Button
$btn_Add = New-Object System.Windows.Forms.Button
$btn_Remove = New-Object System.Windows.Forms.Button
$lbl_Audit_Name = New-Object System.Windows.Forms.Label
$lbl_LHold_Name = New-Object System.Windows.Forms.Label
$lbl_ASync_Name = New-Object System.Windows.Forms.Label
$lbl_Owa_Name = New-Object System.Windows.Forms.Label
$lbl_Pop_Name = New-Object System.Windows.Forms.Label
$lbl_Imap_Name = New-Object System.Windows.Forms.Label
$lbl_Licensed_Name = New-Object System.Windows.Forms.Label
$chkBox_DeliverToMbx = New-Object System.Windows.Forms.CheckBox
$cbox_TDI = New-Object System.Windows.Forms.ComboBox
$lbl_Status = New-Object System.Windows.Forms.Label
$buttonDisable = New-Object System.Windows.Forms.Button
$lbl_Successful = New-Object System.Windows.Forms.Label
$txt_LicenseDetails = New-Object System.Windows.Forms.TextBox
$labelForwardingAddress = New-Object System.Windows.Forms.Label
$labelOutOfOffice = New-Object System.Windows.Forms.Label
$listB_MbxPerm = New-Object System.Windows.Forms.ListBox
$lbl_Msize = New-Object System.Windows.Forms.Label
$txt_Msize = New-Object System.Windows.Forms.TextBox
$btn_Set = New-Object System.Windows.Forms.Button
$btn_OOFenable = New-Object System.Windows.Forms.Button
$txt_OOF = New-Object System.Windows.Forms.TextBox
$lbl_IPHold = New-Object System.Windows.Forms.Label
$lbl_IPHold_Name = New-Object System.Windows.Forms.Label
$btn_CalPerm = New-Object System.Windows.Forms.Button
$listA_CalPerm = New-Object System.Windows.Forms.ListBox
$listB_CalPerm = New-Object System.Windows.Forms.ListBox
$txtA_CalPerm = New-Object System.Windows.Forms.TextBox
$cbox_CalPerm = New-Object System.Windows.Forms.ComboBox
$btn_CalPerm_Add = New-Object System.Windows.Forms.Button
$btn_CalPerm_Remove = New-Object System.Windows.Forms.Button
$groupbox4 = New-Object System.Windows.Forms.GroupBox
$lbl_Ret = New-Object System.Windows.Forms.Label
$cbox_Ret = New-Object System.Windows.Forms.ComboBox
$btn_Ret = New-Object System.Windows.Forms.Button
$btn_Ret_WF = New-Object System.Windows.Forms.Button
$btn_Ret_AF = New-Object System.Windows.Forms.Button


$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion
#------------------------------------------------------------------------
#GeneratedEventScriptBlock
#------------------------------------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$Tool_Load={


    $Livecred = Get-Credential
    $Session = New-PSSession  -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Livecred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    Import-Module MSOnline
    Connect-MsolService -Credential $Livecred
        $lbl_Sync.Text = ""
		$lbl_RecipientType.Text = ""
		$lbl_TDI = ""
		$lbl_Successful.Text = ""
			
		if ((Get-PSSession).state -eq 'opened')
		{
			$lbl_Connected.Text = "Connected to Office 365"
		}
		else
		{
			$lbl_Connected.BackColor = 'Red'
			$lbl_Connected.ForeColor = 'Black'
			$lbl_Connected.Text = "Not Connected to Office 365"
		}
}
	#region Control Helper Functions
	function Update-ListBox
	{
	<#
		.SYNOPSIS
			This functions helps you load items into a ListBox or CheckedListBox.
		
		.DESCRIPTION
			Use this function to dynamically load items into the ListBox control.
		
		.PARAMETER ListBox
			The ListBox control you want to add items to.
		
		.PARAMETER Items
			The object or objects you wish to load into the ListBox's Items collection.
		
		.PARAMETER DisplayMember
			Indicates the property to display for the items in this control.
		
		.PARAMETER Append
			Adds the item(s) to the ListBox without clearing the Items collection.
		
		.EXAMPLE
			Update-ListBox $ListBox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Update-ListBox $listBox1 "Red" -Append
			Update-ListBox $listBox1 "White" -Append
			Update-ListBox $listBox1 "Blue" -Append
		
		.EXAMPLE
			Update-ListBox $listBox1 (Get-Process) "ProcessName"
		
		.NOTES
			Additional information about the function.
	#>
		
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[System.Windows.Forms.ListBox]
			$ListBox,
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			$Items,
			[Parameter(Mandatory = $false)]
			[string]
			$DisplayMember,
			[switch]
			$Append
		)
		
		if (-not $Append)
		{
			$listBox.Items.Clear()
		}
		
		if ($Items -is [System.Windows.Forms.ListBox+ObjectCollection] -or $Items -is [System.Collections.ICollection])
		{
			$listBox.Items.AddRange($Items)
		}
		elseif ($Items -is [System.Collections.IEnumerable])
		{
			$listBox.BeginUpdate()
			foreach ($obj in $Items)
			{
				$listBox.Items.Add($obj)
			}
			$listBox.EndUpdate()
		}
		else
		{
			$listBox.Items.Add($Items)
		}
		
		$listBox.DisplayMember = $DisplayMember
	}
	
	function Update-ListViewColumnSort
	{
	<#
		.SYNOPSIS
			Sort the ListView's item using the specified column.
		
		.DESCRIPTION
			Sort the ListView's item using the specified column.
			This function uses Add-Type to define a class that sort the items.
			The ListView's Tag property is used to keep track of the sorting.
		
		.PARAMETER ListView
			The ListView control to sort.
		
		.PARAMETER ColumnIndex
			The index of the column to use for sorting.
		
		.PARAMETER SortOrder
			The direction to sort the items. If not specified or set to None, it will toggle.
		
		.EXAMPLE
			Update-ListViewColumnSort -ListView $listview1 -ColumnIndex 0
		
		.NOTES
			Additional information about the function.
	#>
		
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[System.Windows.Forms.ListView]
			$ListView,
			[Parameter(Mandatory = $true)]
			[int]
			$ColumnIndex,
			[System.Windows.Forms.SortOrder]
			$SortOrder = 'None'
		)
		
		if (($ListView.Items.Count -eq 0) -or ($ColumnIndex -lt 0) -or ($ColumnIndex -ge $ListView.Columns.Count))
		{
			return;
		}
		
		#region Define ListViewItemComparer
		try
		{
			[ListViewItemComparer] | Out-Null
		}
		catch
		{
			Add-Type -ReferencedAssemblies ('System.Windows.Forms') -TypeDefinition  @" 
	using System;
	using System.Windows.Forms;
	using System.Collections;
	public class ListViewItemComparer : IComparer
	{
	    public int column;
	    public SortOrder sortOrder;
	    public ListViewItemComparer()
	    {
	        column = 0;
			sortOrder = SortOrder.Ascending;
	    }
	    public ListViewItemComparer(int column, SortOrder sort)
	    {
	        this.column = column;
			sortOrder = sort;
	    }
	    public int Compare(object x, object y)
	    {
			if(column >= ((ListViewItem)x).SubItems.Count)
				return  sortOrder == SortOrder.Ascending ? -1 : 1;
		
			if(column >= ((ListViewItem)y).SubItems.Count)
				return sortOrder == SortOrder.Ascending ? 1 : -1;
		
			if(sortOrder == SortOrder.Ascending)
	        	return String.Compare(((ListViewItem)x).SubItems[column].Text, ((ListViewItem)y).SubItems[column].Text);
			else
				return String.Compare(((ListViewItem)y).SubItems[column].Text, ((ListViewItem)x).SubItems[column].Text);
	    }
	}
"@ | Out-Null
		}
		#endregion
		
		if ($ListView.Tag -is [ListViewItemComparer])
		{
			#Toggle the Sort Order
			if ($SortOrder -eq [System.Windows.Forms.SortOrder]::None)
			{
				if ($ListView.Tag.column -eq $ColumnIndex -and $ListView.Tag.sortOrder -eq 'Ascending')
				{
					$ListView.Tag.sortOrder = 'Descending'
				}
				else
				{
					$ListView.Tag.sortOrder = 'Ascending'
				}
			}
			else
			{
				$ListView.Tag.sortOrder = $SortOrder
			}
			
			$ListView.Tag.column = $ColumnIndex
			$ListView.Sort() #Sort the items
		}
		else
		{
			if ($SortOrder -eq [System.Windows.Forms.SortOrder]::None)
			{
				$SortOrder = [System.Windows.Forms.SortOrder]::Ascending
			}
			
			#Set to Tag because for some reason in PowerShell ListViewItemSorter prop returns null
			$ListView.Tag = New-Object ListViewItemComparer ($ColumnIndex, $SortOrder)
			$ListView.ListViewItemSorter = $ListView.Tag #Automatically sorts
		}
	}
	
	function Add-ListViewItem
	{
	<#
		.SYNOPSIS
			Adds the item(s) to the ListView and stores the object in the ListViewItem's Tag property.
	
		.DESCRIPTION
			Adds the item(s) to the ListView and stores the object in the ListViewItem's Tag property.
	
		.PARAMETER ListView
			The ListView control to add the items to.
	
		.PARAMETER Items
			The object or objects you wish to load into the ListView's Items collection.
			
		.PARAMETER  ImageIndex
			The index of a predefined image in the ListView's ImageList.
		
		.PARAMETER  SubItems
			List of strings to add as Subitems.
		
		.PARAMETER Group
			The group to place the item(s) in.
		
		.PARAMETER Clear
			This switch clears the ListView's Items before adding the new item(s).
		
		.EXAMPLE
			Add-ListViewItem -ListView $listview1 -Items "Test" -Group $listview1.Groups[0] -ImageIndex 0 -SubItems "Installed"
	#>
		
		Param( 
		[ValidateNotNull()]
		[Parameter(Mandatory=$true)]
		[System.Windows.Forms.ListView]$ListView,
		[ValidateNotNull()]
		[Parameter(Mandatory=$true)]
		$Items,
		[int]$ImageIndex = -1,
		[string[]]$SubItems,
		$Group,
		[switch]$Clear)
		
		if($Clear)
		{
			$ListView.Items.Clear();
	    }
	    
	    $lvGroup = $null
	    if ($Group -is [System.Windows.Forms.ListViewGroup])
	    {
	        $lvGroup = $Group
	    }
	    elseif ($Group -is [string])
	    {
	        #$lvGroup = $ListView.Group[$Group] # Case sensitive
	        foreach ($groupItem in $ListView.Groups)
	        {
	            if ($groupItem.Name -eq $Group)
	            {
	                $lvGroup = $groupItem
	                break
	            }
	        }
	        
	        if ($null -eq $lvGroup)
	        {
	            $lvGroup = $ListView.Groups.Add($Group, $Group)
	        }
	    }
	    
		if($Items -is [Array])
		{
			$ListView.BeginUpdate()
			foreach ($item in $Items)
			{		
				$listitem  = $ListView.Items.Add($item.ToString(), $ImageIndex)
				#Store the object in the Tag
				$listitem.Tag = $item
				
				if($null -ne $SubItems)
				{
					$listitem.SubItems.AddRange($SubItems)
				}
				
				if($null -ne $lvGroup)
				{
					$listitem.Group = $lvGroup
				}
			}
			$ListView.EndUpdate()
		}
		else
		{
			#Add a new item to the ListView
			$listitem  = $ListView.Items.Add($Items.ToString(), $ImageIndex)
			#Store the object in the Tag
			$listitem.Tag = $Items
			
			if($null -ne $SubItems)
			{
				$listitem.SubItems.AddRange($SubItems)
			}
			
			if($null -ne $lvGroup)
			{
				$listitem.Group = $lvGroup
			}
		}
	}
	
	function Update-ComboBox
	{
	<#
		.SYNOPSIS
			This functions helps you load items into a ComboBox.
		
		.DESCRIPTION
			Use this function to dynamically load items into the ComboBox control.
		
		.PARAMETER ComboBox
			The ComboBox control you want to add items to.
		
		.PARAMETER Items
			The object or objects you wish to load into the ComboBox's Items collection.
		
		.PARAMETER DisplayMember
			Indicates the property to display for the items in this control.
		
		.PARAMETER Append
			Adds the item(s) to the ComboBox without clearing the Items collection.
		
		.EXAMPLE
			Update-ComboBox $combobox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Update-ComboBox $combobox1 "Red" -Append
			Update-ComboBox $combobox1 "White" -Append
			Update-ComboBox $combobox1 "Blue" -Append
		
		.EXAMPLE
			Update-ComboBox $combobox1 (Get-Process) "ProcessName"
		
		.NOTES
			Additional information about the function.
	#>
		
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[System.Windows.Forms.ComboBox]
			$ComboBox,
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			$Items,
			[Parameter(Mandatory = $false)]
			[string]
			$DisplayMember,
			[switch]
			$Append
		)
		
		if (-not $Append)
		{
			$ComboBox.Items.Clear()
		}
		
		if ($Items -is [Object[]])
		{
			$ComboBox.Items.AddRange($Items)
		}
		elseif ($Items -is [System.Collections.IEnumerable])
		{
			$ComboBox.BeginUpdate()
			foreach ($obj in $Items)
			{
				$ComboBox.Items.Add($obj)
			}
			$ComboBox.EndUpdate()
		}
		else
		{
			$ComboBox.Items.Add($Items)
		}
		
		$ComboBox.DisplayMember = $DisplayMember
	}
	#endregion

	$listview1_SelectedIndexChanged={
		#TODO: Place custom script here
		
	}
	
	$checkboxOffice365EnterpriseE_CheckedChanged={
		#TODO: Place custom script here
		
	}

	<#
	$global:Mailbox = Get-Mailbox -$txt_UPN.Text
	$global:MsolUser = Get-MsolUser -UserPrincipalName $txt_UPN.Text
	$global:Recipient = Get-Recipient $txt_UPN.Text
	#>
	
	function User_Details()
	{
		$lbl_Imap.Text = $CASData.ImapEnabled
		$lbl_Pop.Text = $CASData.PopEnabled
		$lbl_Owa.Text = $CASData.OwaEnabled
		$lbl_ASync.Text = $CASData.ActiveSyncEnabled
		$lbl_LHold.Text = $MbxData.LitigationHoldEnabled
		$lbl_Audit.Text = $MbxData.AuditEnabled
		$txt_LicenseDetails.Text = $MsolData.licenses.accountsku.skupartnumber
		$lbl_Licensed.Text = $MsolData.IsLicensed
		$txt_FwdAdd.Text = $MbxData.ForwardingAddress
        $listA_CalPerm.Text = $CalPerm.User
        $listB_CalPerm.Text = $CalPerm.AccessRights
        
        if (![string]::IsNullOrEmpty($MbxData.InPlaceHolds))
        { $lbl_IPHold.Text = $True }
        else { $lbl_IPHold.Text = $False}

		if ($mbxdata.DeliverToMailboxAndForward -eq $true)
		{ $chkBox_DeliverToMbx.Checked = $true }
		else{ $chkBox_DeliverToMbx.Checked = $false}
    }

    function StampPolicyOnFolder($MailboxName)
    {
    Write-host "Stamping Policy on folder for Mailbox Name:" $MailboxName -foregroundcolor  $info
    Add-Content $LogFile ("Stamping Policy on folder for Mailbox Name:" + $MailboxName)

    #Change the user to Impersonate
    $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxName);

    #Search for the folder you want to stamp the property on
    $oFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $oSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName)

    #Uncomment the line below if the folder is in the regular mailbox
    $oFindFolderResults = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$oSearchFilter,$oFolderView)

    #Comment the line below and uncomment the line above if the folder is in the regular mailbox
    #$oFindFolderResults = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$oSearchFilter,$oFolderView)

    if ($oFindFolderResults.TotalCount -eq 0)
    {
         Write-host "Folder does not exist in Mailbox:" $MailboxName -foregroundcolor  $warning
         Add-Content $LogFile ("Folder does not exist in Mailbox:" + $MailboxName)
    }
    else
    {
        Write-host "Folder found in Mailbox:" $MailboxName -foregroundcolor  $info

        #PR_ARCHIVE_TAG 0x3018 – We use the PR_ARCHIVE_TAG instead of the PR_POLICY_TAG
        $PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3018,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);

        #PR_RETENTION_FLAGS 0x301D    
        $RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        
        #PR_ARCHIVE_PERIOD 0x301E - We use the PR_ARCHIVE_PERIOD instead of the PR_RETENTION_PERIOD
        $RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301E,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

        #Bind to the folder found
        $oFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$oFindFolderResults.Folders[0].Id)
       
        #Same as the value in the PR_RETENTION_FLAGS property
        $oFolder.SetExtendedProperty($RetentionFlags, 16)

        #Same as that on the policy - Since this tag is disabled the Period would be 90
        $oFolder.SetExtendedProperty($RetentionPeriod, 60)

        #Change the GUID based on your policy tag
        $PolicyTagGUID = new-Object Guid("{cb1b9cca-4b1c-4dd7-af32-d9c1f6138c94}");

        $oFolder.SetExtendedProperty($PolicyTag, $PolicyTagGUID.ToByteArray())

        $oFolder.Update()

        Write-host "Retention policy stamped!" -foregroundcolor $info
        Add-Content $LogFile ("Retention policy stamped!")
    
    }    

    $service.ImpersonatedUserId = $null
}

    $cbox_MbxPermission_SelectedIndexChanged = {
		#TODO: Place custom script here
		$lbl_Successful.Text = ""
		$txt_Enterid.Text = ""
		$MbxData = Get-Mailbox $txt_UPN.Text
		$RecipientData = Get-RecipientPermission $txt_UPN.Text
        
		if ($cbox_MbxPermission.SelectedItem -eq 'Full Access')
		{
			$listB_MbxPerm.Items.Clear()
			$FAccess = ((Get-MailboxPermission $txt_UPN.Text) | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like $MbxData.UserPrincipalName) }).user
			foreach ($item in $FAccess) { $listB_MbxPerm.Items.Add($item) }
		}
		
		if ($cbox_MbxPermission.SelectedItem -eq 'Send On Behalf')
		{
			$listB_MbxPerm.Items.Clear()
			$SOB = $MbxData.GrantSendOnBehalfTo
			foreach ($item in $SOB) { $listB_MbxPerm.Items.Add($item) }
		}
		
		if ($cbox_MbxPermission.SelectedItem -eq 'Send As')
		{
			$listB_MbxPerm.Items.Clear()
			$SendAs = ($RecipientData | Where {$_.AccessRights -like "SendAs" -and $_.Trustee -notlike "NT AUTHORITY\SELF"}).trustee
			foreach ($item in $SendAs) { $listB_MbxPerm.Items.Add($item) }
		}
		if ([String]::IsNullOrEmpty($txt_Enterid))
        {
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        Import-Csv $OpenFileDialog.FileName | ForEach {
        Set-Mailbox
        }
        }
	}

	$btn_Search_Click={
		#TODO: Place custom script here
		$txt_LicenseDetails.Text = "" 
		$lbl_Sync.Text = ""
		$lbl_RecipientType.Text = ""
		$lbl_TDI.text = ""
		$lbl_Status.Text = ""
		$lbl_Successful.Text = ""
		$lbl_IMAP.Text = ""
		$lbl_POP.Text = ""
		$lbl_OWA.Text = ""
		$lbl_ASync.Text = ""
		$lbl_LHold.Text = ""
		$lbl_Audit.Text = ""
		$lbl_Licensed.Text = ""
		$lbl_OOF.Text = ""
		$txt_FwdAdd.Text = ""
        $lbl_IPHold.Text = ""
		
		$listB_MbxPerm.Items.Clear()
        $listA_CalPerm.Items.Clear()
        $listB_CalPerm.Items.Clear()
        $txt_Enterid.Clear()
        $txtA_CalPerm.Clear()
        $txt_Msize.Clear()

		$RecipientData = Get-Mailbox $txt_UPN.Text
		if (![string]::IsNullorEmpty($RecipientData))
		{
			$lbl_Sync.ForeColor = 'Black'
			$lbl_Sync.Text = "Object Exists in Office 365"
			
			if ($RecipientData.RecipientTypeDetails -eq "usermailbox" -or $RecipientData.RecipientTypeDetails -eq "sharedmailbox" -or $RecipientData.RecipientTypeDetails -eq "roommailbox" -or $RecipientData.RecipientTypeDetails -eq "euipmentmailbox")
			{
				$CASData = Get-CasMailbox $txt_UPN.Text
				$MbxData = Get-Mailbox $txt_UPN.Text
				$MsolData = Get-MsolUser -UserPrincipalName $RecipientData.UserPrincipalName
				$MbxAutoConfigData = Get-MailboxAutoReplyConfiguration $txt_UPN.Text
				
				$lbl_RecipientType.Text = $RecipientData.RecipientTypeDetails
				$lbl_OOF.Text = $MbxAutoConfigData.AutoReplyState
				User_Details
				
			}
			
			else
			{
				$lbl_RecipientType.Text = $RecipientData.RecipientTypeDetails
			}
		}
		
		else
		{
			$lbl_Sync.ForeColor = 'Red'
			$lbl_Sync.Text = "User does not exist"
		}
	}
	
	$groupbox1_Enter={
		#TODO: Place custom script here
		
	}

	$cbox_TDI_SelectedIndexChanged={
		#TODO: Place custom script here
		
		$MbxStatistics_Data = Get-MailboxStatistics $txt_UPN.Text
		$MbxData = Get-Mailbox $txt_UPN.Text
		
		$lbl_TDI.Text = ""
		$lbl_Successful.Text = ""
		if ($cbox_TDI.SelectedItem -eq 'mailbox')
		{
			$lbl_TDI.Text = $MbxStatistics_Data.totalitemsize
		}
		
		if ($cbox_TDI.SelectedItem -eq 'archive')
		{
			if ($MbxData.ArchiveStatus -eq 'Active')
			{
				$lbl_TDI.Text = (get-mailboxstatistics $txt_UPN.Text -archive).totalitemsize
			}
			else
			{
				$lbl_TDI.Text = "Archive Status : Disabled"
			}
		}
		
	}

	$buttonDisable_Click={
		#TODO: Place custom script here
		if ($lbl_OOF.Text -eq 'Disabled')
		{
			$lbl_Successful.Text = "Out of Office Already Disabled"
		}
		else
		{
			Set-MailboxAutoReplyConfiguration $txt_UPN.Text -AutoReplyState Disabled
			$lbl_Successful.Text = "Out of Office Successfully Disabled"
			$lbl_OOF.Text = (Get-MailboxAutoReplyConfiguration $txt_UPN.Text).AutoReplyState
		}
	}

	$btn_Add_Click={
		#TODO: Place custom script here
		$lbl_Successful.Text = ""
		
			if ($cbox_MbxPermission.SelectedItem -eq 'Full Access')
			{
				Add-MailboxPermission $txt_UPN.Text -User $txt_Enterid.Text -AccessRights FullAccess
				$lbl_Successful.Text = "FullAccess provided to $($txt_Enterid.Text)"
				$txt_Enterid.Text = ""
			}
			
			if ($cbox_MbxPermission.SelectedItem -eq 'Send On Behalf')
			{
				Set-Mailbox $txt_UPN.Text -GrantSendOnBehalfTo @{ add = $txt_Enterid.text }
				$lbl_Successful.Text = "SenOnBehalf Access provided to $($txt_Enterid.Text)"
				$txt_Enterid.Text = ""
			}
			
			if ($cbox_MbxPermission.SelectedItem -eq 'Send As')
			{
				Add-RecipientPermission $txt_UPN.Text -AccessRights SendAs -Trustee $txt_Enterid.Text -Confirm:$false
				$lbl_Successful.Text = "SendAs Access provided to $($txt_Enterid.Text)"
				$txt_Enterid.Text = ""
			}
		    if ([String]::IsNullOrEmpty($txt_Enterid.Text))
            {
                $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $OpenFileDialog.initialDirectory = $initialDirectory
                $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
                $OpenFileDialog.ShowDialog() | Out-Null
                Import-Csv $OpenFileDialog.FileName | ForEach {
                        if ($_.Operation -eq "F")
                        {
                        Add-MailboxPermission $txt_UPN.Text -User $_.Name -AccessRights FullAccess
                        $lbl_Successful.Text = "FullAccess provided to $($_.Name)"
                        }
                        if ($_.Operation -eq "A")
                        {
                        Add-RecipientPermission $txt_UPN.Text -Trustee $_.Name -AccessRights SendAs -Confirm:$False
                        $lbl_Successful.Text = "SendAs Access provided to $($_.Name)"
                        }
                        if ($_.Operation -eq "B")
                        {
                        Set-Mailbox $txt_UPN.Text -GrantSendOnBehalfTo @{ add = $_.Name}
                        $lbl_Successful.Text = "SenOnBehalf Access provided to $($_.Name)"
                        }
                    }
            }
            if ([string]::IsNullOrEmpty($txt_UPN.Text) -and [string]::IsNullOrEmpty($txt_Enterid.Text))
            {
                $OpenFileDialog3 = New-Object System.Windows.Forms.OpenFileDialog
                $OpenFileDialog3.initialDirectory = $initialDirectory
                $OpenFileDialog3.filter = "CSV (*.csv)| *.csv"
                $OpenFileDialog3.ShowDialog() | Out-Null
                Import-Csv $OpenFileDialog3.FileName | ForEach {
                        if ($_.Operation -eq "F")
                        {
                        Add-MailboxPermission $_.User -User $_.Name -AccessRights FullAccess
                        $lbl_Successful.Text = "FullAccess provided to $($_.Name)"
                        }
                        if ($_.Operation -eq "A")
                        {
                        Add-RecipientPermission $_.User -Trustee $_.Name -AccessRights SendAs -Confirm:$False
                        $lbl_Successful.Text = "SendAs Access provided to $($_.Name)"
                        }
                        if ($_.Operation -eq "B")
                        {
                        Set-Mailbox $_.User -GrantSendOnBehalfTo @{ add = $_.Name}
                        $lbl_Successful.Text = "SenOnBehalf Access provided to $($_.Name)"
                        }
                    }
            }
	}
	
	$btn_Remove_Click={
		#TODO: Place custom script here
		$lbl_Successful.Text = ""

			if ($cbox_MbxPermission.SelectedItem -eq 'Full Access')
			{
				Remove-MailboxPermission $txt_UPN.Text -User $txt_Enterid.text -AccessRights FullAccess -Confirm:$false
				$lbl_Successful.Text = "Full Access Removed for $($txt_Enterid.Text)"
			}
			
			if ($cbox_MbxPermission.SelectedItem -eq 'Send On Behalf')
			{
				Set-Mailbox $txt_UPN.Text -GrantSendOnBehalfTo @{ Remove = $txt_Enterid.text }
				$lbl_Successful.Text = "SenOnBehalf Removed for $($txt_Enterid.Text)"
			}
			
			if ($cbox_MbxPermission.SelectedItem -eq 'Send As')
			{
				Remove-RecipientPermission $txt_UPN.Text -AccessRights SendAs -Trustee $txt_Enterid.Text -Confirm:$false
				$lbl_Successful.Text = "SendAs Access Removed for $($txt_Enterid.Text)"
			}
            
            if ([string]::IsNullOrEmpty($txt_Enterid.Text))
            {
                $OpenFileDialog2 = New-Object System.Windows.Forms.OpenFileDialog
                $OpenFileDialog2.initialDirectory = $initialDirectory
                $OpenFileDialog2.filter = "CSV (*.csv)| *.csv"
                $OpenFileDialog2.ShowDialog() | Out-Null
                Import-Csv $OpenFileDialog2.FileName | ForEach {
                        if ($_.Operation -eq "F")
                        {
                        Remove-MailboxPermission $txt_UPN.Text -User $_.Name -AccessRights FullAccess -Confirm:$False
                        $lbl_Successful.Text = "FullAccess provided to $($_.Name)"
                        }
                        if ($_.Operation -eq "A")
                        {
                        Remove-RecipientPermission $txt_UPN.Text -Trustee $_.Name -AccessRights SendAs -Confirm:$False
                        $lbl_Successful.Text = "SendAs Access provided to $($_.Name)"
                        }
                        if ($_.Operation -eq "B")
                        {
                        Set-Mailbox $txt_UPN.Text -GrantSendOnBehalfTo @{ Remove = $_.Name}
                        $lbl_Successful.Text = "SenOnBehalf Access provided to $($_.Name)"
                        }
                    }
            }

		
	}

	$btn_AddFwdAdd_Click={
		#TODO: Place custom script here
		if ($txt_FwdAdd.Text -eq "")
		{
			$txt_FwdAdd.Text = "Enter ForwardingAddress"
			Start-Sleep -Seconds 3
			$txt_FwdAdd.Text = ""
			
		}
		else
		{
			Set-Mailbox $txt_UPN.Text -ForwardingAddress $txt_FwdAdd.Text -Confirm:$false
			$lbl_Successful.Text = "ForwardingAddress Added"
			Start-Sleep -Seconds 3
			$txt_FwdAdd.Text = ""
		}
	}
	
    $btn_RemoveFwdAdd_Click={
		#TODO: Place custom script here
		if ($txt_FwdAdd.Text -eq "")
		{
			$txt_FwdAdd.Text = "No ForwardingAddress Found"
			Start-Sleep -Seconds 3
			$txt_FwdAdd.Text = ""
			
		}
		else
		{
			Set-Mailbox $txt_UPN.Text -ForwardingAddress $null -Confirm:$false
			$txt_FwdAdd.Text = "ForwardingAddress Removed"
			Start-Sleep -Seconds 3
			$txt_FwdAdd.Text = ""
		}
	}

	$chkBox_DeliverToMbx_CheckedChanged={
		
		if ($chkBox_DeliverToMbx.Checked -eq $true)
		{
			$lbl_Successful.Text = ""
			Set-Mailbox $txt_UPN.Text -DeliverToMailboxAndForward $true
			$lbl_Status.Text = "Enabled"
		}
		
		if ($chkBox_DeliverToMbx.Checked -eq $false)
		{
			$lbl_Successful.Text = ""
			Set-Mailbox $txt_UPN.Text -DeliverToMailboxAndForward $false
			$lbl_Status.Text = "Disabled"
		}
		
        }

    $btn_Set_Click={
        if (![string]::IsNullOrEmpty($txt_Msize.Text)){
        Set-Mailbox $txt_UPN.Text -ProhibitSendQuota ($txt_Msize.Text)GB -ProhibitSendReceiveQuota ($txt_Msize.Text+0.5)GB -IssueWarningQuota ($txt_Msize-0.5)GB
        $lbl_Successful.Text = "Mailbox Size has been set to $($txt_Msize.Text)"
        }
        else
        {
        $lbl_Successful.Text = "Please enter Mailbox Size value to set."
        }
        }

    $btn_OOFenable_Click={
        Set-MailboxAutoReplyConfiguration -Identity $txt_UPN.Text -InternalMessage $txt_OOF.Text -ExternalMessage $txt_OOF
        $lbl_Status.Text = "Out of Office has been enabled for the user $($txt_UPN.Text)"
        }

    $listB_MbxPerm_SelectedIndexChanged={
		#TODO: Place custom script here
		if (![String]::IsNullOrEmpty($listB_MbxPerm.SelectedIndex))
            {
            $txt_Enterid.Text = $listB_MbxPerm.SelectedIndex
            }
        else
            {
            $txt_Enterid.Text = ""
            }
	}

    $btn_CalPerm_Click={
        $listA_CalPerm.Items.Clear()
        $listB_CalPerm.Items.Clear()
        $CalPerm = Get-MailboxFolderPermission -Identity "$($txt_UPN.Text):\Calendar" | where {$_.User -notlike "Anonymous" -and $_.User -notlike "Default"}
        $User = $CalPerm.User
        $AccessRights = $CalPerm.AccessRights
        foreach ($item in $User) {$listA_CalPerm.Items.Add($item)}
        foreach ($item in $AccessRights) {$listB_CalPerm.Items.Add($item)}
    }

    $btn_CalPerm_Add_Click={
        Add-MailboxFolderPermission -Identity "$($txt_UPN.Text):\Calendar" -User $txtA_CalPerm.Text -AccessRights $cbox_CalPerm.SelectedItem
        $lbl_Successful.Text = "Calendar Permission provided to $($txtA_CalPerm.Text)"
    }

    $btn_CalPerm_Remove_Click={
        Remove-MailboxFolderPermission -Identity "$($txt_UPN.Text):\Calendar" -User $txtA_CalPerm.Text -AccessRights $cbox_CalPerm.SelectedItem -Confirm:$False
        $lbl_Successful.Text = "Calendar Permission removed to $($txtA_CalPerm.Text)"
        }

    $listB_MbxPerm_SelectedItemchanged = {
    $txt_Enterid.Text = $listB_MbxPerm.SelectedItem
    }

    $btn_Ret_Click={
        if ($cbox_Ret.SelectedItem -eq "GRC")
        {
        Set-Mailbox $txt_UPN.Text -RetentionPolicy "GRC Policy"
        $lbl_Successful.Text = "Retention Policy has been set"
        }
        if ([string]::IsNullOrEmpty($txt_UPN.Text))
        {
        $OpenFileDialog4 = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog4.initialDirectory = $initialDirectory
        $OpenFileDialog4.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog4.ShowDialog() | Out-Null
        Import-Csv $OpenFileDialog4.FileName | ForEach {
        Set-Mailbox -Identity $_.Name -RetentionPolicy "GRC Policy"
            }
        }
    }

    $btn_Ret_AF_Click={

    if (![string]::IsNullOrEmpty($txt_UPN.Text)) {
    $FolderName = "Archive"

    Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1 
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchVer)  
 
    # Set the Credentials 
    $credential = Get-Credential
    $service.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $credential.UserName, $credential.GetNetworkCredential().Password
 
    # Change the URL to point to your cas server 
    $service.Url= new-object Uri("https://outlook.office365.com/ews/exchange.asmx") 
 
    # Set $UseAutoDiscover to $true if you want to use AutoDiscover else it will use the URL set above 
    $UseAutoDiscover = $false 
    #$a = get-mailbox "900016232" 
 
    if ($UseAutoDiscover -eq $true) 
    {
        Write-host "Autodiscovering.." -foregroundcolor $info
        $UseAutoDiscover = $false
        $service.AutodiscoverUrl($RecipientData.WindowsEmailAddress)
        Write-host "Autodiscovering Done!" -foregroundcolor $info
        Write-host "EWS URL set to :" $service.Url -foregroundcolor $info

    }
    #To catch the Exceptions generated
    trap [System.Exception] 
    {
        Write-host ("Error: " + $_.Exception.Message) -foregroundcolor $error;
        Add-Content $LogFile ("Error: " + $_.Exception.Message);
        continue;
    } 
    StampPolicyOnFolder($txt_UPN.Text)
    }

    Else
    {
    $OpenFileDialog5 = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog5.initialDirectory = $initialDirectory
    $OpenFileDialog5.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog5.ShowDialog() | Out-Null

    $FolderName = "Archive"

    Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1 
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchVer)  
 
    # Set the Credentials 
    $credential = Get-Credential
    $service.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $credential.UserName, $credential.GetNetworkCredential().Password
 
    # Change the URL to point to your cas server 
    $service.Url= new-object Uri("https://outlook.office365.com/ews/exchange.asmx") 
 
    # Set $UseAutoDiscover to $true if you want to use AutoDiscover else it will use the URL set above 
    $UseAutoDiscover = $false 
    #$a = get-mailbox "900016232" 
 
    Import-Csv $OpenFileDialog5.FileName | ForEach {
        $WindowsEmailAddress = (Get-Mailbox -Identity $_.Name).WindowsEmailAddress
        if ($UseAutoDiscover -eq $true) 
    {
        Write-host "Autodiscovering.." -foregroundcolor $info
        $UseAutoDiscover = $false
        $service.AutodiscoverUrl($WindowsEmailAddress)
        Write-host "Autodiscovering Done!" -foregroundcolor $info
        Write-host "EWS URL set to :" $service.Url -foregroundcolor $info

    }
    #To catch the Exceptions generated
    trap [System.Exception] 
    {
        Write-host ("Error: " + $_.Exception.Message) -foregroundcolor $error;
        Add-Content $LogFile ("Error: " + $_.Exception.Message);
        continue;
    } 
    StampPolicyOnFolder($WindowsEmailAddress)
            }
        }
    }

    $btn_Ret_WF_Click={

    if (![string]::IsNullOrEmpty($txt_UPN.Text)) {
    $FolderName = "Working Folder"

    Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1 
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchVer)  
 
    # Set the Credentials 
    $credential = Get-Credential
    $service.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $credential.UserName, $credential.GetNetworkCredential().Password
 
    # Change the URL to point to your cas server 
    $service.Url= new-object Uri("https://outlook.office365.com/ews/exchange.asmx") 
 
    # Set $UseAutoDiscover to $true if you want to use AutoDiscover else it will use the URL set above 
    $UseAutoDiscover = $false 
    #$a = get-mailbox "900016232" 
 
    if ($UseAutoDiscover -eq $true) 
    {
        Write-host "Autodiscovering.." -foregroundcolor $info
        $UseAutoDiscover = $false
        $service.AutodiscoverUrl($RecipientData.WindowsEmailAddress)
        Write-host "Autodiscovering Done!" -foregroundcolor $info
        Write-host "EWS URL set to :" $service.Url -foregroundcolor $info

    }
    #To catch the Exceptions generated
    trap [System.Exception] 
    {
        Write-host ("Error: " + $_.Exception.Message) -foregroundcolor $error;
        Add-Content $LogFile ("Error: " + $_.Exception.Message);
        continue;
    } 
    StampPolicyOnFolder($txt_UPN.Text)
    }

    Else {

    $OpenFileDialog6 = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog6.initialDirectory = $initialDirectory
    $OpenFileDialog6.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog6.ShowDialog() | Out-Null
    $FolderName = "Archive"

    Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1 
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchVer)  
 
    # Set the Credentials 
    $credential = Get-Credential
    $service.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $credential.UserName, $credential.GetNetworkCredential().Password
 
    # Change the URL to point to your cas server 
    $service.Url= new-object Uri("https://outlook.office365.com/ews/exchange.asmx") 
 
    # Set $UseAutoDiscover to $true if you want to use AutoDiscover else it will use the URL set above 
    $UseAutoDiscover = $false 
    #$a = get-mailbox "900016232" 
 
    Import-Csv $OpenFileDialog6.FileName | ForEach {
        $WindowsEmailAddress = (Get-Mailbox -Identity $_.Name).WindowsEmailAddress
        if ($UseAutoDiscover -eq $true) 
    {
        Write-host "Autodiscovering.." -foregroundcolor $info
        $UseAutoDiscover = $false
        $service.AutodiscoverUrl($WindowsEmailAddress)
        Write-host "Autodiscovering Done!" -foregroundcolor $info
        Write-host "EWS URL set to :" $service.Url -foregroundcolor $info

    }
    #To catch the Exceptions generated
    trap [System.Exception] 
    {
        Write-host ("Error: " + $_.Exception.Message) -foregroundcolor $error;
        Add-Content $LogFile ("Error: " + $_.Exception.Message);
        continue;
    } 
    StampPolicyOnFolder($WindowsEmailAddress)
            }
        }
    }

#----------------------------------------------------------------------------------------
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form_Tool.WindowState = $InitialFormWindowState
	}

	$Form_StoreValues_Closing=
	{
		#Store the control values
		$script:Office_365_Multipurpose_Tool_txt_Enterid = $txt_Enterid.Text
		$script:Office_365_Multipurpose_Tool_cbox_TDI = $cbox_TDI.Text
		$script:Office_365_Multipurpose_Tool_cbox_TDI_SelectedItem = $cbox_TDI.SelectedItem
		$script:Office_365_Multipurpose_Tool_cbox_SelectLicense = $cbox_SelectLicense.Text
		$script:Office_365_Multipurpose_Tool_cbox_SelectLicense_SelectedItem = $cbox_SelectLicense.SelectedItem
		$script:Office_365_Multipurpose_Tool_listB_MbxPerm = $listB_MbxPerm.SelectedItems
		$script:Office_365_Multipurpose_Tool_cbox_MbxPermission = $cbox_MbxPermission.Text
		$script:Office_365_Multipurpose_Tool_cbox_MbxPermission_SelectedItem = $cbox_MbxPermission.SelectedItem
		$script:Office_365_Multipurpose_Tool_txt_FwdAdd = $txt_FwdAdd.Text
		$script:Office_365_Multipurpose_Tool_txt_LicenseDetails = $txt_LicenseDetails.Text
		$script:Office_365_Multipurpose_Tool_txt_UPN = $txt_UPN.Text
		$script:Office_365_Multipurpose_Tool_chkBox_DeliverToMbx = $chkBox_DeliverToMbx.Checked
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$cbox_TDI.remove_SelectedIndexChanged($cbox_TDI_SelectedIndexChanged)
			$btn_AddFwdAdd.remove_Click($btn_AddFwdAdd_Click)
			$btn_RemoveFwdAdd.remove_Click($btn_RemoveFwdAdd_Click)
			$listB_MbxPerm.remove_SelectedIndexChanged($listB_MbxPerm_SelectedIndexChanged)
			$cbox_MbxPermission.remove_SelectedIndexChanged($cbox_MbxPermission_SelectedIndexChanged)
			$buttonDisable.remove_Click($buttonDisable_Click)
			$btn_Remove.remove_Click($btn_Remove_Click)
			$btn_Add.remove_Click($btn_Add_Click)
			$btn_Search.remove_Click($btn_Search_Click)
			$chkBox_DeliverToMbx.remove_CheckedChanged($chkBox_DeliverToMbx_CheckedChanged)
			$groupbox1.remove_Enter($groupbox1_Enter)
			$Tool.remove_Load($Tool_Load)
			$Tool.remove_Load($Form_StateCorrection_Load)
			$Tool.remove_Closing($Form_StoreValues_Closing)
			$Tool.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}

#------------------------------------------------------------------------
#region Generated Form Code
	$Tool.SuspendLayout()
	$groupbox1.SuspendLayout()
	$groupbox2.SuspendLayout()
	$groupbox3.SuspendLayout()
    $groupbox4.SuspendLayout()
	#

	# Tool
	#
	$Tool.Controls.Add($groupbox1)
	$Tool.AutoScaleDimensions = '6, 13'
	$Tool.AutoScaleMode = 'Font'
	$Tool.AutoSize = $True
	$Tool.AutoSizeMode = 'GrowAndShrink'
	$Tool.BackColor = 'Control'
	$Tool.ClientSize = '1037, 563'
	$Tool.FormBorderStyle = 'FixedSingle'
    $string = "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"
    $imageBytes = [Convert]::FromBase64String($string)
    $ims = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
    $alkIcon = [System.Drawing.Image]::FromStream($ims, $True)
    $Tool.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -ArgumentList $ims).GetHicon())

	
    $Tool.Name = 'Office365AutomationTool'
	$Tool.Text = 'Office 365 Automation Tool'
	$Tool.add_Load($Tool_Load)
	#
	# groupbox1
	#
	$groupbox1.Controls.Add($txt_Enterid)
	$groupbox1.Controls.Add($txt_FwdSmtpAdd)
	$groupbox1.Controls.Add($lbl_Connected)
	$groupbox1.Controls.Add($groupbox2)
	$groupbox1.Controls.Add($btn_AddFwdAdd)
	$groupbox1.Controls.Add($btn_RemoveFwdAdd)
	$groupbox1.Controls.Add($buttonOffice365Licenses)
	$groupbox1.Controls.Add($cbox_SelectLicense)
	$groupbox1.Controls.Add($lbl_OOF)
	$groupbox1.Controls.Add($labelOutOfOffice)
	$groupbox1.Controls.Add($listB_MbxPerm)
	$groupbox1.Controls.Add($cbox_MbxPermission)
	$groupbox1.Controls.Add($txt_FwdAdd)
	$groupbox1.Controls.Add($labelForwardingSmtpAddres)
	$groupbox1.Controls.Add($labelConvertTo)
	$groupbox1.Controls.Add($txt_LicenseDetails)
	$groupbox1.Controls.Add($lbl_LicenseDetails)
	$groupbox1.Controls.Add($lbl_Licensed)
	#$groupbox1.Controls.Add($buttonDisable)
	$groupbox1.Controls.Add($labelForwardingAddress)
	$groupbox1.Controls.Add($lbl_Audit_Name)
	$groupbox1.Controls.Add($lbl_Audit)
	$groupbox1.Controls.Add($btn_Remove)
	$groupbox1.Controls.Add($btn_Add)
    $groupbox1.Controls.Add($lbl_LHold_Name)
    $groupbox1.Controls.Add($lbl_Licensed_Name)
	$groupbox1.Controls.Add($lbl_ASync_Name)
	$groupbox1.Controls.Add($lbl_Pop_Name)
	$groupbox1.Controls.Add($lbl_Imap_Name)
	$groupbox1.Controls.Add($lbl_Owa_Name)
	$groupbox1.Controls.Add($lbl_LHold)
	$groupbox1.Controls.Add($lbl_ASync)
	$groupbox1.Controls.Add($lbl_Owa)
	$groupbox1.Controls.Add($lbl_Pop)
	$groupbox1.Controls.Add($lbl_Imap)
	$groupbox1.Controls.Add($lbl_RecipientType)
	$groupbox1.Controls.Add($button_Search)
	$groupbox1.Controls.Add($labelUserid)
	$groupbox1.Controls.Add($lbl_Sync)
	$groupbox1.Controls.Add($txt_UPN)
	$groupbox1.Controls.Add($cbox_MbxType)
	$groupbox1.Controls.Add($labelMailboxType)
	$groupbox1.Controls.Add($groupbox3)
    $groupbox1.Controls.Add($lbl_Msize)
    $groupbox1.Controls.Add($txt_Msize)
    $groupbox1.Controls.Add($btn_Set)
    #$groupbox1.Controls.Add($btn_OOFenable)
    #$groupbox1.Controls.Add($txt_OOF)
    $groupbox1.Controls.Add($lbl_IPHold)
    $groupbox1.Controls.Add($lbl_IPHold_Name)
    $groupbox1.Controls.Add($btn_CalPerm)
    $groupbox1.Controls.Add($btn_CalPerm_Add)
    $groupbox1.Controls.Add($btn_CalPerm_Remove)
    $groupbox1.Controls.Add($listA_CalPerm)
    $groupbox1.Controls.Add($listB_CalPerm)
    $groupbox1.Controls.Add($txtA_CalPerm)
    $groupbox1.Controls.Add($cbox_CalPerm)
    $groupbox1.Controls.Add($groupbox4)


	$groupbox1.BackColor = 'Control'
	$groupbox1.FlatStyle = 'Popup'
	$groupbox1.Location = '7, 6'
	$groupbox1.Name = 'groupbox1'
	$groupbox1.Size = '1022, 547'
	$groupbox1.TabIndex = 30
	$groupbox1.TabStop = $False
	$groupbox1.UseCompatibleTextRendering = $True
	$groupbox1.add_Enter($groupbox1_Enter)

	# lbl_Connected
	#
	$lbl_Connected.AutoSize = $True
	$lbl_Connected.BackColor = '192, 255, 192'
	$lbl_Connected.BorderStyle = 'FixedSingle'
	$lbl_Connected.Font = 'Microsoft Sans Serif, 8.25pt, style=Bold'
	$lbl_Connected.ForeColor = 'LimeGreen'
	$lbl_Connected.Location = '372, 0'
	$lbl_Connected.Name = 'lbl_Connected'
	$lbl_Connected.Size = '178, 20'
	$lbl_Connected.TabIndex = 26
	$lbl_Connected.Text = ' Not Connected '
	$lbl_Connected.TextAlign = 'MiddleCenter'
	$lbl_Connected.UseCompatibleTextRendering = $True
	#

	# groupbox2
	#
	$groupbox2.Controls.Add($lbl_TDI)
	$groupbox2.Controls.Add($cbox_TDI)
	$groupbox2.Location = '221, 128'
	$groupbox2.Name = 'groupbox2'
	$groupbox2.Size = '439, 58'
	$groupbox2.TabIndex = 84
	$groupbox2.TabStop = $False
	$groupbox2.Text = 'Mailbox Size'
	$groupbox2.UseCompatibleTextRendering = $True
	#

    # LabelUserid
    #
	$labelUserid.AutoSize = $True
	$labelUserid.Font = 'Microsoft Sans Serif, 12pt'
	$labelUserid.Location = '10, 25'
	$labelUserid.Name = 'labelUserid'
	$labelUserid.Size = '80,24'
	$labelUserid.TabIndex = 6
	$labelUserid.Text = 'User Name'
	$labelUserid.UseCompatibleTextRendering = $True
    #

	# txt_UPN
	#
	$txt_UPN.Font = 'Microsoft Sans Serif, 11pt'
	$txt_UPN.Location = '100, 25'
	$txt_UPN.Name = 'txt_UPN'
	$txt_UPN.Size = '270, 24'
	$txt_UPN.TabIndex = 1
    #

	# btn_Search
	#
	$button_Search.Font = 'Microsoft Sans Serif, 12pt'
	$button_Search.Location = '380, 25'
	$button_Search.Name = 'btn_Search'
	$button_Search.Size = '80, 30'
	$button_Search.TabIndex = 2
	$button_Search.Text = 'Search'
	$button_Search.UseCompatibleTextRendering = $True
	$button_Search.UseVisualStyleBackColor = $True
	$button_Search.add_Click($btn_Search_Click)
	#

	# labelOutOfOffice
	#
	$labelOutOfOffice.BackColor = '255, 192, 192'
	$labelOutOfOffice.BorderStyle = 'FixedSingle'
	$labelOutOfOffice.Font = 'Microsoft Sans Serif, 11.25pt'
	$labelOutOfOffice.ForeColor = 'Black'
	$labelOutOfOffice.Location = '465,25'
	$labelOutOfOffice.Name = 'labelOutOfOffice'
	$labelOutOfOffice.Size = '100, 30'
	$labelOutOfOffice.TabIndex = 58
	$labelOutOfOffice.Text = 'Out of Office'
	$labelOutOfOffice.TextAlign = 'MiddleCenter'
	$labelOutOfOffice.UseCompatibleTextRendering = $True
	#

    # lbl_OOF
	#
	$lbl_OOF.BorderStyle = 'Fixed3D'
	$lbl_OOF.Font = 'Microsoft Sans Serif, 11.25pt'
	$lbl_OOF.ForeColor = '255, 128, 0'
	$lbl_OOF.Location = '570,25'
	$lbl_OOF.Name = 'lbl_OOF'
	$lbl_OOF.Size = '90, 30'
	$lbl_OOF.TabIndex = 59
	$lbl_OOF.TextAlign = 'MiddleCenter'
	$lbl_OOF.UseCompatibleTextRendering = $True
	#

	# lbl_RecipientType
	#
	$lbl_RecipientType.AutoSize = $True
	$lbl_RecipientType.Font = 'Microsoft Sans Serif, 9pt, style=Italic'
	$lbl_RecipientType.ForeColor = 'Black'
	$lbl_RecipientType.Location = '280, 56'
	$lbl_RecipientType.Name = 'lbl_RecipientType'
	$lbl_RecipientType.Size = '90, 19'
	$lbl_RecipientType.TabIndex = 33
	$lbl_RecipientType.Text = 'Recipient Type'
	$lbl_RecipientType.UseCompatibleTextRendering = $True
	#

    # lbl_Sync
	#
	$lbl_Sync.AutoSize = $True
	$lbl_Sync.Font = 'Microsoft Sans Serif, 9pt, style=Italic'
	$lbl_Sync.ForeColor = '64, 64, 64'
	$lbl_Sync.Location = '120, 56'
	$lbl_Sync.Name = 'lbl_Sync'
	$lbl_Sync.Size = '70, 19'
	$lbl_Sync.TabIndex = 28
	$lbl_Sync.Text = 'Sync Status'
	$lbl_Sync.UseCompatibleTextRendering = $True
	#

	# lbl_TDI
	#
	$lbl_TDI.BackColor = 'Control'
	$lbl_TDI.BorderStyle = 'FixedSingle'
	$lbl_TDI.Font = 'Microsoft Sans Serif, 10pt'
	$lbl_TDI.Location = '170, 20'
	$lbl_TDI.Name = 'lbl_TDI'
	$lbl_TDI.Size = '250, 23'
	$lbl_TDI.TabIndex = 39
	$lbl_TDI.TextAlign = 'MiddleCenter'
	$lbl_TDI.UseCompatibleTextRendering = $True
	#

    # cbox_TDI
	#
	$cbox_TDI.BackColor = 'Control'
	$cbox_TDI.Font = 'Microsoft Sans Serif, 10pt'
	$cbox_TDI.FormattingEnabled = $True
	[void]$cbox_TDI.Items.Add('Mailbox')
	[void]$cbox_TDI.Items.Add('Archive')
	$cbox_TDI.Location = '10, 20'
	$cbox_TDI.Name = 'cbox_TDI'
	$cbox_TDI.Size = '150, 24'
	$cbox_TDI.TabIndex = 7
	$cbox_TDI.Text = 'Mailbox/Archive'
	$cbox_TDI.add_SelectedIndexChanged($cbox_TDI_SelectedIndexChanged)
	#

    # btn_OOFEnable
    #
	#$btn_OOFenable.Font = 'Microsoft Sans Serif, 11pt'
	#$btn_OOFenable.Location = '761, 22'
	#$btn_OOFenable.Name = 'btn_OOFEnable'
	#$btn_OOFenable.Size = '87, 27'
	#$btn_OOFenable.TabIndex = 3
	#$btn_OOFenable.Text = 'Enable'
	#$btn_OOFenable.UseCompatibleTextRendering = $True
	#$btn_OOFenable.UseVisualStyleBackColor = $True
	#$btn_OOFenable.add_Click($btn_OOFenable_Click)
    #

    # txt_OFF
    #
    #$txt_OOF.Font = 'Microsoft Sans Serif, 10pt'
    #$txt_OOF.Location = '761 ,54'
    #$txt_OOF.Multiline = $True
    #$txt_OOF.Name = 'txt_OOF'
    #$txt_OOF.Size = '251, 100'
    #$txt_OOF.TabIndex = 0

	# listB_MbxPerm
	#
	$listB_MbxPerm.Font = 'Microsoft Sans Serif, 9pt'
	$listB_MbxPerm.FormattingEnabled = $True
	$listB_MbxPerm.HorizontalScrollbar = $True
	$listB_MbxPerm.ItemHeight = 15
	$listB_MbxPerm.Location = '670, 261'
	$listB_MbxPerm.Name = 'listB_MbxPerm'
	$listB_MbxPerm.ScrollAlwaysVisible = $True
	$listB_MbxPerm.SelectionMode = 'None'
	$listB_MbxPerm.Size = '340, 139'
	$listB_MbxPerm.TabIndex = 31
	$listB_MbxPerm.add_SelectedIndexChanged($listB_MbxPerm_SelectedIndexChanged)
	#

	# cbox_MbxPermission
	#
	$cbox_MbxPermission.BackColor = 'Control'
	$cbox_MbxPermission.Font = 'Microsoft Sans Serif, 10pt'
	$cbox_MbxPermission.FormattingEnabled = $True
	[void]$cbox_MbxPermission.Items.Add('Full Access')
	[void]$cbox_MbxPermission.Items.Add('Send On Behalf')
	[void]$cbox_MbxPermission.Items.Add('Send As')
	$cbox_MbxPermission.Location = '670, 237'
	$cbox_MbxPermission.Name = 'cbox_MbxPermission'
	$cbox_MbxPermission.Size = '340, 24'
	$cbox_MbxPermission.TabIndex = 36
	$cbox_MbxPermission.Text = 'Mailbox Permission'
	$cbox_MbxPermission.add_SelectedIndexChanged($cbox_MbxPermission_SelectedIndexChanged)
	#

	# txt_Enterid
	#
	$txt_Enterid.Font = 'Microsoft Sans Serif, 11pt'
	$txt_Enterid.Location = '670, 400'
	$txt_Enterid.Name = 'txt_EnterUPN'
	$txt_Enterid.Size = '340, 24'
	$txt_Enterid.TabIndex = 87
    $txt_Enterid.Text = "Enter EmailAddress"
	#

	# btn_Remove
	#
	$btn_Remove.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_Remove.Location = '939, 431'
	$btn_Remove.Name = 'btn_Remove'
	$btn_Remove.Size = '73, 30'
	$btn_Remove.TabIndex = 48
	$btn_Remove.Text = 'Remove'
	$btn_Remove.UseCompatibleTextRendering = $True
	$btn_Remove.UseVisualStyleBackColor = $True
	$btn_Remove.add_Click($btn_Remove_Click)
	#

	# btn_Add
	#
	$btn_Add.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_Add.Location = '886, 431'
	$btn_Add.Name = 'btn_Add'
	$btn_Add.Size = '51, 30'
	$btn_Add.TabIndex = 31
	$btn_Add.Text = 'Add'
	$btn_Add.UseCompatibleTextRendering = $True
	$btn_Add.UseVisualStyleBackColor = $True
	$btn_Add.add_Click($btn_Add_Click)
	#


	# txt_LicenseDetails
	#
	$txt_LicenseDetails.Font = 'Microsoft Sans Serif, 10pt'
	$txt_LicenseDetails.Location = '8, 109'
	$txt_LicenseDetails.Multiline = $True
	$txt_LicenseDetails.Name = 'txt_LicenseDetails'
	$txt_LicenseDetails.ReadOnly = $True
	$txt_LicenseDetails.ScrollBars = 'Both'
	$txt_LicenseDetails.Size = '210, 119'
	$txt_LicenseDetails.TabIndex = 65
	#
	# lbl_LicenseDetails
	#
	$lbl_LicenseDetails.BackColor = '255, 255, 128'
	$lbl_LicenseDetails.BorderStyle = 'FixedSingle'
	$lbl_LicenseDetails.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_LicenseDetails.Location = '8, 78'
	$lbl_LicenseDetails.Name = 'lbl_LicenseDetails'
	$lbl_LicenseDetails.Size = '210, 28'
	$lbl_LicenseDetails.TabIndex = 66
	$lbl_LicenseDetails.Text = 'User License Details'
	$lbl_LicenseDetails.TextAlign = 'MiddleCenter'
	$lbl_LicenseDetails.UseCompatibleTextRendering = $True
	#
	# lbl_Licensed
	#
	$lbl_Licensed.BorderStyle = 'Fixed3D'
	$lbl_Licensed.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_Licensed.ForeColor = '255, 128, 0'
	$lbl_Licensed.Location = '897, 505'
	$lbl_Licensed.Name = 'lbl_Licensed'
	$lbl_Licensed.Size = '123, 27'
	$lbl_Licensed.TabIndex = 64
	$lbl_Licensed.Text = 'Waiting...'
	$lbl_Licensed.TextAlign = 'MiddleCenter'
	$lbl_Licensed.UseCompatibleTextRendering = $True
	#

	# buttonDisable
	#
	#$buttonDisable.Font = 'Microsoft Sans Serif, 11pt'
	#$buttonDisable.Location = '562, 54'
	#$buttonDisable.Name = 'buttonDisable'
	#$buttonDisable.Size = '87, 27'
	#$buttonDisable.TabIndex = 3
	#$buttonDisable.Text = 'Disable'
	#$buttonDisable.UseCompatibleTextRendering = $True
	#$buttonDisable.UseVisualStyleBackColor = $True
	#$buttonDisable.add_Click($buttonDisable_Click)
	#

	# labelForwardingAddress
	#
	$labelForwardingAddress.AutoSize = $True
	$labelForwardingAddress.BackColor = '192, 192, 255'
	$labelForwardingAddress.BorderStyle = 'FixedSingle'
	$labelForwardingAddress.Font = 'Microsoft Sans Serif, 10pt'
	$labelForwardingAddress.ForeColor = 'Black'
	$labelForwardingAddress.Location = '72, 238'
	$labelForwardingAddress.Name = 'labelForwardingAddress'
	$labelForwardingAddress.Size = '126, 23'
	$labelForwardingAddress.TabIndex = 53
	$labelForwardingAddress.Text = 'ForwardingAddress'
	$labelForwardingAddress.TextAlign = 'MiddleCenter'
	$labelForwardingAddress.UseCompatibleTextRendering = $True
	#

	# btn_RemoveFwdAdd
	#
	$btn_RemoveFwdAdd.AutoSize = $True
	$btn_RemoveFwdAdd.Font = 'Microsoft Sans Serif, 8pt'
	$btn_RemoveFwdAdd.Location = '14, 262'
	$btn_RemoveFwdAdd.Name = 'btn_RemoveFwdAdd'
	$btn_RemoveFwdAdd.Size = '55, 24'
	$btn_RemoveFwdAdd.TabIndex = 11
	$btn_RemoveFwdAdd.Text = 'Remove'
	$btn_RemoveFwdAdd.UseCompatibleTextRendering = $True
	$btn_RemoveFwdAdd.UseVisualStyleBackColor = $True
	$btn_RemoveFwdAdd.add_Click($btn_RemoveFwdAdd_Click)
	#
	# btn_AddFwdAdd
	#
	$btn_AddFwdAdd.AutoSize = $True
	$btn_AddFwdAdd.Font = 'Microsoft Sans Serif, 8pt'
	$btn_AddFwdAdd.Location = '14, 237'
	$btn_AddFwdAdd.Name = 'btn_AddFwdAdd'
	$btn_AddFwdAdd.Size = '55, 24'
	$btn_AddFwdAdd.TabIndex = 10
	$btn_AddFwdAdd.Text = 'Add'
	$btn_AddFwdAdd.UseCompatibleTextRendering = $True
	$btn_AddFwdAdd.UseVisualStyleBackColor = $True
	$btn_AddFwdAdd.add_Click($btn_AddFwdAdd_Click)
	#
    
    # txt_FwdAdd
	#
	$txt_FwdAdd.Font = 'Microsoft Sans Serif, 11pt'
	$txt_FwdAdd.Location = '72, 262'
	$txt_FwdAdd.Name = 'txt_FwdAdd'
	$txt_FwdAdd.Size = '298, 24'
	$txt_FwdAdd.TabIndex = 12
	#

	# lbl_Audit
	#
	$lbl_Audit.BorderStyle = 'Fixed3D'
	$lbl_Audit.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_Audit.ForeColor = '255, 128, 0'
	$lbl_Audit.Location = '777, 505'
	$lbl_Audit.Name = 'lbl_Audit'
	$lbl_Audit.Size = '123, 27'
	$lbl_Audit.TabIndex = 35
	$lbl_Audit.Text = 'Waiting...'
	$lbl_Audit.TextAlign = 'MiddleCenter'
	$lbl_Audit.UseCompatibleTextRendering = $True
	#

	# lbl_LHold
	#
	$lbl_LHold.BorderStyle = 'Fixed3D'
	$lbl_LHold.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_LHold.ForeColor = '255, 128, 0'
	$lbl_LHold.Location = '654, 505'
	$lbl_LHold.Name = 'lbl_LHold'
	$lbl_LHold.Size = '123, 27'
	$lbl_LHold.TabIndex = 22
	$lbl_LHold.Text = 'Waiting...'
	$lbl_LHold.TextAlign = 'MiddleCenter'
	$lbl_LHold.UseCompatibleTextRendering = $True
	#

	# lbl_ASync
	#
	$lbl_ASync.BorderStyle = 'Fixed3D'
	$lbl_ASync.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_ASync.ForeColor = '255, 128, 0'
	$lbl_ASync.Location = '492, 505'
	$lbl_ASync.Name = 'lbl_ASync'
	$lbl_ASync.Size = '162, 27'
	$lbl_ASync.TabIndex = 12
	$lbl_ASync.Text = 'Waiting...'
	$lbl_ASync.TextAlign = 'MiddleCenter'
	$lbl_ASync.UseCompatibleTextRendering = $True
	#

	# lbl_Owa
	#
	$lbl_Owa.BorderStyle = 'Fixed3D'
	$lbl_Owa.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_Owa.ForeColor = '255, 128, 0'
	$lbl_Owa.Location = '372, 505'
	$lbl_Owa.Name = 'lbl_Owa'
	$lbl_Owa.Size = '120, 27'
	$lbl_Owa.TabIndex = 11
	$lbl_Owa.Text = 'Waiting...'
	$lbl_Owa.TextAlign = 'MiddleCenter'
	$lbl_Owa.UseCompatibleTextRendering = $True
	#

	# lbl_Pop
	#
	$lbl_Pop.BorderStyle = 'Fixed3D'
	$lbl_Pop.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_Pop.ForeColor = '255, 128, 0'
	$lbl_Pop.Location = '255, 505'
	$lbl_Pop.Name = 'lbl_Pop'
	$lbl_Pop.Size = '116, 27'
	$lbl_Pop.TabIndex = 10
	$lbl_Pop.Text = 'Waiting...'
	$lbl_Pop.TextAlign = 'MiddleCenter'
	$lbl_Pop.UseCompatibleTextRendering = $True
	#

	# lbl_Imap
	#
	$lbl_Imap.BorderStyle = 'Fixed3D'
	$lbl_Imap.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_Imap.ForeColor = '255, 128, 0'
	$lbl_Imap.Location = '134, 505'
	$lbl_Imap.Name = 'lbl_Imap'
	$lbl_Imap.Size = '122, 27'
	$lbl_Imap.TabIndex = 9
	$lbl_Imap.Text = 'Waiting...'
	$lbl_Imap.TextAlign = 'MiddleCenter'
	$lbl_Imap.UseCompatibleTextRendering = $True
	#


    # groupbox3
	#
	$groupbox3.Controls.Add($lbl_Status)
	$groupbox3.Controls.Add($chkBox_DeliverToMbx)
	$groupbox3.Location = '9, 224'
	$groupbox3.Name = 'groupbox3'
	$groupbox3.Size = '651, 70'
	$groupbox3.TabIndex = 86
	$groupbox3.TabStop = $False
	$groupbox3.UseCompatibleTextRendering = $True
	#

	# lbl_Status
	#
	$lbl_Status.AutoSize = $True
	$lbl_Status.Font = 'Microsoft Sans Serif, 10pt, style=Bold, Italic'
	$lbl_Status.ForeColor = '255, 128, 0'
	$lbl_Status.Location = '410, 17'
	$lbl_Status.Name = 'lbl_Status'
	$lbl_Status.Size = '46, 20'
	$lbl_Status.TabIndex = 88
	$lbl_Status.Text = 'Status'
	$lbl_Status.UseCompatibleTextRendering = $True
	#

	# chkBox_DeliverToMbx
	#
	$chkBox_DeliverToMbx.Font = 'Microsoft Sans Serif, 10pt'
	$chkBox_DeliverToMbx.Location = '212, 14'
	$chkBox_DeliverToMbx.Name = 'chkBox_DeliverToMbx'
	$chkBox_DeliverToMbx.Size = '213, 24'
	$chkBox_DeliverToMbx.TabIndex = 13
	$chkBox_DeliverToMbx.Text = 'DeliverToMailboxAndForward :'
	$chkBox_DeliverToMbx.UseCompatibleTextRendering = $True
	$chkBox_DeliverToMbx.UseVisualStyleBackColor = $True
	$chkBox_DeliverToMbx.add_CheckedChanged($chkBox_DeliverToMbx_CheckedChanged)
	#

	# lbl_LHold_Name
	#
    $lbl_LHold_Name.BorderStyle = 'Fixed3D'
	$lbl_LHold_Name.FlatStyle = 'Popup'
	$lbl_LHold_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_LHold_Name.Location = '654, 471'
	$lbl_LHold_Name.Name = 'lbl_LHold_Name'
	$lbl_LHold_Name.Size = '123, 27'
	$lbl_LHold_Name.TabIndex = 32
	$lbl_LHold_Name.Text = 'Litigation Hold'
	$lbl_LHold_Name.UseCompatibleTextRendering = $True
    $lbl_LHold_Name.TextAlign = 'MiddleCenter'
	#

	# lbl_ASync_Name
	#
    $lbl_ASync_Name.BorderStyle = 'Fixed3D'
	$lbl_ASync_Name.FlatStyle = 'Popup'
	$lbl_ASync_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_ASync_Name.Location = '492, 471'
	$lbl_ASync_Name.Name = 'lbl_ASync_Name'
	$lbl_ASync_Name.Size = '162, 27'
	$lbl_ASync_Name.TabIndex = 31
	$lbl_ASync_Name.Text = 'ActiveSync Enabled'
	$lbl_ASync_Name.UseCompatibleTextRendering = $True
	$lbl_ASync_Name.BorderStyle = 'Fixed3D'
    $lbl_ASync_Name.TextAlign = 'MiddleCenter'
	#

	# lbl_Pop_Name
	#
    $lbl_Pop_Name.BorderStyle = 'Fixed3D'
	$lbl_Pop_Name.FlatStyle = 'Popup'
	$lbl_Pop_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_Pop_Name.Location = '255, 471'
	$lbl_Pop_Name.Name = 'lbl_Pop_Name'
	$lbl_Pop_Name.Size = '116, 27'
	$lbl_Pop_Name.TabIndex = 29
	$lbl_Pop_Name.Text = 'POP Enabled'
	$lbl_Pop_Name.UseCompatibleTextRendering = $True
	$lbl_Pop_Name.BorderStyle = 'Fixed3D'
    $lbl_Pop_Name.TextAlign = 'MiddleCenter'
	#

	# lbl_Imap_Name
	#
    $lbl_Imap_Name.BorderStyle = 'Fixed3D'
	$lbl_Imap_Name.FlatStyle = 'Popup'
	$lbl_Imap_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_Imap_Name.Location = '134, 471'
	$lbl_Imap_Name.Name = 'lbl_Imap_Name'
	$lbl_Imap_Name.Size = '122, 27'
	$lbl_Imap_Name.TabIndex = 28
	$lbl_Imap_Name.Text = 'IMAP Enabled'
	$lbl_Imap_Name.UseCompatibleTextRendering = $True
	$lbl_Imap_Name.BorderStyle = 'Fixed3D'
    $lbl_Imap_Name.TextAlign = 'MiddleCenter'
	#

	# lbl_Owa_Name
	#
	$lbl_Owa_Name.BorderStyle = 'Fixed3D'
	$lbl_Owa_Name.FlatStyle = 'Popup'
	$lbl_Owa_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_Owa_Name.Location = '372, 471'
	$lbl_Owa_Name.Name = 'lbl_Owa_Name'
	$lbl_Owa_Name.Size = '120, 27'
	$lbl_Owa_Name.TabIndex = 30
	$lbl_Owa_Name.Text = 'OWA Enabled'
	$lbl_Owa_Name.UseCompatibleTextRendering = $True
	$lbl_Owa_Name.BorderStyle = 'Fixed3D'
    $lbl_Owa_Name.TextAlign = 'MiddleCenter'
	#

    # lbl_Audit_Name
	#
	$lbl_Audit_Name.BorderStyle = 'Fixed3D'
	$lbl_Audit_Name.FlatStyle = 'Popup'
	$lbl_Audit_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_Audit_Name.Location = '777, 471'
	$lbl_Audit_Name.Name = 'lbl_Audit_Name'
	$lbl_Audit_Name.Size = '123, 27'
	$lbl_Audit_Name.TabIndex = 33
	$lbl_Audit_Name.Text = 'Audit Enabled'
	$lbl_Audit_Name.UseCompatibleTextRendering = $True
	$lbl_Audit_Name.BorderStyle = 'Fixed3D'
    $lbl_Audit_Name.TextAlign = 'MiddleCenter'
	#

	# lbl_Licensed_Name
	#
	$lbl_Licensed_Name.BorderStyle = 'Fixed3D'
	$lbl_Licensed_Name.FlatStyle = 'Popup'
	$lbl_Licensed_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_Licensed_Name.Location = '897, 471'
	$lbl_Licensed_Name.Name = 'lbl_Licensed_Name'
	$lbl_Licensed_Name.Size = '123, 27'
	$lbl_Licensed_Name.TabIndex = 34
	$lbl_Licensed_Name.Text = 'Is Licensed'
	$lbl_Licensed_Name.UseCompatibleTextRendering = $True
	$lbl_Licensed_Name.BorderStyle = 'Fixed3D'
    $lbl_Licensed_Name.TextAlign = 'MiddleCenter'
	#

	# lbl_Msize
	#
	$lbl_Msize.AutoSize = $True
	$lbl_Msize.BorderStyle = 'FixedSingle'
	$lbl_Msize.Font = 'Microsoft Sans Serif, 11pt'
	$lbl_Msize.ForeColor = 'Black'
	$lbl_Msize.Location = '221, 193'
	$lbl_Msize.Name = 'lbl_Msize'
	$lbl_Msize.Size = '100, 25'
	$lbl_Msize.TabIndex = 49
	$lbl_Msize.Text = 'Set Mailbox Size'
	$lbl_Msize.TextAlign = 'MiddleCenter'
	$lbl_Msize.UseCompatibleTextRendering = $True
	#

	# txt_Msize
	#
	$txt_Msize.Font = 'Microsoft Sans Serif, 11pt, style=Bold'
	$txt_Msize.Location = '350, 193'
	$txt_Msize.Name = 'txt_Msize'
	$txt_Msize.Size = '100, 25'
	$txt_Msize.TabIndex = 50
	#

    # btn_Set
	#
	$btn_Set.AutoSize = $True
	$btn_Set.Font = 'Microsoft Sans Serif, 11pt'
	$btn_Set.Location = '460, 193'
	$btn_Set.Name = 'btn_Set'
	$btn_Set.Size = '30, 25'
	$btn_Set.TabIndex = 9
	$btn_Set.Text = 'Set'
	$btn_Set.UseCompatibleTextRendering = $True
	$btn_Set.UseVisualStyleBackColor = $True
	$btn_Set.add_Click($btn_Set_Click)
	#


    # lbl_IPHold_Name
	#
	$lbl_IPHold_Name.BorderStyle = 'Fixed3D'
	$lbl_IPHold_Name.FlatStyle = 'Popup'
	$lbl_IPHold_Name.Font = 'Microsoft Sans Serif, 12pt'
	$lbl_IPHold_Name.Location = '12, 471'
	$lbl_IPHold_Name.Name = 'lbl_IPHold_Name'
	$lbl_IPHold_Name.Size = '123, 27'
	$lbl_IPHold_Name.TabIndex = 27
	$lbl_IPHold_Name.Text = 'InPlace Hold'
	$lbl_IPHold_Name.UseCompatibleTextRendering = $True
    $lbl_IPHold_Name.TextAlign = 'MiddleCenter'
	#

	# lbl_IPHold
	#
	$lbl_IPHold.BorderStyle = 'Fixed3D'
	$lbl_IPHold.Font = 'Microsoft Sans Serif, 12pt, style=Bold'
	$lbl_IPHold.ForeColor = '255, 128, 0'
	$lbl_IPHold.Location = '12, 505'
	$lbl_IPHold.Name = 'lbl_IPHold'
	$lbl_IPHold.Size = '123, 27'
	$lbl_IPHold.TabIndex = 17
	$lbl_IPHold.Text = 'Waiting...'
	$lbl_IPHold.TextAlign = 'MiddleCenter'
	$lbl_IPHold.UseCompatibleTextRendering = $True
	#

	# btn_CalPerm
	#
	$btn_CalPerm.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_CalPerm.Location = '670, 25'
	$btn_CalPerm.Name = 'btn_CalPerm'
	$btn_CalPerm.Size = '340, 25'
	$btn_CalPerm.TabIndex = 0
	$btn_CalPerm.Text = 'Calendar Permissions'
	$btn_CalPerm.UseCompatibleTextRendering = $True
    $btn_CalPerm.TextAlign = 'MiddleCenter'
    $btn_CalPerm.UseVisualStyleBackColor = $True
	$btn_CalPerm.add_Click($btn_CalPerm_Click)
	#

    # listA_CalPerm
    #
    $listA_CalPerm.Font = 'Microsoft Sans Serif, 9pt'
	$listA_CalPerm.FormattingEnabled = $True
	$listA_CalPerm.HorizontalScrollbar = $True
	$listA_CalPerm.ItemHeight = 15
	$listA_CalPerm.Location = '670, 50'
	$listA_CalPerm.Name = 'listA_CalPerm'
	$listA_CalPerm.ScrollAlwaysVisible = $True
	$listA_CalPerm.SelectionMode = 'None'
	$listA_CalPerm.Size = '170, 130'
	$listA_CalPerm.TabIndex = 31
	$listA_CalPerm.add_SelectedIndexChanged($listA_CalPerm_SelectedIndexChanged)
	#

    # listB_CalPerm
    #
    $listB_CalPerm.Font = 'Microsoft Sans Serif, 9pt'
	$listB_CalPerm.FormattingEnabled = $True
	$listB_CalPerm.HorizontalScrollbar = $True
	$listB_CalPerm.ItemHeight = 15
	$listB_CalPerm.Location = '840, 50'
	$listB_CalPerm.Name = 'listB_CalPerm'
	$listB_CalPerm.ScrollAlwaysVisible = $True
	$listB_CalPerm.SelectionMode = 'None'
	$listB_CalPerm.Size = '170, 130'
	$listB_CalPerm.TabIndex = 31
	$listB_CalPerm.add_SelectedIndexChanged($listB_CalPerm_SelectedIndexChanged)
	#

    #txtA_CalPerm
    #
    $txtA_CalPerm.Font = 'Microsoft Sans Serif, 11pt'
	$txtA_CalPerm.Location = '670, 180'
	$txtA_CalPerm.Name = 'txtA_CalPerm'
	$txtA_CalPerm.Size = '170, 25'
	$txtA_CalPerm.TabIndex = 50

    #$cbox_CalPerm
    #
    $cbox_CalPerm.BackColor = 'Control'
	$cbox_CalPerm.Font = 'Microsoft Sans Serif, 10pt'
	$cbox_CalPerm.FormattingEnabled = $True
	[void]$cbox_CalPerm.Items.Add('Reviewer')
	[void]$cbox_CalPerm.Items.Add('Editor')
	[void]$cbox_CalPerm.Items.Add('Owner')
    [void]$cbox_CalPerm.Items.Add('Author')
    [void]$cbox_CalPerm.Items.Add('Contributor')
    [void]$cbox_CalPerm.Items.Add('LimitedDetails')
    [void]$cbox_CalPerm.Items.Add('AvailabilityOnly')
	$cbox_CalPerm.Location = '840, 180'
	$cbox_CalPerm.Name = 'cbox_CalPerm'
	$cbox_CalPerm.Size = '170, 25'
	$cbox_CalPerm.TabIndex = 36
	$cbox_CalPerm.Text = 'Permission'
    #

	# btn_CalPerm_Remove
	#
	$btn_CalPerm_Remove.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_CalPerm_Remove.Location = '939, 205'
	$btn_CalPerm_Remove.Name = 'btn_CalPerm_Remove'
	$btn_CalPerm_Remove.Size = '70, 30'
	$btn_CalPerm_Remove.TabIndex = 48
	$btn_CalPerm_Remove.Text = 'Remove'
	$btn_CalPerm_Remove.UseCompatibleTextRendering = $True
	$btn_CalPerm_Remove.UseVisualStyleBackColor = $True
	$btn_CalPerm_Remove.add_Click($btn_CalPerm_Remove_Click)
	#

	# btn_CalPerm_Add
	#
	$btn_CalPerm_Add.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_CalPerm_Add.Location = '886, 205'
	$btn_CalPerm_Add.Name = 'btn_CalPerm_Add'
	$btn_CalPerm_Add.Size = '50, 30'
	$btn_CalPerm_Add.TabIndex = 31
	$btn_CalPerm_Add.Text = 'Add'
	$btn_CalPerm_Add.UseCompatibleTextRendering = $True
	$btn_CalPerm_Add.UseVisualStyleBackColor = $True
	$btn_CalPerm_Add.add_Click($btn_CalPerm_Add_Click)
	#

	# lbl_Successful
	#
	$lbl_Successful.BackColor = '255, 255, 128'
	$lbl_Successful.BorderStyle = 'FixedSingle'
	$lbl_Successful.Font = 'Microsoft Sans Serif, 13pt'
	$lbl_Successful.ForeColor = 'Black'
	$lbl_Successful.Location = '193, 436'
	$lbl_Successful.Name = 'lbl_Successful'
    $lbl_Successful.Text = 'Status'
	$lbl_Successful.Size = '600, 27'
	$lbl_Successful.TabIndex = 23
	$lbl_Successful.TextAlign = 'MiddleCenter'
	$lbl_Successful.UseCompatibleTextRendering = $True
	#

    # groupbox4
    #
    $groupbox4.Controls.Add($lbl_Ret)
    $groupbox4.Controls.Add($btn_Ret)
    $groupbox4.Controls.Add($cbox_Ret)
    $groupbox4.Controls.Add($btn_Ret_WF)
    $groupbox4.Controls.Add($btn_Ret_AF)
	$groupbox4.Location = '9, 300'
	$groupbox4.Name = 'groupbox4'
	$groupbox4.Size = '651, 45'
	$groupbox4.TabIndex = 90
	$groupbox4.TabStop = $False
	$groupbox4.UseCompatibleTextRendering = $True
    #

    # lbl_Ret
    #
    $lbl_Ret.AutoSize = $True
	$lbl_Ret.BorderStyle = 'FixedSingle'
	$lbl_Ret.Font = 'Microsoft Sans Serif, 11pt'
	$lbl_Ret.ForeColor = 'Black'
	$lbl_Ret.Location = '5, 10'
	$lbl_Ret.Name = 'lbl_Ret'
	$lbl_Ret.Size = '120, 25'
	$lbl_Ret.TabIndex = 91
	$lbl_Ret.Text = 'Retention Policy'
	$lbl_Ret.TextAlign = 'MiddleCenter'
	$lbl_Ret.UseCompatibleTextRendering = $True

    # cbox_Ret
    #
    $cbox_Ret.BackColor = 'Control'
	$cbox_Ret.Font = 'Microsoft Sans Serif, 10pt'
	$cbox_Ret.FormattingEnabled = $True
	[void]$cbox_Ret.Items.Add('GRC')
	$cbox_Ret.Location = '130, 10'
	$cbox_Ret.Name = 'cbox_Ret'
	$cbox_Ret.Size = '100, 25'
	$cbox_Ret.TabIndex = 92
	$cbox_Ret.Text = 'Policy'
    #

    # btn_Ret
    #
    $btn_Ret.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_Ret.Location = '240, 10'
	$btn_Ret.Name = 'btn_Ret'
	$btn_Ret.Size = '50, 25'
	$btn_Ret.TabIndex = 93
	$btn_Ret.Text = 'Set'
	$btn_Ret.UseCompatibleTextRendering = $True
    $btn_Ret.TextAlign = 'MiddleCenter'
    $btn_Ret.UseVisualStyleBackColor = $True
	$btn_Ret.add_Click($btn_Ret_Click)
	#

    # btn_Ret_AF
    #
    $btn_Ret_AF.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_Ret_AF.Location = '300, 10'
	$btn_Ret_AF.Name = 'btn_Ret_AF'
	$btn_Ret_AF.Size = '120, 25'
	$btn_Ret_AF.TabIndex = 93
	$btn_Ret_AF.Text = 'Archive Folder'
	$btn_Ret_AF.UseCompatibleTextRendering = $True
    $btn_Ret_AF.TextAlign = 'MiddleCenter'
    $btn_Ret_AF.UseVisualStyleBackColor = $True
	$btn_Ret_AF.add_Click($btn_Ret_AF_Click)
`	#

    # btn_Ret_WF
    #
    $btn_Ret_WF.Font = 'Microsoft Sans Serif, 10pt, style=Bold'
	$btn_Ret_WF.Location = '430, 10'
	$btn_Ret_WF.Name = 'btn_Ret_WF'
	$btn_Ret_WF.Size = '120, 25'
	$btn_Ret_WF.TabIndex = 93
	$btn_Ret_WF.Text = 'Working Folder'
	$btn_Ret_WF.UseCompatibleTextRendering = $True
    $btn_Ret_WF.TextAlign = 'MiddleCenter'
    $btn_Ret_WF.UseVisualStyleBackColor = $True
	$btn_Ret_WF.add_Click($btn_Ret_WF_Click)
`	#

    $groupbox4.ResumeLayout()
	$groupbox3.ResumeLayout()
	$groupbox2.ResumeLayout()
	$groupbox1.ResumeLayout()
	$Tool.ResumeLayout()

#endregion Generated Form Code

#------------------------------------------------------------------

#Save the initial state of the form 
$InitialFormWindowState = $Tool.WindowState 
#Show the Form 
$Tool.ShowDialog()| Out-Null

}#End Function

#Call the Function
GenerateForm