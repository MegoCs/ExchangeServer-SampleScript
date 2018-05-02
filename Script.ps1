$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Stop"
#Transciprt the output of the script to a text file: with -path "######" append mode 
Start-Transcript -path "Ps.log" -append

#pass the account you need access the inbox folder . 
$AccountWithImpersonationRights = "#######@Hotmail.com"
$AccountPassword="#############";
# The desired mail subject to search with could be removed with line 34. in order not to apply the subject search condition 
$MAIL_SUBJECT = "Bill Csv"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
if(![System.IO.File]::Exists($dllpath)){
    $dllpath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
    if(![System.IO.File]::Exists($dllpath)){
        write-host "ERROR:  EWS Client Version 2.0 is not installed Kindly Install and Rerun Script File `r`n";
        exit 1  
    }
}
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$creds = New-Object System.Net.NetworkCredential($AccountWithImpersonationRights,$AccountPassword)
$service.Credentials = $creds
$service.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
#Choose the folder you want to access.
$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$AccountWithImpersonationRights)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$Sfir = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $true)
$Sfsub = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $MAIL_SUBJECT)
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
#Add the filters to the filters collection.
$sfCollection.add($Sfir)
$sfCollection.add($Sfsub)
$sfCollection.add($Sfha)
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(2000)

#apply the filters collection to the provisioned folder.
$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)
$OUTPUT_PATH="Path To The Output Folder"
foreach ($miMailItems in $frFolderResult.Items){
                $miMailItems.Load()
                foreach($attach in $miMailItems.Attachments){
                    $attach.Load()       
                    $fiFile = new-object System.IO.FileStream(($OUTPUT_PATH + "\" + $attach.Name), [System.IO.FileMode]::Create)
                    $fiFile.Write($attach.Content, 0, $attach.Content.Length)
                    $fiFile.Close()
                    write-host "Attachment Extracted TO : "(($OUTPUT_PATH +"\"  + $attach.Name))". `r`n"
           }
                #$miMailItems.isread = $true
                #$miMailItems.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
                #$miMailItems.delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
}
Stop-Transcript
exit 0