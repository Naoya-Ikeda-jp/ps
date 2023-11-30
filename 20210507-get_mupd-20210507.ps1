$OutputEncoding='utf-8'

Function Get-MicrosoftUpdates  
{   
  Param(  
        $NumberOfUpdates,  
        [switch]$all  
       )  
  $Session = New-Object -ComObject Microsoft.Update.Session  
  $Searcher = $Session.CreateUpdateSearcher()  
  if($all)  
    {  
      $HistoryCount = $Searcher.GetTotalHistoryCount()  
      $Searcher.QueryHistory(1,$HistoryCount)  
    }  
  Else { $Searcher.QueryHistory(1,$NumberOfUpdates) }  
} # Get-MicrosoftUpdates 関数の終わり  

Get-MicrosoftUpdates -all >> WinUpdate.log
#pause
