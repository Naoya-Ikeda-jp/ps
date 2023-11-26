function Pin-App {    param(
        [string]$appname,
        [switch]$unpin
    )
    try{
        if ($unpin.IsPresent){
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'タスクバーからピン留めを外す'} | %{$_.DoIt()}
            return "App '$appname' unpinned from Taskbar"
        }else{
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'タスクバーにピン留めする'} | %{$_.DoIt()}
            return "App '$appname' pinned to Taskbar"
        }
    }catch{
        Write-Error "エラーが発生しました。 (アプリの名前は正しいですか？)"
    }
}

#停止一覧
#Pin-App "Microsoft store" -unpin
Pin-App "Excel" -pin
Pin-App "Word" -pin
Pin-App "PowerPoint" -pin
Pin-App "Outlook" -pin

Read-Host "続けるsには Enter キーを押してください..." 