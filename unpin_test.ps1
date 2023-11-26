function Pin-App {    param(
        [string]$appname,
        [switch]$unpin
    )
    try{
        if ($unpin.IsPresent){
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'スタートからピン留めを外す'} | %{$_.DoIt()}
            return "App '$appname' unpinned from Start"
        }else{
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'スタートにピン留めする'} | %{$_.DoIt()}
            return "App '$appname' pinned to Start"
        }
    }catch{
        Write-Error "エラーが発生しました。 (アプリの名前は正しいですか？)"
    }
}

#停止一覧
#インストールされていないアプリには効果なし
#Pin-App "xbox" -unpin
Pin-App "Xbox 本体コンパニオン" -unpin
#Pin-App "カレンダー" -unpin
#Pin-App "天気" -unpin
Pin-App "メール" -unpin
Pin-App "Groove ミュージック" -unpin
Pin-App "映画 & テレビ" -unpin
#Pin-App "ニュース" -unpin
#Pin-App "マネー" -unpin
#Pin-App "ストア" -unpin
Pin-App "フォト" -unpin
#Pin-App "モバイル コンパニオン" -unpin
#Pin-App "新しい Office を始めよう" -unpin

Pin-App "Office" -unpin
Pin-App "OneNote" -unpin
Pin-App "Microsoft store" -unpin
Pin-App "Microsoft Edge" -unpin
Pin-App "Sticky Notes" -unpin
Pin-App "Network Speed Test" -unpin
Pin-App "Skype" -unpin
Pin-App "Microsoft リモート デスクトップ" -unpin
Pin-App "Microsoft ニュース" -unpin
Pin-App "マップ" -unpin
Pin-App "Microsoft Whiteboard" -unpin
Pin-App "Microsoft To-Do:List,Task&Reminder" -unpin
Pin-App "Office Lens" -unpin
Pin-App "Sway" -unpin
Pin-App "アラーム & クロック" -unpin
Pin-App "ボイス レコーダー" -unpin
Pin-App "HP JumpStart" -unpin
Pin-App "HP Power Manager" -unpin
Pin-App "HP WorkWise" -unpin

Pin-App "リモート デスクトップ" -unpin
Pin-App "Microsoft To-Do" -unpin

#pin
Pin-App "PC" -pin
Pin-App "コントロール パネル" -pin

Read-Host "続けるには Enter キーを押してください..." 