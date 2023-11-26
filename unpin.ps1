function Pin-App {    param(
        [string]$appname,
        [switch]$unpin
    )
    try{
        if ($unpin.IsPresent){
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'スタート画面からピン留めを外す'} | %{$_.DoIt()}
            return "App '$appname' unpinned from Start"
        }else{
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'スタート画面にピン留めする'} | %{$_.DoIt()}
            return "App '$appname' pinned to Start"
        }
    }catch{
        Write-Error "エラーが発生しました。 (アプリの名前は正しいですか？)"
    }
}

Pin-App "xbox" -unpin
Pin-App "カレンダー" -unpin
Pin-App "天気" -unpin
Pin-App "メール" -unpin
Pin-App "Groove ミュージック" -unpin
Pin-App "映画 & テレビ" -unpin
Pin-App "ニュース" -unpin
Pin-App "マネー" -unpin
Pin-App "ストア" -unpin
Pin-App "フォト" -unpin
Pin-App "モバイル コンパニオン" -unpin
Pin-App "新しい Office を始めよう" -unpin