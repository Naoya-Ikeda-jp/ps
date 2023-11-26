function Pin-App {    param(
        [string]$appname,
        [switch]$unpin
    )
    try{
        if ($unpin.IsPresent){
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match '�^�X�N�o�[����s�����߂��O��'} | %{$_.DoIt()}
            return "App '$appname' unpinned from Taskbar"
        }else{
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match '�^�X�N�o�[�Ƀs�����߂���'} | %{$_.DoIt()}
            return "App '$appname' pinned to Taskbar"
        }
    }catch{
        Write-Error "�G���[���������܂����B (�A�v���̖��O�͐������ł����H)"
    }
}

#��~�ꗗ
#Pin-App "Microsoft store" -unpin
Pin-App "Excel" -pin
Pin-App "Word" -pin
Pin-App "PowerPoint" -pin
Pin-App "Outlook" -pin

Read-Host "������s�ɂ� Enter �L�[�������Ă�������..." 