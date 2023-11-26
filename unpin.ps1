function Pin-App {    param(
        [string]$appname,
        [switch]$unpin
    )
    try{
        if ($unpin.IsPresent){
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match '�X�^�[�g��ʂ���s�����߂��O��'} | %{$_.DoIt()}
            return "App '$appname' unpinned from Start"
        }else{
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match '�X�^�[�g��ʂɃs�����߂���'} | %{$_.DoIt()}
            return "App '$appname' pinned to Start"
        }
    }catch{
        Write-Error "�G���[���������܂����B (�A�v���̖��O�͐������ł����H)"
    }
}

Pin-App "xbox" -unpin
Pin-App "�J�����_�[" -unpin
Pin-App "�V�C" -unpin
Pin-App "���[��" -unpin
Pin-App "Groove �~���[�W�b�N" -unpin
Pin-App "�f�� & �e���r" -unpin
Pin-App "�j���[�X" -unpin
Pin-App "�}�l�[" -unpin
Pin-App "�X�g�A" -unpin
Pin-App "�t�H�g" -unpin
Pin-App "���o�C�� �R���p�j�I��" -unpin
Pin-App "�V���� Office ���n�߂悤" -unpin