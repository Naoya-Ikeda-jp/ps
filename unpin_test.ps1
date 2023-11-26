function Pin-App {    param(
        [string]$appname,
        [switch]$unpin
    )
    try{
        if ($unpin.IsPresent){
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match '�X�^�[�g����s�����߂��O��'} | %{$_.DoIt()}
            return "App '$appname' unpinned from Start"
        }else{
            ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match '�X�^�[�g�Ƀs�����߂���'} | %{$_.DoIt()}
            return "App '$appname' pinned to Start"
        }
    }catch{
        Write-Error "�G���[���������܂����B (�A�v���̖��O�͐������ł����H)"
    }
}

#��~�ꗗ
#�C���X�g�[������Ă��Ȃ��A�v���ɂ͌��ʂȂ�
#Pin-App "xbox" -unpin
Pin-App "Xbox �{�̃R���p�j�I��" -unpin
#Pin-App "�J�����_�[" -unpin
#Pin-App "�V�C" -unpin
Pin-App "���[��" -unpin
Pin-App "Groove �~���[�W�b�N" -unpin
Pin-App "�f�� & �e���r" -unpin
#Pin-App "�j���[�X" -unpin
#Pin-App "�}�l�[" -unpin
#Pin-App "�X�g�A" -unpin
Pin-App "�t�H�g" -unpin
#Pin-App "���o�C�� �R���p�j�I��" -unpin
#Pin-App "�V���� Office ���n�߂悤" -unpin

Pin-App "Office" -unpin
Pin-App "OneNote" -unpin
Pin-App "Microsoft store" -unpin
Pin-App "Microsoft Edge" -unpin
Pin-App "Sticky Notes" -unpin
Pin-App "Network Speed Test" -unpin
Pin-App "Skype" -unpin
Pin-App "Microsoft �����[�g �f�X�N�g�b�v" -unpin
Pin-App "Microsoft �j���[�X" -unpin
Pin-App "�}�b�v" -unpin
Pin-App "Microsoft Whiteboard" -unpin
Pin-App "Microsoft To-Do:List,Task&Reminder" -unpin
Pin-App "Office Lens" -unpin
Pin-App "Sway" -unpin
Pin-App "�A���[�� & �N���b�N" -unpin
Pin-App "�{�C�X ���R�[�_�[" -unpin
Pin-App "HP JumpStart" -unpin
Pin-App "HP Power Manager" -unpin
Pin-App "HP WorkWise" -unpin

Pin-App "�����[�g �f�X�N�g�b�v" -unpin
Pin-App "Microsoft To-Do" -unpin

#pin
Pin-App "PC" -pin
Pin-App "�R���g���[�� �p�l��" -pin

Read-Host "������ɂ� Enter �L�[�������Ă�������..." 