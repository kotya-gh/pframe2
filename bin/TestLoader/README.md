# TestLoader

## �T�v

TestLoader�̓R�}���h���w��̏����Ŏ��s���A���s���ʁA�R�}���h�̃��^�[���R�[�h���擾���܂��B

�܂��A���s���ʂ̓t�@�C���Ɏ����I�ɋL�^����܂��B

## �C���X�g�[���@�A�n�ߕ�

TestLoader�t�H���_�z���́ulist.json�v��ݒ��A�umain.ps1�v�����s���܂��B

## ���s���ʂɂ���

TestLoader���s��A���̂悤�ȓ��e��W���o�͂��܂��B
```
���̃R�}���h�����s���܂��B: echo 'a'
�R�}���h���s��A���̃��b�Z�[�W���o�͂���܂��B: a
���̃R�}���h�͎��̃��^�[���R�[�h��Ԃ��܂��B: 0
�R�}���h���s�J�n����: 03/10/2023 18:42:54
�R�}���h���s�I������: 03/10/2023 18:42:54
�R�}���h���s����
a
�R�}���h�̃��^�[���R�[�h: 0
```
�܂��A�R�}���h���s���ʂ̓e�L�X�g�t�@�C���A�����JSON�t�@�C���ɋL�^����܂��B

�t�@�C�����͎��̃t�H�[�}�b�g�ŏo�͂���܂��B

0001_hostname_YYYYMMDDHHMMSS_true.txt

0001_hostname_YYYYMMDDHHMMSS_true.json

0002_hostname_YYYYMMDDHHMMSS_false.txt

0002_hostname_YYYYMMDDHHMMSS_false.json

�擪4���͌�q����utestNo�v�Ŏw��̍��Ԃł��B

hostname�́uhostname�v�ݒ�̕�����ł��B

true/false�́ureturnCode�v�A�ureturnMsg�v�Ǝ��s���ʂ��r���A��v�����true�A�s��v�̏ꍇ��false��t�^���܂��B

## list.json �̐ݒ�

EvidenceHomeDir�F���s���ʂ��i�[����f�B���N�g�����t���p�X�Ŏw�肵�܂��B��̏ꍇ�̓X�N���v�g���[�g�z����files�ɕۑ����܂��B

TestConfigure�F���s����R�}���h���i�[����z��ł��B

testId�F���s���ʂ��i�[����T�u�f�B���N�g���ł��BEvidenceHomeDir�z���ɍ쐬����܂��B

testItems�F�z�X�g���ƂɁA���s�������R�}���h���i�[����z��ł��B

testNo�F���Ԃł��B���s���ʂ̃t�@�C�����̐擪�S���i�O�l�߁j�ɕt�^����܂��B

hostname�F�R�}���h�����s����z�X�g���ł��B�X�N���v�g���s�z�X�g�Ɩ{�ݒ肪�قȂ�ꍇ�A�R�}���h���s���X�L�b�v����܂��B

testCommands�F�X�̃R�}���h���i�[����z��ł��B

order�FtestCommands���ł̃R�}���h���s�������w�肵�܂��B

command�F���s�������R�}���h���w�肵�܂��B

returnCode�F���҂��郊�^�[���R�[�h���w�肵�܂��B���Ғl�ƈ�v���邩�`�F�b�N���A���ʂ�Ԃ��܂��B

returnMsg�F���҂���o�̓��b�Z�[�W���w�肵�܂��B���Ғl�ƈ�v���邩�`�F�b�N���A���ʂ�Ԃ��܂��B��̏ꍇ�̓`�F�b�N���X�L�b�v���܂��B

## �J����

pikaru

## ���C�Z���X

MIT