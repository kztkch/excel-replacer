#####################################
Replace-HeaderFooter.ps1�J������
#####################################

�T�v
========
�t�@�C���Ŏw�肳�ꂽExcel�t�@�C���Q�̑S�V�[�g�̃w�b�_�[�ƃt�b�^�[��ϊ�����ps1�X�N���v�g�B
�ʏ퓮��ł́A���ɐݒ�ς݂ł���ꍇ�̓t�@�C���̕ύX�͍s��Ȃ����S�݌v�i�̂͂��j�B

{Left,Center,Right}x{Header,Footer}�̐ݒ肪�\�B

�I�v�V����
=============

* ``- ReplaceHFFileList filename``: �ݒ肵�����w�b�_�[�E�t�b�^�[�̒l�Ɛݒ�Ώۃt�@�C���̃��X�g�������ꂽ�t�@�C����filename�Ŏw�肷��Bfilename���̃t�H�[�}�b�g�͐ݒ肵�����w�b�_�[�E�t�b�^�[�̓��e��1�s�ڂɁA2�s�ڈȍ~�ɐݒ�Ώۂ̃G�N�Z���t�@�C�����L����B1�s�ڂ̃t�H�[�}�b�g��Theader-left,header-center,header-right,footer-left,footer-center,foot-right��TSV�ŋL�ڂ������̂ƂȂ�B�f�t�H���g�l�� ``.\replaceHeaderFooterFile.lst``�B

* ``-Debug``: �{�I�v�V�������w�肵�Ď��s�����ꍇ�AExcel�E�B���h�E��\�����Ȃ��瓮�삷��悤�ɂȂ�B
* ``-Dryrun``: ���ۂɕϊ����������邩�ǂ���������\�����A���ۂ̕ύX�͍s��Ȃ��B

Disclaimer
=================

* Powershell 2.0 �� KingOffice �œ���m�F�����̂�Excel���Ɖ�����肪���邩������Ȃ��B
* Excel��Com�o�R�ő��삵�ē��삷�邽��Excel���C���X�g�[������Ă��Ȃ����ł͓��삵�Ȃ��B

�d�l����
=================

���̂�����o�b�h�m�E�n�E�I�ȑΏ������Ă��镔�������邩������Ȃ��̂Œ��ӁB

* ���ɐݒ�ΏۂɂȂ��Ă���ꍇ�͕ύX�����Ȃ��B

* �l�Ɏg������ꕶ���Ƃ���vba�̎d�l�ɏ�����悤���B�i�����ƒ��ׂĂȂ��B)

* �ŏ��ϊ��Ώۂ̃Z���͈͂� UsedRange �Ƃ��Ă������A�����̊ԈႢ�ŋ���ȃG�N�Z�������������PowerShell���n���O���ċ��낵���̂�``A1:K12``�Ńn�[�h�R�[�f�B���O���Ă���B

* ���܂ɃG�N�Z�������b�N����Ă��܂������������Bfinally��ł�����ReleaseComObject���Ă�B

* Com�I�u�W�F�N�g�Ƀt�@�C�����J�����悤�Ƃ���ƁA���s�����ꏊ�ł͂Ȃ����s���[�U�̃z�[���f�B���N�g�����J�����g�Ɍ�����炵���̂Ńt��PATH�Ŗ����ꍇ�̓t��PATH�ɒ�������������Ă���B

* �����l������ echo �Ƃ� write-host cmdlet�ŏo�͂����Ă���B������logging���������ǂ܂��ǂ����B

* param��[switch]�͎w�肵�Ă��Ȃ��Ƃ���$False�������Ă���̂������ƒ��ׂĂ��Ȃ��̂ŁABool�l�󂯓n���p�̕ϐ���ݒ肵�Ă���i���A�����Ǝd�l�𒲂ׂĖ��ʂȕϐ��Ƃ���������Ȃ疳���������B�j

vim: ft=rst tw=0:
