Attribute VB_Name = "ModBx"
Option Explicit


Public Sub dtgKj(Lb As Integer)
frmFYBX.dtgBx.Columns("����").Visible = True
frmFYBX.dtgBx.Columns("������").Visible = False
frmFYBX.dtgBx.Columns("���ݲ���").Visible = False
frmFYBX.dtgBx.Columns("���η�").Visible = False
frmFYBX.dtgBx.Columns("���·�").Visible = False
frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = False
frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = False
frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = False
frmFYBX.dtgBx.Columns("�˷�").Visible = False
frmFYBX.dtgBx.Columns("ס�޷�").Visible = False
frmFYBX.dtgBx.Columns("�����Ŷӷ�").Visible = False
frmFYBX.dtgBx.Columns("�ͷ�").Visible = False
frmFYBX.dtgBx.Columns("�д���").Visible = False
frmFYBX.dtgBx.Columns("��Ʒ��").Visible = False
frmFYBX.dtgBx.Columns("����").Visible = False
frmFYBX.dtgBx.Columns("��ҵ��").Visible = False
frmFYBX.dtgBx.Columns("ˮ��").Visible = False
frmFYBX.dtgBx.Columns("�绰").Visible = False
frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = False
'frmFYBX.dtgBx.Columns("����").Visible = False
frmFYBX.dtgBx.Columns("�г��ƹ�").Visible = False
frmFYBX.dtgBx.Columns("��Ա��Ƹ").Visible = False
frmFYBX.dtgBx.Columns("��ݷ�").Visible = False
frmFYBX.dtgBx.Columns("��ѵ��").Visible = False
frmFYBX.dtgBx.Columns("����������").Visible = False
frmFYBX.dtgBx.Columns("�Ŷӽ����").Visible = False
frmFYBX.dtgBx.Columns("ͣ����").Visible = False
frmFYBX.dtgBx.Columns("������").Visible = False
frmFYBX.dtgBx.Columns("����ͣ����").Visible = False
frmFYBX.dtgBx.Columns("����������").Visible = False
frmFYBX.dtgBx.Columns("����").Visible = False
frmFYBX.dtgBx.Columns("�׺�").Visible = False
frmFYBX.dtgBx.Columns("����").Visible = False
frmFYBX.dtgBx.Columns("��ͨ����").Visible = False
frmFYBX.dtgBx.Columns("פ�����").Visible = False
frmFYBX.dtgBx.Columns("��λ����").Visible = False
frmFYBX.dtgBx.Columns("�ۺϱ���").Visible = False
frmFYBX.dtgBx.Columns("����").Visible = False
frmFYBX.dtgBx.Columns("������").Visible = False
frmFYBX.dtgBx.Columns("��ͬ���").Visible = False
frmFYBX.dtgBx.Columns("����").Visible = False
'frmFYBX.dtgBx.Columns("����").Visible = False
frmFYBX.dtgBx.Columns("����").Visible = False
frmFYBX.dtgBx.Columns("������").Visible = False
frmFYBX.dtgBx.Columns("������ǩ��").Visible = False
frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = False
frmFYBX.dtgBx.Columns("������").Visible = False
frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = False
frmFYBX.dtgBx.Columns("ǩ������").Visible = False
frmFYBX.dtgBx.Columns("���⳵ע��").Visible = False
frmFYBX.dtgBx.Columns("ǩ������").Visible = False
frmFYBX.frmRen.Visible = False
frmFYBX.dtgNx.Visible = True
frmFYBX.dtgBx.Visible = False
Select Case Lb
Case 7 '��������
    frmFYBX.dtgBx.Columns("����").Visible = True
    'frmFYBX.dtgBx.Columns("��ҵ��").Visible = True
    frmFYBX.dtgBx.Columns("ˮ��").Visible = True
    frmFYBX.dtgBx.Columns("�绰").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    'frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("�г��ƹ�").Visible = True
    frmFYBX.dtgBx.Columns("��Ա��Ƹ").Visible = True
    frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
    frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
    frmFYBX.dtgBx.Columns("����������").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("����ͣ����").Visible = True
    frmFYBX.dtgBx.Columns("����������").Visible = True
Case 8 '�ܾ�����
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
Case 50 '�˷�
    frmFYBX.dtgBx.Columns("�˷�").Visible = True
    frmFYBX.dtgBx.Columns("��ͬ���").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True

Case 51 '�˷�
    frmFYBX.dtgBx.Columns("�˷�").Visible = True
    frmFYBX.dtgBx.Columns("��ͬ���").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True

Case 10 '����

Case 11 '�������
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("��ͬ���").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
    'frmFYBX.dtgBx.Columns("���߷�").Visible = True
    frmFYBX.dtgBx.Columns("�׺�").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
Case 12 '�������
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("��ͬ���").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
    'frmFYBX.dtgBx.Columns("���߷�").Visible = True
    frmFYBX.dtgBx.Columns("�׺�").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True

Case 53 '���۾���
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
    frmFYBX.dtgBx.Columns("�����Ŷӷ�").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    
Case 14 '���ž���
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("�����Ŷӷ�").Visible = True
    
Case 15 'ҵ��Ա
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    'frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
frmFYBX.dtgBx.Columns("���⳵ע��").Visible = True

Case 16 'ҵ��Ա
    'frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    'frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
frmFYBX.dtgBx.Columns("���⳵ע��").Visible = True

Case 17  '��ͨ����

   ' frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ݲ���").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
Case 18 '��ͨ����
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ݲ���").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    
Case 20
Case 21

Case 32 '���ù���
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
    'frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("�׺�").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    'frmFYBX.frmRen.Visible = True
Case 33 '���ù���(Ŀǰ�Ѳ�����)
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
    'frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
    frmFYBX.dtgBx.Columns("�׺�").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
Case 71 '���ù���
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    frmFYBX.dtgBx.Columns("�д���").Visible = True
    frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
    frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
    'frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
Case 35 '����
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("���ݲ���").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���η�").Visible = True
    frmFYBX.dtgBx.Columns("��ͨ����").Visible = True
    frmFYBX.dtgBx.Columns("פ�����").Visible = True
    frmFYBX.dtgBx.Columns("��λ����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
    frmFYBX.dtgBx.Columns("ywyuid").Width = 0
    frmFYBX.dtgBx.Visible = False
    frmFYBX.dtgNx.Visible = True
    frmFYBX.cmdG.Visible = True
Case 54 '���̲�
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    'frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("�׺�").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
Case 70 '���̲�
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
    frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
    frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
    frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
    'frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("�׺�").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
Case 55 '����
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("ywyuid").Width = 0
    frmFYBX.cmdG.Visible = True
        frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 56 '������
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("ywyuid").Width = 0
    frmFYBX.cmdG.Visible = True
        frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 58 '���´���������
    frmFYBX.dtgBx.Columns("����").Visible = True
    'frmFYBX.dtgBx.Columns("��ҵ��").Visible = True
    frmFYBX.dtgBx.Columns("ˮ��").Visible = True
    frmFYBX.dtgBx.Columns("�绰").Visible = True
    frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
    'frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("�г��ƹ�").Visible = True
    frmFYBX.dtgBx.Columns("��Ա��Ƹ").Visible = True
    frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
    'frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
    frmFYBX.dtgBx.Columns("����������").Visible = True
Case 59 '�����ۺϱ���
    frmFYBX.dtgBx.Columns("�ۺϱ���").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.cmdG.Visible = True
    frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 67 '���ݲ���
    frmFYBX.dtgBx.Columns("���ݲ���").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
Case 66 '����
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
Case 72 '���η�

    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
    frmFYBX.dtgBx.Columns("���η�").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = False
Case 84 '��ѵ
    frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
    'frmFYBX.cmdG.Visible = True
        frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 79 '�·��ù���
frmFYBX.dtgBx.Columns("������").Visible = True
frmFYBX.dtgBx.Columns("���ݲ���").Visible = True
frmFYBX.dtgBx.Columns("���η�").Visible = True
frmFYBX.dtgBx.Columns("���·�").Visible = True
frmFYBX.dtgBx.Columns("ͨ�ŷ�").Visible = True
frmFYBX.dtgBx.Columns("���ڽ�ͨ��").Visible = True
frmFYBX.dtgBx.Columns("���⽻ͨ��").Visible = True
frmFYBX.dtgBx.Columns("�˷�").Visible = True
frmFYBX.dtgBx.Columns("ס�޷�").Visible = True
frmFYBX.dtgBx.Columns("�����Ŷӷ�").Visible = True
frmFYBX.dtgBx.Columns("�ͷ�").Visible = True
frmFYBX.dtgBx.Columns("�д���").Visible = True
frmFYBX.dtgBx.Columns("��Ʒ��").Visible = True
frmFYBX.dtgBx.Columns("����").Visible = True
frmFYBX.dtgBx.Columns("��ҵ��").Visible = True
frmFYBX.dtgBx.Columns("ˮ��").Visible = True
frmFYBX.dtgBx.Columns("�绰").Visible = True
frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
'frmFYBX.dtgBx.Columns("����").Visible = True
frmFYBX.dtgBx.Columns("�г��ƹ�").Visible = True
frmFYBX.dtgBx.Columns("��Ա��Ƹ").Visible = True
frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
frmFYBX.dtgBx.Columns("��ѵ��").Visible = True
frmFYBX.dtgBx.Columns("����������").Visible = True
frmFYBX.dtgBx.Columns("�Ŷӽ����").Visible = True
frmFYBX.dtgBx.Columns("ͣ����").Visible = True
frmFYBX.dtgBx.Columns("������").Visible = True
frmFYBX.dtgBx.Columns("����ͣ����").Visible = True
frmFYBX.dtgBx.Columns("����������").Visible = True
'frmFYBX.dtgBx.Columns("����").Visible = True
frmFYBX.dtgBx.Columns("�׺�").Visible = True
frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = False
    frmFYBX.dtgBx.Columns("����").Visible = False
    frmFYBX.dtgBx.Columns("������").Visible = False
    frmFYBX.dtgBx.Columns("������ǩ��").Visible = False
    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = False
    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = False
    frmFYBX.dtgBx.Columns("ǩ������").Visible = False
frmFYBX.frmRen.Visible = True
Case 82 '�ڲ�����
frmFYBX.dtgBx.Columns("��ݷ�").Visible = True
frmFYBX.dtgBx.Columns("�칫��Ʒ").Visible = True
frmFYBX.dtgBx.Columns("������").Visible = True
frmFYBX.dtgBx.Columns("ͣ����").Visible = True
frmFYBX.dtgBx.Columns("������").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("����").Visible = True
    frmFYBX.dtgBx.Columns("������").Visible = True
'    frmFYBX.dtgBx.Columns("������ǩ��").Visible = True
'    frmFYBX.dtgBx.Columns("ǩ��ʱ��").Visible = True
'    frmFYBX.dtgBx.Columns("���ž���ǩ��").Visible = True
'    frmFYBX.dtgBx.Columns("ǩ������").Visible = True
    frmFYBX.cmdG.Visible = True
   frmFYBX.dtgBx.Visible = True
   frmFYBX.dtgNx.Visible = False
End Select


frmFYBX.dtgBx.Refresh

End Sub

Public Sub fydBound(Bxid As String)
Dim tt As String
Dim oo As Integer
Dim Lcou As Integer
On Error Resume Next
Lcou = 0
frmFYBX.lblBh.Caption = Bxid
frmFYBX.cmdSave.Enabled = False
'��¼����־
Call mod1.zhuDa(2, Bxid)

frmFYBX.Kd = False '�ǳ��ο���,�Ա㱣��ʱ������Ա��ǩ������

        tt = "fydOpen(" & Bxid & ")"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        frmFYBX.lblBh.Caption = mod1.HTP.Fields("BxId").Value
        frmFYBX.LblTrq.Caption = mod1.HTP.Fields("Trq").Value
        frmFYBX.comQy.Caption = mod1.HTP.Fields("qy").Value
        frmFYBX.lblBM.Caption = mod1.HTP.Fields("bm").Value
        frmFYBX.txtHg.Text = mod1.HTP.Fields("hG").Value
        frmFYBX.lblDx.Caption = mod1.HTP.Fields("hGD").Value
        frmFYBX.lblFR.Caption = mod1.HTP.Fields("fRQ").Value
        frmFYBX.lblLR.Caption = mod1.HTP.Fields("lRQ").Value
        frmFYBX.lblRq.Caption = mod1.HTP.Fields("QrQ").Value
        frmFYBX.txtQc.Text = mod1.HTP.Fields("QMin").Value
        frmFYBX.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
        frmFYBX.txtBz.Text = mod1.HTP.Fields("Bz").Value
        frmFYBX.lblBt.Caption = mod1.HTP.Fields("Fbt").Value

        frmFYBX.txtCwBZ.Text = mod1.HTP.Fields("CWBZ").Value
        frmFYBX.lblLc.Caption = mod1.HTP.Fields("LC").Value
        frmFYBX.lblLcRen.Caption = mod1.HTP.Fields("LCren").Value
        frmFYBX.lblLcUid.Caption = mod1.HTP.Fields("LCuid").Value
        frmFYBX.lblYwy.Caption = mod1.HTP.Fields("ywy").Value  '����������
        frmFYBX.lblUid.Caption = mod1.HTP.Fields("Uid").Value
        frmFYBX.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value '��ǰ��ӦNewFuwu���ID
        Lcou = mod1.HTP.Fields("Lcou").Value '��������
        frmFYBX.lblYqf.Caption = mod1.HTP.Fields("yqf").Value  'ҵ����˵ĸ���Ա�Ƿ�ǩ��
        frmFYBX.lblGui.Caption = mod1.HTP.Fields("GRen").Value '������
        frmFYBX.lblGuid.Caption = mod1.HTP.Fields("Grid").Value
        frmFYBX.optFp1.Value = mod1.HTP.Fields("fp").Value
        frmFYBX.lblNewF.Caption = mod1.HTP.Fields("newF").Value
        If mod1.HTP.Fields("czf").Value = True Then '�Ƿ���ʾ����ǩ��
            frmFYBX.frmZQ.Visible = True
            frmFYBX.cmdFQ.Caption = mod1.HTP.Fields("zjin").Value
            frmFYBX.lblFT.Caption = mod1.HTP.Fields("tc").Value
        Else
            frmFYBX.frmZQ.Visible = False
        End If
        If frmFYBX.optFp1.Value = False Then
            frmFYBX.optFp2.Value = True
            frmFYBX.txtFP.Text = mod1.HTP.Fields("fpnr").Value
        End If
        '���ϰ���°���ʾ��ͬ��ǩ�ְ�ť
        If IsNull(mod1.HTP.Fields("Lcou").Value) = True Then
            frmFYBX.frmQm.Visible = True
            frmFYBX.cmdBxr.Caption = mod1.HTP.Fields("yWy").Value
            frmFYBX.cmdJc.Caption = mod1.HTP.Fields("Jian").Value
            frmFYBX.cmdJl.Caption = mod1.HTP.Fields("JinLi").Value
            frmFYBX.cmdZj.Caption = mod1.HTP.Fields("zJin").Value
            frmFYBX.lblTa.Caption = mod1.HTP.Fields("ta").Value
            frmFYBX.lblTb.Caption = mod1.HTP.Fields("tb").Value
            frmFYBX.lblTC.Caption = mod1.HTP.Fields("tc").Value
            frmFYBX.lblTd.Caption = mod1.HTP.Fields("td").Value
        Else                                           '�°�
            frmFYBX.frmNewQ.Visible = True
            'Call ModBx.AddLcBut(mod1.HTP.Fields("Nlb").Value)

            tt = "FydQmOpen('" & frmFYBX.lblBh.Caption & "'," & 23 & ")" '23ΪworkBl�еı�����������
            mod1.HTT.Close
            mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
            mod1.HTT.MoveFirst
            For oo = 0 To mod1.HTT.RecordCount - 1
                frmFYBX.cmdQm(oo).Caption = mod1.HTT.Fields("QRen").Value
                frmFYBX.lblTm(oo).Caption = mod1.HTT.Fields("QRQ").Value
                mod1.HTT.MoveNext
            Next
        End If
        If frmFYBX.txtQc.Text <> "" Then
            frmFYBX.txtQc.PasswordChar = ""
            frmFYBX.txtQc.Enabled = False
        Else
            frmFYBX.txtQc.PasswordChar = "*"
            frmFYBX.txtQc.Enabled = True
        End If
        If Val(frmFYBX.lblBh.Caption) > 124571 Then
            frmFYBX.frmG.Visible = True
        End If

        
        '�򿪷����ܱ�
    tt = "FydMxOpen(" & Val(Bxid) & ")"
 
    Call ModBx.dtgKj(frmFYBX.lblNlb.Caption)
    If IsNull(mod1.HTP.Fields("lcou").Value) = True And frmFYBX.lblNlb.Caption = 9 Then '�ϰ��з��ݲ���
        frmFYBX.dtgBx.Columns("���ݲ���").Visible = True
        frmFYBX.dtgBx.Columns("����").Visible = True
        frmFYBX.dtgBx.Columns("����").Visible = True
        frmFYBX.dtgBx.Columns("������").Visible = True
    End If
        frmFYBX.cmdAdd.Visible = False
        frmFYBX.cmdDel.Visible = False
        frmFYBX.cmdSave.Enabled = False
        frmFYBX.dtgBx.AllowUpdate = False
        
       
    If mod1.HTP.Fields("lc") = 1 Or mod1.HTP.Fields("lc") = 0 Then   '����ǳ��ο���
        frmFYBX.cmdMod.Enabled = True
        frmFYBX.adoF2.Recordset.Close
        frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
    Else
        frmFYBX.cmdMod.Enabled = False
        frmFYBX.adoF2.Recordset.Close
        frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    End If

        Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
        tt = "Select atime as ����,khmc as ��������,sj as ����,fwbt as ���ݲ���,lyf as ���η�,gwf as ���·�,txf as ͨ�ŷ�,njtf as ���ڽ�ͨ��,wjtf as ���⽻ͨ��," & _
        "tcf as ͣ����,clf as ������,yf as �˷�,zcf as ס�޷�,bmtd as �����Ŷӷ�,cf as �ͷ�,ZDF as �д���,LPF as ��Ʒ��,fz as ����,WYF as ��ҵ��," & _
        "sd as ˮ��,DW as �绰,BGYP as �칫��Ʒ,YZ as ����,SZTG as �г��ƹ�,RYZP as ��Ա��Ƹ,KDF as ��ݷ�,PXF as ��ѵ��,CWSX as ����������,TDJS as �Ŷӽ����," & _
        "GTCF as ����ͣ����,GCLF as ����������,gg as ����,yH as �׺�,wl as ����,qtf as ������,gjj as ������,zhbx as �ۺϱ���,jtbt as ��ͨ����,zwbt as פ�����,gwbt as ��λ����,bm as ����,qy as ����,ywy as ����," & _
        "bid,gzdh as ���⳵ע��,xg as �ϼ�,qrq as ǩ������,GongF,GBM from fyBx where Bxid=" & Val(Bxid) & " order by bm,bid"
        frmFYBX.Fmx.Close
        frmFYBX.Fmx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Call ModBx.DiZ










        
    If (mod1.DName = "�ľ�" Or mod1.DName = "�Ǽ���") And frmFYBX.lblLc.Caption > 1 Then
        frmFYBX.txtCwBZ.Enabled = True
        frmFYBX.txtCwBZ.Locked = False
        frmFYBX.cmdSave.Enabled = True
        frmFYBX.txtBz.Locked = True
    End If
    
        '�����̰�ť.
        Call OpenAN
    'If mod1.Bq2 = True And frmFYBX.txtQM = "" And frmFYBX.lblLcRen.Caption = mod1.DName Then
    'If frmFYBX.lblLc.Caption = Lcou Then '��������������,���������ǩ��
    If mod1.Bq2 = True Then
        frmFYBX.txtQc.Enabled = True
    Else
        frmFYBX.txtQc.Enabled = False
    End If
    If Val(frmFYBX.lblNlb.Caption) = 79 Then
        frmFYBX.cmdMod.Enabled = True
    End If
    frmFYBX.frmEd.Visible = False
    frmFYBX.cmdG.Visible = False
    Call frmFYBX.QMBound(Val(Bxid))
End Sub
















Public Sub FyQing() 'Ӫ�������������
Dim oo As Integer
On Error Resume Next
    frmFYBX.frmNewQ.Visible = False
    frmFYBX.lblBh.Caption = ""
    frmFYBX.comQy.Caption = "�Ϻ�"
    frmFYBX.txtHg.Text = ""
    frmFYBX.lblDx.Caption = ""
    frmFYBX.lblFR.Caption = ""
    frmFYBX.lblLR.Caption = ""
    frmFYBX.lblRq.Caption = ""
    frmFYBX.cmdBxr.Caption = ""
    frmFYBX.cmdJc.Caption = ""
    frmFYBX.cmdJl.Caption = ""
    frmFYBX.cmdZj.Caption = ""
    frmFYBX.comDQ.Text = ""
    frmFYBX.txtQc.Text = ""
    frmFYBX.txtCwBZ.Text = ""
    frmFYBX.txtBz.Text = ""
    frmFYBX.lblTa.Caption = ""
    frmFYBX.lblTb.Caption = ""
    frmFYBX.lblTC.Caption = ""
    frmFYBX.lblTd.Caption = ""
    frmFYBX.LblTrq.Caption = ""
    frmFYBX.lblNlb.Caption = ""
    frmFYBX.frmQm.Visible = False
    frmFYBX.frmNewQ.Visible = False
    frmFYBX.frmYf.Visible = False
    frmFYBX.frmWd.Visible = False
    frmFYBX.lblLc.Caption = ""
    frmFYBX.lblLcRen.Caption = ""
    frmFYBX.lblLcUid.Caption = ""
    frmFYBX.lblBt.Caption = ""
    For oo = 5 To 0 Step -1
        Unload frmFYBX.lblQM(oo)
        Unload frmFYBX.cmdQm(oo)
        Unload frmFYBX.lblTm(oo)
    Next
    frmFYBX.lblQM(0).Caption = "������"
    frmFYBX.cmdQm(0).Caption = ""
    frmFYBX.lblTm(0).Caption = ""
    frmFYBX.txtCwBZ.Enabled = False '����עֻ���ڲ������ʱ�ܱ༭
    frmFYBX.lblYwy.Caption = "" '����������
    frmFYBX.lblUid.Caption = ""
    frmFYBX.lblFwid.Caption = "" '��ǰ��ӦNewFuwu���ID
    frmFYBX.lblYqf.Caption = "" 'ҵ����˵ĸ���Ա�Ƿ�ǩ��
    frmFYBX.frmRen.Visible = False
    frmFYBX.lblGui.Caption = ""
    frmFYBX.lblGuid.Caption = ""
    frmFYBX.cmdGui.Visible = False
    frmFYBX.cmdDao.Visible = False
    frmFYBX.optFp1.Value = False
    frmFYBX.optFp2.Value = False
    frmFYBX.txtFP.Text = ""
    frmFYBX.lblBid.Caption = ""
    frmFYBX.lblNewF.Caption = ""
    frmFYBX.lblTx.Visible = False
    frmFYBX.lblGZDH.Visible = False
    frmFYBX.txtGZDH.Visible = False
    frmFYBX.frmZQ.Visible = False
    frmFYBX.cmdFQ.Caption = ""
    frmFYBX.lblFT.Caption = ""
    frmFYBX.lbl1.Caption = "" '��������
    frmFYBX.lbl2.Caption = "" '���˷���
    frmFYBX.frmG.Visible = False
    frmFYBX.txtBm.Text = ""
    Call frmFYBX.dtgPFF
End Sub
Public Sub AddLcBut(Nlb As Integer)  '�������ǩ�ְ�ť
Dim tt As String
Dim oo As Integer
On Error Resume Next
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    mod1.HTP.MoveNext '��һ�����鰴ť�������,����,������һ��¼
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmFYBX.lblQM(oo)
        Load frmFYBX.cmdQm(oo)
        Load frmFYBX.lblTm(oo)
        frmFYBX.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmFYBX.lblQM(oo).Visible = True
        frmFYBX.lblQM(oo).Left = frmFYBX.lblQM(oo - 1).Left + 1100
        frmFYBX.cmdQm(oo).Caption = ""
        frmFYBX.lblTm(oo).Caption = ""
        frmFYBX.cmdQm(oo).Visible = True
        frmFYBX.lblTm(oo).Visible = True
        frmFYBX.cmdQm(oo).Left = frmFYBX.cmdQm(oo - 1).Left + 1100
        frmFYBX.lblTm(oo).Left = frmFYBX.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
    Next

'��ӽ�QMRZ��
'tt = "QMrzOpen('����')"
'mod1.HTT.Close
'mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
'mod1.HTP.MoveFirst
'Do While Not mod1.HTP.EOF
'    mod1.HTT.AddNew "Qlabel", mod1.HTP.Fields("LNR").Value
'    mod1.HTT.Update "BTZ", 23  '������
'    mod1.HTT.Update "QDBh", frmFYBX.lblBh.Caption '���
'    mod1.HTT.Update "Zid", mod1.HTP.Fields("zid").Value '˳��
'    If mod1.HTP.Fields("mid").Value = 38 Or mod1.HTP.Fields("mid").Value = 43 Or _
'       mod1.HTP.Fields("mid").Value = 48 Then                                    '�Ƿ�Ϊҵ�������ϸǩ��
'       mod1.HTT.Update "MXQF", 1
'    End If
'    mod1.HTT.UpdateBatch
'    mod1.HTP.MoveNext
'Loop
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "QMRZAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@NLb") = Nlb
        mod1.cmd.Parameters("@btz") = mod1.BTZ
        mod1.cmd.Parameters("@QDBH") = frmFYBX.lblBh.Caption '���
        mod1.cmd.Execute
        Set cmd = Nothing
'        If Nlb = 79 Then
'            frmFYBX.lblQM(0).Caption = "������"
'        End If
End Sub

Public Sub OpenAN()
Dim tt As String
Dim oo As Integer
On Error Resume Next
    For oo = 10 To 1 Step -1
        Unload frmFYBX.cmdQm(oo)
        Unload frmFYBX.lblQM(oo)
        Unload frmFYBX.lblTm(oo)
    Next

      'tt = "qmrzOpen(" & mod1.BTZ & ",'" & frmFYBX.lblBh.Caption & "')"
      tt = "qmrzOpen(23,'" & frmFYBX.lblBh.Caption & "')"
      Set mod1.HTP = CreateObject("adodb.recordset")
      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
      mod1.HTP.MoveFirst
      frmFYBX.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
        If mod1.HTP.Fields("xf").Value = True Then
            frmFYBX.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
            frmFYBX.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value
        Else
            frmFYBX.cmdQm(0).Caption = ""
            frmFYBX.lblTm(0).Caption = ""
        End If
      frmFYBX.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
      mod1.HTP.MoveNext
      For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmFYBX.lblQM(oo)
        frmFYBX.lblQM(oo).Caption = ""
        Load frmFYBX.cmdQm(oo)
        frmFYBX.cmdQm(oo).Caption = ""
        Load frmFYBX.lblTm(oo)
        frmFYBX.lblTm(oo).Caption = ""
        frmFYBX.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
        If mod1.HTP.Fields("xf").Value = True Then
            frmFYBX.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
            If frmFYBX.cmdQm(oo).Caption = "�Ͼ��쾭��" Then
                frmFYBX.cmdQm(oo).Caption = "�Ͼ��쾭��"
            End If
            frmFYBX.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value
        End If

        frmFYBX.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
        frmFYBX.lblQM(oo).Visible = True
        frmFYBX.cmdQm(oo).Visible = True
        frmFYBX.lblTm(oo).Visible = True
        frmFYBX.lblQM(oo).Left = frmFYBX.lblQM(oo - 1).Left + 1100
        frmFYBX.cmdQm(oo).Left = frmFYBX.cmdQm(oo - 1).Left + 1100
        frmFYBX.lblTm(oo).Left = frmFYBX.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
        
     Next
End Sub

Public Sub DiZ()
Dim oo As Integer
Dim rr As Integer
Dim F1 As Single '�������úϼ�
Dim F2 As Single '���˷��úϼ�
On Error Resume Next
F1 = 0: F2 = 0
        frmFYBX.Fmx.Requery
        'Set frmFYBX.dtgNx.DataSource = frmFYBX.Fmx
        
If frmFYBX.Fmx.RecordCount = 0 Then
    Set frmFYBX.dtgNx.DataSource = frmFYBX.Fmx
    frmFYBX.dtgNx.Rows = 2
    frmFYBX.dtgNx.FixedRows = 0
    frmFYBX.dtgNx.FixedRows = 1

Else
    frmFYBX.dtgNx.Rows = 2
    frmFYBX.dtgNx.FixedRows = 1
    Set frmFYBX.dtgNx.DataSource = frmFYBX.Fmx
End If
        '��ʾ��ֵ�ֶ�
        For oo = 3 To 40
            frmFYBX.dtgNx.ColWidth(oo) = 0
        Next


        For oo = 3 To 40
            rr = 1
            frmFYBX.dtgNx.Col = oo
            Do While Not rr >= frmFYBX.dtgNx.Rows
                frmFYBX.dtgNx.Row = rr
                If Val(frmFYBX.dtgNx.Text) > 0 Then
                    frmFYBX.dtgNx.ColWidth(oo) = 1000
                    
                    frmFYBX.dtgNx.Col = 48
                    If Val(frmFYBX.dtgNx.Text) = 1 Then
                        frmFYBX.dtgNx.Col = oo
                        frmFYBX.dtgNx.CellForeColor = &HFF&
                        F1 = F1 + Val(frmFYBX.dtgNx.Text)
                    ElseIf Val(frmFYBX.dtgNx.Text) = 2 Then
                        frmFYBX.dtgNx.Col = oo
                        frmFYBX.dtgNx.CellForeColor = &HC00000
                        F2 = F2 + Val(frmFYBX.dtgNx.Text)
                    End If
                    
                    'Exit Do
                End If
                rr = rr + 1
            Loop
        Next
        frmFYBX.lbl1.Caption = F2: frmFYBX.lbl2.Caption = F1
'        If frmFYBX.dtgNx.ColWidth(3) = 1005 Or frmFYBX.dtgNx.ColWidth(36) = 1005 Then
'            frmFYBX.dtgNx.ColWidth(36) = 1000
'            frmFYBX.dtgNx.ColWidth(3) = 0
'        End If
        frmFYBX.dtgNx.FixedRows = 0
        frmFYBX.dtgNx.MergeCol(1) = True
        frmFYBX.dtgNx.MergeCol(2) = True
        frmFYBX.dtgNx.MergeCol(41) = True
        frmFYBX.dtgNx.MergeCol(42) = True
        frmFYBX.dtgNx.MergeCol(43) = True
        frmFYBX.dtgNx.MergeCells = 3
        frmFYBX.dtgNx.FixedRows = 1
        'If frmFYBX.lblBm.Caption = "���̲�" Then
            frmFYBX.dtgNx.ColWidth(45) = 1000
        'Else
            'frmFYBX.dtgNx.ColWidth(41) = 0
        'End If
        If frmFYBX.lblNlb.Caption = 35 Then
            frmFYBX.dtgNx.ColWidth(45) = 0
            'frmFYBX.dtgNx.ColWidth(40) = 1000
        Else
            frmFYBX.dtgNx.ColWidth(40) = 0
        End If
        
        
End Sub
