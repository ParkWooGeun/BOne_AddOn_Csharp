namespace PSH_BOne_AddOn
{
    /// <summary>
    /// �̺�Ʈ ����
    /// </summary>
    internal static class PSH_SetFilter
    {
        public static void Execute()
		{
			SAPbouiCOM.EventFilters oFilters = null;
			SAPbouiCOM.EventFilter oFilter = null;

			oFilters = new SAPbouiCOM.EventFilters();

			ITEM_PRESSED(ref oFilter, ref oFilters); //1
            KEY_DOWN(ref oFilter, ref oFilters); //2
			GOT_FOCUS(ref oFilter, ref oFilters); //3
			LOST_FOCUS(ref oFilter, ref oFilters); //4
			COMBO_SELECT(ref oFilter, ref oFilters); //5
			CLICK(ref oFilter, ref oFilters); //6
			DOUBLE_CLICK(ref oFilter, ref oFilters); //7
			MATRIX_LINK_PRESSED(ref oFilter, ref oFilters); //8
			//MATRIX_COLLAPSE_PRESSED(oFilter, oFilters); //9
			VALIDATE(ref oFilter, ref oFilters); //10
			MATRIX_LOAD(ref oFilter, ref oFilters); //11
			//DATASOURCE_LOAD(oFilter, oFilters); //12
			FORM_LOAD(ref oFilter, ref oFilters); //16
			FORM_UNLOAD(ref oFilter, ref oFilters); //17
			//FORM_ACTIVATE(oFilter, oFilters); //18
			//FORM_DEACTIVATE(oFilter, oFilters); //19
			//FORM_CLOSE(oFilter, oFilters); //20
			FORM_RESIZE(ref oFilter, ref oFilters); //21
			//FORM_KEY_DOWN(oFilter, oFilters); //22
			//FORM_MENU_HILIGHT(oFilter, oFilters); //23
			//PRINT(oFilter, oFilters); //24
			//PRINT_DATA(oFilter, oFilters); //25
			CHOOSE_FROM_LIST(ref oFilter, ref oFilters); //27
			RIGHT_CLICK(ref oFilter, ref oFilters); //28
			MENU_CLICK(ref oFilter, ref oFilters); //32
			FORM_DATA_ADD(ref oFilter, ref oFilters); //33
			FORM_DATA_UPDATE(ref oFilter, ref oFilters); //34
			//FORM_DATA_DELETE(oFilter, oFilters); //35
			FORM_DATA_LOAD(ref oFilter, ref oFilters); //36

			//Setting the application with the EventFilters object
			PSH_Globals.SBO_Application.SetFilter(oFilters);

			oFilter = null;
			oFilters = null;
		}

		private static void ITEM_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
			//System Form Type

			//�λ����
			//�޿�����

			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

            //�λ����
            oFilter.AddEx("PH_PY001"); //�λ縶���� ���
            oFilter.AddEx("PH_PY002"); //���½ð����� ���
            oFilter.AddEx("PH_PY003"); //���¿��¼���
            oFilter.AddEx("PH_PY004"); //�ٹ��������
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY006"); //��ȣ�۾����
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY009"); //�����ڷ�UPLOAD
			oFilter.AddEx("PH_PY010"); //���ϱ���ó��
			oFilter.AddEx("PH_PY011"); //������ ȣĪ �ϰ� ����(2013.07.05 �۸�� �߰�)
			oFilter.AddEx("PH_PY012"); //������
			oFilter.AddEx("PH_PY013"); //�����ϼ����
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY015"); //������ġ���
			oFilter.AddEx("PH_PY016"); //�⺻�������
			oFilter.AddEx("PH_PY017"); //����������
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)
			oFilter.AddEx("PH_PY019"); //�ݺ�����
			oFilter.AddEx("PH_PY020"); //�ϱ��� ����������
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����
			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY202"); //�����ӹ��� �ް���� ��ȸ
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�λ� - ����Ʈ
			oFilter.AddEx("PH_PY501"); //���ǹ߱���Ȳ
			oFilter.AddEx("PH_PY505"); //�Ի��ڴ���
			oFilter.AddEx("PH_PY510"); //������
			oFilter.AddEx("PH_PY515"); //�����ڻ�����
			oFilter.AddEx("PH_PY520"); //���������������ڴ���
			oFilter.AddEx("PH_PY525"); //�зº��ο���Ȳ
			oFilter.AddEx("PH_PY530"); //���ɺ��ο���Ȳ
			oFilter.AddEx("PH_PY535"); //�ټӳ�����ο���Ȳ
			oFilter.AddEx("PH_PY540"); //�ο���Ȳ(��ܿ�)
			oFilter.AddEx("PH_PY545"); //�ο���Ȳ(�볻��)
			oFilter.AddEx("PH_PY550"); //��ü�ο���Ȳ
			oFilter.AddEx("PH_PY555"); //���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY560"); //�������Ȳ
			oFilter.AddEx("PH_PY565"); //����ٹ�����Ȳ
			oFilter.AddEx("PH_PY570"); //����/���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY575"); //���±�����Ȳ
			oFilter.AddEx("PH_PY580"); //���κ����¿���
			oFilter.AddEx("PH_PY585"); //������ٱ�Ϻ�
			oFilter.AddEx("PH_PY590"); //�Ⱓ����������ǥ
			oFilter.AddEx("PH_PY595"); //�ټӳ����Ȳ
			oFilter.AddEx("PH_PY600"); //���ں�����ٹ���Ȳ
			oFilter.AddEx("PH_PY605"); //�ټӺ����ް��߻��׻�볻��
			oFilter.AddEx("PH_PY610"); //���±��к���볻��
			oFilter.AddEx("PH_PY615"); //�����ٹ���Ȳ
			oFilter.AddEx("PH_PY620"); //���������ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY635"); //����,��������Ȳ
			oFilter.AddEx("PH_PY640"); //���ο���������ȯ����Ȳ
			oFilter.AddEx("PH_PY645"); //�ڰݼ���������Ȳ
			oFilter.AddEx("PH_PY650"); //�뵿���հ�����Ȳ
			oFilter.AddEx("PH_PY655"); //���ƴ������Ȳ
			oFilter.AddEx("PH_PY660"); //��ֱٷ�����Ȳ
			oFilter.AddEx("PH_PY665"); //����ڳ���Ȳ
			oFilter.AddEx("PH_PY670"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY675"); //�ٹ�����Ȳ
			oFilter.AddEx("PH_PY676"); //���½ð�������ȸ
			oFilter.AddEx("PH_PY677"); //���ϱ����̻�����ȸ
			oFilter.AddEx("PH_PY679"); //���κ� �������� ��ȸ
			oFilter.AddEx("PH_PY680"); //�����Ȳ
			oFilter.AddEx("PH_PY685"); //���󰡱���Ȳ
			oFilter.AddEx("PH_PY690"); //��������Ȳ
			oFilter.AddEx("PH_PY695"); //�λ���ī��
			oFilter.AddEx("PH_PY705"); //��������ޱ���Ȯ��
			oFilter.AddEx("PH_PY860"); //ȣ��ǥ��ȸ
			oFilter.AddEx("PH_PY503"); //��������ڸ��
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���
			oFilter.AddEx("PH_PY507"); //��������Ȳ
			oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ
			oFilter.AddEx("PH_PY935"); //�����ȣ��Ȳ
			oFilter.AddEx("PH_PY551"); //����ο���ȸ
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY522"); //�ӱ���ũ�������Ȳ
			oFilter.AddEx("PH_PY523"); //�ӱ���ũ����ڿ���������Ȳ
			oFilter.AddEx("PH_PY524"); //������ �߰� ���곻��
			oFilter.AddEx("PH_PY683"); //����ٹ�������Ȳ
			oFilter.AddEx("PH_PYA65"); //������Ȳ (����)
			oFilter.AddEx("PH_PY583"); //���κ� �������� ��ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY108"); //�������޼���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY117"); //�޻󿩸����۾�
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�
			oFilter.AddEx("PH_PY119"); //�޻��������ϻ���
			oFilter.AddEx("PH_PY120"); //�޻󿩼ұ�����ó��
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY125"); //�������� ����
			oFilter.AddEx("PH_PY127"); //���κ� 4�뺸�� �������� �� ����ݾ��Է�
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

			//�޿����� - ����Ʈ
			oFilter.AddEx("PH_PY625"); //��Ź�ڸ��
			oFilter.AddEx("PH_PY630"); //����������������Ȳ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿��μ����������
			oFilter.AddEx("PH_PY720"); //�󿩺μ����������
			oFilter.AddEx("PH_PY725"); //�޿����޺��������
			oFilter.AddEx("PH_PY740"); //�����޺��������
			oFilter.AddEx("PH_PY730"); //�޿��������
			oFilter.AddEx("PH_PY735"); //�󿩺������
			oFilter.AddEx("PH_PY745"); //����������Ȳ
			oFilter.AddEx("PH_PY750"); //�ٷμҵ�¡����Ȳ
			oFilter.AddEx("PH_PY755"); //��ȣȸ������Ȳ
			oFilter.AddEx("PH_PY760"); //����ӱݹ������ݻ��⳻����
			oFilter.AddEx("PH_PY765"); //�޿�����������
			oFilter.AddEx("PH_PY770"); //�����ҵ��õ¡�����������
			oFilter.AddEx("PH_PY775"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY776"); //�ܿ�������Ȳ
			oFilter.AddEx("PH_PY780"); //����뺸�賻��
			oFilter.AddEx("PH_PY785"); //�����ο��ݳ���
			oFilter.AddEx("PH_PY790"); //���ǰ����賻��
			oFilter.AddEx("PH_PY795"); //�����μ����޿�����
			oFilter.AddEx("PH_PY800"); //�ΰǺ������ڷ�
			oFilter.AddEx("PH_PY805"); //�޿����纯������
			oFilter.AddEx("PH_PY810"); //���޺�����ӱݳ���
			oFilter.AddEx("PH_PY815"); //����ӱݳ���
			oFilter.AddEx("PH_PY820"); //����ӱݳ���
			oFilter.AddEx("PH_PY825"); //������O/T��Ȳ
			oFilter.AddEx("PH_PY830"); //�μ����ΰǺ���Ȳ (��ȹ)
			oFilter.AddEx("PH_PY835"); //���޺�O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY840"); //ǳ�����ڰ����ڷ�
			oFilter.AddEx("PH_PY845"); //�Ⱓ���޿����޳���
			oFilter.AddEx("PH_PY850"); //�ұ޺����޸���
			oFilter.AddEx("PH_PY855"); //���κ��ӱ����޴���
			oFilter.AddEx("PH_PY865"); //��뺸����Ȳ (����)
			oFilter.AddEx("PH_PY870"); //��纰��O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY875"); //���޺������������
			oFilter.AddEx("PH_PY716"); //�Ⱓ���޿��μ����������
			oFilter.AddEx("PH_PY721"); //�Ⱓ���󿩺μ����������
			oFilter.AddEx("PH_PY717"); //�Ⱓ���޿��ݺ��������
			oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
			oFilter.AddEx("PH_PY701");//�޿����޴��� (������)

			oFilter.AddEx("PH_PYA10"); //�޿����޴���(�μ�)
			oFilter.AddEx("PH_PYA20"); //�޿��μ����������(�μ�)
			oFilter.AddEx("PH_PYA30"); //�����޴���(�μ�)
			oFilter.AddEx("PH_PYA40"); //�󿩺μ����������(�μ�)
			oFilter.AddEx("PH_PYA50"); //DC��ȯ�ںδ�����޳���
			oFilter.AddEx("PH_PYA75"); //�����ܼ������޴���
            oFilter.AddEx("PH_PYA80"); //�ٹ��ð�ǥ���
            oFilter.AddEx("PH_PYA90"); //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
            oFilter.AddEx("PH_PY526"); //�ӱ���ũ�ο���Ȳ
            oFilter.AddEx("PH_PY999"); //����ڱ��Ѱ���
			oFilter.AddEx("PH_PY998"); //����ڱ�����ȸ

			//�������
			oFilter.AddEx("PH_PY401"); //���ٹ������
			oFilter.AddEx("PH_PY402"); //��������ڷ� ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���
			oFilter.AddEx("PH_PY415"); //������
			oFilter.AddEx("PH_PY417"); //���� �������ϻ���
			oFilter.AddEx("PH_PY980"); //�Ű�_�ٷμҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ű�_�Ƿ�����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //�Ű�_��αݸ����ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�Ű�_�����ҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY419"); //ǥ�ؼ����������ڵ��

			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��
			oFilter.AddEx("ZPY504"); //��������ȸ
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY507"); //��������ȸ(��ü)
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�
			oFilter.AddEx("ZPY510"); //�����ٹ��� �ϰ�����
			oFilter.AddEx("ZPY521"); //�ٷμҵ������ü����
			oFilter.AddEx("ZPY522"); //�Ƿ�� ��α� �����ü����

			oFilter.AddEx("RPY401"); //������õ¡�� ������
			oFilter.AddEx("RPY501"); //�����ڷ���Ȳ
			oFilter.AddEx("RPY502"); //�����ٹ�����Ȳ
			oFilter.AddEx("RPY503"); //�ٷμҵ� ��õ¡����
			oFilter.AddEx("RPY504"); //�ٷμҵ� ��õ������
			oFilter.AddEx("RPY505"); //�ҵ��ڷ�����ǥ
			oFilter.AddEx("RPY506"); //����¡��ȯ�޴���
			oFilter.AddEx("RPY508"); //������������ǥ
			oFilter.AddEx("RPY509"); //���ټ��Ű����ǥ
			oFilter.AddEx("RPY510"); //������ٷμҵ����
			oFilter.AddEx("RPY511");//��αݸ���

			//��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY302"); //���ڱ����޿Ϸ�ó��
			oFilter.AddEx("PH_PY303"); //���ڱ��������ϻ���
			oFilter.AddEx("PH_PY305"); //���ڱݽ�û��
			oFilter.AddEx("PH_PY306"); //���ڱݽ�û����(���κ�)
			oFilter.AddEx("PH_PY307"); //���ڱݽ�û����(�б⺰)
			oFilter.AddEx("PH_PY309"); //��αݵ��
			oFilter.AddEx("PH_PY310"); //��αݰ�����ȯ
			oFilter.AddEx("PH_PY311"); //��ٹ���������
			oFilter.AddEx("PH_PY312"); //������� ���κ����
			oFilter.AddEx("PH_PY313"); //��αݰ��
			oFilter.AddEx("PH_PY314"); //��αݰ�� ���� ��ȸ(�޿������ڷ��)
			oFilter.AddEx("PH_PY030"); //������
			oFilter.AddEx("PH_PY031"); //������
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY315"); //���κ���α��ܾ���Ȳ
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
			oFilter.AddEx("PS_DateChange"); //��¥ ���� ���
			oFilter.AddEx("PS_DateCommit"); //��¥ ����

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO600"); //�����繫��ǥ
			oFilter.AddEx("PS_CO606"); //�����繫��ǥ ���� ����
			oFilter.AddEx("PS_CO605"); //���ռ��Һ�
			oFilter.AddEx("PS_CO120"); //������ ���� ���
			oFilter.AddEx("PS_CO130"); //��ǰ�� ���� ���
			oFilter.AddEx("PS_CO131"); //������������Ȳ
			oFilter.AddEx("PS_CO510"); //����������������ȸ
			oFilter.AddEx("PS_CO520"); //��ǰ���� �����׸� ��ȸ
			oFilter.AddEx("PS_CO501"); //ǰ�񺰿������
			oFilter.AddEx("PS_CO502"); //ǰ����տ����׸���
		}

		private static void KEY_DOWN(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);

            //System Form Type
			//�λ����
			//�޿�����

			//�������
			//��Ÿ����

			//AddOn Form Type
			//�λ����
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY011"); //������ ȣĪ �ϰ� ����(2013.07.05 �۸�� �߰�)
			oFilter.AddEx("PH_PY012"); //������
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY015"); //������ġ���
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�(2015.05.18 �۸�� �߰�)
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����

			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ
			oFilter.AddEx("PH_PY695"); //�λ���ī��

			//���°��� - ����Ʈ
			oFilter.AddEx("PH_PY580"); //���α��¿���
			oFilter.AddEx("PH_PY575"); //�����Ȳ
            oFilter.AddEx("PH_PY680"); //���±�����Ȳ
            oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ

			//�޿�����
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���κ��������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY676"); //���½ð�������ȸ
			oFilter.AddEx("PH_PY677"); //���ϱ����̻�����ȸ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿����޴���(�μ�����)
			oFilter.AddEx("PH_PY720"); //�����޴���(�μ�����)
            oFilter.AddEx("PH_PY730"); //�޿�����
            oFilter.AddEx("PH_PY735"); //�󿩺���
            oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���

            //�޿�����-����Ʈ
            oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
            oFilter.AddEx("PH_PY701"); //�޿����޴��� (������)

            //�������
            oFilter.AddEx("PH_PY401"); //���ٹ������
			oFilter.AddEx("PH_PY402"); //��������ڷ� ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���
			oFilter.AddEx("PH_PY415"); //������
			oFilter.AddEx("PH_PY980"); //�Ű�_�ٷμҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ű�_�Ƿ�����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //�Ű�_��αݸ����ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�Ű�_�����ҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY419"); //ǥ�ؼ����������ڵ��

			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��
			oFilter.AddEx("ZPY504"); //��������ȸ
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�
			oFilter.AddEx("ZPY510"); //�����ٹ��� �ϰ�����
			oFilter.AddEx("ZPY521"); //�ٷμҵ������ü����
			oFilter.AddEx("ZPY522"); //�Ƿ�� ��α� �����ü����

			oFilter.AddEx("RPY401"); //������õ¡�� ������
			oFilter.AddEx("RPY501"); //�����ڷ���Ȳ
			oFilter.AddEx("RPY502"); //�����ٹ�����Ȳ
			oFilter.AddEx("RPY503"); //�ٷμҵ� ��õ¡����
			oFilter.AddEx("RPY504"); //�ٷμҵ� ��õ������
			oFilter.AddEx("RPY505"); //�ҵ��ڷ�����ǥ
			oFilter.AddEx("RPY506"); //����¡��ȯ�޴���
			oFilter.AddEx("RPY508"); //������������ǥ
			oFilter.AddEx("RPY509"); //���ټ��Ű����ǥ
			oFilter.AddEx("RPY510"); //������ٷμҵ����
			oFilter.AddEx("RPY511"); //��αݸ���

			//��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY305"); //���ڱݽ�û��
			oFilter.AddEx("PH_PY306"); //���ڱݽ�û����(���κ�)
			oFilter.AddEx("PH_PY309"); //��αݵ��
			oFilter.AddEx("PH_PY310"); //��αݰ�����ȯ
			oFilter.AddEx("PH_PY313"); //��αݰ��
			oFilter.AddEx("PH_PY314"); //��αݰ�� ���� ��ȸ(�޿������ڷ��)
			oFilter.AddEx("PH_PY030"); //������
			oFilter.AddEx("PH_PY031"); //������
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY315"); //���κ���α��ܾ���Ȳ
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����
			oFilter.AddEx("PS_DateCommit"); //��¥ ����
			oFilter.AddEx("PH_PY998"); //����ڱ�����ȸ

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO120"); //������ ���� ���
			oFilter.AddEx("PS_CO131"); //7.������������Ȳ
			oFilter.AddEx("PS_CO510"); // ������������Ȳ
			oFilter.AddEx("PS_CO520"); //��ǰ���� �����׸� ��ȸ
			oFilter.AddEx("PS_CO501"); //ǰ�񺰿������
			oFilter.AddEx("PS_CO502"); //ǰ����տ����׸���
		}

		private static void GOT_FOCUS(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);

			//System Form Type
			//�λ����
			//�޿�����

			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���� ���
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY004"); //�ٹ��������
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY006"); //��ȣ�۾����
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY009"); //�����ڷ�UPLOAD
			oFilter.AddEx("PH_PY011"); //������ ȣĪ �ϰ� ����(2013.07.05 �۸�� �߰�)
			oFilter.AddEx("PH_PY013"); //�����ϼ����
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY015"); //������ġ���
			oFilter.AddEx("PH_PY016"); //�⺻�������
			oFilter.AddEx("PH_PY017"); //����������
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)
			oFilter.AddEx("PH_PY019"); //�ݺ�����
			oFilter.AddEx("PH_PY020"); //�ϱ��� ����������

			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�λ� - ����Ʈ
			oFilter.AddEx("PH_PY501"); //���ǹ߱���Ȳ
			oFilter.AddEx("PH_PY505"); //�Ի��ڴ���
			oFilter.AddEx("PH_PY510"); //������
			oFilter.AddEx("PH_PY515"); //�����ڻ�����
			oFilter.AddEx("PH_PY520"); //���������������ڴ���
			oFilter.AddEx("PH_PY525"); //�зº��ο���Ȳ
			oFilter.AddEx("PH_PY530"); //���ɺ��ο���Ȳ
			oFilter.AddEx("PH_PY535"); //�ټӳ�����ο���Ȳ
			oFilter.AddEx("PH_PY540"); //�ο���Ȳ(��ܿ�)
			oFilter.AddEx("PH_PY545"); //�ο���Ȳ(�볻��)
			oFilter.AddEx("PH_PY550"); //��ü�ο���Ȳ
			oFilter.AddEx("PH_PY555"); //���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY560"); //�������Ȳ
			oFilter.AddEx("PH_PY565"); //����ٹ�����Ȳ
			oFilter.AddEx("PH_PY570"); //����/���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY575"); //���±�����Ȳ
			oFilter.AddEx("PH_PY580"); //���κ����¿���
			oFilter.AddEx("PH_PY585"); //������ٱ�Ϻ�
			oFilter.AddEx("PH_PY590"); //�Ⱓ����������ǥ
			oFilter.AddEx("PH_PY595"); //�ټӳ����Ȳ
			oFilter.AddEx("PH_PY600"); //���ں�����ٹ���Ȳ
			oFilter.AddEx("PH_PY605"); //�ټӺ����ް��߻��׻�볻��
			oFilter.AddEx("PH_PY610"); //���±��к���볻��
			oFilter.AddEx("PH_PY615"); //�����ٹ���Ȳ
			oFilter.AddEx("PH_PY620"); //���������ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY635"); //����,��������Ȳ
			oFilter.AddEx("PH_PY640"); //���ο���������ȯ����Ȳ
			oFilter.AddEx("PH_PY645"); //�ڰݼ���������Ȳ
			oFilter.AddEx("PH_PY650"); //�뵿���հ�����Ȳ
			oFilter.AddEx("PH_PY655"); //���ƴ������Ȳ
			oFilter.AddEx("PH_PY660"); //��ֱٷ�����Ȳ
			oFilter.AddEx("PH_PY665"); //����ڳ���Ȳ
			oFilter.AddEx("PH_PY670"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY675"); //�ٹ�����Ȳ
			oFilter.AddEx("PH_PY679"); //���κ� �������� ��ȸ
			oFilter.AddEx("PH_PY680"); //�����Ȳ
			oFilter.AddEx("PH_PY685"); //���󰡱���Ȳ
			oFilter.AddEx("PH_PY690"); //��������Ȳ
			oFilter.AddEx("PH_PY695"); //�λ���ī��
			oFilter.AddEx("PH_PY705"); //��������ޱ���Ȯ��
			oFilter.AddEx("PH_PY860"); //ȣ��ǥ��ȸ
			oFilter.AddEx("PH_PY503"); //��������ڸ��
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���
			oFilter.AddEx("PH_PY507"); //��������Ȳ
			oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ
			oFilter.AddEx("PH_PY935"); //�����ȣ��Ȳ
			oFilter.AddEx("PH_PY551"); //����ο���ȸ
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY522"); //�ӱ���ũ�������Ȳ
			oFilter.AddEx("PH_PY523"); //�ӱ���ũ����ڿ���������Ȳ
			oFilter.AddEx("PH_PY524"); //������ �߰� ���곻��
			oFilter.AddEx("PH_PY683"); //����ٹ�������Ȳ
			oFilter.AddEx("PH_PYA65"); //������Ȳ (����)
			oFilter.AddEx("PH_PY583"); //���κ� �������� ��ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY108"); //�������޼���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY117"); //�޻󿩸����۾�
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�
			oFilter.AddEx("PH_PY120"); //�޻󿩼ұ�����ó��
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY125"); //�������� ����
			oFilter.AddEx("PH_PY127"); //���κ� 4�뺸�� �������� �� ����ݾ��Է�
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���

			//�޿����� - ����Ʈ
			oFilter.AddEx("PH_PY625"); //��Ź�ڸ��
			oFilter.AddEx("PH_PY630"); //����������������Ȳ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿��μ����������
			oFilter.AddEx("PH_PY720"); //�󿩺μ����������
			oFilter.AddEx("PH_PY725"); //�޿����޺��������
			oFilter.AddEx("PH_PY740"); //�����޺��������
			oFilter.AddEx("PH_PY730"); //�޿��������
			oFilter.AddEx("PH_PY735"); //�󿩺������
			oFilter.AddEx("PH_PY745"); //����������Ȳ
			oFilter.AddEx("PH_PY750"); //�ٷμҵ�¡����Ȳ
			oFilter.AddEx("PH_PY755"); //��ȣȸ������Ȳ
			oFilter.AddEx("PH_PY760"); //����ӱݹ������ݻ��⳻����
			oFilter.AddEx("PH_PY765"); //�޿�����������
			oFilter.AddEx("PH_PY770"); //�����ҵ��õ¡�����������
			oFilter.AddEx("PH_PY775"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY776"); //�ܿ�������Ȳ
			oFilter.AddEx("PH_PY780"); //����뺸�賻��
			oFilter.AddEx("PH_PY785"); //�����ο��ݳ���
			oFilter.AddEx("PH_PY790"); //���ǰ����賻��
			oFilter.AddEx("PH_PY795"); //�����μ����޿�����
			oFilter.AddEx("PH_PY800"); //�ΰǺ������ڷ�
			oFilter.AddEx("PH_PY805"); //�޿����纯������
			oFilter.AddEx("PH_PY810"); //���޺�����ӱݳ���
			oFilter.AddEx("PH_PY815"); //����ӱݳ���
			oFilter.AddEx("PH_PY820"); //����ӱݳ���
			oFilter.AddEx("PH_PY825"); //������O/T��Ȳ
			oFilter.AddEx("PH_PY830"); //�μ����ΰǺ���Ȳ (��ȹ)
			oFilter.AddEx("PH_PY835"); //���޺�O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY840"); //ǳ�����ڰ����ڷ�
			oFilter.AddEx("PH_PY845"); //�Ⱓ���޿����޳���
			oFilter.AddEx("PH_PY850"); //�ұ޺����޸���
			oFilter.AddEx("PH_PY855"); //���κ��ӱ����޴���
			oFilter.AddEx("PH_PY865"); //��뺸����Ȳ (����)
			oFilter.AddEx("PH_PY870"); //��纰��O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY875"); //���޺������������
			oFilter.AddEx("PH_PY716"); //�Ⱓ���޿��μ����������
			oFilter.AddEx("PH_PY721"); //�Ⱓ���󿩺μ����������
			oFilter.AddEx("PH_PY717"); //�Ⱓ���޿��ݺ��������
			oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
			oFilter.AddEx("PH_PY701"); //�޿����޴��� (������)

			oFilter.AddEx("PH_PYA10"); //�޿����޴���(�μ�)
			oFilter.AddEx("PH_PYA20"); //�޿��μ����������(�μ�)
			oFilter.AddEx("PH_PYA30"); //�����޴���(�μ�)
			oFilter.AddEx("PH_PYA40"); //�󿩺μ����������(�μ�)
			oFilter.AddEx("PH_PYA50"); //DC��ȯ�ںδ�����޳���
			oFilter.AddEx("PH_PYA75"); //�����ܼ������޴���
            oFilter.AddEx("PH_PYA80"); //�ٹ��ð�ǥ���
            oFilter.AddEx("PH_PYA90"); //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
            oFilter.AddEx("PH_PY526"); //�ӱ���ũ�ο���Ȳ
            oFilter.AddEx("PH_PY999"); //����ڱ��Ѱ���
			oFilter.AddEx("PH_PY998"); //����ڱ�����ȸ

			//�������
			oFilter.AddEx("PH_PY401"); //���ٹ������
			oFilter.AddEx("PH_PY402"); //��������ڷ� ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���
			oFilter.AddEx("PH_PY415"); //������
			oFilter.AddEx("PH_PY980"); //�Ű�_�ٷμҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ű�_�Ƿ�����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //�Ű�_��αݸ����ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�Ű�_�����ҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY419"); //ǥ�ؼ����������ڵ��

			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��
			oFilter.AddEx("ZPY504"); //��������ȸ
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�
			oFilter.AddEx("ZPY521"); //�ٷμҵ������ü����
			oFilter.AddEx("ZPY522"); //�Ƿ�� ��α� �����ü����

			//��Ÿ����
			oFilter.AddEx("PH_PY311"); //��ٹ���������
			oFilter.AddEx("PH_PY312"); //������� ���κ����
			oFilter.AddEx("PH_PY309"); //��αݵ��
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
			oFilter.AddEx("PS_DateChange"); //��¥ ���� ���
			oFilter.AddEx("PS_DateCommit"); //��¥ ����

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO600"); //�����繫��ǥ
			oFilter.AddEx("PS_CO605"); //���ռ��Һ�
			oFilter.AddEx("PS_CO501"); //ǰ�񺰿������
			oFilter.AddEx("PS_CO502"); //ǰ����տ����׸���
		}

		private static void LOST_FOCUS(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);

			//System Form Type
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����
		}

		private static void COMBO_SELECT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);

			//System Form Type
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���� ���
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY004"); //�ٹ��������
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY006"); //��ȣ�۾����
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY009"); //�����ڷ�UPLOAD
			oFilter.AddEx("PH_PY012"); //������
			oFilter.AddEx("PH_PY013"); //�����ϼ����
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY016"); //�⺻�������
			oFilter.AddEx("PH_PY017"); //����������
			oFilter.AddEx("PH_PY018"); //���ϱٹ���üũ(������)
			oFilter.AddEx("PH_PY019"); //�ݺ�����
			oFilter.AddEx("PH_PY020"); //�ϱ��� ����������
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����
			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�λ� - ����Ʈ
			oFilter.AddEx("PH_PY501"); //���ǹ߱���Ȳ
			oFilter.AddEx("PH_PY505"); //�Ի��ڴ���
			oFilter.AddEx("PH_PY510"); //������
			oFilter.AddEx("PH_PY515"); //�����ڻ�����
			oFilter.AddEx("PH_PY520"); //���������������ڴ���
			oFilter.AddEx("PH_PY525"); //�зº��ο���Ȳ
			oFilter.AddEx("PH_PY530"); //���ɺ��ο���Ȳ
			oFilter.AddEx("PH_PY535"); //�ټӳ�����ο���Ȳ
			oFilter.AddEx("PH_PY540"); //�ο���Ȳ(��ܿ�)
			oFilter.AddEx("PH_PY545"); //�ο���Ȳ(�볻��)
			oFilter.AddEx("PH_PY550"); //��ü�ο���Ȳ
			oFilter.AddEx("PH_PY555"); //���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY560"); //�������Ȳ
			oFilter.AddEx("PH_PY565"); //����ٹ�����Ȳ
			oFilter.AddEx("PH_PY570"); //����/���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY575"); //���±�����Ȳ
			oFilter.AddEx("PH_PY580"); //���κ����¿���
			oFilter.AddEx("PH_PY585"); //������ٱ�Ϻ�
			oFilter.AddEx("PH_PY590"); //�Ⱓ����������ǥ
			oFilter.AddEx("PH_PY595"); //�ټӳ����Ȳ
			oFilter.AddEx("PH_PY600"); //���ں�����ٹ���Ȳ
			oFilter.AddEx("PH_PY605"); //�ټӺ����ް��߻��׻�볻��
			oFilter.AddEx("PH_PY610"); //���±��к���볻��
			oFilter.AddEx("PH_PY615"); //�����ٹ���Ȳ
			oFilter.AddEx("PH_PY620"); //���������ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY635"); //����,��������Ȳ
			oFilter.AddEx("PH_PY640"); //���ο���������ȯ����Ȳ
			oFilter.AddEx("PH_PY645"); //�ڰݼ���������Ȳ
			oFilter.AddEx("PH_PY650"); //�뵿���հ�����Ȳ
			oFilter.AddEx("PH_PY655"); //���ƴ������Ȳ
			oFilter.AddEx("PH_PY660"); //��ֱٷ�����Ȳ
			oFilter.AddEx("PH_PY665"); //����ڳ���Ȳ
			oFilter.AddEx("PH_PY670"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY675"); //�ٹ�����Ȳ
			oFilter.AddEx("PH_PY676"); //���½ð�������ȸ
			oFilter.AddEx("PH_PY677"); //���ϱ����̻�����ȸ
			oFilter.AddEx("PH_PY679"); //���κ� �������� ��ȸ
			oFilter.AddEx("PH_PY680"); //�����Ȳ
			oFilter.AddEx("PH_PY685"); //���󰡱���Ȳ
			oFilter.AddEx("PH_PY690"); //��������Ȳ
			oFilter.AddEx("PH_PY695"); //�λ���ī��
			oFilter.AddEx("PH_PY705"); //��������ޱ���Ȯ��
			oFilter.AddEx("PH_PY860"); //ȣ��ǥ��ȸ
			oFilter.AddEx("PH_PY503"); //��������ڸ��
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���
			oFilter.AddEx("PH_PY507"); //��������Ȳ
			oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ
			oFilter.AddEx("PH_PY935"); //�����ȣ��Ȳ
			oFilter.AddEx("PH_PY551"); //����ο���ȸ
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY522"); //�ӱ���ũ�������Ȳ
			oFilter.AddEx("PH_PY523"); //�ӱ���ũ����ڿ���������Ȳ
			oFilter.AddEx("PH_PY524"); //������ �߰� ���곻��
			oFilter.AddEx("PH_PY683"); //����ٹ�������Ȳ
			oFilter.AddEx("PH_PYA65"); //������Ȳ (����)
			oFilter.AddEx("PH_PY583"); //���κ� �������� ��ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY108"); //�������޼���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY117"); //�޻󿩸����۾�
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�
			oFilter.AddEx("PH_PY119"); //�޻��������ϻ���
			oFilter.AddEx("PH_PY120"); //�޻󿩼ұ�����ó��
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY127"); ////���κ� 4�뺸�� �������� �� ����ݾ��Է�
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

            //�޿����� - ����Ʈ
            oFilter.AddEx("PH_PY625"); //��Ź�ڸ��
			oFilter.AddEx("PH_PY630"); //����������������Ȳ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿��μ����������
			oFilter.AddEx("PH_PY720"); //�󿩺μ����������
			oFilter.AddEx("PH_PY725"); //�޿����޺��������
			oFilter.AddEx("PH_PY740"); //�����޺��������
			oFilter.AddEx("PH_PY730"); //�޿��������
			oFilter.AddEx("PH_PY735"); //�󿩺������
			oFilter.AddEx("PH_PY745"); //����������Ȳ
			oFilter.AddEx("PH_PY750"); //�ٷμҵ�¡����Ȳ
			oFilter.AddEx("PH_PY755"); //��ȣȸ������Ȳ
			oFilter.AddEx("PH_PY760"); //����ӱݹ������ݻ��⳻����
			oFilter.AddEx("PH_PY765"); //�޿�����������
			oFilter.AddEx("PH_PY770"); //�����ҵ��õ¡�����������
			oFilter.AddEx("PH_PY775"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY776"); //�ܿ�������Ȳ
			oFilter.AddEx("PH_PY780"); //����뺸�賻��
			oFilter.AddEx("PH_PY785"); //�����ο��ݳ���
			oFilter.AddEx("PH_PY790"); //���ǰ����賻��
			oFilter.AddEx("PH_PY795"); //�����μ����޿�����
			oFilter.AddEx("PH_PY800"); //�ΰǺ������ڷ�
			oFilter.AddEx("PH_PY805"); //�޿����纯������
			oFilter.AddEx("PH_PY810"); //���޺�����ӱݳ���
			oFilter.AddEx("PH_PY815"); //����ӱݳ���
			oFilter.AddEx("PH_PY820"); //����ӱݳ���
			oFilter.AddEx("PH_PY825"); //������O/T��Ȳ
			oFilter.AddEx("PH_PY830"); //�μ����ΰǺ���Ȳ (��ȹ)
			oFilter.AddEx("PH_PY835"); //���޺�O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY840"); //ǳ�����ڰ����ڷ�
			oFilter.AddEx("PH_PY845"); //�Ⱓ���޿����޳���
			oFilter.AddEx("PH_PY850"); //�ұ޺����޸���
			oFilter.AddEx("PH_PY855"); //���κ��ӱ����޴���
			oFilter.AddEx("PH_PY865"); //��뺸����Ȳ (����)
			oFilter.AddEx("PH_PY870"); //��纰��O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY875"); //���޺������������
			oFilter.AddEx("PH_PY716"); //�Ⱓ���޿��μ����������
			oFilter.AddEx("PH_PY721"); //�Ⱓ���󿩺μ����������
			oFilter.AddEx("PH_PY717"); //�Ⱓ���޿��ݺ��������
			oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
			oFilter.AddEx("PH_PY701"); //�޿����޴��� (������)
			oFilter.AddEx("PH_PYA10"); //�޿����޴���(�μ�)
			oFilter.AddEx("PH_PYA20"); //�޿��μ����������(�μ�)
			oFilter.AddEx("PH_PYA30"); //�����޴���(�μ�)
			oFilter.AddEx("PH_PYA40"); //�󿩺μ����������(�μ�)
			oFilter.AddEx("PH_PYA50"); //DC��ȯ�ںδ�����޳���
			oFilter.AddEx("PH_PYA75"); //�����ܼ������޴���
            oFilter.AddEx("PH_PYA80"); //�ٹ��ð�ǥ���
            oFilter.AddEx("PH_PYA90"); //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
            oFilter.AddEx("PH_PY526"); //�ӱ���ũ�ο���Ȳ
            oFilter.AddEx("PH_PY999"); //����ڱ��Ѱ���

			//�������
			oFilter.AddEx("PH_PY402"); //��������ڷ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���

			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����
			oFilter.AddEx("PH_PY980"); //�ٷμҵ����޸���_�����ü�ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ƿ�����޸���_�����ü�ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //��α����޸���_�����ü�ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�����ҵ����޸���_�����ü�ڷ��ۼ�

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��

			//��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY307"); //���ڱݽ�û����(�б⺰)
			oFilter.AddEx("PH_PY030"); //������
			oFilter.AddEx("PH_PY031"); //������
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
			oFilter.AddEx("PS_DateChange"); //��¥ ���� ���
			oFilter.AddEx("PS_DateCommit"); //��¥ ����

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO600"); //�����繫��ǥ
			oFilter.AddEx("PS_CO510"); //����������������ȸ
			oFilter.AddEx("PS_CO520"); //��ǰ���� �����׸� ��ȸ
		}

		private static void CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);

			//System Form Type
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���� ���
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY004"); //�ٹ��������
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY006"); //��ȣ�۾����
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY009"); //�����ڷ�UPLOAD
			oFilter.AddEx("PH_PY011"); //������ ȣĪ �ϰ� ����(2013.07.05 �۸�� �߰�)
			oFilter.AddEx("PH_PY013"); //�����ϼ����
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY015"); //������ġ���
			oFilter.AddEx("PH_PY016"); //�⺻�������
			oFilter.AddEx("PH_PY017"); //����������
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)
			oFilter.AddEx("PH_PY019"); //�ݺ�����
			oFilter.AddEx("PH_PY020"); //�ϱ��� ����������
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����
			oFilter.AddEx("PH_PY202"); //�����ӹ��� �ް���� ��� ��ȸ
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�λ� - ����Ʈ
			oFilter.AddEx("PH_PY501"); //���ǹ߱���Ȳ
			oFilter.AddEx("PH_PY505"); //�Ի��ڴ���
			oFilter.AddEx("PH_PY510"); //������
			oFilter.AddEx("PH_PY515"); //�����ڻ�����
			oFilter.AddEx("PH_PY520"); //���������������ڴ���
			oFilter.AddEx("PH_PY525"); //�зº��ο���Ȳ
			oFilter.AddEx("PH_PY530"); //���ɺ��ο���Ȳ
			oFilter.AddEx("PH_PY535"); //�ټӳ�����ο���Ȳ
			oFilter.AddEx("PH_PY540"); //�ο���Ȳ(��ܿ�)
			oFilter.AddEx("PH_PY545"); //�ο���Ȳ(�볻��)
			oFilter.AddEx("PH_PY550"); //��ü�ο���Ȳ
			oFilter.AddEx("PH_PY555"); //���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY560"); //�������Ȳ
			oFilter.AddEx("PH_PY565"); //����ٹ�����Ȳ
			oFilter.AddEx("PH_PY570"); //����/���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY575"); //���±�����Ȳ
			oFilter.AddEx("PH_PY580"); //���κ����¿���
			oFilter.AddEx("PH_PY585"); //������ٱ�Ϻ�
			oFilter.AddEx("PH_PY590"); //�Ⱓ����������ǥ
			oFilter.AddEx("PH_PY595"); //�ټӳ����Ȳ
			oFilter.AddEx("PH_PY600"); //���ں�����ٹ���Ȳ
			oFilter.AddEx("PH_PY605"); //�ټӺ����ް��߻��׻�볻��
			oFilter.AddEx("PH_PY610"); //���±��к���볻��
			oFilter.AddEx("PH_PY615"); //�����ٹ���Ȳ
			oFilter.AddEx("PH_PY620"); //���������ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY635"); //����,��������Ȳ
			oFilter.AddEx("PH_PY640"); //���ο���������ȯ����Ȳ
			oFilter.AddEx("PH_PY645"); //�ڰݼ���������Ȳ
			oFilter.AddEx("PH_PY650"); //�뵿���հ�����Ȳ
			oFilter.AddEx("PH_PY655"); //���ƴ������Ȳ
			oFilter.AddEx("PH_PY660"); //��ֱٷ�����Ȳ
			oFilter.AddEx("PH_PY665"); //����ڳ���Ȳ
			oFilter.AddEx("PH_PY670"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY675"); //�ٹ�����Ȳ
			oFilter.AddEx("PH_PY676"); //���½ð�������ȸ
			oFilter.AddEx("PH_PY677"); //���ϱ����̻�����ȸ
			oFilter.AddEx("PH_PY679"); //���κ� �������� ��ȸ
			oFilter.AddEx("PH_PY680"); //�����Ȳ
			oFilter.AddEx("PH_PY685"); //���󰡱���Ȳ
			oFilter.AddEx("PH_PY690"); //��������Ȳ
			oFilter.AddEx("PH_PY695"); //�λ���ī��
			oFilter.AddEx("PH_PY705"); //��������ޱ���Ȯ��
			oFilter.AddEx("PH_PY860"); //ȣ��ǥ��ȸ
			oFilter.AddEx("PH_PY503"); //��������ڸ��
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���
			oFilter.AddEx("PH_PY507"); //��������Ȳ
			oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ
			oFilter.AddEx("PH_PY935"); //�����ȣ��Ȳ
			oFilter.AddEx("PH_PY551"); //����ο���ȸ
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY522"); //�ӱ���ũ�������Ȳ
			oFilter.AddEx("PH_PY523"); //�ӱ���ũ����ڿ���������Ȳ
			oFilter.AddEx("PH_PY524"); //������ �߰� ���곻��
			oFilter.AddEx("PH_PY683"); //����ٹ�������Ȳ
			oFilter.AddEx("PH_PYA65"); //������Ȳ (����)
			oFilter.AddEx("PH_PY583"); //���κ� �������� ��ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY108"); //�������޼���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY117"); //�޻󿩸����۾�
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�
			oFilter.AddEx("PH_PY119"); //�޻��������ϻ���
			oFilter.AddEx("PH_PY120"); //�޻󿩼ұ�����ó��
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY125"); //�������� ����
			oFilter.AddEx("PH_PY127"); //���κ� 4�뺸�� �������� �� ����ݾ��Է�
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

            //�޿����� - ����Ʈ
            oFilter.AddEx("PH_PY625"); //��Ź�ڸ��
			oFilter.AddEx("PH_PY630"); //����������������Ȳ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿��μ����������
			oFilter.AddEx("PH_PY720"); //�󿩺μ����������
			oFilter.AddEx("PH_PY725"); //�޿����޺��������
			oFilter.AddEx("PH_PY740"); //�����޺��������
			oFilter.AddEx("PH_PY730"); //�޿��������
			oFilter.AddEx("PH_PY735"); //�󿩺������
			oFilter.AddEx("PH_PY745"); //����������Ȳ
			oFilter.AddEx("PH_PY750"); //�ٷμҵ�¡����Ȳ
			oFilter.AddEx("PH_PY755"); //��ȣȸ������Ȳ
			oFilter.AddEx("PH_PY760"); //����ӱݹ������ݻ��⳻����
			oFilter.AddEx("PH_PY765"); //�޿�����������
			oFilter.AddEx("PH_PY770"); //�����ҵ��õ¡�����������
			oFilter.AddEx("PH_PY775"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY776"); //�ܿ�������Ȳ
			oFilter.AddEx("PH_PY780"); //����뺸�賻��
			oFilter.AddEx("PH_PY785"); //�����ο��ݳ���
			oFilter.AddEx("PH_PY790"); //���ǰ����賻��
			oFilter.AddEx("PH_PY795"); //�����μ����޿�����
			oFilter.AddEx("PH_PY800"); //�ΰǺ������ڷ�
			oFilter.AddEx("PH_PY805"); //�޿����纯������
			oFilter.AddEx("PH_PY810"); //���޺�����ӱݳ���
			oFilter.AddEx("PH_PY815"); //����ӱݳ���
			oFilter.AddEx("PH_PY820"); //����ӱݳ���
			oFilter.AddEx("PH_PY825"); //������O/T��Ȳ
			oFilter.AddEx("PH_PY830"); //�μ����ΰǺ���Ȳ (��ȹ)
			oFilter.AddEx("PH_PY835"); //���޺�O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY840"); //ǳ�����ڰ����ڷ�
			oFilter.AddEx("PH_PY845"); //�Ⱓ���޿����޳���
			oFilter.AddEx("PH_PY850"); //�ұ޺����޸���
			oFilter.AddEx("PH_PY855"); //���κ��ӱ����޴���
			oFilter.AddEx("PH_PY865"); //��뺸����Ȳ (����)
			oFilter.AddEx("PH_PY870"); //��纰��O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY875"); //���޺������������
			oFilter.AddEx("PH_PY716"); //�Ⱓ���޿��μ����������
			oFilter.AddEx("PH_PY721"); //�Ⱓ���󿩺μ����������
			oFilter.AddEx("PH_PY717"); //�Ⱓ���޿��ݺ��������
			oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
			oFilter.AddEx("PH_PY701"); //�޿����޴��� (������)
			oFilter.AddEx("PH_PYA10"); //�޿����޴���(�μ�)
			oFilter.AddEx("PH_PYA20"); //�޿��μ����������(�μ�)
			oFilter.AddEx("PH_PYA30"); //�����޴���(�μ�)
			oFilter.AddEx("PH_PYA40"); //�󿩺μ����������(�μ�)
			oFilter.AddEx("PH_PYA50"); //DC��ȯ�ںδ�����޳���
			oFilter.AddEx("PH_PYA75"); //�����ܼ������޴���
            oFilter.AddEx("PH_PYA80"); //�ٹ��ð�ǥ���
            oFilter.AddEx("PH_PYA90"); //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
            oFilter.AddEx("PH_PY526"); //�ӱ���ũ�ο���Ȳ
            oFilter.AddEx("PH_PY999"); //����ڱ��Ѱ���
			oFilter.AddEx("PH_PY998"); //����ڱ�����ȸ

			//�������
			oFilter.AddEx("PH_PY401"); //���ٹ������
			oFilter.AddEx("PH_PY402"); //��������ڷ� ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���
			oFilter.AddEx("PH_PY980"); //�Ű�_�ٷμҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ű�_�Ƿ�����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //�Ű�_��αݸ����ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�Ű�_�����ҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY419"); //ǥ�ؼ����������ڵ��
			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��
			oFilter.AddEx("ZPY504"); //��������ȸ
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�
			oFilter.AddEx("ZPY521"); //�ٷμҵ������ü����
			oFilter.AddEx("ZPY522"); //�Ƿ�� ��α� �����ü����

			//��Ÿ����
			oFilter.AddEx("PH_PY307"); //���ڱݽ�û����(�б⺰)
			oFilter.AddEx("PH_PY309"); //��αݵ��
			oFilter.AddEx("PH_PY311"); //��ٹ���������
			oFilter.AddEx("PH_PY312"); //������� ���κ����
			oFilter.AddEx("PH_PY030"); //������
			oFilter.AddEx("PH_PY031"); //������
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY315"); //���κ���α��ܾ���Ȳ
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
			oFilter.AddEx("PS_DateChange"); //��¥ ���� ���
			oFilter.AddEx("PS_DateCommit"); //��¥ ����

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO600"); //�����繫��ǥ
			oFilter.AddEx("PS_CO120"); //������ ���� ���
			oFilter.AddEx("PS_CO510"); //����������������ȸ
			oFilter.AddEx("PS_CO520"); //��ǰ���� �����׸� ��ȸ
		}

		private static void DOUBLE_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);

			//System Form Type
            //�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

			//�λ����
			//�޿�����
			oFilter.AddEx("PH_PY104"); //������������ݾ� �ϰ����
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�

			//�������
			oFilter.AddEx("PH_PY402"); //��������ڷ���
            oFilter.AddEx("PH_PY407"); //�������ڷ���
            oFilter.AddEx("PH_PY405"); //�����Ƿ���ڷ���

			//��Ÿ����

		}

		private static void MATRIX_LINK_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			oFilter.AddEx("ZPY507"); //��������ȸ(��ü)

			//��Ÿ����
		}

        private static void MATRIX_COLLAPSE_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED);
        }

		private static void VALIDATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���� ���
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY012"); //������
			oFilter.AddEx("PH_PY013"); //�����ϼ����
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY015"); //������ġ���
			oFilter.AddEx("PH_PY016"); //�⺻�������
			oFilter.AddEx("PH_PY017"); //����������
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)
			oFilter.AddEx("PH_PY019"); //�ݺ�����
			oFilter.AddEx("PH_PY020"); //�ϱ��� ����������
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����

			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY202"); //�����ӹ��� �ް���� ��ȸ
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�λ� - ����Ʈ
			oFilter.AddEx("PH_PY501"); //���ǹ߱���Ȳ
			oFilter.AddEx("PH_PY505"); //�Ի��ڴ���
			oFilter.AddEx("PH_PY510"); //������
			oFilter.AddEx("PH_PY515"); //�����ڻ�����
			oFilter.AddEx("PH_PY520"); //���������������ڴ���
			oFilter.AddEx("PH_PY525"); //�зº��ο���Ȳ
			oFilter.AddEx("PH_PY530"); //���ɺ��ο���Ȳ
			oFilter.AddEx("PH_PY535"); //�ټӳ�����ο���Ȳ
			oFilter.AddEx("PH_PY540"); //�ο���Ȳ(��ܿ�)
			oFilter.AddEx("PH_PY545"); //�ο���Ȳ(�볻��)
			oFilter.AddEx("PH_PY550"); //��ü�ο���Ȳ
			oFilter.AddEx("PH_PY555"); //���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY560"); //�������Ȳ
			oFilter.AddEx("PH_PY565"); //����ٹ�����Ȳ
			oFilter.AddEx("PH_PY570"); //����/���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY575"); //���±�����Ȳ
			oFilter.AddEx("PH_PY580"); //���κ����¿���
			oFilter.AddEx("PH_PY585"); //������ٱ�Ϻ�
			oFilter.AddEx("PH_PY590"); //�Ⱓ����������ǥ
			oFilter.AddEx("PH_PY595"); //�ټӳ����Ȳ
			oFilter.AddEx("PH_PY600"); //���ں�����ٹ���Ȳ
			oFilter.AddEx("PH_PY605"); //�ټӺ����ް��߻��׻�볻��
			oFilter.AddEx("PH_PY610"); //���±��к���볻��
			oFilter.AddEx("PH_PY615"); //�����ٹ���Ȳ
			oFilter.AddEx("PH_PY620"); //���������ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY635"); //����,��������Ȳ
			oFilter.AddEx("PH_PY640"); //���ο���������ȯ����Ȳ
			oFilter.AddEx("PH_PY645"); //�ڰݼ���������Ȳ
			oFilter.AddEx("PH_PY650"); //�뵿���հ�����Ȳ
			oFilter.AddEx("PH_PY655"); //���ƴ������Ȳ
			oFilter.AddEx("PH_PY660"); //��ֱٷ�����Ȳ
			oFilter.AddEx("PH_PY665"); //����ڳ���Ȳ
			oFilter.AddEx("PH_PY670"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY675"); //�ٹ�����Ȳ
			oFilter.AddEx("PH_PY676"); //���½ð�������ȸ
			oFilter.AddEx("PH_PY677"); //���ϱ����̻�����ȸ
			oFilter.AddEx("PH_PY679"); //���κ� �������� ��ȸ
			oFilter.AddEx("PH_PY680"); //�����Ȳ
			oFilter.AddEx("PH_PY685"); //���󰡱���Ȳ
			oFilter.AddEx("PH_PY690"); //��������Ȳ
			oFilter.AddEx("PH_PY695"); //�λ���ī��
			oFilter.AddEx("PH_PY705"); //��������ޱ���Ȯ��
			oFilter.AddEx("PH_PY860"); //ȣ��ǥ��ȸ
			oFilter.AddEx("PH_PY503"); //��������ڸ��
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���
			oFilter.AddEx("PH_PY507"); //��������Ȳ
			oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ
			oFilter.AddEx("PH_PY935"); //�����ȣ��Ȳ
			oFilter.AddEx("PH_PY551"); //����ο���ȸ
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY522"); //�ӱ���ũ�������Ȳ
			oFilter.AddEx("PH_PY523"); //�ӱ���ũ����ڿ���������Ȳ
			oFilter.AddEx("PH_PY524"); //������ �߰� ���곻��
			oFilter.AddEx("PH_PY683"); //����ٹ�������Ȳ
			oFilter.AddEx("PH_PYA65"); //������Ȳ (����)
			oFilter.AddEx("PH_PY583"); //���κ� �������� ��ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY108"); //�������޼���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY117"); //�޻󿩸����۾�
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�
			oFilter.AddEx("PH_PY120"); //�޻󿩼ұ�����ó��
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

            //�޿����� - ����Ʈ
            oFilter.AddEx("PH_PY625"); //��Ź�ڸ��
			oFilter.AddEx("PH_PY630"); //����������������Ȳ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿��μ����������
			oFilter.AddEx("PH_PY720"); //�󿩺μ����������
			oFilter.AddEx("PH_PY725"); //�޿����޺��������
			oFilter.AddEx("PH_PY740"); //�����޺��������
			oFilter.AddEx("PH_PY730"); //�޿��������
			oFilter.AddEx("PH_PY735"); //�󿩺������
			oFilter.AddEx("PH_PY745"); //����������Ȳ
			oFilter.AddEx("PH_PY750"); //�ٷμҵ�¡����Ȳ
			oFilter.AddEx("PH_PY755"); //��ȣȸ������Ȳ
			oFilter.AddEx("PH_PY760"); //����ӱݹ������ݻ��⳻����
			oFilter.AddEx("PH_PY765"); //�޿�����������
			oFilter.AddEx("PH_PY770"); //�����ҵ��õ¡�����������
			oFilter.AddEx("PH_PY775"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY776"); //�ܿ�������Ȳ
			oFilter.AddEx("PH_PY780"); //����뺸�賻��
			oFilter.AddEx("PH_PY785"); //�����ο��ݳ���
			oFilter.AddEx("PH_PY790"); //���ǰ����賻��
			oFilter.AddEx("PH_PY795"); //�����μ����޿�����
			oFilter.AddEx("PH_PY800"); //�ΰǺ������ڷ�
			oFilter.AddEx("PH_PY805"); //�޿����纯������
			oFilter.AddEx("PH_PY810"); //���޺�����ӱݳ���
			oFilter.AddEx("PH_PY815"); //����ӱݳ���
			oFilter.AddEx("PH_PY820"); //����ӱݳ���
			oFilter.AddEx("PH_PY825"); //������O/T��Ȳ
			oFilter.AddEx("PH_PY830"); //�μ����ΰǺ���Ȳ (��ȹ)
			oFilter.AddEx("PH_PY835"); //���޺�O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY840"); //ǳ�����ڰ����ڷ�
			oFilter.AddEx("PH_PY845"); //�Ⱓ���޿����޳���
			oFilter.AddEx("PH_PY850"); //�ұ޺����޸���
			oFilter.AddEx("PH_PY855"); //���κ��ӱ����޴���
			oFilter.AddEx("PH_PY865"); //��뺸����Ȳ (����)
			oFilter.AddEx("PH_PY870"); //��纰��O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY875"); //���޺������������
			oFilter.AddEx("PH_PY716"); //�Ⱓ���޿��μ����������
			oFilter.AddEx("PH_PY721"); //�Ⱓ���󿩺μ����������
			oFilter.AddEx("PH_PY717"); //�Ⱓ���޿��ݺ��������
			oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
			oFilter.AddEx("PH_PY701"); //�޿����޴��� (������)

			oFilter.AddEx("PH_PYA10"); //�޿����޴���(�μ�)
			oFilter.AddEx("PH_PYA20"); //�޿��μ����������(�μ�)
			oFilter.AddEx("PH_PYA30"); //�����޴���(�μ�)
			oFilter.AddEx("PH_PYA40"); //�󿩺μ����������(�μ�)
			oFilter.AddEx("PH_PYA50"); //DC��ȯ�ںδ�����޳���
			oFilter.AddEx("PH_PYA75"); //�����ܼ������޴���
            oFilter.AddEx("PH_PYA80"); //�ٹ��ð�ǥ���
            oFilter.AddEx("PH_PYA90"); //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
            oFilter.AddEx("PH_PY526"); //�ӱ���ũ�ο���Ȳ
            oFilter.AddEx("PH_PY999"); //����ڱ��Ѱ���
			oFilter.AddEx("PH_PY998"); //����ڱ�����ȸ

			//�������
			oFilter.AddEx("PH_PY401"); //���ٹ������
			oFilter.AddEx("PH_PY402"); //��������ڷ� ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���
			oFilter.AddEx("PH_PY415"); //������
			oFilter.AddEx("PH_PY417"); //���� �������ϻ���
			oFilter.AddEx("PH_PY980"); //�Ű�_�ٷμҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ű�_�Ƿ�����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //�Ű�_��αݸ����ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�Ű�_�����ҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY419"); //ǥ�ؼ����������ڵ��

			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��
			oFilter.AddEx("ZPY504"); //��������ȸ
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY507"); //��������ȸ(��ü)
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�
			oFilter.AddEx("ZPY510"); //�����ٹ��� �ϰ�����
			oFilter.AddEx("ZPY521"); //�ٷμҵ������ü����
			oFilter.AddEx("ZPY522"); //�Ƿ�� ��α� �����ü����

			//��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY305"); //���ڱݽ�û��
			oFilter.AddEx("PH_PY306"); //���ڱݽ�û����(���κ�)
			oFilter.AddEx("PH_PY309"); //��αݵ��
			oFilter.AddEx("PH_PY310"); //��αݰ�����ȯ
			oFilter.AddEx("PH_PY311"); //��ٹ���������
			oFilter.AddEx("PH_PY313"); //��αݰ��
			oFilter.AddEx("PH_PY314"); //��αݰ�� ���� ��ȸ(�޿������ڷ��)
			oFilter.AddEx("PH_PY030"); //������
			oFilter.AddEx("PH_PY031"); //������
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY315"); //���κ���α��ܾ���Ȳ
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
			oFilter.AddEx("PS_DateChange"); //��¥ ���� ���
			oFilter.AddEx("PS_DateCommit"); //��¥ ����

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO600"); //�����繫��ǥ
			oFilter.AddEx("PS_CO120"); //������ ���� ���
			oFilter.AddEx("PS_CO131"); //7.������������Ȳ
			oFilter.AddEx("PS_CO501"); //ǰ�񺰿������
			oFilter.AddEx("PS_CO502"); //ǰ����տ����׸���
		}

		private static void MATRIX_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD);

			//System Form Type
            //�����
			//�ǸŰ���
			//���Ű���
			//������
			//�������

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���� ���
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)

			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY105"); //ȣ�����ǥ
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

            //�������
            oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�

			//��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY309"); //��αݵ��
			oFilter.AddEx("PH_PY310"); //��αݰ�����ȯ
			oFilter.AddEx("PH_PY311"); //��ٹ���������
			oFilter.AddEx("PH_PY313"); //��αݰ��
			oFilter.AddEx("PH_PY012"); //������
			oFilter.AddEx("PH_PY315"); //���κ���α��ܾ���Ȳ
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO510"); // ������������Ȳ
			oFilter.AddEx("PS_CO520"); //��ǰ���� �����׸� ��ȸ
			oFilter.AddEx("PS_CO501"); //ǰ�񺰿������
			oFilter.AddEx("PS_CO502"); //ǰ����տ����׸���
		}

		private static void DATASOURCE_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD);

			//System Form Type
			//�����
			//�ǸŰ���
			//���Ű���
			//������
			//�������

			//AddOn Form Type
            //�����
			//�ǸŰ���
			//���Ű���
			//������
			//�������
		}

		private static void FORM_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���� ���
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY004"); //�ٹ��������
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY006"); //��ȣ�۾����
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY009"); //�����ڷ�UPLOAD
			oFilter.AddEx("PH_PY010"); //���ϱ���ó��
			oFilter.AddEx("PH_PY011"); //������ ȣĪ �ϰ� ����(2013.07.05 �۸�� �߰�)
			oFilter.AddEx("PH_PY013"); //�����ϼ����
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY015"); //������ġ���
			oFilter.AddEx("PH_PY016"); //�⺻�������
			oFilter.AddEx("PH_PY017"); //����������
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)
			oFilter.AddEx("PH_PY019"); //�ݺ�����
			oFilter.AddEx("PH_PY020"); //�ϱ��� ����������
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����
			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�λ� - ����Ʈ
			oFilter.AddEx("PH_PY501"); //���ǹ߱���Ȳ
			oFilter.AddEx("PH_PY505"); //�Ի��ڴ���
			oFilter.AddEx("PH_PY510"); //������
			oFilter.AddEx("PH_PY515"); //�����ڻ�����
			oFilter.AddEx("PH_PY520"); //���������������ڴ���
			oFilter.AddEx("PH_PY525"); //�зº��ο���Ȳ
			oFilter.AddEx("PH_PY530"); //���ɺ��ο���Ȳ
			oFilter.AddEx("PH_PY535"); //�ټӳ�����ο���Ȳ
			oFilter.AddEx("PH_PY540"); //�ο���Ȳ(��ܿ�)
			oFilter.AddEx("PH_PY545"); //�ο���Ȳ(�볻��)
			oFilter.AddEx("PH_PY550"); //��ü�ο���Ȳ
			oFilter.AddEx("PH_PY555"); //���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY560"); //�������Ȳ
			oFilter.AddEx("PH_PY565"); //����ٹ�����Ȳ
			oFilter.AddEx("PH_PY570"); //����/���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY575"); //���±�����Ȳ
			oFilter.AddEx("PH_PY580"); //���κ����¿���
			oFilter.AddEx("PH_PY585"); //������ٱ�Ϻ�
			oFilter.AddEx("PH_PY590"); //�Ⱓ����������ǥ
			oFilter.AddEx("PH_PY595"); //�ټӳ����Ȳ
			oFilter.AddEx("PH_PY600"); //���ں�����ٹ���Ȳ
			oFilter.AddEx("PH_PY605"); //�ټӺ����ް��߻��׻�볻��
			oFilter.AddEx("PH_PY610"); //���±��к���볻��
			oFilter.AddEx("PH_PY615"); //�����ٹ���Ȳ
			oFilter.AddEx("PH_PY620"); //���������ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY635"); //����,��������Ȳ
			oFilter.AddEx("PH_PY640"); //���ο���������ȯ����Ȳ
			oFilter.AddEx("PH_PY645"); //�ڰݼ���������Ȳ
			oFilter.AddEx("PH_PY650"); //�뵿���հ�����Ȳ
			oFilter.AddEx("PH_PY655"); //���ƴ������Ȳ
			oFilter.AddEx("PH_PY660"); //��ֱٷ�����Ȳ
			oFilter.AddEx("PH_PY665"); //����ڳ���Ȳ
			oFilter.AddEx("PH_PY670"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY675"); //�ٹ�����Ȳ
			oFilter.AddEx("PH_PY676"); //���½ð�������ȸ
			oFilter.AddEx("PH_PY677"); //���ϱ����̻�����ȸ
			oFilter.AddEx("PH_PY679"); //���κ� �������� ��ȸ
			oFilter.AddEx("PH_PY680"); //�����Ȳ
			oFilter.AddEx("PH_PY685"); //���󰡱���Ȳ
			oFilter.AddEx("PH_PY690"); //��������Ȳ
			oFilter.AddEx("PH_PY695"); //�λ���ī��
			oFilter.AddEx("PH_PY705"); //��������ޱ���Ȯ��
			oFilter.AddEx("PH_PY860"); //ȣ��ǥ��ȸ
			oFilter.AddEx("PH_PY503"); //��������ڸ��
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���
			oFilter.AddEx("PH_PY507"); //��������Ȳ
			oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ
			oFilter.AddEx("PH_PY935"); //�����ȣ��Ȳ
			oFilter.AddEx("PH_PY551"); //����ο���ȸ
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY522"); //�ӱ���ũ�������Ȳ
			oFilter.AddEx("PH_PY523"); //�ӱ���ũ����ڿ���������Ȳ
			oFilter.AddEx("PH_PY524"); //������ �߰� ���� ����
			oFilter.AddEx("PH_PY683"); //����ٹ�������Ȳ
			oFilter.AddEx("PH_PYA65"); //������Ȳ (����)
			oFilter.AddEx("PH_PY583"); //���κ� �������� ��ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY108"); //�������޼���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY117"); //�޻󿩸����۾�
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�
			oFilter.AddEx("PH_PY119"); //�޻��������ϻ���
			oFilter.AddEx("PH_PY120"); //�޻󿩼ұ�����ó��
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY125"); //�������� ����
			oFilter.AddEx("PH_PY127"); //���κ� 4�뺸�� �������� �� ����ݾ��Է�
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

            //�޿����� - ����Ʈ
            oFilter.AddEx("PH_PY625"); //��Ź�ڸ��
			oFilter.AddEx("PH_PY630"); //����������������Ȳ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿��μ����������
			oFilter.AddEx("PH_PY720"); //�󿩺μ����������
			oFilter.AddEx("PH_PY725"); //�޿����޺��������
			oFilter.AddEx("PH_PY740"); //�����޺��������
			oFilter.AddEx("PH_PY730"); //�޿��������
			oFilter.AddEx("PH_PY735"); //�󿩺������
			oFilter.AddEx("PH_PY745"); //����������Ȳ
			oFilter.AddEx("PH_PY750"); //�ٷμҵ�¡����Ȳ
			oFilter.AddEx("PH_PY755"); //��ȣȸ������Ȳ
			oFilter.AddEx("PH_PY760"); //����ӱݹ������ݻ��⳻����
			oFilter.AddEx("PH_PY765"); //�޿�����������
			oFilter.AddEx("PH_PY770"); //�����ҵ��õ¡�����������
			oFilter.AddEx("PH_PY775"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY776"); //�ܿ�������Ȳ
			oFilter.AddEx("PH_PY780"); //����뺸�賻��
			oFilter.AddEx("PH_PY785"); //�����ο��ݳ���
			oFilter.AddEx("PH_PY790"); //���ǰ����賻��
			oFilter.AddEx("PH_PY795"); //�����μ����޿�����
			oFilter.AddEx("PH_PY800"); //�ΰǺ������ڷ�
			oFilter.AddEx("PH_PY805"); //�޿����纯������
			oFilter.AddEx("PH_PY810"); //���޺�����ӱݳ���
			oFilter.AddEx("PH_PY815"); //����ӱݳ���
			oFilter.AddEx("PH_PY820"); //����ӱݳ���
			oFilter.AddEx("PH_PY825"); //������O/T��Ȳ
			oFilter.AddEx("PH_PY830"); //�μ����ΰǺ���Ȳ (��ȹ)
			oFilter.AddEx("PH_PY835"); //���޺�O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY840"); //ǳ�����ڰ����ڷ�
			oFilter.AddEx("PH_PY845"); //�Ⱓ���޿����޳���
			oFilter.AddEx("PH_PY850"); //�ұ޺����޸���
			oFilter.AddEx("PH_PY855"); //���κ��ӱ����޴���
			oFilter.AddEx("PH_PY865"); //��뺸����Ȳ (����)
			oFilter.AddEx("PH_PY870"); //��纰��O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY875"); //���޺������������
			oFilter.AddEx("PH_PY716"); //�Ⱓ���޿��μ����������
			oFilter.AddEx("PH_PY721"); //�Ⱓ���󿩺μ����������
			oFilter.AddEx("PH_PY717"); //�Ⱓ���޿��ݺ��������
			oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
			oFilter.AddEx("PH_PY701"); //�޿����޴��� (������)

			oFilter.AddEx("PH_PYA10"); //�޿����޴���(�μ�)
			oFilter.AddEx("PH_PYA20"); //�޿��μ����������(�μ�)
			oFilter.AddEx("PH_PYA30"); //�����޴���(�μ�)
			oFilter.AddEx("PH_PYA40"); //�󿩺μ����������(�μ�)
			oFilter.AddEx("PH_PYA50"); //DC��ȯ�ںδ�����޳���
			oFilter.AddEx("PH_PYA75"); //�����ܼ������޴���
            oFilter.AddEx("PH_PYA80"); //�ٹ��ð�ǥ���
            oFilter.AddEx("PH_PYA90"); //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
            oFilter.AddEx("PH_PY526"); //�ӱ���ũ�ο���Ȳ
            oFilter.AddEx("PH_PY999"); //����ڱ��Ѱ���

			//�������
			oFilter.AddEx("PH_PY401"); //���ٹ������
			oFilter.AddEx("PH_PY402"); //��������ڷ� ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���
			oFilter.AddEx("PH_PY415"); //������
			oFilter.AddEx("PH_PY417"); //���� �������ϻ���
			oFilter.AddEx("PH_PY980"); //�Ű�_�ٷμҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ű�_�Ƿ�����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //�Ű�_��αݸ����ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�Ű�_�����ҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY419"); //ǥ�ؼ����������ڵ��

			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��
			oFilter.AddEx("ZPY504"); //��������ȸ
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY507"); //��������ȸ(��ü)
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�
			oFilter.AddEx("ZPY510"); //�����ٹ��� �ϰ�����
			oFilter.AddEx("ZPY521"); //�ٷμҵ������ü����
			oFilter.AddEx("ZPY522"); //�Ƿ�� ��α� �����ü����

			oFilter.AddEx("RPY401"); //������õ¡�� ������
			oFilter.AddEx("RPY501"); //�����ڷ���Ȳ
			oFilter.AddEx("RPY502"); //�����ٹ�����Ȳ
			oFilter.AddEx("RPY503"); //�ٷμҵ� ��õ¡����
			oFilter.AddEx("RPY504"); //�ٷμҵ� ��õ������
			oFilter.AddEx("RPY505"); //�ҵ��ڷ�����ǥ
			oFilter.AddEx("RPY506"); //����¡��ȯ�޴���
			oFilter.AddEx("RPY508"); //������������ǥ
			oFilter.AddEx("RPY509"); //���ټ��Ű����ǥ
			oFilter.AddEx("RPY510"); //������ٷμҵ����
			oFilter.AddEx("RPY511"); //��αݸ���

			//��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY302"); //���ڱ����޿Ϸ�ó��
			oFilter.AddEx("PH_PY303"); //���ڱ��������ϻ���
			oFilter.AddEx("PH_PY305"); //���ڱݽ�û��
			oFilter.AddEx("PH_PY306"); //���ڱݽ�û����(���κ�)
			oFilter.AddEx("PH_PY307"); //���ڱݽ�û����(�б⺰)
			oFilter.AddEx("PH_PY309"); //��αݵ��
			oFilter.AddEx("PH_PY310"); //��αݰ�����ȯ
			oFilter.AddEx("PH_PY311"); //��ٹ���������
			oFilter.AddEx("PH_PY312"); //������� ���κ����
			oFilter.AddEx("PH_PY313"); //��αݰ��
			oFilter.AddEx("PH_PY030"); //������
			oFilter.AddEx("PH_PY031"); //������
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY315"); //���κ���α��ܾ���Ȳ
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
			oFilter.AddEx("PS_DateChange"); //��¥ ���� ���
			oFilter.AddEx("PS_DateCommit"); //��¥ ����
        }

		private static void FORM_UNLOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���� ���
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY004"); //�ٹ��������
			oFilter.AddEx("PH_PY005"); //������������
			oFilter.AddEx("PH_PY006"); //��ȣ�۾����
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY008"); //�ϱ��µ��
			oFilter.AddEx("PH_PY011"); //������ ȣĪ �ϰ� ����(2013.07.05 �۸�� �߰�)
			oFilter.AddEx("PH_PY013"); //�����ϼ����
			oFilter.AddEx("PH_PY014"); //�����ϼ�����
			oFilter.AddEx("PH_PY015"); //������ġ���
			oFilter.AddEx("PH_PY016"); //�⺻�������
			oFilter.AddEx("PH_PY017"); //����������
			oFilter.AddEx("PH_PY018"); //���ϱٹ�üũ(������)
			oFilter.AddEx("PH_PY019"); //�ݺ�����
			oFilter.AddEx("PH_PY020"); //�ϱ��� ����������
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����
			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�λ� - ����Ʈ
			oFilter.AddEx("PH_PY501"); //���ǹ߱���Ȳ
			oFilter.AddEx("PH_PY505"); //�Ի��ڴ���
			oFilter.AddEx("PH_PY510"); //������
			oFilter.AddEx("PH_PY515"); //�����ڻ�����
			oFilter.AddEx("PH_PY520"); //���������������ڴ���
			oFilter.AddEx("PH_PY525"); //�зº��ο���Ȳ
			oFilter.AddEx("PH_PY530"); //���ɺ��ο���Ȳ
			oFilter.AddEx("PH_PY535"); //�ټӳ�����ο���Ȳ
			oFilter.AddEx("PH_PY540"); //�ο���Ȳ(��ܿ�)
			oFilter.AddEx("PH_PY545"); //�ο���Ȳ(�볻��)
			oFilter.AddEx("PH_PY550"); //��ü�ο���Ȳ
			oFilter.AddEx("PH_PY555"); //���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY560"); //�������Ȳ
			oFilter.AddEx("PH_PY565"); //����ٹ�����Ȳ
			oFilter.AddEx("PH_PY570"); //����/���ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY575"); //���±�����Ȳ
			oFilter.AddEx("PH_PY580"); //���κ����¿���
			oFilter.AddEx("PH_PY585"); //������ٱ�Ϻ�
			oFilter.AddEx("PH_PY590"); //�Ⱓ����������ǥ
			oFilter.AddEx("PH_PY595"); //�ټӳ����Ȳ
			oFilter.AddEx("PH_PY600"); //���ں�����ٹ���Ȳ
			oFilter.AddEx("PH_PY605"); //�ټӺ����ް��߻��׻�볻��
			oFilter.AddEx("PH_PY610"); //���±��к���볻��
			oFilter.AddEx("PH_PY615"); //�����ٹ���Ȳ
			oFilter.AddEx("PH_PY620"); //���������ϱٹ�����Ȳ
			oFilter.AddEx("PH_PY635"); //����,��������Ȳ
			oFilter.AddEx("PH_PY640"); //���ο���������ȯ����Ȳ
			oFilter.AddEx("PH_PY645"); //�ڰݼ���������Ȳ
			oFilter.AddEx("PH_PY650"); //�뵿���հ�����Ȳ
			oFilter.AddEx("PH_PY655"); //���ƴ������Ȳ
			oFilter.AddEx("PH_PY660"); //��ֱٷ�����Ȳ
			oFilter.AddEx("PH_PY665"); //����ڳ���Ȳ
			oFilter.AddEx("PH_PY670"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY675"); //�ٹ�����Ȳ
			oFilter.AddEx("PH_PY676"); //���½ð�������ȸ
			oFilter.AddEx("PH_PY677"); //���ϱ����̻�����ȸ
			oFilter.AddEx("PH_PY679"); //���κ� �������� ��ȸ
			oFilter.AddEx("PH_PY680"); //�����Ȳ
			oFilter.AddEx("PH_PY685"); //���󰡱���Ȳ
			oFilter.AddEx("PH_PY690"); //��������Ȳ
			oFilter.AddEx("PH_PY695"); //�λ���ī��
			oFilter.AddEx("PH_PY705"); //��������ޱ���Ȯ��
			oFilter.AddEx("PH_PY860"); //ȣ��ǥ��ȸ
			oFilter.AddEx("PH_PY503"); //��������ڸ��
			oFilter.AddEx("PH_PY678"); //�����ٹ��� �ϰ� ���
			oFilter.AddEx("PH_PY507"); //��������Ȳ
			oFilter.AddEx("PH_PY681"); //��ٹ��ϼ���Ȳ
			oFilter.AddEx("PH_PY935"); //�����ȣ��Ȳ
			oFilter.AddEx("PH_PY551"); //����ο���ȸ
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY522"); //�ӱ���ũ�������Ȳ
			oFilter.AddEx("PH_PY523"); //�ӱ���ũ����ڿ���������Ȳ
			oFilter.AddEx("PH_PY524"); //������ �߰� ���곻��
			oFilter.AddEx("PH_PY683"); //����ٹ�������Ȳ
			oFilter.AddEx("PH_PYA65"); //������Ȳ (����)
			oFilter.AddEx("PH_PY583"); //���κ� �������� ��ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY102"); //�����׸���
			oFilter.AddEx("PH_PY103"); //�����׸���
			oFilter.AddEx("PH_PY104"); //������������ݾ��ϰ����
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY107"); //�޻󿩱����ϼ���
			oFilter.AddEx("PH_PY108"); //�������޼���
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���
			oFilter.AddEx("PH_PY109_1"); //�޻󿩺����ڷ� �׸����
			oFilter.AddEx("PH_PY110"); //���λ������
			oFilter.AddEx("PH_PY111"); //�޻󿩰��
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����
			oFilter.AddEx("PH_PY113"); //�޻󿩺а��ڷ����
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY115"); //�����ݰ��
			oFilter.AddEx("PH_PY116"); //�����ݺа��ڷ����
			oFilter.AddEx("PH_PY117"); //�޻󿩸����۾�
			oFilter.AddEx("PH_PY118"); //�޻�Email�߼�
			oFilter.AddEx("PH_PY119"); //�޻��������ϻ���
			oFilter.AddEx("PH_PY120"); //�޻󿩼ұ�����ó��
			oFilter.AddEx("PH_PY121"); //�򰡰��޾� ���
			oFilter.AddEx("PH_PY122"); //�޻���� ���κμ��������
			oFilter.AddEx("PH_PY123"); //���з����
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���
			oFilter.AddEx("PH_PY125"); //�������� ����
			oFilter.AddEx("PH_PY127"); //���κ� 4�뺸�� �������� �� ����ݾ��Է�
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

            //�޿����� - ����Ʈ
            oFilter.AddEx("PH_PY625"); //��Ź�ڸ��
			oFilter.AddEx("PH_PY630"); //����������������Ȳ
			oFilter.AddEx("PH_PY700"); //�޿����޴���
			oFilter.AddEx("PH_PY710"); //�����޴���
			oFilter.AddEx("PH_PY715"); //�޿��μ����������
			oFilter.AddEx("PH_PY720"); //�󿩺μ����������
			oFilter.AddEx("PH_PY725"); //�޿����޺��������
			oFilter.AddEx("PH_PY740"); //�����޺��������
			oFilter.AddEx("PH_PY730"); //�޿��������
			oFilter.AddEx("PH_PY735"); //�󿩺������
			oFilter.AddEx("PH_PY745"); //����������Ȳ
			oFilter.AddEx("PH_PY750"); //�ٷμҵ�¡����Ȳ
			oFilter.AddEx("PH_PY755"); //��ȣȸ������Ȳ
			oFilter.AddEx("PH_PY760"); //����ӱݹ������ݻ��⳻����
			oFilter.AddEx("PH_PY765"); //�޿�����������
			oFilter.AddEx("PH_PY770"); //�����ҵ��õ¡�����������
			oFilter.AddEx("PH_PY775"); //���κ�������Ȳ
			oFilter.AddEx("PH_PY776"); //�ܿ�������Ȳ
			oFilter.AddEx("PH_PY780"); //����뺸�賻��
			oFilter.AddEx("PH_PY785"); //�����ο��ݳ���
			oFilter.AddEx("PH_PY790"); //���ǰ����賻��
			oFilter.AddEx("PH_PY795"); //�����μ����޿�����
			oFilter.AddEx("PH_PY800"); //�ΰǺ������ڷ�
			oFilter.AddEx("PH_PY805"); //�޿����纯������
			oFilter.AddEx("PH_PY810"); //���޺�����ӱݳ���
			oFilter.AddEx("PH_PY815"); //����ӱݳ���
			oFilter.AddEx("PH_PY820"); //����ӱݳ���
			oFilter.AddEx("PH_PY825"); //������O/T��Ȳ
			oFilter.AddEx("PH_PY830"); //�μ����ΰǺ���Ȳ (��ȹ)
			oFilter.AddEx("PH_PY835"); //���޺�O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY840"); //ǳ�����ڰ����ڷ�
			oFilter.AddEx("PH_PY845"); //�Ⱓ���޿����޳���
			oFilter.AddEx("PH_PY850"); //�ұ޺����޸���
			oFilter.AddEx("PH_PY855"); //���κ��ӱ����޴���
			oFilter.AddEx("PH_PY865"); //��뺸����Ȳ (����)
			oFilter.AddEx("PH_PY870"); //��纰��O/T�׼�����Ȳ
			oFilter.AddEx("PH_PY875"); //���޺������������
			oFilter.AddEx("PH_PY716"); //�Ⱓ���޿��μ����������
			oFilter.AddEx("PH_PY721"); //�Ⱓ���󿩺μ����������
			oFilter.AddEx("PH_PY717"); //�Ⱓ���޿��ݺ��������
			oFilter.AddEx("PH_PY718"); //����Ϸ�ݾ״��O/T��Ȳ
			oFilter.AddEx("PH_PY701"); //�޿����޴��� (������)

			oFilter.AddEx("PH_PYA10"); //�޿����޴���(�μ�)
			oFilter.AddEx("PH_PYA20"); //�޿��μ����������(�μ�)
			oFilter.AddEx("PH_PYA30"); //�����޴���(�μ�)
			oFilter.AddEx("PH_PYA40"); //�󿩺μ����������(�μ�)
			oFilter.AddEx("PH_PYA50"); //DC��ȯ�ںδ�����޳���
			oFilter.AddEx("PH_PYA75"); //�����ܼ������޴���
            oFilter.AddEx("PH_PYA80"); //�ٹ��ð�ǥ���
            oFilter.AddEx("PH_PYA90"); //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
            oFilter.AddEx("PH_PY526"); //�ӱ���ũ�ο���Ȳ
            oFilter.AddEx("PH_PY999"); //����ڱ��Ѱ���
			oFilter.AddEx("PH_PY998"); //����ڱ�����ȸ

			//�������
			oFilter.AddEx("PH_PY401"); //���ٹ������
			oFilter.AddEx("PH_PY402"); //��������ڷ� ���
			oFilter.AddEx("PH_PY405"); //�Ƿ����
			oFilter.AddEx("PH_PY407"); //��αݵ��
			oFilter.AddEx("PH_PY409"); //��α����������
			oFilter.AddEx("PH_PY411"); //����.�����ҵ�������
			oFilter.AddEx("PH_PY413"); //������.�����������Ա��ڷ� ���
			oFilter.AddEx("PH_PY415"); //������
			oFilter.AddEx("PH_PY417"); //���� �������ϻ���
			oFilter.AddEx("PH_PY980"); //�Ű�_�ٷμҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY985"); //�Ű�_�Ƿ�����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY990"); //�Ű�_��αݸ����ڷ��ۼ�
			oFilter.AddEx("PH_PY995"); //�Ű�_�����ҵ����޸����ڷ��ۼ�
			oFilter.AddEx("PH_PY419"); //ǥ�ؼ����������ڵ��

			oFilter.AddEx("PH_PY910"); //�ҵ�����Ű����
			oFilter.AddEx("PH_PY915"); //�ٷμҵ��õ¡�������
			oFilter.AddEx("PH_PY920"); //��õ¡�����������
			oFilter.AddEx("PH_PY925"); //��αݸ������
			oFilter.AddEx("PH_PY930"); //����¡����ȯ�޴���
			oFilter.AddEx("PH_PY931"); //ǥ�ؼ�������������ȸ
			oFilter.AddEx("PH_PY932"); //���ٹ��������Ȳ
			oFilter.AddEx("PH_PY933"); //�����Ѿ׽Ű�����ڷ�
			oFilter.AddEx("PH_PYA55"); //����¡����ȯ�޴���(����)
			oFilter.AddEx("PH_PYA70"); //�ҵ漼��õ¡������������û�����

			oFilter.AddEx("ZPY341"); //���� �����ڷ� ����
			oFilter.AddEx("ZPY343"); //���� �ڷ� ����
			oFilter.AddEx("ZPY421"); //�����ҵ������ü����
			oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���
			oFilter.AddEx("ZPY502"); //��(��) �ٹ��� ���
			oFilter.AddEx("ZPY503"); //���꼼�װ��
			oFilter.AddEx("ZPY504"); //��������ȸ
			oFilter.AddEx("ZPY505"); //��αݸ����
			oFilter.AddEx("ZPY506"); //�Ƿ������
			oFilter.AddEx("ZPY507"); //��������ȸ(��ü)
			oFilter.AddEx("ZPY508"); //�������� �ҵ���� �� ���
			oFilter.AddEx("ZPY509"); //�����ڷ� �����۾�
			oFilter.AddEx("ZPY510"); //�����ٹ��� �ϰ�����
			oFilter.AddEx("ZPY521"); //�ٷμҵ������ü����
			oFilter.AddEx("ZPY522"); //�Ƿ�� ��α� �����ü����

			oFilter.AddEx("RPY401"); //������õ¡�� ������
			oFilter.AddEx("RPY501"); //�����ڷ���Ȳ
			oFilter.AddEx("RPY502"); //�����ٹ�����Ȳ
			oFilter.AddEx("RPY503"); //�ٷμҵ� ��õ¡����
			oFilter.AddEx("RPY504"); //�ٷμҵ� ��õ������
			oFilter.AddEx("RPY505"); //�ҵ��ڷ�����ǥ
			oFilter.AddEx("RPY506"); //����¡��ȯ�޴���
			oFilter.AddEx("RPY508"); //������������ǥ
			oFilter.AddEx("RPY509"); //���ټ��Ű����ǥ
			oFilter.AddEx("RPY510"); //������ٷμҵ����
			oFilter.AddEx("RPY511"); //��αݸ���

			//��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY302"); //���ڱ����޿Ϸ�ó��
			oFilter.AddEx("PH_PY303"); //���ڱ��������ϻ���
			oFilter.AddEx("PH_PY305"); //���ڱݽ�û��
			oFilter.AddEx("PH_PY306"); //���ڱݽ�û����(���κ�)
			oFilter.AddEx("PH_PY307"); //���ڱݽ�û����(�б⺰)
			oFilter.AddEx("PH_PY311"); //��ٹ���������
			oFilter.AddEx("PH_PY312"); //������� ���κ����
			oFilter.AddEx("PH_PY313"); //��αݰ��
			oFilter.AddEx("PH_PY030"); //������
			oFilter.AddEx("PH_PY031"); //������
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY315"); //���κ���α��ܾ���Ȳ
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)
			oFilter.AddEx("PS_DateChange"); //��¥ ���� ���
			oFilter.AddEx("PS_DateCommit"); //��¥ ����

			//�繫����
			oFilter.AddEx("PS_CO658"); //�����繫��ǥ�� ��������
			oFilter.AddEx("PS_CO600"); //�����繫��ǥ
			oFilter.AddEx("PS_CO605"); //���ռ��Һ�
			oFilter.AddEx("PS_CO120"); //������ ���� ���
			oFilter.AddEx("PS_CO130"); //��ǰ�� ���� ���
			oFilter.AddEx("PS_CO131"); //������������Ȳ
			oFilter.AddEx("PS_CO606"); //�����繫��ǥ ���� ����
			oFilter.AddEx("PS_CO510"); //����������������ȸ
			oFilter.AddEx("PS_CO520"); //��ǰ���� �����׸� ��ȸ
			oFilter.AddEx("PS_CO501"); //ǰ�񺰿������
			oFilter.AddEx("PS_CO502"); //ǰ����տ����׸���
		}

		private static void FORM_ACTIVATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����
		}

		private static void FORM_DEACTIVATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}
		
		private static void FORM_CLOSE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
		}

		private static void FORM_RESIZE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���͵��
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���
			oFilter.AddEx("PH_PY007"); //�����ܰ����
			oFilter.AddEx("PH_PY508"); //�������� ��� �� �߱�
			oFilter.AddEx("PH_PY021"); //�����󿬶�ó����
			oFilter.AddEx("PH_PY201"); //�����ӹ��� �ް���� ���
			oFilter.AddEx("PH_PY203"); //�����������
			oFilter.AddEx("PH_PY204"); //������ȹ���
			oFilter.AddEx("PH_PY205"); //������ȹVS������ȸ

			//�޿�����
			oFilter.AddEx("PH_PY100"); //���ؼ��׼���
			oFilter.AddEx("PH_PY101"); //��������
			oFilter.AddEx("PH_PY106"); //������ļ���
			oFilter.AddEx("PH_PY114"); //�����ݱ��ؼ���
			oFilter.AddEx("PH_PY130"); //���� ���������� ��޵��
			oFilter.AddEx("PH_PY131"); //���������� ������
			oFilter.AddEx("PH_PY132"); //�������� ���κ� ���
			oFilter.AddEx("PH_PY133"); //������ Ƚ�� ����
			oFilter.AddEx("PH_PY134"); //�ҵ漼/�ֹμ� ��������
			oFilter.AddEx("PH_PY129"); //���κ���������(DC��) ���
            oFilter.AddEx("PH_PY135"); //�޻󿩺а�ó��
            oFilter.AddEx("PH_PY136"); //�޻󿩺а�ó�� ��α�Ģ����

            //�������
            oFilter.AddEx("ZPY501"); //�ҵ�����׸� ���

            //��Ÿ����
			oFilter.AddEx("PH_PY301"); //���ڱݽ�û���
			oFilter.AddEx("PH_PY302"); //���ڱ����޿Ϸ�ó��
			oFilter.AddEx("PH_PY305"); //���ڱݽ�û��
			oFilter.AddEx("PH_PY306"); //���ڱݽ�û����(���κ�)
			oFilter.AddEx("PH_PY307"); //���ڱݽ�û����(�б⺰)
			oFilter.AddEx("PH_PY032"); //��������
			oFilter.AddEx("PH_PY034"); //����а�ó��
			oFilter.AddEx("PH_PYA60"); //���ڱݽ�û����(����)

			//���°���
			oFilter.AddEx("PH_PY677"); //���±����̻��� ����

			//�繫����
			oFilter.AddEx("PS_CO605"); //���ռ��Һ�
		}

		private static void FORM_KEY_DOWN(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void FORM_MENU_HILIGHT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void vPRINT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void PRINT_DATA(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT_DATA);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

            //AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void CHOOSE_FROM_LIST(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���͵��
			oFilter.AddEx("PH_PY005"); //������������

			//�޿�����
			oFilter.AddEx("PH_PY103"); //�����׸���

			//�������
			//��Ÿ����
		}

		private static void RIGHT_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);

	        //System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���͵��
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY003"); //���¿��¼���

			//�޿�����
			oFilter.AddEx("PH_PY109"); //�޻󿩺����ڷ���

			//�������
			//��Ÿ����
		}

		private static void MENU_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void FORM_DATA_ADD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void FORM_DATA_UPDATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);

		    //System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void FORM_DATA_DELETE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE);

			//System Form Type
			//�����
			//�λ����
			//�޿�����
			//�������
			//��Ÿ����


			//AddOn Form Type
			//�����
			//�λ����
			//�޿�����
            //�������
			//��Ÿ����
		}

		private static void FORM_DATA_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		{
			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);

			//System Form Type
			//�����
			//�λ����
			//�޿�����

			//�������
			//��Ÿ����

			//AddOn Form Type
			//�����
			oFilter.AddEx("PH_PY000"); //����ڱ��Ѱ���

            //�λ����
			oFilter.AddEx("PH_PY001"); //�λ縶���͵��
			oFilter.AddEx("PH_PY002"); //���½ð����� ���
			oFilter.AddEx("PH_PY105"); //ȣ��ǥ���
			oFilter.AddEx("PH_PY124"); //�����Ǿ� �ݾ� ���

			//�޿�����
			oFilter.AddEx("PH_PY112"); //�޻��ڷ����

			//�������
			//��Ÿ����
		}
	}
}
