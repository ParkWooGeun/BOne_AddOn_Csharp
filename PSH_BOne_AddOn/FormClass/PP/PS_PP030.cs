using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작업지시등록
	/// </summary>
	internal class PS_PP030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Matrix oMat03;
		private SAPbouiCOM.DBDataSource oDS_PS_USERDS01; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP030H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP030L; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP030M; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMat01Row01;
		private int oMat02Row02;
		private int oMat03Row03;

		////사용자구조체
		private struct ItemInformations
		{
			public string itemCode;
			public string BatchNum;
			public int Quantity;
			public int OPORNo;
			public int POR1No;
			public bool Check;
			public int OPDNNo;
			public int PDN1No;
		}
		private ItemInformations[] ItemInformation;
		private int ItemInformationCount;

		private string oDocEntry01;
		private string oSCardCod01;
		private string oMark;
		private SAPbouiCOM.BoFormMode oFormMode01;
		private bool oHasMatrix01;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				//PS_PP030_CreateItems();
				//PS_PP030_ComboBox_Setting();
				//PS_PP030_CF_ChooseFromList();
				//PS_PP030_EnableMenus();
				//PS_PP030_SetDocument(oFormDocEntry);
				//PS_PP030_FormResize();
				//Initialization();

				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1286", true); //닫기
				oForm.EnableMenu("1284", true); //취소
				oForm.EnableMenu("1293", true); //행삭제
				oForm.EnableMenu("1299", false); //행닫기
			}
			catch(Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

	//	public void Initialization()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		SAPbouiCOM.ComboBox oCombo = null;

	//		////아이디별 사업장 세팅
	//		oCombo = oForm01.Items.Item("SBPLId").Specific;
	//		oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);

	//		////아이디별 사번 세팅
	//		//    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD

	//		////아이디별 부서 세팅
	//		//    Set oCombo = oForm01.Items("DeptCode").Specific
	//		//    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
	//		//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oCombo = null;
	//		return;
	//		Initialization_Error:
	//		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//		//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oCombo = null;
	//		MDC_Com.MDC_GF_Message(ref "Initialization_Error:" + Err().Number + " - " + Err().Description, ref "E");
	//	}

	//	public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		int i = 0;
	//		decimal Total = default(decimal);
	//		switch (pval.EventType) {
	//			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
	//				////1
	//				Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
	//				////2
	//				Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
	//				////5
	//				Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_CLICK:
	//				////6
	//				Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
	//				////7
	//				Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
	//				////8
	//				Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
	//				////10
	//				Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
	//				////11

	//				// 공정금액 합계 추가 S

	//				Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);

	//				for (i = 0; i <= oMat03.VisualRowCount - 1; i++) {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					Total = Total + oMat03.Columns.Item("CpPrice").Cells.Item(i + 1).Specific.VALUE;
	//				}

	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				oForm01.Items.Item("Total").Specific.VALUE = Total;
	//				break;
	//			// 공정금액 합계 추가 E


	//			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
	//				////18
	//				break;
	//			////et_FORM_ACTIVATE
	//			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
	//				////19
	//				break;
	//			////et_FORM_DEACTIVATE
	//			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
	//				////20
	//				Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
	//				////27
	//				Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
	//				////3
	//				Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
	//				////4
	//				break;
	//			////et_LOST_FOCUS
	//			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
	//				////17
	//				Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
	//				break;
	//		}
	//		return;
	//		Raise_ItemEvent_Error:
	//		///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}


	//	public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		int i = 0;
	//		////BeforeAction = True
	//		if ((pval.BeforeAction == true)) {
	//			switch (pval.MenuUID) {
	//				case "1284":
	//					//취소
	//					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
	//						if ((PS_PP030_Validate("취소") == false)) {
	//							BubbleEvent = false;
	//							return;
	//						}
	//						if (SubMain.Sbo_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1")) {
	//							BubbleEvent = false;
	//							return;
	//						}
	//					} else {
	//						MDC_Com.MDC_GF_Message(ref "현재 모드에서는 취소할수 없습니다.", ref "W");
	//						BubbleEvent = false;
	//						return;
	//					}
	//					break;
	//				case "1286":
	//					//닫기

	//					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
	//						if ((PS_PP030_Validate("닫기") == false)) {
	//							BubbleEvent = false;
	//							return;
	//						}
	//						if (SubMain.Sbo_Application.MessageBox("문서를 닫기(종료) 처리하겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1")) {
	//							BubbleEvent = false;
	//							return;
	//						}
	//					} else {
	//						MDC_Com.MDC_GF_Message(ref "현재 모드에서는 닫기(종료) 처리할 수 없습니다.", ref "W");
	//						BubbleEvent = false;
	//						return;
	//					}
	//					break;


	//				case "1292":
	//					//행추가
	//					break;
	//				case "1293":
	//					//행삭제
	//					Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
	//					break;
	//				case "1281":
	//					//찾기
	//					break;
	//				case "1282":
	//					//추가
	//					break;
	//				case "1288":
	//				case "1289":
	//				case "1290":
	//				case "1291":
	//					//레코드이동버튼
	//					break;
	//			}
	//		////BeforeAction = False
	//		} else if ((pval.BeforeAction == false)) {
	//			switch (pval.MenuUID) {
	//				case "1284":
	//					//취소
	//					break;
	//				case "1286":
	//					//닫기
	//					break;
	//				case "1292":
	//					//행추가
	//					if (oLastItemUID01 == "Mat03") {
	//						//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////멀티인경우만
	//						if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104") {
	//							////행추가가 가능검사
	//							if ((PS_PP030_Validate("행추가03") == false)) {
	//								BubbleEvent = false;
	//								return;
	//							}
	//							oMat03.AddRow(1, oMat03Row03 - 1);
	//							for (i = 1; i <= oMat03.VisualRowCount; i++) {
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oMat03.Columns.Item("Sequence").Cells.Item(i).Specific.VALUE = i;

	//								////새로추가된 행의 값 설정
	//								if (oMat03Row03 == i) {
	//									//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oMat03.Columns.Item("ReWorkYN").Cells.Item(i).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//									////PK/탈지일때 재작업여부 예

	//									//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
	//									////실적여부 아니오
	//									//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//									////일보여부 예
	//								}
	//							}
	//							oMat03.FlushToDataSource();
	//							oMat03.LoadFromDataSource();
	//						} else {
	//							MDC_Com.MDC_GF_Message(ref "멀티인 경우만 행추가 가능합니다.", ref "W");
	//						}
	//					}
	//					break;
	//				case "1293":
	//					//행삭제
	//					Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
	//					break;
	//				case "1281":
	//					//찾기
	//					PS_PP030_FormItemEnabled();
	//					////UDO방식
	//					oForm01.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//					break;
	//				case "1282":
	//					//추가
	//					PS_PP030_FormItemEnabled();
	//					////UDO방식
	//					PS_PP030_AddMatrixRow01(0, ref true);
	//					////UDO방식
	//					PS_PP030_AddMatrixRow02(0, ref true);
	//					////UDO방식
	//					break;
	//				case "1288":
	//				case "1289":
	//				case "1290":
	//				case "1291":
	//					//레코드이동버튼
	//					PS_PP030_FormItemEnabled();
	//					break;
	//			}
	//		}
	//		return;
	//		Raise_MenuEvent_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		////BeforeAction = True
	//		if ((BusinessObjectInfo.BeforeAction == true)) {
	//			switch (BusinessObjectInfo.EventType) {
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
	//					////33
	//					break;
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
	//					////34
	//					break;
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
	//					////35
	//					break;
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
	//					////36
	//					break;
	//			}
	//		////BeforeAction = False
	//		} else if ((BusinessObjectInfo.BeforeAction == false)) {
	//			switch (BusinessObjectInfo.EventType) {
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
	//					////33
	//					break;
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
	//					////34
	//					break;
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
	//					////35
	//					break;
	//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
	//					////36
	//					break;
	//			}
	//		}
	//		return;
	//		Raise_FormDataEvent_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		if (pval.BeforeAction == true) {
	//			//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
	//			//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
	//			//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
	//			//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
	//			//            MenuCreationParams01.uniqueID = "MenuUID"
	//			//            MenuCreationParams01.String = "메뉴명"
	//			//            MenuCreationParams01.Enabled = True
	//			//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
	//			//        End If
	//		} else if (pval.BeforeAction == false) {
	//			//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
	//			//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
	//			//        End If
	//		}
	//		if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02" | pval.ItemUID == "Mat03") {
	//			if (pval.Row > 0) {
	//				oLastItemUID01 = pval.ItemUID;
	//				oLastColUID01 = pval.ColUID;
	//				oLastColRow01 = pval.Row;
	//			}
	//		} else {
	//			oLastItemUID01 = pval.ItemUID;
	//			oLastColUID01 = "";
	//			oLastColRow01 = 0;
	//		}
	//		if (pval.ItemUID == "Mat01") {
	//			if (pval.Row > 0) {
	//				oMat01Row01 = pval.Row;
	//			}
	//		} else if (pval.ItemUID == "Mat02") {
	//			if (pval.Row > 0) {
	//				oMat02Row02 = pval.Row;
	//			}
	//		} else if (pval.ItemUID == "Mat03") {
	//			if (pval.Row > 0) {
	//				oMat03Row03 = pval.Row;
	//			}
	//		}
	//		return;
	//		Raise_RightClickEvent_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		int i = 0;
	//		string query01 = null;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		string oOrdGbn01 = null;
	//		string oProcType01 = null;

	//		short li_Cnt = 0;
	//		short li_LineId = 0;

	//		object lChildForm = null;
	//		//팝업창 호출 용 변수(2012.04.12 송명규)

	//		if (pval.BeforeAction == true) {
	//			if (pval.ItemUID == "Button01") {
	//				if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
	//					PS_PP030_MTX01();
	//					////조회
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
	//				}
	//			}
	//			if (pval.ItemUID == "1") {
	//				if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

	//					if (PS_PP030_DataValidCheck() == false) {
	//						BubbleEvent = false;
	//						return;
	//					}

	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oDocEntry01 = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE);
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oSCardCod01 = Strings.Trim(oForm01.Items.Item("SCardCod").Specific.VALUE);

	//					oFormMode01 = oForm01.Mode;
	//					////멀티게이지 일괄생성기능구현 , 엔드베이링 추가 - 류영조
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//						if (PS_PP030_AutoCreateMultiGage() == false) {
	//							PS_PP030_AddMatrixRow01(oMat02.VisualRowCount);
	//							PS_PP030_AddMatrixRow02(oMat03.VisualRowCount);
	//							BubbleEvent = false;
	//							return;
	//						}
	//						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
	//						PS_PP030_FormItemEnabled();
	//						SubMain.Sbo_Application.ActivateMenuItem(("1282"));
	//						BubbleEvent = false;
	//						return;
	//					} else {
	//						////멀티게이지를 제외한 나머지 경우는 자동으로 입력
	//					}
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
	//					if (PS_PP030_DataValidCheck() == false) {
	//						BubbleEvent = false;
	//						return;
	//					}
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oDocEntry01 = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE);
	//					oFormMode01 = oForm01.Mode;
	//					/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//					if (oMat02.VisualRowCount == 0) {
	//						oMat02.Clear();
	//						oMat02.AddRow();
	//						oMat02.FlushToDataSource();
	//						oMat02.LoadFromDataSource();
	//					}
	//					////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
	//				}
	//			}

	//			//표준공수조회, 품목별공수조회 버튼 클릭 시(2012.04.12 송명규)
	//			//표준공수조회 버튼 클릭
	//			if (pval.ItemUID == "btnWkSrch") {

	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "선택") {

	//					MDC_Com.MDC_GF_Message(ref "작업구분을 선택하십시오.", ref "W");
	//					return;

	//				} else {

	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {

	//						lChildForm = new PS_PP033();
	//						//UPGRADE_WARNING: lChildForm.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						lChildForm.LoadForm(this);

	//					} else {

	//						MDC_Com.MDC_GF_Message(ref "작업구분이 [제품_기계공구] 또는 [제품_몰드] 일 경우에만 사용이 가능합니다.", ref "W");
	//						return;

	//					}

	//				}


	//			//품목별공수조회 버튼 클릭
	//			} else if (pval.ItemUID == "btnItmSrch") {

	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "선택") {

	//					MDC_Com.MDC_GF_Message(ref "작업구분을 선택하십시오.", ref "W");
	//					return;

	//				} else {

	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {

	//						lChildForm = new PS_PP031();
	//						//UPGRADE_WARNING: lChildForm.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						lChildForm.LoadForm();

	//					} else {

	//						MDC_Com.MDC_GF_Message(ref "작업구분이 [제품_기계공구] 또는 [제품_몰드] 일 경우에만 사용이 가능합니다.", ref "W");
	//						return;

	//					}

	//				}

	//			}
	//			//표준공수조회, 품목별공수조회 버튼 클릭 시(2012.04.12 송명규)

	//		} else if (pval.BeforeAction == false) {
	//			if (pval.ItemUID == "Button01") {
	//				if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
	//				}
	//			}
	//			if (pval.ItemUID == "1") {
	//				if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
	//					if (pval.ActionSuccess == true) {
	//						RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
	//						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oOrdGbn01 = MDC_PS_Common.GetValue("SELECT U_OrdGbn FROM [@PS_PP030H] WHERE DocEntry = '" + oDocEntry01 + "'");
	//						////기계공구, 몰드
	//						if ((oOrdGbn01 == "105" | oOrdGbn01 == "106")) {
	//							query01 = "SELECT U_ProcType, DocEntry, LineId FROM [@PS_PP030L] WHERE DocEntry = '" + oDocEntry01 + "'";
	//							RecordSet01.DoQuery(query01);
	//							for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
	//								if ((RecordSet01.Fields.Item(0).Value == "10")) {
	//									PS_PP030_PurchaseRequest(RecordSet01.Fields.Item(1).Value, RecordSet01.Fields.Item(2).Value);
	//								}
	//								RecordSet01.MoveNext();
	//							}
	//						}
	//						//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//						RecordSet01 = null;
	//						PS_PP030_FormItemEnabled();
	//						PS_PP030_AddMatrixRow01(0, ref true);
	//						////UDO방식일때
	//						PS_PP030_AddMatrixRow02(0, ref true);
	//						////UDO방식일때
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SCardCod").Specific.VALUE = oSCardCod01;
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("Total").Specific.VALUE = 0;
	//						//공정금액 합계 초기화

	//						oForm01.Items.Item("Button01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//					}
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
	//					if (pval.ActionSuccess == true) {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("Total").Specific.VALUE = 0;
	//						//공정금액 합계 초기화
	//					}
	//				} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
	//					if (pval.ActionSuccess == true) {
	//						if ((oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)) {
	//							RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
	//							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oOrdGbn01 = MDC_PS_Common.GetValue("SELECT U_OrdGbn FROM [@PS_PP030H] WHERE DocEntry = '" + oDocEntry01 + "'");
	//							////기계공구, 몰드
	//							if ((oOrdGbn01 == "105" | oOrdGbn01 == "106")) {
	//								query01 = "SELECT U_ProcType, DocEntry, LineId FROM [@PS_PP030L] WHERE DocEntry = '" + oDocEntry01 + "'";
	//								RecordSet01.DoQuery(query01);
	//								for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
	//									if ((RecordSet01.Fields.Item(0).Value == "10")) {
	//										PS_PP030_PurchaseRequest(RecordSet01.Fields.Item(1).Value, RecordSet01.Fields.Item(2).Value);
	//									}
	//									RecordSet01.MoveNext();
	//								}
	//							}
	//							if (oOrdGbn01 == "104") {
	//								query01 = "Update [@PS_PP030M] set VisOrder = U_Sequence - 1, LineId = U_Sequence, U_LineId = U_Sequence WHERE LineId <> U_Sequence And DocEntry = '" + oDocEntry01 + "'";
	//								RecordSet01.DoQuery(query01);

	//								query01 = "SELECT Count(*), Min(LineId) FROM [@PS_PP030M] WHERE DocEntry = '" + oDocEntry01 + "' and U_CpCode = 'CP50107'";
	//								RecordSet01.DoQuery(query01);

	//								//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								li_Cnt = RecordSet01.Fields.Item(0).Value;
	//								//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								li_LineId = RecordSet01.Fields.Item(1).Value;

	//								if ((li_Cnt > 1)) {
	//									query01 = "Update [@PS_PP030M] set U_ResultYN = 'N' WHERE DocEntry = '" + oDocEntry01 + "' and LineId = '" + li_LineId + "'";
	//									RecordSet01.DoQuery(query01);
	//								} else {
	//									query01 = "Update [@PS_PP030M] set U_ResultYN = 'Y' WHERE DocEntry = '" + oDocEntry01 + "' and LineId = '" + li_LineId + "'";
	//									RecordSet01.DoQuery(query01);
	//								}


	//							}
	//							oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
	//							oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
	//							PS_PP030_FormItemEnabled();
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("DocEntry").Specific.VALUE = oDocEntry01;
	//							oForm01.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//						}
	//						PS_PP030_FormItemEnabled();
	//					}
	//				}
	//			}
	//		}
	//		return;
	//		Raise_EVENT_ITEM_PRESSED_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		string ordGbn = null;
	//		string InputGbn = null;
	//		object ChildForm01 = null;
	//		if (pval.BeforeAction == true) {
	//			MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "CntcCode", "");
	//			////사용자값활성
	//			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
	//				////찾기모드는 입력가능하도록
	//				MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "ItemCode", "");
	//				////사용자값활성 입력가능하도록
	//			} else {
	//				MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "ItemCode", "");
	//				////사용자값활성 입력은 안됨
	//			}
	//			//        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
	//			MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "Mat02", "CntcCode");
	//			////사용자값활성
	//			MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "Mat03", "CpBCode");
	//			////사용자값활성
	//			MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "Mat03", "CpCode");
	//			////사용자값활성
	//			if (pval.ItemUID == "Mat02") {
	//				if (pval.ColUID == "ItemCode") {
	//					//UPGRADE_WARNING: oMat02.Columns(InputGbn).Cells(pval.Row).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oMat02.Columns.Item("InputGbn").Cells.Item(pval.Row).Specific.Selected == null) {
	//						MDC_Com.MDC_GF_Message(ref "투입구분을 선택하세요", ref "W");
	//						oMat02.Columns.Item("InputGbn").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//						BubbleEvent = false;
	//						return;
	//					} else {
	//						if ((PS_PP030_Validate("수정02") == false)) {
	//							BubbleEvent = false;
	//							return;
	//						}
	//						//[2011.2.14] 추가 Begin------------------------------------------------------------------------------------------------------------
	//						//107010002(END BEARING #44),107010004(END BEARING #2) 일경우에는 투입자재를 직접 입력한다.
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE) == "107010002" | Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE) == "107010004") {
	//							//                        BubbleEvent = False
	//							return;
	//						}
	//						//[2011.2.14] 추가 End--------------------------------------------------------------------------------------------------------------

	//						//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						ordGbn = Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE);
	//						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						InputGbn = Strings.Trim(oMat02.Columns.Item("InputGbn").Cells.Item(pval.Row).Specific.Selected.VALUE);
	//						ChildForm01 = new PS_SM021();
	//						//UPGRADE_WARNING: ChildForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, ordGbn, InputGbn, Strings.Trim(oDS_PS_PP030H.GetValue("U_BPLId", 0)));
	//						BubbleEvent = false;
	//						return;
	//					}
	//				}
	//			} else if (pval.ItemUID == "Mat03") {
	//				if (pval.ColUID == "FailCode") {
	//					//UPGRADE_WARNING: oMat03.Columns(FailCode).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(pval.Row).Specific.VALUE)) {
	//						//If oForm01.Items("FailCode").Specific.VALUE = "" Then
	//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
	//						BubbleEvent = false;
	//					}
	//				}
	//			}
	//		} else if (pval.BeforeAction == false) {

	//		}
	//		return;
	//		Raise_EVENT_KEY_DOWN_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		oForm01.Freeze(true);
	//		if (pval.BeforeAction == true) {

	//		} else if (pval.BeforeAction == false) {
	//			if (pval.ItemChanged == true) {
	//				oForm01.Freeze(true);
	//				if ((pval.ItemUID == "Mat02")) {
	//					if ((PS_PP030_Validate("수정02") == false)) {
	//						oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, Strings.Trim(oDS_PS_PP030L.GetValue("U_" + pval.ColUID, pval.Row - 1)));
	//					} else {
	//						if ((pval.ColUID == "특정컬럼")) {
	//						} else {
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//						}
	//					}
	//				} else if ((pval.ItemUID == "Mat03")) {
	//					if ((PS_PP030_Validate("수정03") == false)) {
	//						oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, Strings.Trim(oDS_PS_PP030M.GetValue("U_" + pval.ColUID, pval.Row - 1)));
	//					} else {
	//						if ((pval.ColUID == "WorkGbn")) {
	//							//UPGRADE_WARNING: oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if (oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE == "10") {
	//								//UPGRADE_WARNING: oMat03.Columns(StdHour).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_CpPrice", pval.Row - 1, Convert.ToString(MDC_PS_Common.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1) * oMat03.Columns.Item("StdHour").Cells.Item(pval.Row).Specific.VALUE));
	//								//UPGRADE_WARNING: oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							} else if (oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE == "20") {
	//								//UPGRADE_WARNING: oMat03.Columns(StdHour).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_CpPrice", pval.Row - 1, Convert.ToString(MDC_PS_Common.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1) * oMat03.Columns.Item("StdHour").Cells.Item(pval.Row).Specific.VALUE));
	//							}

	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);

	//						} else if ((pval.ColUID == "특정컬럼")) {
	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//							if (oMat03.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP030M.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
	//								PS_PP030_AddMatrixRow02((pval.Row));
	//							}
	//						} else {
	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//						}
	//					}
	//				} else {
	//					if ((pval.ItemUID == "OrdGbn")) {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE);
	//						if (oHasMatrix01 == true) {
	//							//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							////작업구분이 멀티일때만
	//							if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "104") {
	//								oForm01.Items.Item("BasicGub").Enabled = true;
	//								oForm01.Items.Item("MulGbn1").Enabled = true;
	//								oForm01.Items.Item("MulGbn2").Enabled = true;
	//								oForm01.Items.Item("MulGbn3").Enabled = true;
	//							} else {
	//								oForm01.Items.Item("BasicGub").Enabled = false;
	//								oForm01.Items.Item("MulGbn1").Enabled = false;
	//								oForm01.Items.Item("MulGbn2").Enabled = false;
	//								oForm01.Items.Item("MulGbn3").Enabled = false;
	//							}
	//							//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							////엔드베어링일때
	//							if ((oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "107")) {
	//								//                            oMat02.Columns("InputGbn").Editable = True
	//								//                            Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn2")
	//								oMat02.Columns.Item("InputGbn").Editable = true;
	//							} else {
	//								//                            oMat02.Columns("InputGbn").Editable = True
	//								//                            Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn")
	//								oMat02.Columns.Item("InputGbn").Editable = false;
	//							}
	//							//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if ((oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "105" | oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "106")) {
	//								oMat02.Columns.Item("Weight").Editable = true;
	//							} else {
	//								oMat02.Columns.Item("Weight").Editable = false;
	//							}
	//							//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							////멀티,엔드베어링이면
	//							if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//								//                            oMat03.Columns("CpBCode").Editable = False
	//								//                            oMat03.Columns("CpCode").Editable = False
	//								//                            oMat03.Columns("ResultYN").Editable = False
	//								//                            oMat03.Columns("ReportYN").Editable = False
	//							} else {
	//								//                            oMat03.Columns("CpBCode").Editable = True
	//								//                            oMat03.Columns("CpCode").Editable = True
	//								//                            oMat03.Columns("ResultYN").Editable = True
	//								//                            oMat03.Columns("ReportYN").Editable = True
	//							}
	//							oMat02.Clear();
	//							oMat02.FlushToDataSource();
	//							oMat02.LoadFromDataSource();
	//							PS_PP030_AddMatrixRow01(0, ref true);
	//							oMat03.Clear();
	//							oMat03.FlushToDataSource();
	//							oMat03.LoadFromDataSource();
	//							PS_PP030_AddMatrixRow02(0, ref true);
	//							//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("DocDate").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("DueDate").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("ItemCode").Specific.VALUE = "";
	//							////공정리스트 매트릭스를 초기화 시킨다.
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("ItemName").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("JakMyung").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("JakSize").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("JakUnit").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("CntcCode").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("CntcName").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("OrdMgNum").Specific.VALUE = "";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("ReqWt").Specific.VALUE = 0;
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("SelWt").Specific.VALUE = 0;
	//						} else {
	//							////그냥선택시
	//							//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							////작업구분이 멀티일때만
	//							if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "104" | oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "107") {
	//								//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								////작업구분이 멀티일때만
	//								if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "104") {
	//									oForm01.Items.Item("BasicGub").Enabled = true;
	//									oForm01.Items.Item("MulGbn1").Enabled = true;
	//									oForm01.Items.Item("MulGbn2").Enabled = true;
	//									oForm01.Items.Item("MulGbn3").Enabled = true;
	//								} else {
	//									oForm01.Items.Item("BasicGub").Enabled = false;
	//									oForm01.Items.Item("MulGbn1").Enabled = false;
	//									oForm01.Items.Item("MulGbn2").Enabled = false;
	//									oForm01.Items.Item("MulGbn3").Enabled = false;
	//								}
	//								//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								////엔드베어링일때
	//								if ((oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "107")) {
	//									//                                oMat02.Columns("InputGbn").Editable = True
	//									//                                Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn2")
	//									oMat02.Columns.Item("InputGbn").Editable = true;
	//								} else {
	//									//                                oMat02.Columns("InputGbn").Editable = True
	//									//                                Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn")
	//									oMat02.Columns.Item("InputGbn").Editable = false;
	//								}
	//								//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								if ((oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "105" | oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "106")) {
	//									oMat02.Columns.Item("Weight").Editable = true;
	//								} else {
	//									oMat02.Columns.Item("Weight").Editable = false;
	//								}
	//								//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								////멀티,엔드베어링이면
	//								if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//									//                                oMat03.Columns("CpBCode").Editable = False
	//									//                                oMat03.Columns("CpCode").Editable = False
	//									//                                oMat03.Columns("ResultYN").Editable = False
	//									//                                oMat03.Columns("ReportYN").Editable = False
	//								} else {
	//									//                                oMat03.Columns("CpBCode").Editable = True
	//									//                                oMat03.Columns("CpCode").Editable = True
	//									//                                oMat03.Columns("ResultYN").Editable = True
	//									//                                oMat03.Columns("ReportYN").Editable = True
	//								}
	//								oMat02.Clear();
	//								oMat02.FlushToDataSource();
	//								oMat02.LoadFromDataSource();
	//								PS_PP030_AddMatrixRow01(0, ref true);
	//								oMat03.Clear();
	//								oMat03.FlushToDataSource();
	//								oMat03.LoadFromDataSource();
	//								PS_PP030_AddMatrixRow02(0, ref true);
	//								//Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
	//								//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);
	//								//UPGRADE_WARNING: oForm01.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("DocDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
	//								//UPGRADE_WARNING: oForm01.Items(DueDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("DueDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("ItemCode").Specific.VALUE = "";
	//								////공정리스트 매트릭스를 초기화 시킨다.
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("ItemName").Specific.VALUE = "";
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("JakMyung").Specific.VALUE = "";
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("JakSize").Specific.VALUE = "";
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("JakUnit").Specific.VALUE = "";
	//								//                            oForm01.Items("CntcCode").Specific.Value = ""
	//								//                            oForm01.Items("CntcName").Specific.Value = ""
	//								//UPGRADE_WARNING: oForm01.Items(CntcCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD();
	//								//UPGRADE_WARNING: oForm01.Items(OrdMgNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("OrdMgNum").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("ReqWt").Specific.VALUE = 0;
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("SelWt").Specific.VALUE = 0;
	//								//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							} else if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "선택") {
	//								////아무행위도 하지 않음
	//							////멀티랑 엔드베어링일때
	//							} else {
	//								MDC_Com.MDC_GF_Message(ref "멀티,엔드베어링작업만 선택할수 있습니다.", ref "W");
	//								//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//							}
	//						}
	//						//                ElseIf (pval.ItemUID = "CardCode") Then
	//						//                    Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
	//						//                    Call oDS_PS_PP030H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
	//						//                Else
	//						//                    Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
	//					}
	//				}
	//				oMat02.LoadFromDataSource();
	//				oMat03.LoadFromDataSource();
	//				oMat02.AutoResizeColumns();
	//				oMat03.AutoResizeColumns();
	//				oForm01.Update();
	//				if (pval.ItemUID == "Mat01") {
	//					oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
	//				} else if (pval.ItemUID == "Mat02") {
	//					oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
	//				} else if (pval.ItemUID == "Mat03") {
	//					oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
	//				} else {

	//				}
	//				oForm01.Freeze(false);
	//			}
	//		}
	//		oForm01.Freeze(false);
	//		return;
	//		Raise_EVENT_COMBO_SELECT_Error:
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		string True_False = null;

	//		if (pval.BeforeAction == true) {
	//			if (pval.ItemUID == "Opt01") {
	//				oForm01.Freeze(true);
	//				True_False = Convert.ToString(oMat02.Columns.Item("Weight").Editable);
	//				oForm01.Settings.MatrixUID = "Mat01";
	//				oForm01.Settings.EnableRowFormat = true;
	//				oForm01.Settings.Enabled = true;
	//				oMat02.Columns.Item("Weight").Editable = Convert.ToBoolean(True_False);
	//				oMat01.AutoResizeColumns();
	//				oMat02.AutoResizeColumns();
	//				oMat03.AutoResizeColumns();
	//				oForm01.Freeze(false);
	//			}
	//			if (pval.ItemUID == "Opt02") {
	//				oForm01.Freeze(true);
	//				True_False = Convert.ToString(oMat02.Columns.Item("Weight").Editable);
	//				oForm01.Settings.MatrixUID = "Mat02";
	//				oForm01.Settings.EnableRowFormat = true;
	//				oForm01.Settings.Enabled = true;
	//				oMat02.Columns.Item("Weight").Editable = Convert.ToBoolean(True_False);
	//				oMat01.AutoResizeColumns();
	//				oMat02.AutoResizeColumns();
	//				oMat03.AutoResizeColumns();
	//				oForm01.Freeze(false);
	//			}
	//			if (pval.ItemUID == "Opt03") {
	//				oForm01.Freeze(true);
	//				True_False = Convert.ToString(oMat02.Columns.Item("Weight").Editable);
	//				oForm01.Settings.MatrixUID = "Mat03";
	//				oForm01.Settings.EnableRowFormat = true;
	//				oForm01.Settings.Enabled = true;
	//				oMat02.Columns.Item("Weight").Editable = Convert.ToBoolean(True_False);
	//				oMat01.AutoResizeColumns();
	//				oMat02.AutoResizeColumns();
	//				oMat03.AutoResizeColumns();
	//				oForm01.Freeze(false);
	//			}
	//			if (pval.ItemUID == "Mat01") {
	//				if (pval.Row > 0) {
	//					oMat01.SelectRow(pval.Row, true, false);
	//					oMat01Row01 = pval.Row;
	//				}
	//			}
	//			if (pval.ItemUID == "Mat02") {
	//				if (pval.Row > 0) {
	//					oMat02.SelectRow(pval.Row, true, false);
	//					oMat02Row02 = pval.Row;
	//				}
	//			}
	//			if (pval.ItemUID == "Mat03") {
	//				if (pval.Row > 0) {
	//					oMat03.SelectRow(pval.Row, true, false);
	//					oMat03Row03 = pval.Row;
	//				}
	//			}
	//		} else if (pval.BeforeAction == false) {

	//		}
	//		return;
	//		Raise_EVENT_CLICK_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		if (pval.BeforeAction == true) {
	//			if (pval.ItemUID == "Mat01") {
	//				if (pval.Row == 0) {

	//					oMat01.Columns.Item(pval.ColUID).TitleObject.Sortable = true;
	//					oMat01.FlushToDataSource();

	//				} else if (pval.Row > 0) {
	//					oHasMatrix01 = true;
	//					oForm01.Freeze(true);
	//					oMat02.Clear();
	//					oMat02.FlushToDataSource();
	//					oMat02.LoadFromDataSource();
	//					PS_PP030_AddMatrixRow01(0, ref true);
	//					////아이템활성화하여 Validate 발생
	//					oForm01.Items.Item("OrdGbn").Enabled = true;
	//					oForm01.Items.Item("BPLId").Enabled = true;
	//					oForm01.Items.Item("ItemCode").Enabled = true;
	//					oForm01.Items.Item("OrdMgNum").Enabled = true;
	//					//UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(pval.Row).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oMat01.Columns.Item("ItmBsort").Cells.Item(pval.Row).Specific.Selected == null) {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//						//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//						//UPGRADE_WARNING: oForm01.Items(CntcCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD();
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("DocDate").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("DueDate").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("BaseType").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("BaseNum").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("OrdMgNum").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("OrdNum").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("OrdSub1").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("OrdSub2").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("ItemCode").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("JakMyung").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("JakSize").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("JakUnit").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("ReqWt").Specific.VALUE = 0;
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SelWt").Specific.VALUE = 0;
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SjNum").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SjLine").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("LotNo").Specific.VALUE = "";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SjPrice").Specific.VALUE = 0;
	//					} else {
	//						//UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(pval.Row).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("OrdGbn").Specific.Select(oMat01.Columns.Item("ItmBsort").Cells.Item(pval.Row).Specific.Selected.VALUE, SAPbouiCOM.BoSearchKey.psk_ByValue);
	//						//UPGRADE_WARNING: oMat01.Columns(BPLId).Cells(pval.Row).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("BPLId").Specific.Select(oMat01.Columns.Item("BPLId").Cells.Item(pval.Row).Specific.Selected.VALUE, SAPbouiCOM.BoSearchKey.psk_ByValue);
	//						//UPGRADE_WARNING: oForm01.Items(CntcCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD();
	//						//UPGRADE_WARNING: oForm01.Items(ItemCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("ItemCode").Specific.VALUE = oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("DocDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
	//						//UPGRADE_WARNING: oForm01.Items(DueDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("DueDate").Specific.VALUE = oMat01.Columns.Item("ReqDate").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items(BaseType).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("BaseType").Specific.VALUE = oMat01.Columns.Item("BaseType").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items(BaseNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("BaseNum").Specific.VALUE = oMat01.Columns.Item("BaseNum").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("OrdMgNum").Specific.VALUE = "";
	//						////기계공구,몰드일경우 작번이 생성되어있음
	//						//UPGRADE_WARNING: oMat01.Columns(BaseType).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (oMat01.Columns.Item("BaseType").Cells.Item(pval.Row).Specific.VALUE == "작번요청") {
	//							////If oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Selected.Value = "105" Or oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Selected.Value = "106" Then
	//							//                        oForm01.Items("OrdNum").Specific.Value = oMat01.Columns("OrdNum").Cells(pval.Row).Specific.Value
	//							//                        oForm01.Items("OrdSub1").Specific.Value = oMat01.Columns("OrdSub1").Cells(pval.Row).Specific.Value
	//							//                        oForm01.Items("OrdSub2").Specific.Value = oMat01.Columns("OrdSub2").Cells(pval.Row).Specific.Value
	//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030H.SetValue("U_OrdNum", 0, oMat01.Columns.Item("OrdNum").Cells.Item(pval.Row).Specific.VALUE);
	//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030H.SetValue("U_OrdSub1", 0, oMat01.Columns.Item("OrdSub1").Cells.Item(pval.Row).Specific.VALUE);
	//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030H.SetValue("U_OrdSub2", 0, oMat01.Columns.Item("OrdSub2").Cells.Item(pval.Row).Specific.VALUE);

	//							//UPGRADE_WARNING: oMat01.Columns(OrdSub1).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if (oMat01.Columns.Item("OrdSub1").Cells.Item(pval.Row).Specific.VALUE == "00") {
	//								//// 메인작번일경우 작명과 규격에 품목명, 규격으로
	//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030H.SetValue("U_JakMyung", 0, MDC_PS_Common.GetValue("SELECT FrgnName FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("OrdNum").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
	//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030H.SetValue("U_JakSize", 0, MDC_PS_Common.GetValue("SELECT U_Size FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("OrdNum").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
	//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030H.SetValue("U_JakUnit", 0, MDC_PS_Common.GetValue("SELECT salUnitMsr FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("OrdNum").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
	//							} else {
	//								//// 서브작번일경우 작명과 규격에 서브작번명, 규격으로
	//								//UPGRADE_WARNING: oForm01.Items(JakMyung).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("JakMyung").Specific.VALUE = oMat01.Columns.Item("JakMyung").Cells.Item(pval.Row).Specific.VALUE;
	//								//UPGRADE_WARNING: oForm01.Items(JakSize).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("JakSize").Specific.VALUE = oMat01.Columns.Item("JakSize").Cells.Item(pval.Row).Specific.VALUE;
	//								//UPGRADE_WARNING: oForm01.Items(JakUnit).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oForm01.Items.Item("JakUnit").Specific.VALUE = oMat01.Columns.Item("JakUnit").Cells.Item(pval.Row).Specific.VALUE;
	//							}


	//							//UPGRADE_WARNING: oMat01.Columns(BaseType).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						} else if (oMat01.Columns.Item("BaseType").Cells.Item(pval.Row).Specific.VALUE == "생산요청") {
	//							////Else '//생산요청번호
	//							//UPGRADE_WARNING: oForm01.Items(OrdMgNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("OrdMgNum").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyyMMdd");
	//							//UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP030_01 ' & oForm01.Items(OrdNum).Specific.VALUE & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("OrdNum").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyyMMdd") + MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" + oForm01.Items.Item("OrdNum").Specific.VALUE + "'");
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("OrdSub1").Specific.VALUE = "00";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("OrdSub2").Specific.VALUE = "000";

	//							//UPGRADE_WARNING: oForm01.Items(JakMyung).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("JakMyung").Specific.VALUE = oMat01.Columns.Item("JakMyung").Cells.Item(pval.Row).Specific.VALUE;
	//							//UPGRADE_WARNING: oForm01.Items(JakSize).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("JakSize").Specific.VALUE = oMat01.Columns.Item("JakSize").Cells.Item(pval.Row).Specific.VALUE;
	//							//UPGRADE_WARNING: oForm01.Items(JakUnit).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("JakUnit").Specific.VALUE = oMat01.Columns.Item("JakUnit").Cells.Item(pval.Row).Specific.VALUE;
	//						}

	//						//UPGRADE_WARNING: oForm01.Items(ReqWt).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("ReqWt").Specific.VALUE = oMat01.Columns.Item("RemainWt").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items(SelWt).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SelWt").Specific.VALUE = oMat01.Columns.Item("RemainWt").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items(SjNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SjNum").Specific.VALUE = oMat01.Columns.Item("ORDRNum").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items(SjLine).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SjLine").Specific.VALUE = oMat01.Columns.Item("RDR1Num").Cells.Item(pval.Row).Specific.VALUE;
	//						//UPGRADE_WARNING: oForm01.Items(LotNo).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("LotNo").Specific.VALUE = MDC_PS_Common.GetValue("SELECT U_LotNo FROM [ORDR] WHERE DocEntry = '" + oMat01.Columns.Item("ORDRNum").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1);
	//						//UPGRADE_WARNING: oForm01.Items(SjPrice).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns(RDR1Num).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("SjPrice").Specific.VALUE = MDC_PS_Common.GetValue("SELECT LineTotal FROM [RDR1] WHERE DocEntry = '" + oMat01.Columns.Item("ORDRNum").Cells.Item(pval.Row).Specific.VALUE + "' AND LineNum = '" + oMat01.Columns.Item("RDR1Num").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1);
	//					}
	//					oForm01.Items.Item("OrdGbn").Enabled = false;
	//					oForm01.Items.Item("BPLId").Enabled = false;
	//					oForm01.Items.Item("ItemCode").Enabled = false;
	//					//UPGRADE_WARNING: oMat01.Columns(BaseType).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oMat01.Columns.Item("BaseType").Cells.Item(pval.Row).Specific.VALUE == "작번요청") {
	//						oForm01.Items.Item("OrdMgNum").Enabled = false;
	//						//UPGRADE_WARNING: oMat01.Columns(BaseType).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					} else if (oMat01.Columns.Item("BaseType").Cells.Item(pval.Row).Specific.VALUE == "생산요청") {
	//						oForm01.Items.Item("OrdMgNum").Enabled = true;
	//					}
	//					oForm01.Freeze(false);
	//					oHasMatrix01 = false;
	//				}
	//			}
	//		} else if (pval.BeforeAction == false) {

	//		}
	//		return;
	//		Raise_EVENT_DOUBLE_CLICK_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		if (pval.BeforeAction == true) {

	//		} else if (pval.BeforeAction == false) {

	//		}
	//		return;
	//		Raise_EVENT_MATRIX_LINK_PRESSED_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		oForm01.Freeze(true);
	//		int i = 0;
	//		bool Exist = false;
	//		string sQry = null;
	//		int TotalAmt = 0;
	//		object ReqCod = null;

	//		SAPbobsCOM.Recordset oRecordSet01 = null;

	//		oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string zQry = null;
	//		SAPbobsCOM.Recordset oRecordset02 = null;
	//		double TotalQty = 0;
	//		decimal useMkg = default(decimal);
	//		if (pval.BeforeAction == true) {
	//			if (pval.ItemChanged == true) {
	//				if ((pval.ItemUID == "Mat02")) {
	//					if ((PS_PP030_Validate("수정02") == false)) {
	//						oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, Strings.Trim(oDS_PS_PP030L.GetValue("U_" + pval.ColUID, pval.Row - 1)));
	//					} else {
	//						if ((pval.ColUID == "ItemCode")) {
	//							////기타작업
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//							if (oMat02.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP030L.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
	//								PS_PP030_AddMatrixRow01((pval.Row));
	//							}
	//							//[2011.2.14] 추가 Begin------------------------------------------------------------------------------------------------------------
	//							//107010002(END BEARING #44),107010004(END BEARING #2) 일경우에는 투입자재를 직접 입력한다.
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if (Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE) == "107010002" | Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE) == "107010004") {
	//								oMat02.Columns.Item("BatchNum").Editable = true;
	//							}
	//							//[2011.2.14] 추가 End--------------------------------------------------------------------------------------------------------------
	//							//[2011.2.14] 추가 Begin------------------------------------------------------------------------------------------------------------
	//						} else if (pval.ColUID == "BatchNum") {
	//							oRecordset02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
	//							oMat02.FlushToDataSource();

	//							//UPGRADE_WARNING: oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							zQry = "EXEC [PS_PP030_06] '" + oMat02.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE + "', '" + oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE + "'";
	//							oRecordset02.DoQuery(zQry);

	//							//                        oDS_PS_PP030L.setValue "U_ItemCode", oRow - 1, oRecordSet02.Fields(0).VALUE
	//							oDS_PS_PP030L.SetValue("U_ItemName", pval.Row - 1, oRecordset02.Fields.Item("ItemName").Value);
	//							oDS_PS_PP030L.SetValue("U_ItemGpCd", pval.Row - 1, oRecordset02.Fields.Item("ItmsGrpCod").Value);
	//							oDS_PS_PP030L.SetValue("U_Unit", pval.Row - 1, oRecordset02.Fields.Item("InvntryUom").Value);
	//							oDS_PS_PP030L.SetValue("U_Weight", pval.Row - 1, oRecordset02.Fields.Item("Quantity").Value);
	//							oMat02.SetLineData(pval.Row);

	//							//UPGRADE_NOTE: oRecordset02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//							oRecordset02 = null;
	//							//[2011.2.14] 추가 End--------------------------------------------------------------------------------------------------------------

	//						} else if (pval.ColUID == "Weight") {
	//							//UPGRADE_WARNING: oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if (oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE < 0) {
	//								oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(0));
	//							} else {
	//								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//							}
	//						} else if ((pval.ColUID == "CntcCode")) {
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030L.SetValue("U_CntcName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
	//							////비고란 추가 안되는 것 수정 - 류영조
	//							//                    ElseIf (pval.ColUID = "Comments") Then
	//							//                        If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Or Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107" Then
	//							//                            Dim TotalQty As Double
	//							//                            For i = 1 To oDS_PS_PP030L.Size - 1
	//							//                                TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
	//							//                            Next
	//							//                            Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty)
	//							//                        End If
	//						} else if ((pval.ColUID == "Comments")) {
	//							if (Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "104" | Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "107") {
	//								//류영조


	//								if (Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "107") {
	//									sQry = "Select IsNull(U_useMkg, 0) From [OITM] Where ItemCode = '" + Strings.Trim(oDS_PS_PP030H.GetValue("U_ItemCode", 0)) + "'";
	//									oRecordSet01.DoQuery(sQry);
	//									useMkg = oRecordSet01.Fields.Item(0).Value / 1000;

	//									for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++) {
	//										TotalQty = TotalQty + Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
	//									}
	//									if (useMkg == 0) {
	//										oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(System.Math.Round(TotalQty, 0)));
	//										//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//										oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//									} else {
	//										oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(System.Math.Round(TotalQty / useMkg, 0)));
	//										//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//										oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//									}
	//								} else {
	//									for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++) {
	//										TotalQty = TotalQty + Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
	//									}
	//									oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty));
	//									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//								}

	//								//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//								oRecordSet01 = null;
	//							}
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//						} else {
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//						}
	//					}
	//				} else if ((pval.ItemUID == "Mat03")) {

	//					if (pval.ColUID == "StdHour" | pval.ColUID == "ReDate") {
	//						oMat03.FlushToDataSource();
	//						//표준공수와 완료요구일은 수정이 가능해야 하므로 Flush 를 함

	//						//표준공수 등록 시
	//						if (pval.ColUID == "StdHour") {
	//							//공정단가 계산_S
	//							//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if (oMat03.Columns.Item("WorkGbn").Cells.Item(pval.Row).Specific.VALUE == "10") {
	//								//UPGRADE_WARNING: oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_CpPrice", pval.Row - 1, Convert.ToString(MDC_PS_Common.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1) * oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE));
	//								//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							} else if (oMat03.Columns.Item("WorkGbn").Cells.Item(pval.Row).Specific.VALUE == "20") {
	//								//UPGRADE_WARNING: oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_CpPrice", pval.Row - 1, Convert.ToString(MDC_PS_Common.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1) * oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE));
	//							}
	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//							//공정단가 계산_E

	//							//합계 계산_S
	//							for (i = 0; i <= oMat03.VisualRowCount - 1; i++) {

	//								//                            Call Sbo_Application.MessageBox(oDS_PS_PP030M.GetValue("U_CpPrice", i))
	//								TotalAmt = TotalAmt + Convert.ToDouble(oDS_PS_PP030M.GetValue("U_CpPrice", i));

	//							}

	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oForm01.Items.Item("Total").Specific.VALUE = TotalAmt;
	//							//합계 계산_E

	//						}

	//					}

	//					//작업일보가 등록된 작지 중에서 공정대분류와 공정중분류는 수정 불가
	//					if (pval.ColUID == "CpBCode" | pval.ColUID == "CpCode") {

	//						if ((PS_PP030_Validate("수정03") == false)) {
	//							oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, Strings.Trim(oDS_PS_PP030M.GetValue("U_" + pval.ColUID, pval.Row - 1)));
	//						} else {

	//							if ((pval.ColUID == "CpBCode")) {
	//								////기타작업
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_CpBName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT Name FROM [@PS_PP001H] WHERE Code = '" + oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
	//								if (oMat03.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP030M.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
	//									PS_PP030_AddMatrixRow02((pval.Row));
	//								}
	//							} else if ((pval.ColUID == "CpCode")) {
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//								//UPGRADE_WARNING: oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_CpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE Code = '" + oMat03.Columns.Item("CpBCode").Cells.Item(pval.Row).Specific.VALUE + "' AND U_CpCode = '" + oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
	//								oDS_PS_PP030M.SetValue("U_StdHour", pval.Row - 1, Convert.ToString(0));
	//								oDS_PS_PP030M.SetValue("U_CpPrice", pval.Row - 1, Convert.ToString(0));
	//								oDS_PS_PP030M.SetValue("U_ResultYN", pval.Row - 1, "Y");
	//								//UPGRADE_WARNING: oMat03.Columns(CpCode).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								if (oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE == "CP50103" | oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE == "CP50106") {
	//									oDS_PS_PP030M.SetValue("U_ReWorkYN", pval.Row - 1, "Y");

	//									//Call oMat03.Columns("ReWorkYN").Cells(pval.Row).Specific.Select(0, psk_Index) '//PK/탈지일때 재작업여부 예
	//								} else {
	//									oDS_PS_PP030M.SetValue("U_ReWorkYN", pval.Row - 1, "N");
	//									//Call oMat03.Columns("ReWorkYN").Cells(pval.Row).Specific.Select(1, psk_Index) '//PK/탈지일때 재작업여부 아니오
	//								}

	//								//                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때
	//								//                            Exist = False
	//								//                            For i = 1 To oMat03.RowCount - 1
	//								//                                If Trim(oDS_PS_PP030M.GetValue("U_CpCode", i - 1)) = "CP50106" Then '//탈지공정이 있으면
	//								//                                    Exist = True
	//								//                                End If
	//								//                            Next
	//								//                            If Exist = True Then
	//								//                                Call oForm01.Items("MulGbn1").Specific.Select("10", psk_ByValue)
	//								//                            ElseIf Exist = False Then
	//								//                                Call oForm01.Items("MulGbn1").Specific.Select("20", psk_ByValue)
	//								//                            End If
	//								//                        End If
	//							////표준공수 입력시
	//							} else if ((pval.ColUID == "StdHour")) {
	//								////공정단가 표시
	//								//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								if (oMat03.Columns.Item("WorkGbn").Cells.Item(pval.Row).Specific.VALUE == "10") {
	//									//UPGRADE_WARNING: oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oDS_PS_PP030M.SetValue("U_CpPrice", pval.Row - 1, Convert.ToString(MDC_PS_Common.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1) * oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE));
	//									//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								} else if (oMat03.Columns.Item("WorkGbn").Cells.Item(pval.Row).Specific.VALUE == "20") {
	//									//UPGRADE_WARNING: oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oDS_PS_PP030M.SetValue("U_CpPrice", pval.Row - 1, Convert.ToString(MDC_PS_Common.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1) * oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE));
	//								}
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);

	//							} else {
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//							}
	//						}
	//					} else if ((pval.ColUID == "FailCode")) {
	//						//UPGRADE_WARNING: oMat03.Columns(ReWorkYN).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (oMat03.Columns.Item("ReWorkYN").Cells.Item(pval.Row).Specific.VALUE == "Y") {
	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030M.SetValue("U_FailName", pval.Row - 1, MDC_PS_Common.GetValue("Select U_SmalName From [@PS_PP003L] Where U_SmalCode = '" + oMat03.Columns.Item("FailCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
	//						}
	//					}
	//				} else {
	//					if ((pval.ItemUID == "DocEntry")) {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oDS_PS_PP030H.SetValue(pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
	//					} else if ((pval.ItemUID == "OrdMgNum")) {
	//						////생산요청,작번요청
	//						//UPGRADE_WARNING: oForm01.Items(BaseType).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if ((oForm01.Items.Item("BaseType").Specific.VALUE == "작번요청")) {
	//							oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oDS_PS_PP030H.GetValue("U_" + pval.ItemUID, 0));
	//						////생산요청이나, 기준문서타입이 없는경우
	//						} else {
	//							//UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if (string.IsNullOrEmpty(oForm01.Items.Item(pval.ItemUID).Specific.VALUE)) {
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
	//								oDS_PS_PP030H.SetValue("U_OrdNum", 0, "");
	//								oDS_PS_PP030H.SetValue("U_OrdSub1", 0, "");
	//								oDS_PS_PP030H.SetValue("U_OrdSub2", 0, "");
	//							} else {
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP030_01 ' & oForm01.Items(pval.ItemUID).Specific.VALUE & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030H.SetValue("U_OrdNum", 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE + MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'"));
	//								oDS_PS_PP030H.SetValue("U_OrdSub1", 0, "00");
	//								oDS_PS_PP030H.SetValue("U_OrdSub2", 0, "000");
	//							}
	//						}
	//					} else if ((pval.ItemUID == "ItemCode")) {
	//						//                    If oForm01.Mode = fm_UPDATE_MODE Then
	//						//                        Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
	//						//                        Call oDS_PS_PP030H.setValue("U_ItemName", 0, MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
	//						//                    Else
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oDS_PS_PP030H.SetValue("U_ItemName", 0, MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'"));
	//						//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////멀티일경우
	//						if ((oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104")) {
	//							oDS_PS_PP030H.SetValue("U_MulGbn1", 0, "");
	//							oDS_PS_PP030H.SetValue("U_MulGbn2", 0, "");
	//							oDS_PS_PP030H.SetValue("U_MulGbn3", 0, "");
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030H.SetValue("U_MulGbn1", 0, MDC_PS_Common.GetValue("SELECT U_Jakup1 FROM [OITM] WHERE ItemCode = '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'", 0, 1));
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030H.SetValue("U_MulGbn2", 0, MDC_PS_Common.GetValue("SELECT U_Jakup2 FROM [OITM] WHERE ItemCode = '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'", 0, 1));
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030H.SetValue("U_MulGbn3", 0, MDC_PS_Common.GetValue("SELECT U_Jakup3 FROM [OITM] WHERE ItemCode = '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'", 0, 1));
	//						} else {
	//							oDS_PS_PP030H.SetValue("U_MulGbn1", 0, "");
	//							oDS_PS_PP030H.SetValue("U_MulGbn2", 0, "");
	//							oDS_PS_PP030H.SetValue("U_MulGbn3", 0, "");
	//						}
	//						PS_PP030_MTX03();
	//						////공정리스트 처리
	//						//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////휘팅,부품이면
	//						if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "102" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "111" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "601" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "602") {
	//							PS_PP030_MTX02();
	//							////투입자재 처리
	//						}
	//						////멀티,엔드베어링의 경우 작명을 업데이트
	//						//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if ((oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107")) {
	//							oDS_PS_PP030H.SetValue("U_JakMyung", 0, oDS_PS_PP030H.GetValue("U_ItemName", 0));
	//						}
	//						//                    End If
	//					} else if ((pval.ItemUID == "SelWt")) {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (Conversion.Val(oForm01.Items.Item(pval.ItemUID).Specific.VALUE) < 0) {
	//							oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(0));
	//							MDC_Com.MDC_GF_Message(ref "수,중량이 올바르지 않습니다.", ref "W");
	//						}
	//						//UPGRADE_WARNING: oForm01.Items(BaseType).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (!string.IsNullOrEmpty(oForm01.Items.Item("BaseType").Specific.VALUE)) {
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							////투입수량이 수주수량보다 크면
	//							if (Conversion.Val(oForm01.Items.Item(pval.ItemUID).Specific.VALUE) > Conversion.Val(oForm01.Items.Item("ReqWt").Specific.VALUE)) {
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oDS_PS_PP030H.SetValue("U_SelWt", 0, oForm01.Items.Item("ReqWt").Specific.VALUE);
	//								MDC_Com.MDC_GF_Message(ref "수,중량이 올바르지 않습니다.", ref "W");
	//							}
	//						}
	//					// 요청자 추가 20180726 황영수
	//					} else if ((pval.ItemUID == "ReqCod")) {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						sQry = "SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + Strings.Trim(oForm01.Items.Item("ReqCod").Specific.VALUE) + "'";
	//						oRecordSet01.DoQuery(sQry);
	//						//UPGRADE_WARNING: oForm01.Items(ReqNam).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oRecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oForm01.Items.Item("ReqNam").Specific.VALUE = oRecordSet01.Fields.Item(0).Value;
	//					} else if ((pval.ItemUID == "CntcCode")) {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oDS_PS_PP030H.SetValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'", 0, 1));
	//					} else {
	//						if (pval.ItemUID == "SItemCod" | pval.ItemUID == "SCardCod") {
	//						} else {
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							oDS_PS_PP030H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
	//						}
	//					}
	//				}
	//				oMat02.LoadFromDataSource();
	//				oMat03.LoadFromDataSource();
	//				oMat02.AutoResizeColumns();
	//				oMat02.AutoResizeColumns();
	//				oForm01.Update();
	//				if (pval.ItemUID == "Mat01") {
	//					oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				} else if (pval.ItemUID == "Mat02") {
	//					oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				} else if (pval.ItemUID == "Mat03") {
	//					oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				} else {
	//					oForm01.Items.Item(pval.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				}
	//			}
	//		} else if (pval.BeforeAction == false) {

	//		}
	//		oForm01.Freeze(false);
	//		return;
	//		Raise_EVENT_VALIDATE_Error:
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		int i = 0;
	//		if (pval.BeforeAction == true) {

	//		} else if (pval.BeforeAction == false) {
	//			PS_PP030_FormItemEnabled();
	//			if (pval.ItemUID == "Mat01") {
	//				oMat01.Clear();
	//				oMat01.FlushToDataSource();
	//				oMat01.LoadFromDataSource();
	//			} else if (pval.ItemUID == "Mat02") {
	//				////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//				for (i = 1; i <= oMat02.VisualRowCount; i++) {
	//					if (i <= oMat02.VisualRowCount) {
	//						//UPGRADE_WARNING: oMat02.Columns(InputGbn).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (string.IsNullOrEmpty(oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.VALUE)) {
	//							oMat02.DeleteRow((i));
	//							i = i - 1;
	//						}
	//					}
	//				}
	//				for (i = 1; i <= oMat02.VisualRowCount; i++) {
	//					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
	//				}
	//				oMat02.FlushToDataSource();
	//				if (oMat02.VisualRowCount == 0) {
	//					PS_PP030_AddMatrixRow01(oMat02.VisualRowCount, ref true);
	//					////UDO방식
	//				} else {
	//					////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
	//					PS_PP030_AddMatrixRow01(oMat02.VisualRowCount);
	//					////UDO방식
	//					////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//				}
	//				////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
	//			} else if (pval.ItemUID == "Mat03") {
	//				PS_PP030_AddMatrixRow02(oMat03.VisualRowCount);
	//				////UDO방식
	//				if (Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "104") {
	//					oMat03.Columns.Item("Sequence").TitleObject.Sortable = true;
	//					oMat03.Columns.Item("Sequence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
	//					oMat03.Columns.Item("Sequence").TitleObject.Sortable = false;
	//					oMat03.FlushToDataSource();
	//				}
	//			}
	//		}
	//		return;
	//		Raise_EVENT_MATRIX_LOAD_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		if (pval.BeforeAction == true) {

	//		} else if (pval.BeforeAction == false) {
	//			PS_PP030_FormResize();
	//		}
	//		return;
	//		Raise_EVENT_RESIZE_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		SAPbouiCOM.DataTable oDataTable01 = null;
	//		if (pval.BeforeAction == true) {

	//		} else if (pval.BeforeAction == false) {
	//			if ((pval.ItemUID == "SItemCod")) {
	//				//UPGRADE_WARNING: pval.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (pval.SelectedObjects == null) {
	//				} else {
	//					//UPGRADE_WARNING: pval.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oDataTable01 = pval.SelectedObjects;
	//					//UPGRADE_WARNING: oDataTable01.Columns().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oForm01.DataSources.UserDataSources.Item("SItemCod").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
	//					//UPGRADE_WARNING: oDataTable01.Columns().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oForm01.DataSources.UserDataSources.Item("SItemNam").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
	//					//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//					oDataTable01 = null;
	//				}
	//			}

	//			if ((pval.ItemUID == "SCardCod")) {
	//				//UPGRADE_WARNING: pval.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (pval.SelectedObjects == null) {
	//				} else {
	//					//UPGRADE_WARNING: pval.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oDataTable01 = pval.SelectedObjects;
	//					//UPGRADE_WARNING: oDataTable01.Columns().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oForm01.DataSources.UserDataSources.Item("SCardCod").Value = oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value;
	//					//UPGRADE_WARNING: oDataTable01.Columns().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					oForm01.DataSources.UserDataSources.Item("SCardNam").Value = oDataTable01.Columns.Item("CardName").Cells.Item(0).Value;
	//					//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//					oDataTable01 = null;
	//				}
	//			}
	//			//        If (pval.ItemUID = "CntcCode") Then
	//			//            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030H", "U_CntcCode,U_CntcName")
	//			//        End If
	//			//        If (pval.ItemUID = "Mat02") Then
	//			//            If (pval.ColUID = "CntcCode") Then
	//			//                If pval.SelectedObjects Is Nothing Then
	//			//                Else
	//			//                    Set oDataTable01 = pval.SelectedObjects
	//			//                    Call oDS_PS_PP030L.setValue("U_CntcCode", pval.Row - 1, oDataTable01.Columns("empID").Cells(0).Value)
	//			//                    Call oDS_PS_PP030L.setValue("U_CntcName", pval.Row - 1, oDataTable01.Columns("firstName").Cells(0).Value & oDataTable01.Columns("lastName").Cells(0).Value)
	//			//                    Set oDataTable01 = Nothing
	//			//                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
	//			//                    oMat02.LoadFromDataSource
	//			//                End If
	//			//            End If
	//			//        End If
	//		}
	//		return;
	//		Raise_EVENT_CHOOSE_FROM_LIST_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}


	//	private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02" | pval.ItemUID == "Mat03") {
	//			if (pval.Row > 0) {
	//				oLastItemUID01 = pval.ItemUID;
	//				oLastColUID01 = pval.ColUID;
	//				oLastColRow01 = pval.Row;
	//			}
	//		} else {
	//			oLastItemUID01 = pval.ItemUID;
	//			oLastColUID01 = "";
	//			oLastColRow01 = 0;
	//		}
	//		if (pval.ItemUID == "Mat01") {
	//			if (pval.Row > 0) {
	//				oMat01Row01 = pval.Row;
	//			}
	//		} else if (pval.ItemUID == "Mat02") {
	//			if (pval.Row > 0) {
	//				oMat02Row02 = pval.Row;
	//			}
	//		} else if (pval.ItemUID == "Mat03") {
	//			if (pval.Row > 0) {
	//				oMat03Row03 = pval.Row;
	//			}
	//		}
	//		return;
	//		Raise_EVENT_GOT_FOCUS_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		if (pval.BeforeAction == true) {
	//		} else if (pval.BeforeAction == false) {
	//			SubMain.RemoveForms(oFormUniqueID01);
	//			//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//			oForm01 = null;
	//			//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//			oMat01 = null;
	//		}
	//		return;
	//		Raise_EVENT_FORM_UNLOAD_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		int i = 0;
	//		double TotalQty = 0;
	//		decimal useMkg = default(decimal);
	//		SAPbobsCOM.Recordset oRecordSet01 = null;
	//		string sQry = null;
	//		if ((oLastColRow01 > 0)) {
	//			if (pval.BeforeAction == true) {
	//				if (oLastItemUID01 == "Mat02") {
	//					if ((PS_PP030_Validate("행삭제02") == false)) {
	//						BubbleEvent = false;
	//						return;
	//					}
	//				} else if (oLastItemUID01 == "Mat03") {
	//					//                If oForm01.Items("OrdGbn").Specific.Selected.Value = "104" Then '//멀티일경우
	//					//                    Call MDC_Com.MDC_GF_Message("멀티게이지는 공정을 변경할수 없습니다.", "W")
	//					//                    BubbleEvent = False
	//					//                    Exit Sub
	//					//                ElseIf oForm01.Items("OrdGbn").Specific.Selected.Value = "107" Then '//엔드베어링일경우
	//					//                    Call MDC_Com.MDC_GF_Message("엔드베어링은 공정을 변경할수 없습니다.", "W")
	//					//                    BubbleEvent = False
	//					//                    Exit Sub
	//					//                End If
	//					//                If oMat03.Columns("CpCode").Cells(oMat03Row03).Specific.Value = "CP30112" Then
	//					//                    Call MDC_Com.MDC_GF_Message("바렐공정은 변경할수 없습니다.", "W")
	//					//                    BubbleEvent = False
	//					//                    Exit Sub
	//					//                End If
	//					//                If oMat03.Columns("CpCode").Cells(oMat03Row03).Specific.Value = "CP30114" Then
	//					//                    Call MDC_Com.MDC_GF_Message("포장공정은 변경할수 없습니다.", "W")
	//					//                    BubbleEvent = False
	//					//                    Exit Sub
	//					//                End If

	//					if ((PS_PP030_Validate("행삭제03") == false)) {
	//						BubbleEvent = false;
	//						return;
	//					}
	//				}
	//			} else if (pval.BeforeAction == false) {
	//				if (oLastItemUID01 == "Mat02") {
	//					for (i = 1; i <= oMat02.VisualRowCount; i++) {
	//						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
	//						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//					}
	//					oMat02.FlushToDataSource();
	//					oDS_PS_PP030L.RemoveRecord(oDS_PS_PP030L.Size - 1);
	//					oMat02.LoadFromDataSource();
	//					if (oMat02.RowCount == 0) {
	//						PS_PP030_AddMatrixRow01(0);
	//					} else {
	//						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP030L.GetValue("U_ItemCode", oMat02.RowCount - 1)))) {
	//							PS_PP030_AddMatrixRow01(oMat02.RowCount);
	//						}
	//					}
	//					if (Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "104" | Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "107") {
	//						//류영조
	//						oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//						if (Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "107") {
	//							sQry = "Select IsNull(U_useMkg, 0) From [OITM] Where ItemCode = '" + Strings.Trim(oDS_PS_PP030H.GetValue("U_ItemCode", 0)) + "'";
	//							oRecordSet01.DoQuery(sQry);
	//							useMkg = oRecordSet01.Fields.Item(0).Value / 1000;

	//							for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++) {
	//								TotalQty = TotalQty + Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
	//							}
	//							if (useMkg == 0) {
	//								oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty));
	//							} else {
	//								oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty / useMkg));
	//							}
	//						} else {
	//							for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++) {
	//								TotalQty = TotalQty + Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
	//							}
	//							oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty));
	//						}

	//						//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//						oRecordSet01 = null;

	//						//                    Dim TotalQty As Double
	//						//                    For i = 1 To oDS_PS_PP030L.Size - 1
	//						//                        TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
	//						//                    Next
	//						//                    Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty)
	//						oMat02.LoadFromDataSource();
	//						oForm01.Update();
	//					}
	//				} else if (oLastItemUID01 == "Mat03") {
	//					for (i = 1; i <= oMat03.VisualRowCount; i++) {
	//						//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
	//						//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						oMat03.Columns.Item("Sequence").Cells.Item(i).Specific.VALUE = i;
	//					}
	//					oMat03.FlushToDataSource();
	//					oDS_PS_PP030M.RemoveRecord(oDS_PS_PP030M.Size - 1);
	//					oMat03.LoadFromDataSource();
	//					if (oMat03.RowCount == 0) {
	//						PS_PP030_AddMatrixRow02(0);
	//					} else {
	//						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP030M.GetValue("U_CpBCode", oMat03.RowCount - 1)))) {
	//							PS_PP030_AddMatrixRow02(oMat03.RowCount);
	//						}
	//					}

	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					////멀티일때
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104") {
	//						//                    Call oForm01.Items("MulGbn1").Specific.Select("20", psk_ByValue)
	//						//                    For i = 1 To oMat03.RowCount - 1
	//						//                        If oMat03.Columns("CpCode").Cells(i).Specific.VALUE = "CP50106" Then '//탈지공정이 있으면
	//						//                            Call oForm01.Items("MulGbn1").Specific.Select("10", psk_ByValue)
	//						//                            Exit For
	//						//                        End If
	//						//                    Next
	//					}
	//				}
	//			}
	//		}
	//		return;
	//		Raise_EVENT_ROW_DELETE_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}


	//	private bool PS_PP030_CreateItems()
	//	{
	//		bool functionReturnValue = false;
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		oForm01.Freeze(true);
	//		string oQuery01 = null;
	//		SAPbobsCOM.Recordset oRecordSet01 = null;
	//		oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01");
	//		oDS_PS_PP030H = oForm01.DataSources.DBDataSources("@PS_PP030H");
	//		oDS_PS_PP030L = oForm01.DataSources.DBDataSources("@PS_PP030L");
	//		oDS_PS_PP030M = oForm01.DataSources.DBDataSources("@PS_PP030M");

	//		oMat01 = oForm01.Items.Item("Mat01").Specific;
	//		oMat02 = oForm01.Items.Item("Mat02").Specific;
	//		oMat03 = oForm01.Items.Item("Mat03").Specific;
	//		oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
	//		oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
	//		oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
	//		oMat01.AutoResizeColumns();
	//		oMat02.AutoResizeColumns();
	//		oMat03.AutoResizeColumns();

	//		oForm01.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
	//		oForm01.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
	//		oForm01.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
	//		oForm01.DataSources.UserDataSources.Add("ReqType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
	//		oForm01.DataSources.UserDataSources.Add("SItemCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
	//		oForm01.DataSources.UserDataSources.Add("SItemNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
	//		oForm01.DataSources.UserDataSources.Add("SCardCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
	//		oForm01.DataSources.UserDataSources.Add("SCardNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
	//		oForm01.DataSources.UserDataSources.Add("ReqCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
	//		oForm01.DataSources.UserDataSources.Add("ReqNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
	//		oForm01.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
	//		oForm01.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
	//		oForm01.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);

	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("ReqType").Specific.DataBind.SetBound(true, "", "ReqType");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("SItemCod").Specific.DataBind.SetBound(true, "", "SItemCod");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("SItemNam").Specific.DataBind.SetBound(true, "", "SItemNam");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("SCardCod").Specific.DataBind.SetBound(true, "", "SCardCod");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("SCardNam").Specific.DataBind.SetBound(true, "", "SCardNam");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("ReqCod").Specific.DataBind.SetBound(true, "", "ReqCod");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("ReqNam").Specific.DataBind.SetBound(true, "", "ReqNam");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.GroupWith 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("Opt01").Specific.GroupWith("Opt02");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.GroupWith 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("Opt01").Specific.GroupWith("Opt03");

	//		oForm01.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");
	//		//    oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = 0

	//		//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oRecordSet01 = null;
	//		oForm01.Freeze(false);
	//		return functionReturnValue;
	//		PS_PP030_CreateItems_Error:
	//		//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oRecordSet01 = null;
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		return functionReturnValue;
	//	}

	//	public void PS_PP030_ComboBox_Setting()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement


	//		string sQry = null;

	//		oForm01.Freeze(true);
	//		////콤보에 기본값설정
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ReqType", "", "10", "계획생산요청");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ReqType", "", "20", "수주생산요청");
	//		MDC_PS_Common.Combo_ValidValues_SetValueItem(ref (oForm01.Items.Item("ReqType").Specific), "PS_PP030", "ReqType", true);

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ReqType", "10", "계획생산요청");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ReqType", "20", "수주생산요청");
	//		MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("ReqType"), "PS_PP030", "Mat01", "ReqType");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "BasicGub", "", "10", "통합");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "BasicGub", "", "20", "비통합");
	//		MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm01.Items.Item("BasicGub").Specific), "PS_PP030", "BasicGub");
	//		//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("BasicGub").Specific.Select("비통합", SAPbouiCOM.BoSearchKey.psk_ByValue);

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn1", "", "10", "탈지");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn1", "", "20", "비탈지");
	//		MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm01.Items.Item("MulGbn1").Specific), "PS_PP030", "MulGbn1");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn2", "", "10", "시계");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn2", "", "20", "반시계");
	//		MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm01.Items.Item("MulGbn2").Specific), "PS_PP030", "MulGbn2");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn3", "", "10", "배면");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn3", "", "20", "상면");
	//		MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm01.Items.Item("MulGbn3").Specific), "PS_PP030", "MulGbn3");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "10", "일반");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "20", "원재료");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "30", "스크랩");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn2", "20", "원재료");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn2", "30", "스크랩");
	//		MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns.Item("InputGbn"), "PS_PP030", "Mat02", "InputGbn");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "10", "청구");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "20", "잔재");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "30", "취소");
	//		MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns.Item("ProcType"), "PS_PP030", "Mat02", "ProcType");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "10", "자가");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "20", "정밀");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "30", "외주");
	//		MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("WorkGbn"), "PS_PP030", "Mat03", "WorkGbn");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ResultYN", "Y", "예");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ResultYN", "N", "아니오");
	//		MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ResultYN"), "PS_PP030", "Mat03", "ResultYN");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReWorkYN", "Y", "예");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReWorkYN", "N", "아니오");
	//		MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ReWorkYN"), "PS_PP030", "Mat03", "ReWorkYN");

	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReportYN", "Y", "예");
	//		MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReportYN", "N", "아니오");
	//		MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ReportYN"), "PS_PP030", "Mat03", "ReportYN");

	//		//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ItemCode", "01", "완제품")
	//		//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ItemCode", "02", "반제품")
	//		//    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP030", "Mat01", "ItemCode")
	//		//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ItemCode", "", "01", "완제품")
	//		//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ItemCode", "", "02", "반제품")
	//		//    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP030", "ItemCode")

	//		MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("SBPLId").Specific), ref "SELECT BPLId, BPLName FROM OBPL order by BPLId", ref "", ref false, ref true);
	//		MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItmBsort").Specific), ref "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", ref "", ref false, ref true);
	//		MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItmMsort").Specific), ref "SELECT PSH_ITMMSORT.U_Code, PSH_ITMMSORT.U_CodeName FROM [@PSH_ITMMSORT] PSH_ITMMSORT LEFT JOIN [@PSH_ITMBSORT] PSH_ITMBSORT ON PSH_ITMBSORT.Code = PSH_ITMMSORT.U_rCode WHERE PSH_ITMBSORT.U_PudYN = 'Y'", ref "", ref false, ref true);
	//		MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("Mark").Specific), ref "SELECT Code, Name FROM [@PSH_MARK] order by Code", ref "", ref false, ref true);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
	//		MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("OrdGbn").Specific), ref "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", ref "", ref false, ref false);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oForm01.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
	//		MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("BPLId").Specific), ref "SELECT BPLId, BPLName FROM OBPL order by BPLId", ref "", ref false, ref false);

	//		//    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
	//		//    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
	//		MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId");
	//		MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code");
	//		MDC_Com.MDC_GP_MatrixSetMatComboList(oMat02.Columns.Item("ItemGpCd"), "SELECT ItmsGrpCod, ItmsGrpNam FROM [OITB]");
	//		MDC_Com.MDC_GP_MatrixSetMatComboList(oMat03.Columns.Item("Unit"), "SELECT Code, Name FROM [@PSH_CPUOM]");

	//		//재청구사유(라인)
	//		sQry = "        SELECT      U_Minor,";
	//		sQry = sQry + "             U_CdName";
	//		sQry = sQry + " FROM        [@PS_SY001L]";
	//		sQry = sQry + " WHERE       Code = 'P203'";
	//		sQry = sQry + "             AND U_UseYN = 'Y'";
	//		sQry = sQry + "             AND U_Minor <> 'A'";
	//		sQry = sQry + " ORDER BY    U_Seq";
	//		MDC_Com.MDC_GP_MatrixSetMatComboList(oMat02.Columns.Item("RCode"), sQry);

	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_ComboBox_Setting_Error:
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public void PS_PP030_CF_ChooseFromList()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		////ChooseFromList 설정
	//		SAPbouiCOM.ChooseFromListCollection oCFLs = null;
	//		SAPbouiCOM.Conditions oCons = null;
	//		SAPbouiCOM.Condition oCon = null;
	//		SAPbouiCOM.ChooseFromList oCFL = null;
	//		SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
	//		SAPbouiCOM.EditText oEdit = null;
	//		SAPbouiCOM.Column oColumn = null;

	//		//    Set oEdit = oForm01.Items("CntcCode").Specific
	//		//    Set oCFLs = oForm01.ChooseFromLists
	//		//    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
	//		//
	//		//    oCFLCreationParams.ObjectType = lf_Employee
	//		//    oCFLCreationParams.uniqueID = "CFLEMPLOYEE"
	//		//    oCFLCreationParams.MultiSelection = False
	//		//    Set oCFL = oCFLs.Add(oCFLCreationParams)
	//		//
	//		//'    Set oCons = oCFL.GetConditions()
	//		//'    Set oCon = oCons.Add()
	//		//'    oCon.Alias = "CardType"
	//		//'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
	//		//'    oCon.CondVal = "C"
	//		//'    oCFL.SetConditions oCons
	//		//
	//		//    oEdit.ChooseFromListUID = "CFLEMPLOYEE"
	//		//    oEdit.ChooseFromListAlias = "empID"
	//		//
	//		//    Set oColumn = oMat02.Columns("CntcCode")
	//		//    Set oCFLs = oForm01.ChooseFromLists
	//		//    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
	//		//
	//		//    oCFLCreationParams.ObjectType = lf_Employee
	//		//    oCFLCreationParams.uniqueID = "CFLEMPLOYEE2"
	//		//    oCFLCreationParams.MultiSelection = False
	//		//    Set oCFL = oCFLs.Add(oCFLCreationParams)
	//		//
	//		//'    Set oCons = oCFL.GetConditions()
	//		//'    Set oCon = oCons.Add()
	//		//'    oCon.Alias = "CardType"
	//		//'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
	//		//'    oCon.CondVal = "C"
	//		//'    oCFL.SetConditions oCons
	//		//
	//		//    oColumn.ChooseFromListUID = "CFLEMPLOYEE2"
	//		//    oColumn.ChooseFromListAlias = "empID"

	//		oEdit = oForm01.Items.Item("SCardCod").Specific;
	//		oCFLs = oForm01.ChooseFromLists;
	//		oCFLCreationParams = SubMain.Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

	//		oCFLCreationParams.ObjectType = "2";
	//		oCFLCreationParams.UniqueID = "CFLSCARDCOD";
	//		oCFLCreationParams.MultiSelection = false;
	//		oCFL = oCFLs.Add(oCFLCreationParams);

	//		//    Set oCons = oCFL.GetConditions()
	//		//    Set oCon = oCons.Add()
	//		//    oCon.Alias = "CardType"
	//		//    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
	//		//    oCon.CondVal = "C"
	//		//    oCFL.SetConditions oCons

	//		oEdit.ChooseFromListUID = "CFLSCARDCOD";
	//		oEdit.ChooseFromListAlias = "CardCode";

	//		oEdit = oForm01.Items.Item("SItemCod").Specific;
	//		oCFLs = oForm01.ChooseFromLists;
	//		oCFLCreationParams = SubMain.Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

	//		oCFLCreationParams.ObjectType = "4";
	//		oCFLCreationParams.UniqueID = "CFLSITEMCOD";
	//		oCFLCreationParams.MultiSelection = false;
	//		oCFL = oCFLs.Add(oCFLCreationParams);

	//		//    Set oCons = oCFL.GetConditions()
	//		//    Set oCon = oCons.Add()
	//		//    oCon.Alias = "CardType"
	//		//    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
	//		//    oCon.CondVal = "C"
	//		//    oCFL.SetConditions oCons

	//		oEdit.ChooseFromListUID = "CFLSITEMCOD";
	//		oEdit.ChooseFromListAlias = "ItemCode";

	//		return;
	//		PS_PP030_CF_ChooseFromList_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public void PS_PP030_FormItemEnabled()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement


	//		string sQry01 = null;
	//		string sQry02 = null;

	//		oForm01.Freeze(true);
	//		if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
	//			////각모드에따른 아이템설정
	//			oForm01.EnableMenu("1281", true);
	//			////찾기
	//			oForm01.EnableMenu("1282", false);
	//			////추가
	//			oForm01.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			oForm01.Items.Item("OrdGbn").Enabled = true;
	//			oForm01.Items.Item("BPLId").Enabled = true;
	//			oForm01.Items.Item("DocDate").Enabled = true;
	//			//2011.12.05 송명규 수정(True에서 False로) False에서 True로 환원(2017.02.21 송명규)
	//			oForm01.Items.Item("DueDate").Enabled = true;
	//			oForm01.Items.Item("ItemCode").Enabled = true;
	//			oForm01.Items.Item("CntcCode").Enabled = true;
	//			oForm01.Items.Item("BasicGub").Enabled = true;
	//			oForm01.Items.Item("MulGbn1").Enabled = false;
	//			oForm01.Items.Item("MulGbn2").Enabled = false;
	//			oForm01.Items.Item("MulGbn3").Enabled = false;
	//			oForm01.Items.Item("DocEntry").Enabled = false;
	//			oForm01.Items.Item("OrdMgNum").Enabled = true;
	//			oForm01.Items.Item("ReqWt").Enabled = false;
	//			oForm01.Items.Item("SelWt").Enabled = true;
	//			oForm01.Items.Item("Mat01").Enabled = true;
	//			oForm01.Items.Item("Mat02").Enabled = true;
	//			oForm01.Items.Item("Mat03").Enabled = true;
	//			oForm01.Items.Item("Button01").Enabled = true;
	//			oForm01.Items.Item("1").Enabled = true;

	//			//        Call oForm01.Items("SBPLId").Specific.Select(0, psk_Index)
	//			//        Call oForm01.Items("ItmBsort").Specific.Select(0, psk_Index)
	//			//        Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)
	//			//        Call oForm01.Items("ReqType").Specific.Select(0, psk_Index)
	//			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
	//			//UPGRADE_WARNING: oForm01.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("DocDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
	//			//UPGRADE_WARNING: oForm01.Items(DueDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("DueDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("SItemCod").Specific.VALUE = "";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("SCardCod").Specific.VALUE = "";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("OrdMgNum").Specific.VALUE = "";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("OrdNum").Specific.VALUE = "";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("OrdSub1").Specific.VALUE = "";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("OrdSub2").Specific.VALUE = "";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("BasicGub").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
	//			//UPGRADE_WARNING: oForm01.Items(OrdMgNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("OrdMgNum").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
	//			//UPGRADE_WARNING: oForm01.Items(CntcCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD();
	//			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);
	//			PS_PP030_FormClear();
	//			////UDO방식
	//			oMat02.Columns.Item("BatchNum").Editable = false;
	//		} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
	//			////각모드에따른 아이템설정
	//			oForm01.EnableMenu("1281", false);
	//			////찾기
	//			oForm01.EnableMenu("1282", true);
	//			////추가
	//			oForm01.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			oForm01.Items.Item("OrdGbn").Enabled = true;
	//			oForm01.Items.Item("BPLId").Enabled = true;
	//			oForm01.Items.Item("DocDate").Enabled = true;
	//			//2011.12.05 송명규 수정(True에서 False로) False에서 True로 환원(2017.02.21 송명규)
	//			oForm01.Items.Item("DueDate").Enabled = true;
	//			oForm01.Items.Item("ItemCode").Enabled = true;
	//			oForm01.Items.Item("CntcCode").Enabled = true;
	//			oForm01.Items.Item("BasicGub").Enabled = true;
	//			oForm01.Items.Item("MulGbn1").Enabled = false;
	//			oForm01.Items.Item("MulGbn2").Enabled = false;
	//			oForm01.Items.Item("MulGbn3").Enabled = false;
	//			oForm01.Items.Item("DocEntry").Enabled = true;
	//			oForm01.Items.Item("OrdMgNum").Enabled = true;
	//			oForm01.Items.Item("ReqWt").Enabled = false;
	//			oForm01.Items.Item("SelWt").Enabled = true;
	//			oForm01.Items.Item("Mat01").Enabled = false;
	//			oForm01.Items.Item("Mat02").Enabled = false;
	//			oForm01.Items.Item("Mat03").Enabled = false;
	//			oForm01.Items.Item("Button01").Enabled = true;
	//			oForm01.Items.Item("1").Enabled = true;
	//			oMat02.Columns.Item("BatchNum").Editable = false;
	//		} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
	//			oForm01.EnableMenu("1281", true);
	//			////찾기
	//			oForm01.EnableMenu("1282", true);
	//			////추가
	//			oMat02.Columns.Item("BatchNum").Editable = false;
	//			////각모드에따른 아이템설정
	//			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP030H] WHERE DocEntry = ' & Trim(oDS_PS_PP030H.GetValue(DocEntry, 0)) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" + Strings.Trim(oDS_PS_PP030H.GetValue("DocEntry", 0)) + "'", 0, 1) == "Y") {
	//				oForm01.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				oForm01.Items.Item("OrdGbn").Enabled = false;
	//				oForm01.Items.Item("BPLId").Enabled = false;
	//				oForm01.Items.Item("DocDate").Enabled = false;
	//				oForm01.Items.Item("DueDate").Enabled = false;
	//				oForm01.Items.Item("ItemCode").Enabled = false;
	//				oForm01.Items.Item("CntcCode").Enabled = false;
	//				oForm01.Items.Item("MulGbn1").Enabled = false;
	//				oForm01.Items.Item("MulGbn2").Enabled = false;
	//				oForm01.Items.Item("MulGbn3").Enabled = false;
	//				oForm01.Items.Item("DocEntry").Enabled = false;
	//				oForm01.Items.Item("OrdMgNum").Enabled = false;
	//				oForm01.Items.Item("ReqWt").Enabled = false;
	//				oForm01.Items.Item("SelWt").Enabled = false;
	//				oForm01.Items.Item("Mat01").Enabled = false;
	//				oForm01.Items.Item("Mat02").Enabled = false;
	//				oForm01.Items.Item("Mat03").Enabled = false;
	//				oForm01.Items.Item("Button01").Enabled = false;
	//				oForm01.Items.Item("1").Enabled = false;
	//			} else {
	//				oForm01.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				oForm01.Items.Item("Mat01").Enabled = true;
	//				oForm01.Items.Item("Mat02").Enabled = true;
	//				oForm01.Items.Item("Mat03").Enabled = true;
	//				oForm01.Items.Item("Button01").Enabled = true;
	//				oForm01.Items.Item("1").Enabled = true;

	//				oForm01.Items.Item("DocEntry").Enabled = false;
	//				oForm01.Items.Item("OrdGbn").Enabled = false;
	//				oForm01.Items.Item("BPLId").Enabled = false;
	//				oForm01.Items.Item("ItemCode").Enabled = false;
	//				oForm01.Items.Item("OrdMgNum").Enabled = false;

	//				//실적(작업일보)문서가 없고 원가 상의 재공에 투입된 자료가 아니라면 아래 필드의 데이터는 수정(2017.02.21 송명규)
	//				//실적 자료 조회용 쿼리
	//				sQry01 = "          SELECT  COUNT(*)";
	//				sQry01 = sQry01 + " FROM    [@PS_PP040H] AS T0";
	//				sQry01 = sQry01 + "         INNER JOIN";
	//				sQry01 = sQry01 + "         [@PS_PP040L] AS T1";
	//				sQry01 = sQry01 + "             ON T0.DocEntry = T1.DocEntry";
	//				sQry01 = sQry01 + " WHERE   T0.Canceled = 'N'";
	//				sQry01 = sQry01 + "         AND T1.U_PP030HNo = " + Strings.Trim(oDS_PS_PP030H.GetValue("DocEntry", 0));

	//				//원가 자료 조회용 쿼리
	//				sQry02 = "          SELECT  COUNT(*)";
	//				sQry02 = sQry02 + " FROM    [@PS_CO130L] AS T0";
	//				sQry02 = sQry02 + " WHERE   T0.U_POEntry = " + Strings.Trim(oDS_PS_PP030H.GetValue("DocEntry", 0));

	//				//If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oDS_PS_PP030H.GetValue("DocEntry", 0)) & "'") = 0 Then
	//				//UPGRADE_WARNING: MDC_PS_Common.GetValue(sQry02) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: MDC_PS_Common.GetValue(sQry01) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//실적(작업일보) 및 원가계산된 자료가 없으면 아래 자료는 수정가능
	//				if (MDC_PS_Common.GetValue(sQry01) == 0 | MDC_PS_Common.GetValue(sQry02) == 0) {
	//					oForm01.Items.Item("DocDate").Enabled = true;
	//					//2011.12.05 송명규 수정(True에서 False로), False에서 True로 환원(2017.02.21 송명규)
	//					oForm01.Items.Item("DueDate").Enabled = true;
	//					oForm01.Items.Item("CntcCode").Enabled = true;
	//					oForm01.Items.Item("SelWt").Enabled = true;
	//					oForm01.Items.Item("ReqWt").Enabled = true;
	//					////작업구분이 멀티일때만
	//					if (Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "104") {
	//						oForm01.Items.Item("BasicGub").Enabled = true;
	//						oForm01.Items.Item("MulGbn1").Enabled = true;
	//						oForm01.Items.Item("MulGbn2").Enabled = true;
	//						oForm01.Items.Item("MulGbn3").Enabled = true;
	//					} else {
	//						oForm01.Items.Item("BasicGub").Enabled = false;
	//						oForm01.Items.Item("MulGbn1").Enabled = false;
	//						oForm01.Items.Item("MulGbn2").Enabled = false;
	//						oForm01.Items.Item("MulGbn3").Enabled = false;
	//					}
	//				////실적 등록 및 원가 계산된 자료가 하나라도 있으면
	//				} else {
	//					//멀티 작번인 경우
	//					if (Strings.Left(Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdNum", 0)), 1) == "E") {
	//						oForm01.Items.Item("DocDate").Enabled = false;
	//						oForm01.Items.Item("DueDate").Enabled = false;
	//						oForm01.Items.Item("CntcCode").Enabled = false;
	//						oForm01.Items.Item("ReqWt").Enabled = true;
	//						oForm01.Items.Item("SelWt").Enabled = true;
	//						oForm01.Items.Item("MulGbn1").Enabled = true;
	//						oForm01.Items.Item("MulGbn2").Enabled = false;
	//						oForm01.Items.Item("MulGbn3").Enabled = false;
	//					} else {
	//						oForm01.Items.Item("DocDate").Enabled = false;
	//						oForm01.Items.Item("DueDate").Enabled = false;
	//						oForm01.Items.Item("CntcCode").Enabled = false;
	//						oForm01.Items.Item("ReqWt").Enabled = false;
	//						oForm01.Items.Item("SelWt").Enabled = false;
	//						oForm01.Items.Item("MulGbn1").Enabled = true;
	//						oForm01.Items.Item("MulGbn2").Enabled = false;
	//						oForm01.Items.Item("MulGbn3").Enabled = false;
	//					}

	//				}

	//				if ((Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "107")) {
	//					oMat02.Columns.Item("InputGbn").Editable = true;
	//				} else {
	//					oMat02.Columns.Item("InputGbn").Editable = false;
	//				}

	//				if ((Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "105" | Strings.Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) == "106")) {
	//					oMat02.Columns.Item("Weight").Editable = true;
	//				} else {
	//					oMat02.Columns.Item("Weight").Editable = false;
	//				}

	//				//            If (Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Or Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107") Then
	//				//                oMat03.Columns("CpBCode").Editable = False
	//				//                oMat03.Columns("CpCode").Editable = False
	//				//                oMat03.Columns("ResultYN").Editable = False
	//				//                oMat03.Columns("ReportYN").Editable = False
	//				//            Else
	//				//                oMat03.Columns("CpBCode").Editable = True
	//				//                oMat03.Columns("CpCode").Editable = True
	//				//                oMat03.Columns("ResultYN").Editable = True
	//				//                oMat03.Columns("ReportYN").Editable = True
	//				//            End If
	//			}
	//		}
	//		//    oMat01.AutoResizeColumns
	//		//    oMat02.AutoResizeColumns
	//		//    oMat03.AutoResizeColumns
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_FormItemEnabled_Error:
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public void PS_PP030_AddMatrixRow01(int oRow, ref bool RowIserted = false)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		oForm01.Freeze(true);
	//		////행추가여부
	//		if (RowIserted == false) {
	//			oDS_PS_PP030L.InsertRecord((oRow));
	//		}
	//		oMat02.AddRow();
	//		oDS_PS_PP030L.Offset = oRow;
	//		oDS_PS_PP030L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
	//		//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		////엔드베어링은 투입구분,원재료,스크랩
	//		if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//			oDS_PS_PP030L.SetValue("U_InputGbn", oRow, "20");
	//		////나머지경우는 일반으로 선택
	//		} else {
	//			oDS_PS_PP030L.SetValue("U_InputGbn", oRow, "10");
	//		}
	//		oDS_PS_PP030L.SetValue("U_ProcType", oRow, "20");
	//		oMat02.LoadFromDataSource();
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_AddMatrixRow01_Error:
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_AddMatrixRow01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public void PS_PP030_AddMatrixRow02(int oRow, ref bool RowIserted = false)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		oForm01.Freeze(true);
	//		////행추가여부
	//		if (RowIserted == false) {
	//			oDS_PS_PP030M.InsertRecord((oRow));
	//		}
	//		oMat03.AddRow();
	//		oDS_PS_PP030M.Offset = oRow;
	//		oDS_PS_PP030M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
	//		oDS_PS_PP030M.SetValue("U_Sequence", oRow, Convert.ToString(oRow + 1));
	//		oDS_PS_PP030M.SetValue("U_WorkGbn", oRow, "10");
	//		oDS_PS_PP030M.SetValue("U_ReWorkYN", oRow, "N");
	//		oDS_PS_PP030M.SetValue("U_ResultYN", oRow, "N");
	//		oDS_PS_PP030M.SetValue("U_ReportYN", oRow, "Y");
	//		oMat03.LoadFromDataSource();
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_AddMatrixRow02_Error:
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_AddMatrixRow02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public void PS_PP030_FormClear()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		string DocEntry = null;
	//		//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_PP030'", ref "");
	//		if (Convert.ToDouble(DocEntry) == 0) {
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("DocEntry").Specific.VALUE = 1;
	//		} else {
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("DocEntry").Specific.VALUE = DocEntry;
	//		}
	//		//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if ((oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "105" & oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "106" & oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "선택")) {
	//			//UPGRADE_WARNING: oForm01.Items(OrdMgNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (!string.IsNullOrEmpty(oForm01.Items.Item("OrdMgNum").Specific.VALUE)) {
	//				//UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP030_01 ' & oForm01.Items(OrdMgNum).Specific.VALUE & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				oForm01.Items.Item("OrdNum").Specific.VALUE = oForm01.Items.Item("OrdMgNum").Specific.VALUE + MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" + oForm01.Items.Item("OrdMgNum").Specific.VALUE + "'");
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				oForm01.Items.Item("OrdSub1").Specific.VALUE = "00";
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				oForm01.Items.Item("OrdSub2").Specific.VALUE = "000";
	//			}
	//		}
	//		return;
	//		PS_PP030_FormClear_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void PS_PP030_EnableMenus()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		////메뉴활성화
	//		//    Call oForm01.EnableMenu("1288", True)
	//		//    Call oForm01.EnableMenu("1289", True)
	//		//    Call oForm01.EnableMenu("1290", True)
	//		//    Call oForm01.EnableMenu("1291", True)
	//		////Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
	//		MDC_Com.MDC_GP_EnableMenus(ref oForm01, false, false, true, true, false, true, true, true, true,
	//		true, true, false, false, false, false, false);
	//		////메뉴설정
	//		return;
	//		PS_PP030_EnableMenus_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void PS_PP030_SetDocument(string oFromDocEntry01)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		if ((string.IsNullOrEmpty(oFromDocEntry01))) {
	//			PS_PP030_FormItemEnabled();
	//			PS_PP030_AddMatrixRow01(0, ref true);
	//			////UDO방식일때
	//			PS_PP030_AddMatrixRow02(0, ref true);
	//			////UDO방식일때
	//		} else {
	//			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
	//			PS_PP030_FormItemEnabled();
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oForm01.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
	//			oForm01.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//		}
	//		return;
	//		PS_PP030_SetDocument_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public bool PS_PP030_DataValidCheck()
	//	{
	//		bool functionReturnValue = false;
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		functionReturnValue = false;
	//		object i = null;
	//		int j = 0;
	//		bool CP30112 = false;
	//		bool CP30114 = false;

	//		short Lot104Exsits = 0;
	//		//멀티구전산 이전 lotno
	//		string query01 = null;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		//    Dim sQry As String
	//		//    Dim oRecordSet01 As SAPbobsCOM.Recordset
	//		//
	//		//    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

	//		if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
	//			PS_PP030_FormClear();
	//		}

	//		//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "선택") {
	//			SubMain.Sbo_Application.SetStatusBarMessage("작지구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//			oForm01.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}
	//		//UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (oForm01.Items.Item("BPLId").Specific.Selected.VALUE == "선택") {
	//			SubMain.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//			oForm01.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}
	//		//UPGRADE_WARNING: oForm01.Items(OrdSub2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		//UPGRADE_WARNING: oForm01.Items(OrdSub1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		//UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (string.IsNullOrEmpty(oForm01.Items.Item("OrdNum").Specific.VALUE) | string.IsNullOrEmpty(oForm01.Items.Item("OrdSub1").Specific.VALUE) | string.IsNullOrEmpty(oForm01.Items.Item("OrdSub2").Specific.VALUE)) {
	//			SubMain.Sbo_Application.SetStatusBarMessage("작지번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//			oForm01.Items.Item("OrdMgNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}
	//		//UPGRADE_WARNING: oForm01.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (string.IsNullOrEmpty(oForm01.Items.Item("DocDate").Specific.VALUE)) {
	//			SubMain.Sbo_Application.SetStatusBarMessage("지시일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//			oForm01.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}
	//		//UPGRADE_WARNING: oForm01.Items(ItemCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (string.IsNullOrEmpty(oForm01.Items.Item("ItemCode").Specific.VALUE)) {
	//			SubMain.Sbo_Application.SetStatusBarMessage("품목코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//			oForm01.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (Conversion.Val(oForm01.Items.Item("SelWt").Specific.VALUE) <= 0) {
	//			SubMain.Sbo_Application.SetStatusBarMessage("지시수,중량이 올바르지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//			oForm01.Items.Item("SelWt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}

	//		//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104") {
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = "Exec [PS_PP030_09] '" + oForm01.Items.Item("ItemCode").Specific.VALUE + "','";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + Strings.Right(Strings.Left(oForm01.Items.Item("DocDate").Specific.VALUE, 6), 4) + Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE) + Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.VALUE) + "'";
	//			RecordSet01.DoQuery(query01);
	//			if (RecordSet01.Fields.Item(0).Value == 1) {
	//				SubMain.Sbo_Application.MessageBox("원재료 사용량을 초과하였습니다. 담당자에게 문의하세요. (" + RecordSet01.Fields.Item(1).Value + " kg)");
	//				functionReturnValue = false;
	//				return functionReturnValue;
	//			}
	//		}

	//		//기계공구일경우 작번등록과 일자 비교_S(2017.02.21 송명규 추가)
	//		//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105") {

	//			if (PS_PP030_CheckDate() == false) {
	//				SubMain.Sbo_Application.SetStatusBarMessage("작업지시등록일은 작번등록일과 같거나 늦어야합니다. 확인하십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//				functionReturnValue = false;
	//				return functionReturnValue;
	//			}

	//		}
	//		//기계공구일경우 작번등록과 일자 비교_E(2017.02.21 송명규 추가)

	//		//    '마감상태 체크_S(2017.11.23 송명규 추가), 순환품인 경우 공정등록을 3달 후에도 등록하는 경우가 있어 주석 처리(2017.11.28 송명규)
	//		//    If MDC_PS_Common.Check_Finish_Status(oForm01.Items("BPLId").Specific.VALUE, oForm01.Items("DocDate").Specific.VALUE) = False Then
	//		//        Call Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 지시일자를 확인하고, 회계부서로 문의하세요.", bmt_Short, True)
	//		//        PS_PP030_DataValidCheck = False
	//		//        Exit Function
	//		//    End If
	//		//    '마감상태 체크_E(2017.11.23 송명규 추가)

	//		//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		////휘팅일경우
	//		if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101") {
	//			//        CP30112 = False
	//			//        CP30114 = False
	//			//        For i = 1 To oMat03.VisualRowCount
	//			//            If (oMat03.Columns("CpBCode").Cells(i).Specific.Value = "CP301") Then '//공정이 휘팅이면
	//			//                If oMat03.Columns("CpCode").Cells(i).Specific.Value = "CP30112" Then '//바렐공정
	//			//                    CP30112 = True
	//			//                End If
	//			//                If oMat03.Columns("CpCode").Cells(i).Specific.Value = "CP30114" Then '//포장공정
	//			//                    CP30114 = True
	//			//                End If
	//			//            End If
	//			//        Next
	//			//        If CP30112 <> True Or CP30114 <> True Then
	//			//            MDC_Com.MDC_GF_Message "휘팅공정은 바렐,포장공정은 필수 입니다.", "W"
	//			//            PS_PP030_DataValidCheck = False
	//			//            Exit Function
	//			//        End If
	//		}
	//		////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//		//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		////멀티는 투입자재라인이 존재해야함
	//		if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104") {
	//			////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = "SELECT Count(*) FROM Z_DSMDFRY Where lotno = '" + oForm01.Items.Item("OrdNum").Specific.VALUE + "'";
	//			RecordSet01.DoQuery(query01);
	//			//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			Lot104Exsits = RecordSet01.Fields.Item(0).Value;

	//			if (Lot104Exsits == 0) {
	//				if (oMat02.VisualRowCount <= 1) {
	//					SubMain.Sbo_Application.SetStatusBarMessage("투입자재라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//					functionReturnValue = false;
	//					return functionReturnValue;
	//				}
	//			}
	//			////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//		}
	//		////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
	//		if (oMat03.VisualRowCount <= 1) {
	//			SubMain.Sbo_Application.SetStatusBarMessage("공정리스트 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}

	//		for (i = 1; i <= oMat02.VisualRowCount - 1; i++) {
	//			//UPGRADE_WARNING: oMat02.Columns(InputGbn).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if ((oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.Selected == null)) {
	//				SubMain.Sbo_Application.SetStatusBarMessage("투입구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//				oMat02.Columns.Item("InputGbn").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				functionReturnValue = false;
	//				return functionReturnValue;
	//			}
	//			//UPGRADE_WARNING: oMat02.Columns(ItemCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if ((string.IsNullOrEmpty(oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE))) {
	//				SubMain.Sbo_Application.SetStatusBarMessage("품목은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//				oMat02.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				functionReturnValue = false;
	//				return functionReturnValue;
	//			}
	//			//UPGRADE_WARNING: oMat02.Columns(ItemGpCd).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if ((oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected == null)) {
	//				SubMain.Sbo_Application.SetStatusBarMessage("품목그룹은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//				functionReturnValue = false;
	//				return functionReturnValue;
	//			}
	//			////휘팅,부품,엔드베어링일경우
	//			//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "104" & oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "105" & oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "106" & oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "107") {
	//				if ((oMat02.VisualRowCount > 2)) {
	//					SubMain.Sbo_Application.SetStatusBarMessage("해당작지는 투입자재 한품목만 입력가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//					functionReturnValue = false;
	//					return functionReturnValue;
	//				}
	//			}
	//			////기계공구,몰드인경우
	//			//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {
	//				//UPGRADE_WARNING: oMat02.Columns(ProcType).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if ((oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected == null)) {
	//					SubMain.Sbo_Application.SetStatusBarMessage("조달방식은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//					oMat02.Columns.Item("ProcType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//					functionReturnValue = false;
	//					return functionReturnValue;
	//				}
	//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (Conversion.Val(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.VALUE) <= 0) {
	//					SubMain.Sbo_Application.SetStatusBarMessage("수,중량은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//					oMat02.Columns.Item("Weight").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//					functionReturnValue = false;
	//					return functionReturnValue;
	//				}
	//				//원재료 중복 청구 시(2018.09.17 송명규, 김석태 과장 요청)
	//				//UPGRADE_WARNING: oMat02.Columns(LineId).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: oMat02.Columns(ItemCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if ((PS_PP030_Check_DupReq(oForm01.Items.Item("DocEntry").Specific.VALUE, oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE, oMat02.Columns.Item("LineId").Cells.Item(i).Specific.VALUE)) == true) {
	//					//UPGRADE_WARNING: oMat02.Columns(RCode).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((oMat02.Columns.Item("RCode").Cells.Item(i).Specific.Selected == null)) {
	//						//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						SubMain.Sbo_Application.SetStatusBarMessage(i + "행의 원재료 청구가 중복되어 재청구사유를 필수로 입력하여야 합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//						oMat02.Columns.Item("RCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//						functionReturnValue = false;
	//						return functionReturnValue;
	//					}
	//				}
	//			}
	//			////멀티,엔드베어링인경우
	//			//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (Conversion.Val(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.VALUE) <= 0) {
	//					SubMain.Sbo_Application.SetStatusBarMessage("수,중량은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//					functionReturnValue = false;
	//					return functionReturnValue;
	//				}
	//				//UPGRADE_WARNING: oMat02.Columns(BatchNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (string.IsNullOrEmpty(oMat02.Columns.Item("BatchNum").Cells.Item(i).Specific.VALUE)) {
	//					SubMain.Sbo_Application.SetStatusBarMessage("배치번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//					functionReturnValue = false;
	//					return functionReturnValue;
	//				}
	//			}
	//		}

	//		for (i = 1; i <= oMat03.VisualRowCount - 1; i++) {
	//			//UPGRADE_WARNING: oMat03.Columns(CpBCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if ((string.IsNullOrEmpty(oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.VALUE))) {
	//				SubMain.Sbo_Application.SetStatusBarMessage("공정대분류는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//				oMat03.Columns.Item("CpBCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				functionReturnValue = false;
	//				return functionReturnValue;
	//			}
	//			//UPGRADE_WARNING: oMat03.Columns(CpCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if ((string.IsNullOrEmpty(oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.VALUE))) {
	//				SubMain.Sbo_Application.SetStatusBarMessage("공정중분류는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//				oMat03.Columns.Item("CpCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
	//				functionReturnValue = false;
	//				return functionReturnValue;
	//			}
	//			//        For j = i + 1 To oMat03.VisualRowCount - 1
	//			//            If (oMat03.Columns("CpBCode").Cells(i).Specific.Value = oMat03.Columns("CpBCode").Cells(j).Specific.Value And oMat03.Columns("CpCode").Cells(i).Specific.Value = oMat03.Columns("CpCode").Cells(j).Specific.Value) Then
	//			//                Sbo_Application.SetStatusBarMessage "중복된 공정이 존재합니다.", bmt_Short, True
	//			//                oMat03.Columns("CpBCode").Cells(j).Click ct_Regular
	//			//                PS_PP030_DataValidCheck = False
	//			//                Exit Function
	//			//            End If
	//			//        Next
	//		}

	//		if ((PS_PP030_Validate("검사01") == false)) {
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}

	//		if ((PS_PP030_Validate("검사02") == false)) {
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}

	//		if ((PS_PP030_Validate("검사03") == false)) {
	//			functionReturnValue = false;
	//			return functionReturnValue;
	//		}

	//		oDS_PS_PP030L.RemoveRecord(oDS_PS_PP030L.Size - 1);
	//		oDS_PS_PP030M.RemoveRecord(oDS_PS_PP030M.Size - 1);
	//		oMat02.LoadFromDataSource();
	//		oMat03.LoadFromDataSource();

	//		if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
	//			PS_PP030_FormClear();
	//		}
	//		functionReturnValue = true;
	//		return functionReturnValue;
	//		PS_PP030_DataValidCheck_Error:
	//		functionReturnValue = false;
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		return functionReturnValue;
	//	}

	//	private void PS_PP030_MTX01()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		////메트릭스에 데이터 로드
	//		oForm01.Freeze(true);
	//		int i = 0;
	//		string query01 = null;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string Param01 = null;
	//		string Param02 = null;
	//		string Param03 = null;
	//		string Param04 = null;
	//		string Param05 = null;
	//		string Param06 = null;
	//		string Param07 = null;
	//		string Param08 = null;

	//		//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param01 = Strings.Trim(oForm01.Items.Item("SBPLId").Specific.Selected.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param02 = Strings.Trim(oForm01.Items.Item("ItmBsort").Specific.Selected.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param03 = Strings.Trim(oForm01.Items.Item("ItmMsort").Specific.Selected.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param04 = Strings.Trim(oForm01.Items.Item("ReqType").Specific.Selected.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param05 = Strings.Trim(oForm01.Items.Item("SItemCod").Specific.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param06 = Strings.Trim(oForm01.Items.Item("SCardCod").Specific.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param07 = Strings.Trim(oForm01.Items.Item("Mark").Specific.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param08 = Strings.Trim(oForm01.Items.Item("ReqCod").Specific.VALUE);


	//		SAPbouiCOM.ProgressBar ProgressBar01 = null;
	//		ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

	//		query01 = "EXEC PS_PP030_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "'";
	//		RecordSet01.DoQuery(query01);

	//		oMat01.Clear();
	//		oMat01.FlushToDataSource();
	//		oMat01.LoadFromDataSource();

	//		if ((RecordSet01.RecordCount == 0)) {
	//			MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
	//			goto PS_PP030_MTX01_Exit;
	//		}

	//		for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
	//			if (i != 0) {
	//				oDS_PS_USERDS01.InsertRecord((i));
	//			}
	//			oDS_PS_USERDS01.Offset = i;
	//			oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
	//			oDS_PS_USERDS01.SetValue("U_ColReg01", i, RecordSet01.Fields.Item(0).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg02", i, RecordSet01.Fields.Item(1).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg03", i, RecordSet01.Fields.Item(2).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg04", i, RecordSet01.Fields.Item(3).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg05", i, RecordSet01.Fields.Item(4).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg06", i, RecordSet01.Fields.Item(5).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColDt01", i, RecordSet01.Fields.Item(6).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg07", i, RecordSet01.Fields.Item(7).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg08", i, RecordSet01.Fields.Item(8).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColQty01", i, RecordSet01.Fields.Item(9).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColQty02", i, RecordSet01.Fields.Item(10).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg09", i, RecordSet01.Fields.Item(11).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg10", i, RecordSet01.Fields.Item(12).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg11", i, RecordSet01.Fields.Item(13).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg12", i, RecordSet01.Fields.Item(14).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg13", i, RecordSet01.Fields.Item(15).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg14", i, RecordSet01.Fields.Item(16).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg15", i, RecordSet01.Fields.Item(17).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg16", i, RecordSet01.Fields.Item(18).Value);
	//			oDS_PS_USERDS01.SetValue("U_ColReg17", i, RecordSet01.Fields.Item(19).Value);
	//			RecordSet01.MoveNext();
	//			ProgressBar01.Value = ProgressBar01.Value + 1;
	//			ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
	//		}
	//		oMat01.LoadFromDataSource();
	//		oMat01.AutoResizeColumns();
	//		oForm01.Update();

	//		ProgressBar01.Stop();
	//		//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		ProgressBar01 = null;
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_MTX01_Exit:
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		if ((ProgressBar01 != null)) {
	//			ProgressBar01.Stop();
	//		}
	//		return;
	//		PS_PP030_MTX01_Error:
	//		ProgressBar01.Stop();
	//		//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		ProgressBar01 = null;
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void PS_PP030_MTX02()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		////메트릭스에 데이터 로드
	//		oForm01.Freeze(true);
	//		int i = 0;
	//		string query01 = null;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string Param01 = null;
	//		string Param02 = null;
	//		string Param03 = null;
	//		string Param04 = null;
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		Param01 = Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE);
	//		//    Param02 = Trim(oForm01.Items("Param01").Specific.VALUE)
	//		//    Param03 = Trim(oForm01.Items("Param01").Specific.VALUE)
	//		//    Param04 = Trim(oForm01.Items("Param01").Specific.VALUE)
	//		//    If oForm01.Items("OrdGbn").Specific.Selected.Value = "104" Then '//멀티인경우
	//		//        Query01 = "SELECT PS_PP001H.U_CpBCode, PS_PP001H.U_CpBName, PS_PP001L.U_CpCode, PS_PP001L.U_CpName FROM [@PS_PP001H] PS_PP001H LEFT JOIN [@PS_PP001L] PS_PP001L ON PS_PP001H.Code  =PS_PP001L.Code WHERE PS_PP001H.Code = 'CP501'"
	//		//    ElseIf oForm01.Items("OrdGbn").Specific.Selected.Value = "107" Then '//엔드베어링인경우
	//		//        Query01 = "SELECT PS_PP001H.U_CpBCode, PS_PP001H.U_CpBName, PS_PP001L.U_CpCode, PS_PP001L.U_CpName FROM [@PS_PP001H] PS_PP001H LEFT JOIN [@PS_PP001L] PS_PP001L ON PS_PP001H.Code  =PS_PP001L.Code WHERE PS_PP001H.Code = 'CP101'"
	//		//    Else
	//		query01 = "SELECT PS_PP005H.U_ItemCod2, PS_PP005H.U_ItemNam2, OITM.ItmsGrpCod FROM [@PS_PP005H] PS_PP005H LEFT JOIN [OITM] OITM ON PS_PP005H.U_ItemCod2 = OITM.ItemCode WHERE U_ItemCod1 = '" + Param01 + "'";
	//		//    End If
	//		RecordSet01.DoQuery(query01);

	//		oMat02.Clear();
	//		oMat02.FlushToDataSource();
	//		oMat02.LoadFromDataSource();

	//		if ((RecordSet01.RecordCount == 0)) {
	//			PS_PP030_AddMatrixRow01(0, ref true);
	//			goto PS_PP030_MTX02_Exit;
	//		}

	//		for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
	//			if (i != 0) {
	//				oDS_PS_PP030L.InsertRecord((i));
	//			}
	//			oDS_PS_PP030L.Offset = i;
	//			oDS_PS_PP030L.SetValue("U_LineNum", i, Convert.ToString(i + 1));

	//			oDS_PS_PP030L.SetValue("U_InputGbn", i, "10");
	//			////투입구분 '//휘팅,부품의경우만 실행되므로 항상 10이다
	//			oDS_PS_PP030L.SetValue("U_ItemCode", i, RecordSet01.Fields.Item(0).Value);
	//			////품목코드
	//			oDS_PS_PP030L.SetValue("U_ItemName", i, RecordSet01.Fields.Item(1).Value);
	//			////품목이름
	//			oDS_PS_PP030L.SetValue("U_ItemGpCd", i, RecordSet01.Fields.Item(2).Value);
	//			////품목그룹
	//			oDS_PS_PP030L.SetValue("U_BatchNum", i, "");
	//			////배치번호
	//			oDS_PS_PP030L.SetValue("U_Weight", i, Convert.ToString(0));
	//			////중량
	//			oDS_PS_PP030L.SetValue("U_DueDate", i, "");
	//			oDS_PS_PP030L.SetValue("U_CntcCode", i, "");
	//			oDS_PS_PP030L.SetValue("U_CntcName", i, "");
	//			oDS_PS_PP030L.SetValue("U_ProcType", i, "20");
	//			oDS_PS_PP030L.SetValue("U_Comments", i, "");
	//			oDS_PS_PP030L.SetValue("U_LineId", i, "");
	//			if (i == RecordSet01.RecordCount - 1) {
	//				PS_PP030_AddMatrixRow01(i + 1);
	//				////마지막행에 한줄추가
	//			}
	//			RecordSet01.MoveNext();
	//		}
	//		oMat02.LoadFromDataSource();
	//		oMat02.AutoResizeColumns();
	//		oForm01.Update();

	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_MTX02_Exit:
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_MTX02_Error:
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private void PS_PP030_MTX03()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		////메트릭스에 데이터 로드
	//		oForm01.Freeze(true);
	//		int i = 0;
	//		string query01 = null;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string Param01 = null;
	//		string Param02 = null;
	//		string Param03 = null;
	//		string Param04 = null;

	//		string itemCode = null;
	//		string BasicGub = null;

	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		itemCode = Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE);

	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		BasicGub = Strings.Trim(oForm01.Items.Item("BasicGub").Specific.VALUE);

	//		query01 = "         EXEC [PS_PP030_07] '";
	//		query01 = query01 + itemCode + "','";
	//		query01 = query01 + BasicGub + "'";

	//		RecordSet01.DoQuery(query01);

	//		oMat03.Clear();
	//		oMat03.FlushToDataSource();
	//		oMat03.LoadFromDataSource();

	//		if ((RecordSet01.RecordCount == 0)) {
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.VALUE) == "105" | Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.VALUE) == "106") {
	//				oForm01.Items.Item("Mat03").Enabled = true;
	//			} else {
	//				oForm01.Items.Item("Mat03").Enabled = false;
	//				////휘팅,부품,멀티,엔베는 표준공정이 등록되지 않으면 진행불가능
	//			}
	//			PS_PP030_AddMatrixRow02(0, ref true);
	//			////GoTo PS_PP030_MTX03_Exit
	//		} else {
	//			oForm01.Items.Item("Mat03").Enabled = true;
	//			//        If oForm01.Items("OrdGbn").Specific.VALUE = "104" Then
	//			//            Call oForm01.Items("MulGbn1").Specific.Select("10", psk_ByValue)
	//			//        End If
	//		}

	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.VALUE) != "105") {
	//			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
	//				if (i != 0) {
	//					oDS_PS_PP030M.InsertRecord((i));
	//				}
	//				oDS_PS_PP030M.Offset = i;
	//				oDS_PS_PP030M.SetValue("U_LineNum", i, Convert.ToString(i + 1));
	//				oDS_PS_PP030M.SetValue("U_Sequence", i, Convert.ToString(i + 1));
	//				oDS_PS_PP030M.SetValue("U_CpBCode", i, RecordSet01.Fields.Item(0).Value);
	//				oDS_PS_PP030M.SetValue("U_CpBName", i, RecordSet01.Fields.Item(1).Value);
	//				oDS_PS_PP030M.SetValue("U_CpCode", i, RecordSet01.Fields.Item(2).Value);
	//				oDS_PS_PP030M.SetValue("U_CpName", i, RecordSet01.Fields.Item(3).Value);
	//				oDS_PS_PP030M.SetValue("U_Unit", i, RecordSet01.Fields.Item(4).Value);
	//				oDS_PS_PP030M.SetValue("U_ReWorkYN", i, "N");
	//				oDS_PS_PP030M.SetValue("U_ResultYN", i, RecordSet01.Fields.Item(5).Value);
	//				oDS_PS_PP030M.SetValue("U_ReportYN", i, RecordSet01.Fields.Item(6).Value);
	//				oDS_PS_PP030M.SetValue("U_WorkGbn", i, "10");
	//				if (i == RecordSet01.RecordCount - 1) {
	//					PS_PP030_AddMatrixRow02(i + 1);
	//					////마지막행에 한줄추가
	//				}
	//				RecordSet01.MoveNext();
	//			}
	//		} else {
	//			////기계공구류는 검사공정을 기본적으로 입력
	//			oDS_PS_PP030M.Offset = 0;
	//			oDS_PS_PP030M.SetValue("U_LineNum", 0, Convert.ToString(1));
	//			oDS_PS_PP030M.SetValue("U_Sequence", 0, Convert.ToString(1));
	//			oDS_PS_PP030M.SetValue("U_CpBCode", 0, "CP204");
	//			oDS_PS_PP030M.SetValue("U_CpBName", 0, "검사");
	//			oDS_PS_PP030M.SetValue("U_CpCode", 0, "CP20402");
	//			oDS_PS_PP030M.SetValue("U_CpName", 0, "최종검사");
	//			oDS_PS_PP030M.SetValue("U_Unit", 0, "");
	//			oDS_PS_PP030M.SetValue("U_ReWorkYN", 0, "N");
	//			oDS_PS_PP030M.SetValue("U_ResultYN", 0, "N");
	//			oDS_PS_PP030M.SetValue("U_ReportYN", 0, "N");
	//			oDS_PS_PP030M.SetValue("U_WorkGbn", 0, "10");
	//			PS_PP030_AddMatrixRow02(1);
	//			////마지막행에 한줄추가
	//		}
	//		oMat03.LoadFromDataSource();
	//		oMat03.AutoResizeColumns();
	//		oForm01.Update();

	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_MTX03_Exit:
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		return;
	//		PS_PP030_MTX03_Error:
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		oForm01.Freeze(false);
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_MTX03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private bool PS_PP030_DI_API()
	//	{
	//		bool functionReturnValue = false;
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		functionReturnValue = true;
	//		object i = null;
	//		int j = 0;
	//		SAPbobsCOM.Documents oDIObject = null;
	//		int RetVal = 0;
	//		int LineNumCount = 0;
	//		int ResultDocNum = 0;
	//		if (SubMain.Sbo_Company.InTransaction == true) {
	//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
	//		}
	//		SubMain.Sbo_Company.StartTransaction();

	//		ItemInformation = new ItemInformations[1];
	//		ItemInformationCount = 0;
	//		for (i = 1; i <= oMat01.VisualRowCount; i++) {
	//			Array.Resize(ref ItemInformation, ItemInformationCount + 1);
	//			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			ItemInformation[ItemInformationCount].itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE;
	//			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			ItemInformation[ItemInformationCount].BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.VALUE;
	//			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			ItemInformation[ItemInformationCount].Quantity = oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.VALUE;
	//			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			ItemInformation[ItemInformationCount].OPORNo = oMat01.Columns.Item("OPORNo").Cells.Item(i).Specific.VALUE;
	//			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			ItemInformation[ItemInformationCount].POR1No = oMat01.Columns.Item("POR1No").Cells.Item(i).Specific.VALUE;
	//			ItemInformation[ItemInformationCount].Check = false;
	//			ItemInformationCount = ItemInformationCount + 1;
	//		}

	//		LineNumCount = 0;
	//		oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm01.Items.Item("BPLId").Specific.Selected.VALUE));
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oDIObject.CardCode = Strings.Trim(oForm01.Items.Item("CardCode").Specific.VALUE);
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm01.Items.Item("InDate").Specific.VALUE, "&&&&-&&-&&"));
	//		for (i = 0; i <= ItemInformationCount - 1; i++) {
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (ItemInformation[i].Check == true) {
	//				goto Continue_First;
	//			}
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (i != 0) {
	//				oDIObject.Lines.Add();
	//			}
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oDIObject.Lines.ItemCode = ItemInformation[i].itemCode;
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oDIObject.Lines.WarehouseCode = Strings.Trim(oForm01.Items.Item("WhsCode").Specific.VALUE);
	//			oDIObject.Lines.BaseType = Convert.ToInt32("22");
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oDIObject.Lines.BaseEntry = ItemInformation[i].OPORNo;
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			oDIObject.Lines.BaseLine = ItemInformation[i].POR1No;
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			for (j = i; j <= Information.UBound(ItemInformation); j++) {
	//				if (ItemInformation[j].Check == true) {
	//					goto Continue_Second;
	//				}
	//				//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if ((ItemInformation[i].itemCode != ItemInformation[j].itemCode | ItemInformation[i].OPORNo != ItemInformation[j].OPORNo | ItemInformation[i].POR1No != ItemInformation[j].POR1No)) {
	//					goto Continue_Second;
	//				}
	//				////같은것
	//				oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation[j].Quantity;
	//				oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation[j].BatchNum;
	//				oDIObject.Lines.BatchNumbers.Quantity = ItemInformation[j].Quantity;
	//				oDIObject.Lines.BatchNumbers.Add();
	//				ItemInformation[j].PDN1No = LineNumCount;
	//				ItemInformation[j].Check = true;
	//				Continue_Second:
	//			}
	//			LineNumCount = LineNumCount + 1;
	//			Continue_First:
	//		}
	//		RetVal = oDIObject.Add();
	//		if (RetVal == 0) {
	//			ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
	//			for (i = 0; i <= Information.UBound(ItemInformation); i++) {
	//				//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				oDS_PS_PP030L.SetValue("U_OPDNNo", i, Convert.ToString(ResultDocNum));
	//				//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				oDS_PS_PP030L.SetValue("U_PDN1No", i, Convert.ToString(ItemInformation[i].PDN1No));
	//			}
	//		} else {
	//			goto PS_PP030_DI_API_Error;
	//		}

	//		if (SubMain.Sbo_Company.InTransaction == true) {
	//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
	//		}
	//		oMat01.LoadFromDataSource();
	//		oMat01.AutoResizeColumns();

	//		//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oDIObject = null;
	//		return functionReturnValue;
	//		PS_PP030_DI_API_DI_Error:
	//		if (SubMain.Sbo_Company.InTransaction == true) {
	//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
	//		}
	//		SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		functionReturnValue = false;
	//		//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oDIObject = null;
	//		return functionReturnValue;
	//		PS_PP030_DI_API_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_DI_API_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		functionReturnValue = false;
	//		return functionReturnValue;
	//	}

	//	private void PS_PP030_FormResize()
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement


	//		//생산요청, 작지리스트(Option)
	//		oForm01.Items.Item("Opt01").Left = 10;

	//		//생산요청, 작지리스트(Matrix)
	//		oForm01.Items.Item("Mat01").Top = 58;
	//		oForm01.Items.Item("Mat01").Height = oForm01.Height / 2 - 120;
	//		oForm01.Items.Item("Mat01").Left = oForm01.Items.Item("Opt01").Left;
	//		oForm01.Items.Item("Mat01").Width = oForm01.Width - 30;

	//		//작업구분(Label)
	//		oForm01.Items.Item("9").Top = oForm01.Items.Item("Mat01").Height + oForm01.Items.Item("Mat01").Top + 5;
	//		oForm01.Items.Item("9").Left = oForm01.Items.Item("Opt01").Left;
	//		//작업구분(TextBox)
	//		oForm01.Items.Item("OrdGbn").Top = oForm01.Items.Item("9").Top;
	//		oForm01.Items.Item("OrdGbn").Left = oForm01.Items.Item("9").Left + oForm01.Items.Item("9").Width;
	//		oForm01.Items.Item("BasicGub").Top = oForm01.Items.Item("9").Top;
	//		oForm01.Items.Item("BasicGub").Left = oForm01.Items.Item("9").Left + oForm01.Items.Item("9").Width + oForm01.Items.Item("9").Width;
	//		//제품코드(Label)
	//		oForm01.Items.Item("17").Top = oForm01.Items.Item("9").Top + oForm01.Items.Item("9").Height + 1;
	//		oForm01.Items.Item("17").Left = oForm01.Items.Item("9").Left;
	//		//제품코드(Link)
	//		oForm01.Items.Item("1000001").Top = oForm01.Items.Item("17").Top + 1;
	//		oForm01.Items.Item("1000001").Left = oForm01.Items.Item("17").Left + oForm01.Items.Item("17").Width - 15;

	//		//제품코드(TextBox)
	//		oForm01.Items.Item("ItemCode").Top = oForm01.Items.Item("17").Top;
	//		oForm01.Items.Item("ItemCode").Left = oForm01.Items.Item("OrdGbn").Left;
	//		//제품명(TextBox)
	//		oForm01.Items.Item("ItemName").Top = oForm01.Items.Item("ItemCode").Top;
	//		oForm01.Items.Item("ItemName").Left = oForm01.Items.Item("ItemCode").Left + oForm01.Items.Item("ItemCode").Width;

	//		//기준일자(Label)
	//		oForm01.Items.Item("14").Top = oForm01.Items.Item("17").Top + oForm01.Items.Item("17").Height + 1;
	//		oForm01.Items.Item("14").Left = oForm01.Items.Item("17").Left;
	//		//기준일자(TextBox)
	//		oForm01.Items.Item("OrdMgNum").Top = oForm01.Items.Item("14").Top;
	//		oForm01.Items.Item("OrdMgNum").Left = oForm01.Items.Item("ItemCode").Left;

	//		//작업지시번호(Label)
	//		oForm01.Items.Item("67").Top = oForm01.Items.Item("14").Top + oForm01.Items.Item("14").Height + 1;
	//		oForm01.Items.Item("67").Left = oForm01.Items.Item("14").Left;
	//		//작업지시번호(TextBox)
	//		oForm01.Items.Item("OrdNum").Top = oForm01.Items.Item("67").Top;
	//		oForm01.Items.Item("OrdNum").Left = oForm01.Items.Item("OrdMgNum").Left;
	//		//작업지시번호(Sub)(TextBox)
	//		oForm01.Items.Item("OrdSub1").Top = oForm01.Items.Item("67").Top;
	//		oForm01.Items.Item("OrdSub1").Left = oForm01.Items.Item("OrdNum").Left + oForm01.Items.Item("OrdNum").Width;
	//		oForm01.Items.Item("OrdSub2").Top = oForm01.Items.Item("67").Top;
	//		oForm01.Items.Item("OrdSub2").Left = oForm01.Items.Item("OrdSub1").Left + oForm01.Items.Item("OrdSub1").Width;

	//		//지시,완료일자(Label)
	//		oForm01.Items.Item("18").Top = oForm01.Items.Item("67").Top + oForm01.Items.Item("67").Height + 1;
	//		oForm01.Items.Item("18").Left = oForm01.Items.Item("67").Left;
	//		//지시일자(TextBox)
	//		oForm01.Items.Item("DocDate").Top = oForm01.Items.Item("18").Top;
	//		oForm01.Items.Item("DocDate").Left = oForm01.Items.Item("OrdNum").Left;
	//		//완료일자(TextBox)
	//		oForm01.Items.Item("DueDate").Top = oForm01.Items.Item("18").Top;
	//		oForm01.Items.Item("DueDate").Left = oForm01.Items.Item("DocDate").Left + oForm01.Items.Item("DocDate").Width;

	//		//담당자(Label)
	//		oForm01.Items.Item("15").Top = oForm01.Items.Item("18").Top + oForm01.Items.Item("18").Height + 1;
	//		oForm01.Items.Item("15").Left = oForm01.Items.Item("18").Left;
	//		//담당자(TextBox)
	//		oForm01.Items.Item("CntcCode").Top = oForm01.Items.Item("15").Top;
	//		oForm01.Items.Item("CntcCode").Left = oForm01.Items.Item("DocDate").Left;
	//		//담당자명(TextBox)
	//		oForm01.Items.Item("CntcName").Top = oForm01.Items.Item("15").Top;
	//		oForm01.Items.Item("CntcName").Left = oForm01.Items.Item("CntcCode").Left + oForm01.Items.Item("CntcCode").Width;

	//		//수주번호(Label)
	//		oForm01.Items.Item("13").Top = oForm01.Items.Item("15").Top + oForm01.Items.Item("15").Height + 1;
	//		oForm01.Items.Item("13").Left = oForm01.Items.Item("15").Left;
	//		//수주번호(TextBox)
	//		oForm01.Items.Item("SjNum").Top = oForm01.Items.Item("13").Top;
	//		oForm01.Items.Item("SjNum").Left = oForm01.Items.Item("CntcCode").Left;
	//		//수주라인(TextBox)
	//		oForm01.Items.Item("SjLine").Top = oForm01.Items.Item("13").Top;
	//		oForm01.Items.Item("SjLine").Left = oForm01.Items.Item("SjNum").Left + oForm01.Items.Item("SjNum").Width;

	//		//수주LOT번호(Label)
	//		oForm01.Items.Item("39").Top = oForm01.Items.Item("13").Top + oForm01.Items.Item("13").Height + 1;
	//		oForm01.Items.Item("39").Left = oForm01.Items.Item("13").Left;
	//		//수주LOT번호(TextBox)
	//		oForm01.Items.Item("LotNo").Top = oForm01.Items.Item("39").Top;
	//		oForm01.Items.Item("LotNo").Left = oForm01.Items.Item("SjNum").Left;

	//		//멀티작업구분(Label)
	//		oForm01.Items.Item("1000005").Top = oForm01.Items.Item("39").Top + oForm01.Items.Item("39").Height + 1;
	//		oForm01.Items.Item("1000005").Left = oForm01.Items.Item("39").Left;
	//		//멀티작업구분1(TextBox)
	//		oForm01.Items.Item("MulGbn1").Top = oForm01.Items.Item("1000005").Top;
	//		oForm01.Items.Item("MulGbn1").Left = oForm01.Items.Item("LotNo").Left;
	//		//멀티작업구분2(TextBox)
	//		oForm01.Items.Item("MulGbn2").Top = oForm01.Items.Item("1000005").Top;
	//		oForm01.Items.Item("MulGbn2").Left = oForm01.Items.Item("MulGbn1").Left + oForm01.Items.Item("MulGbn1").Width;
	//		//멀티작업구분3(TextBox)
	//		oForm01.Items.Item("MulGbn3").Top = oForm01.Items.Item("1000005").Top;
	//		oForm01.Items.Item("MulGbn3").Left = oForm01.Items.Item("MulGbn2").Left + oForm01.Items.Item("MulGbn2").Width;

	//		//기준문서구분(Label)
	//		oForm01.Items.Item("63").Top = oForm01.Items.Item("1000005").Top + oForm01.Items.Item("1000005").Height + 1;
	//		oForm01.Items.Item("63").Left = oForm01.Items.Item("1000005").Left;
	//		//기준문서구분(TextBox)
	//		oForm01.Items.Item("BaseType").Top = oForm01.Items.Item("63").Top;
	//		oForm01.Items.Item("BaseType").Left = oForm01.Items.Item("MulGbn1").Left;

	//		//기준문서번호(Label)
	//		oForm01.Items.Item("65").Top = oForm01.Items.Item("63").Top;
	//		oForm01.Items.Item("65").Left = oForm01.Items.Item("BaseType").Left + oForm01.Items.Item("BaseType").Width;
	//		//기준문서번호(TextBox)
	//		oForm01.Items.Item("BaseNum").Top = oForm01.Items.Item("65").Top;
	//		oForm01.Items.Item("BaseNum").Left = oForm01.Items.Item("65").Left + oForm01.Items.Item("65").Width;

	//		//투입자재(Option)
	//		oForm01.Items.Item("Opt02").Top = oForm01.Items.Item("63").Top + oForm01.Items.Item("63").Height + 15;
	//		oForm01.Items.Item("Opt02").Left = oForm01.Items.Item("63").Left;

	//		//투입자재(Matrix)
	//		oForm01.Items.Item("Mat02").Top = oForm01.Items.Item("Opt02").Top + oForm01.Items.Item("Opt02").Height + 1;
	//		oForm01.Items.Item("Mat02").Left = oForm01.Items.Item("63").Left;
	//		oForm01.Items.Item("Mat02").Width = oForm01.Width / 2 - 25;
	//		oForm01.Items.Item("Mat02").Height = oForm01.Height - oForm01.Items.Item("Mat02").Top - 60;

	//		//문서번호(Label)
	//		oForm01.Items.Item("11").Top = oForm01.Items.Item("9").Top;
	//		oForm01.Items.Item("11").Left = 320;
	//		//문서번호(TextBox)
	//		oForm01.Items.Item("DocEntry").Top = oForm01.Items.Item("9").Top;
	//		oForm01.Items.Item("DocEntry").Left = oForm01.Items.Item("11").Left + oForm01.Items.Item("11").Width;

	//		//사업장(Label)
	//		oForm01.Items.Item("1000002").Top = oForm01.Items.Item("14").Top;
	//		oForm01.Items.Item("1000002").Left = 255;
	//		//사업장(TextBox)
	//		oForm01.Items.Item("BPLId").Top = oForm01.Items.Item("14").Top;
	//		oForm01.Items.Item("BPLId").Left = 335;

	//		//작번이름(Label)
	//		oForm01.Items.Item("70").Top = oForm01.Items.Item("1000002").Top + oForm01.Items.Item("1000002").Height + 1;
	//		oForm01.Items.Item("70").Left = oForm01.Items.Item("11").Left;
	//		//작번이름(TextBox)
	//		oForm01.Items.Item("JakMyung").Top = oForm01.Items.Item("70").Top;
	//		oForm01.Items.Item("JakMyung").Left = oForm01.Items.Item("70").Left + oForm01.Items.Item("70").Width;

	//		//작번규격,단위(Label)
	//		oForm01.Items.Item("72").Top = oForm01.Items.Item("70").Top + oForm01.Items.Item("70").Height + 1;
	//		oForm01.Items.Item("72").Left = oForm01.Items.Item("70").Left;
	//		//작번규격(TextBox)
	//		oForm01.Items.Item("JakSize").Top = oForm01.Items.Item("72").Top;
	//		oForm01.Items.Item("JakSize").Left = oForm01.Items.Item("72").Left + oForm01.Items.Item("72").Width;
	//		//작번단위(TextBox)
	//		oForm01.Items.Item("JakUnit").Top = oForm01.Items.Item("72").Top;
	//		oForm01.Items.Item("JakUnit").Left = oForm01.Items.Item("JakSize").Left + oForm01.Items.Item("JakSize").Width;

	//		//요청수,중량(Label)
	//		oForm01.Items.Item("42").Top = oForm01.Items.Item("72").Top + oForm01.Items.Item("72").Height + 1;
	//		oForm01.Items.Item("42").Left = oForm01.Items.Item("72").Left;
	//		//요청수, 중량
	//		oForm01.Items.Item("ReqWt").Top = oForm01.Items.Item("42").Top;
	//		oForm01.Items.Item("ReqWt").Left = oForm01.Items.Item("42").Left + oForm01.Items.Item("42").Width;

	//		//지시수,중량(Label)
	//		oForm01.Items.Item("40").Top = oForm01.Items.Item("42").Top + oForm01.Items.Item("42").Height + 1;
	//		oForm01.Items.Item("40").Left = oForm01.Items.Item("42").Left;
	//		//지시수,중량
	//		oForm01.Items.Item("SelWt").Top = oForm01.Items.Item("40").Top;
	//		oForm01.Items.Item("SelWt").Left = oForm01.Items.Item("40").Left + oForm01.Items.Item("40").Width;

	//		//수주금액(Label)
	//		oForm01.Items.Item("38").Top = oForm01.Items.Item("40").Top + oForm01.Items.Item("40").Height + 1;
	//		oForm01.Items.Item("38").Left = oForm01.Items.Item("40").Left;
	//		//수주금액
	//		oForm01.Items.Item("SjPrice").Top = oForm01.Items.Item("38").Top;
	//		oForm01.Items.Item("SjPrice").Left = oForm01.Items.Item("38").Left + oForm01.Items.Item("38").Width;

	//		//문서상태(Label)
	//		oForm01.Items.Item("79").Top = oForm01.Items.Item("38").Top + oForm01.Items.Item("38").Height + 1;
	//		oForm01.Items.Item("79").Left = oForm01.Items.Item("38").Left + 65;
	//		//문서상태(TextBox)
	//		oForm01.Items.Item("Status").Top = oForm01.Items.Item("79").Top;
	//		oForm01.Items.Item("Status").Left = oForm01.Items.Item("79").Left + oForm01.Items.Item("79").Width;

	//		//취소여부(Label)
	//		oForm01.Items.Item("71").Top = oForm01.Items.Item("79").Top + oForm01.Items.Item("79").Height + 1;
	//		oForm01.Items.Item("71").Left = oForm01.Items.Item("79").Left;
	//		//취소여부(TextBox)
	//		oForm01.Items.Item("Canceled").Top = oForm01.Items.Item("71").Top;
	//		oForm01.Items.Item("Canceled").Left = oForm01.Items.Item("71").Left + oForm01.Items.Item("71").Width;

	//		//공정리스트(Option)
	//		oForm01.Items.Item("Opt03").Top = oForm01.Items.Item("9").Top;
	//		oForm01.Items.Item("Opt03").Left = oForm01.Width / 2;

	//		//표준공수조회(BUTTON)
	//		oForm01.Items.Item("btnWkSrch").Top = oForm01.Items.Item("9").Top - 2;
	//		oForm01.Items.Item("btnWkSrch").Left = oForm01.Items.Item("Opt03").Left + oForm01.Items.Item("Opt03").Width + 3;

	//		//품목별공수조회(BUTTON)
	//		oForm01.Items.Item("btnItmSrch").Top = oForm01.Items.Item("btnWkSrch").Top;
	//		oForm01.Items.Item("btnItmSrch").Left = oForm01.Items.Item("btnWkSrch").Left + oForm01.Items.Item("btnWkSrch").Width + 3;

	//		//공정금액합계(Label)
	//		oForm01.Items.Item("77").Top = oForm01.Items.Item("9").Top;
	//		oForm01.Items.Item("77").Left = oForm01.Items.Item("btnItmSrch").Left + oForm01.Items.Item("btnItmSrch").Width + 5;

	//		//공정금액합계(TextBox)
	//		oForm01.Items.Item("Total").Top = oForm01.Items.Item("9").Top;
	//		oForm01.Items.Item("Total").Left = oForm01.Items.Item("77").Left + oForm01.Items.Item("77").Width;

	//		//공정리스트(Matrix)
	//		oForm01.Items.Item("Mat03").Left = oForm01.Items.Item("Opt03").Left;
	//		oForm01.Items.Item("Mat03").Top = oForm01.Items.Item("9").Top + 18;
	//		oForm01.Items.Item("Mat03").Height = oForm01.Height - oForm01.Items.Item("Mat03").Top - 60;
	//		oForm01.Items.Item("Mat03").Width = oForm01.Width - oForm01.Items.Item("Mat03").Left - 20;

	//		return;
	//		PS_PP030_FormResize_Error:
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	public bool PS_PP030_Validate(string ValidateType)
	//	{
	//		bool functionReturnValue = false;
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		functionReturnValue = true;
	//		object i = null;
	//		int j = 0;
	//		string query01 = null;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
	//		bool Exist = false;

	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP030H] WHERE DocEntry = ' & oForm01.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" + oForm01.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
	//			MDC_Com.MDC_GF_Message(ref "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", ref "W");
	//			functionReturnValue = false;
	//			goto PS_PP030_Validate_Exit;
	//		}

	//		string QueryString = null;
	//		if (ValidateType == "검사01") {

	//		////투입자재 매트릭스에 대한 검사
	//		} else if (ValidateType == "검사02") {
	//			////삭제된 행을 찾아서 삭제가능성 검사
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = "SELECT PS_PP030L.DocEntry,PS_PP030L.LineId,PS_PP030L.U_ProcType FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry WHERE PS_PP030H.Canceled = 'N' AND PS_PP030L.DocEntry = '" + oForm01.Items.Item("DocEntry").Specific.VALUE + "'";
	//			RecordSet01.DoQuery(query01);
	//			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
	//				Exist = false;
	//				for (j = 1; j <= oMat02.RowCount - 1; j++) {
	//					//UPGRADE_WARNING: oMat02.Columns(LineId).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					////새로추가된 행인경우, 검사할필요없다
	//					if ((string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(j).Specific.VALUE))) {
	//					} else {
	//						////라인번호가 같고, 문서번호가 같으면 존재하는행
	//						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (Conversion.Val(RecordSet01.Fields.Item(0).Value) == Conversion.Val(oForm01.Items.Item("DocEntry").Specific.VALUE) & Conversion.Val(RecordSet01.Fields.Item(1).Value) == Conversion.Val(oMat02.Columns.Item("LineId").Cells.Item(j).Specific.VALUE)) {
	//							Exist = true;
	//							//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							////몰드,기계공구
	//							if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {
	//								////DB상에는 청구이고 매트릭스의 조달방법이 잔재로 변경된경우 수정할수 없다.
	//								//UPGRADE_WARNING: oMat02.Columns(ProcType).Cells(j).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								if (RecordSet01.Fields.Item(2).Value == "10" & oMat02.Columns.Item("ProcType").Cells.Item(j).Specific.Selected.VALUE != "10") {
	//									MDC_Com.MDC_GF_Message(ref "구매요청이 청구에서 잔재,취소로 변경되었습니다. 수정할수 없습니다.", ref "W");
	//									functionReturnValue = false;
	//									goto PS_PP030_Validate_Exit;
	//								}
	//							}
	//						}
	//					}
	//				}
	//				////삭제된 행중 구매요청에 아직 존재하면 수정할수 없다.
	//				if (Exist == false) {
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					////몰드,기계공구
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {
	//						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + RecordSet01.Fields.Item(0).Value + "' AND U_PP030LNo = '" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0) {
	//							MDC_Com.MDC_GF_Message(ref "삭제된행이 구매요청문서 입니다. 적용할수 없습니다.", ref "W");
	//							functionReturnValue = false;
	//							goto PS_PP030_Validate_Exit;
	//						}
	//					}
	//					////삭제된 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정할수 없다.
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + RecordSet01.Fields.Item(0).Value + "'", 0, 1)) > 0) {
	//							MDC_Com.MDC_GF_Message(ref "작업일보 등록된 행입니다. 수정할수 없습니다.", ref "W");
	//							functionReturnValue = false;
	//							goto PS_PP030_Validate_Exit;
	//						}
	//					}
	//					////휘팅,부품은 삭제되는데 제약이 없다.
	//				}
	//				RecordSet01.MoveNext();
	//			}

	//			for (i = 1; i <= oMat02.RowCount - 1; i++) {
	//				//UPGRADE_WARNING: oMat02.Columns(LineId).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				////새로추가된 행인경우, 검사할필요없다
	//				if ((string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(i).Specific.VALUE))) {
	//				} else {
	//					////기존에 있던 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정할수 없다.
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'", 0, 1)) > 0) {
	//							query01 = "SELECT ";
	//							query01 = query01 + " PS_PP030L.U_ItemCode,";
	//							query01 = query01 + " PS_PP030L.U_ItemName,";
	//							query01 = query01 + " PS_PP030L.U_ItemGpCd,";
	//							query01 = query01 + " PS_PP030L.U_Weight,";
	//							query01 = query01 + " PS_PP030H.U_BPLId,";
	//							query01 = query01 + " CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),";
	//							query01 = query01 + " PS_PP030L.U_CntcCode,";
	//							query01 = query01 + " PS_PP030L.U_CntcName,";
	//							query01 = query01 + " PS_PP030L.U_ProcType,";
	//							query01 = query01 + " PS_PP030L.U_Comments";
	//							query01 = query01 + " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry";
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							query01 = query01 + " WHERE PS_PP030H.DocEntry = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'";
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							query01 = query01 + " AND PS_PP030L.LineId = '" + Strings.Trim(oMat02.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'";
	//							query01 = query01 + " AND PS_PP030H.Canceled = 'N'";
	//							RecordSet01.DoQuery(query01);
	//							//UPGRADE_WARNING: oMat02.Columns(Comments).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns(ProcType).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns(CntcName).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns(CntcCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns(DueDate).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns(ItemGpCd).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns(ItemName).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat02.Columns(ItemCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if ((RecordSet01.Fields.Item(0).Value == oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(1).Value == oMat02.Columns.Item("ItemName").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(2).Value == oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected.VALUE & Conversion.Val(RecordSet01.Fields.Item(3).Value) == Conversion.Val(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.VALUE) & RecordSet01.Fields.Item(4).Value == oForm01.Items.Item("BPLId").Specific.Selected.VALUE & RecordSet01.Fields.Item(5).Value == oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(6).Value == oMat02.Columns.Item("CntcCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(7).Value == oMat02.Columns.Item("CntcName").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(8).Value == oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.VALUE & RecordSet01.Fields.Item(9).Value == oMat02.Columns.Item("Comments").Cells.Item(i).Specific.VALUE)) {
	//							////값이 변경된 행의경우
	//							} else {
	//								MDC_Com.MDC_GF_Message(ref "작업일보가 등록된 행은 수정할수 없습니다.", ref "W");
	//								functionReturnValue = false;
	//								goto PS_PP030_Validate_Exit;
	//							}
	//						}
	//					}
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					////몰드,기계공구
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {
	//						//UPGRADE_WARNING: oMat02.Columns(ProcType).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////잔재인 행은 제외
	//						if (oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.VALUE == "20") {
	//							//UPGRADE_WARNING: oMat02.Columns(ProcType).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////취소인 행은 제외
	//						} else if (oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.VALUE == "30") {
	//						////청구인행에 대해
	//						} else {
	//							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if ((MDC_PS_Common.GetValue("SELECT U_OKYN FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND U_PP030LNo = '" + Strings.Trim(oMat02.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'", 0, 1)) == "Y") {
	//								////결재가 완료된 값중
	//								query01 = "SELECT ";
	//								query01 = query01 + " PS_PP030L.U_ItemCode,";
	//								query01 = query01 + " PS_PP030L.U_ItemName,";
	//								query01 = query01 + " PS_PP030L.U_ItemGpCd,";
	//								query01 = query01 + " Round(PS_PP030L.U_Weight,2),";
	//								query01 = query01 + " PS_PP030H.U_BPLId,";
	//								query01 = query01 + " CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),";
	//								query01 = query01 + " PS_PP030L.U_CntcCode,";
	//								query01 = query01 + " PS_PP030L.U_CntcName,";
	//								query01 = query01 + " PS_PP030L.U_ProcType,";
	//								query01 = query01 + " PS_PP030L.U_Comments";
	//								query01 = query01 + " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry";
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								query01 = query01 + " WHERE PS_PP030H.DocEntry = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'";
	//								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								query01 = query01 + " AND PS_PP030L.LineId = '" + Strings.Trim(oMat02.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'";
	//								query01 = query01 + " AND PS_PP030H.Canceled = 'N'";
	//								RecordSet01.DoQuery(query01);

	//								//UPGRADE_WARNING: oMat02.Columns(Comments).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns(ProcType).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns(CntcName).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns(CntcCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns(DueDate).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns(ItemGpCd).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns(ItemName).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat02.Columns(ItemCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								if ((RecordSet01.Fields.Item(0).Value == oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(1).Value == oMat02.Columns.Item("ItemName").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(2).Value == oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected.VALUE & Conversion.Val(RecordSet01.Fields.Item(3).Value) == Conversion.Val(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.VALUE) & RecordSet01.Fields.Item(4).Value == oForm01.Items.Item("BPLId").Specific.Selected.VALUE & RecordSet01.Fields.Item(5).Value == oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(6).Value == oMat02.Columns.Item("CntcCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(7).Value == oMat02.Columns.Item("CntcName").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(8).Value == oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.VALUE & RecordSet01.Fields.Item(9).Value == oMat02.Columns.Item("Comments").Cells.Item(i).Specific.VALUE)) {
	//								////값이 변경된 행의경우
	//								} else {
	//									MDC_Com.MDC_GF_Message(ref "구매요청 결재가 완료된 행은 수정할수 없습니다.", ref "W");
	//									functionReturnValue = false;
	//									goto PS_PP030_Validate_Exit;
	//								}
	//							}
	//						}
	//					}
	//				}
	//			}
	//		////공정 매트릭스에 대한 검사
	//		} else if (ValidateType == "검사03") {
	//			////삭제된 행을 찾아서 삭제가능성 검사
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = "SELECT PS_PP030M.DocEntry,PS_PP030M.LineId,PS_PP030M.U_Sequence, PS_PP030M.U_WorkGbn FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030M.DocEntry = '" + oForm01.Items.Item("DocEntry").Specific.VALUE + "'";
	//			RecordSet01.DoQuery(query01);
	//			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
	//				Exist = false;
	//				for (j = 1; j <= oMat03.RowCount - 1; j++) {
	//					//UPGRADE_WARNING: oMat03.Columns(LineId).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					////새로추가된 행인경우, 검사할필요없다
	//					if ((string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(j).Specific.VALUE))) {
	//					} else {
	//						////라인번호가 같고, 문서번호가 같으면 존재하는행,시퀀스도 같아야 한다. 행을 삭제할경우 시퀀스가 변경될수 있기때문에.
	//						//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						if (Conversion.Val(RecordSet01.Fields.Item(0).Value) == Conversion.Val(oForm01.Items.Item("DocEntry").Specific.VALUE) & Conversion.Val(RecordSet01.Fields.Item(1).Value) == Conversion.Val(oMat03.Columns.Item("LineId").Cells.Item(j).Specific.VALUE) & Conversion.Val(RecordSet01.Fields.Item(2).Value) == Conversion.Val(oMat03.Columns.Item("Sequence").Cells.Item(j).Specific.VALUE)) {
	//							Exist = true;
	//							//                        If oForm01.Items("OrdGbn").Specific.Selected.Value = "101" Then '//휘팅
	//							//                            '//DB상에는 외주이고 매트릭스의 조달방법이 외주가 아닌경우 수정할수 없다.
	//							//                            If RecordSet01.Fields(3).Value = "30" And oMat03.Columns("WorkGbn").Cells(j).Specific.Selected.Value <> "30" Then
	//							//                                Call MDC_Com.MDC_GF_Message("작업구분이 외주에서 자가,정밀로 변경되었습니다. 수정할수 없습니다.", "W")
	//							//                                PS_PP030_Validate = False
	//							//                                GoTo PS_PP030_Validate_Exit
	//							//                            End If
	//							//                        End If
	//							// 검사조건은 필요가 없을 듯 한데 왜 넣었을까?주석처리(2017.12.07 송명규)
	//							//                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구
	//							//                            '//DB상에는 외주이고 매트릭스의 조달방법이 외주가 아닌경우 수정할수 없다.
	//							//                            If RecordSet01.Fields(3).VALUE = "30" And oMat03.Columns("WorkGbn").Cells(j).Specific.Selected.VALUE <> "30" Then
	//							//                                Call MDC_Com.MDC_GF_Message("작업구분이 외주에서 자가,정밀로 변경되었습니다. 수정할수 없습니다.", "W")
	//							//                                PS_PP030_Validate = False
	//							//                                GoTo PS_PP030_Validate_Exit
	//							//                            End If
	//							//                        End If
	//							//위 검사조건은 필요가 없을 듯 한데 왜 넣었을까?주석처리(2017.12.07 송명규)
	//						}
	//					}
	//				}
	//				////삭제된 행중 작업일보에 등록된행
	//				if (Exist == false) {
	//					//                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040L.U_PP030MNo = '" & RecordSet01.Fields(1).VALUE & "'", 0, 1)) > 0 Then
	//					//                    MDC_Com.MDC_GF_Message "삭제된행이 작업일보 등록된 행입니다. 적용할수 없습니다.", "W"
	//					//                    PS_PP030_Validate = False
	//					//                    GoTo PS_PP030_Validate_Exit
	//					//                End If
	//					////삭제된행중에 외주반출등록된행
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101") {
	//						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + RecordSet01.Fields.Item(0).Value + "' AND PS_MM130L.U_PP030MNo = '" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0) {
	//							MDC_Com.MDC_GF_Message(ref "삭제된행이 외주반출 등록된 행입니다. 적용할수 없습니다.", ref "W");
	//							functionReturnValue = false;
	//							goto PS_PP030_Validate_Exit;
	//						}
	//					}
	//					////삭제된행중에 외주등록된행
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					////기계공구
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105") {
	//						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE PS_MM005H.U_OrdType in ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + RecordSet01.Fields.Item(0).Value + "-" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0) {
	//							MDC_Com.MDC_GF_Message(ref "삭제된행이 외주청구 등록된 행입니다. 적용할수 없습니다.", ref "W");
	//							functionReturnValue = false;
	//							goto PS_PP030_Validate_Exit;
	//						}
	//					}
	//				}
	//				RecordSet01.MoveNext();
	//			}

	//			for (i = 1; i <= oMat03.RowCount - 1; i++) {
	//				//UPGRADE_WARNING: oMat03.Columns(LineId).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				////새로추가된 행인경우, 검사할필요없다
	//				if ((string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.VALUE))) {
	//				} else {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP040L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						////작업일보등록된문서중에 수정이 된문서를 구함
	//						query01 = "SELECT ";
	//						query01 = query01 + " PS_PP030M.U_CpBCode,";
	//						query01 = query01 + " PS_PP030M.U_CpCode,";
	//						query01 = query01 + " PS_PP030M.U_ResultYN,";
	//						query01 = query01 + " PS_PP030M.U_ReportYN";
	//						query01 = query01 + " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry";
	//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						query01 = query01 + " WHERE PS_PP030H.DocEntry = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'";
	//						//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						query01 = query01 + " AND PS_PP030M.LineId = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'";
	//						query01 = query01 + " AND PS_PP030H.Canceled = 'N'";
	//						RecordSet01.DoQuery(query01);
	//						// CP40101,2 공정코드는 일보,실적 수정가능 배병관대리 요청 20200603
	//						if ((RecordSet01.Fields.Item(1).Value == "CP40101" | RecordSet01.Fields.Item(1).Value == "CP40102")) {
	//						} else {
	//							//UPGRADE_WARNING: oMat03.Columns(ReportYN).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat03.Columns(ResultYN).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat03.Columns(CpCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oMat03.Columns(CpBCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if ((RecordSet01.Fields.Item(0).Value == oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(1).Value == oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(2).Value == oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Selected.VALUE & RecordSet01.Fields.Item(3).Value == oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Selected.VALUE)) {
	//							////값이 변경된 행의경우
	//							} else {
	//								MDC_Com.MDC_GF_Message(ref "작업일보가 등록된 행은 수정할수 없습니다.", ref "W");
	//								//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								oMat03.SelectRow(i, true, false);
	//								functionReturnValue = false;
	//								goto PS_PP030_Validate_Exit;
	//							}
	//						}
	//					}
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101") {
	//						//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////자가인 행은 제외
	//						if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.VALUE == "10") {
	//							//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////정밀인 행은 제외
	//						} else if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.VALUE == "20") {
	//						////외주
	//						} else {
	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_MM130L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'", 0, 1)) > 0) {
	//								////외주반출등록된문서중에 수정이 된문서를 구함
	//								query01 = "SELECT ";
	//								query01 = query01 + " PS_PP030M.U_CpBCode,";
	//								query01 = query01 + " PS_PP030M.U_CpCode,";
	//								query01 = query01 + " PS_PP030M.U_ResultYN,";
	//								query01 = query01 + " PS_PP030M.U_ReportYN,";
	//								query01 = query01 + " PS_PP030M.U_WorkGbn";
	//								query01 = query01 + " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry";
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								query01 = query01 + " WHERE PS_PP030H.DocEntry = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'";
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								query01 = query01 + " AND PS_PP030M.LineId = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'";
	//								query01 = query01 + " AND PS_PP030H.Canceled = 'N'";
	//								RecordSet01.DoQuery(query01);
	//								//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(ReportYN).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(ResultYN).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(CpCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(CpBCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								if ((RecordSet01.Fields.Item(0).Value == oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(1).Value == oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(2).Value == oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Selected.VALUE & RecordSet01.Fields.Item(3).Value == oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Selected.VALUE & RecordSet01.Fields.Item(4).Value == oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.VALUE)) {
	//								////값이 변경된 행의경우
	//								} else {
	//									MDC_Com.MDC_GF_Message(ref "외주반출이 등록된 행은 수정할수 없습니다.", ref "W");
	//									//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oMat03.SelectRow(i, true, false);
	//									functionReturnValue = false;
	//									goto PS_PP030_Validate_Exit;
	//								}
	//							}
	//						}
	//					}
	//					//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					////기계공구일대
	//					if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105") {
	//						//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////자가인 행은 제외
	//						if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.VALUE == "10") {
	//							//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//						////정밀인 행은 제외
	//						} else if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.VALUE == "20") {
	//						////외주
	//						} else {
	//							//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//							if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "-" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'", 0, 1)) > 0) {
	//								////외주청구등록된문서중에 수정이 된문서를 구함
	//								query01 = "SELECT ";
	//								query01 = query01 + " PS_PP030M.U_CpBCode,";
	//								query01 = query01 + " PS_PP030M.U_CpCode,";
	//								query01 = query01 + " PS_PP030M.U_ResultYN,";
	//								query01 = query01 + " PS_PP030M.U_ReportYN,";
	//								query01 = query01 + " PS_PP030M.U_WorkGbn";
	//								query01 = query01 + " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry";
	//								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								query01 = query01 + " WHERE PS_PP030H.DocEntry = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'";
	//								//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								query01 = query01 + " AND PS_PP030M.LineId = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'";
	//								query01 = query01 + " AND PS_PP030H.Canceled = 'N'";
	//								RecordSet01.DoQuery(query01);
	//								//UPGRADE_WARNING: oMat03.Columns(WorkGbn).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(ReportYN).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(ResultYN).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(CpCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								//UPGRADE_WARNING: oMat03.Columns(CpBCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//								if ((RecordSet01.Fields.Item(0).Value == oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(1).Value == oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(2).Value == oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Selected.VALUE & RecordSet01.Fields.Item(3).Value == oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Selected.VALUE & RecordSet01.Fields.Item(4).Value == oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.VALUE)) {
	//								////값이 변경된 행의경우
	//								} else {
	//									MDC_Com.MDC_GF_Message(ref "외주청구가 등록된 행은 수정할수 없습니다.", ref "W");
	//									//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//									oMat03.SelectRow(i, true, false);
	//									functionReturnValue = false;
	//									goto PS_PP030_Validate_Exit;
	//								}
	//							}
	//						}
	//					}
	//				}
	//			}
	//			////모든값의 변경에 대해 검사하여 변경이 되었을시 수정가능검사를 하여 체크한다.
	//		} else if (ValidateType == "수정02") {
	//			//UPGRADE_WARNING: oMat02.Columns(LineId).Cells(oMat02Row02).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			////새로추가된 행인경우, 수정하여도 무방하다
	//			if ((string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.VALUE))) {
	//			} else {
	//				////삭제된 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정할수 없다.
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "작업일보 등록된 행입니다. 수정할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				////몰드,기계공구
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {
	//					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT U_OKYN FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND U_PP030LNo = '" + Strings.Trim(oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.VALUE) + "'", 0, 1)) == "Y") {
	//						MDC_Com.MDC_GF_Message(ref "구매요청 결재가 진행된 행입니다. 수정할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//			}
	//		} else if (ValidateType == "행삭제02") {
	//			////행삭제전 행삭제가능여부검사
	//			//UPGRADE_WARNING: oMat02.Columns(LineId).Cells(oMat02Row02).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			////새로추가된 행인경우, 삭제하여도 무방하다
	//			if ((string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.VALUE))) {
	//			} else {
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "104" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "107") {
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "작업일보 등록된 행입니다. 삭제할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				////몰드,기계공구
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {
	//					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND U_PP030LNo = '" + Strings.Trim(oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "구매요청된행 입니다. 삭제할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//			}
	//			////모든값의 변경에 대해 검사하여 변경이 되었을시 수정가능검사를 하여 체크한다.
	//		} else if (ValidateType == "수정03") {
	//			//UPGRADE_WARNING: oMat03.Columns(LineId).Cells(oMat03Row03).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			////새로추가된 행인경우, 수정하여도 무방하다
	//			if ((string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE))) {

	//			} else {
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE != "102") {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP040L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "작업일보 등록된 행 입니다. 수정할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//				////삭제된행중에 외주반출등록된행
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101") {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_MM130L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "외주반출 등록된 행입니다. 수정할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//				////삭제된행중에 외주청구등록된행
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				////기계공구일때
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105") {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "-" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "외주청구 등록된 행입니다. 수정할수 없습니다1.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//			}
	//		} else if (ValidateType == "행삭제03") {
	//			////행삭제전 행삭제가능여부검사
	//			//UPGRADE_WARNING: oMat03.Columns(LineId).Cells(oMat03Row03).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			////새로추가된 행인경우, 삭제하여도 무방하다
	//			if ((string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE))) {
	//			} else {
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP040L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//					MDC_Com.MDC_GF_Message(ref "작업일보 등록된 행 입니다. 삭제할수 없습니다.", ref "W");
	//					functionReturnValue = false;
	//					goto PS_PP030_Validate_Exit;
	//				}
	//				////삭제된행중에 외주반출등록된행
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101") {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_MM130L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "외주반출 등록된 행입니다. 수정할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//				////삭제된행중에 외주청구등록된행
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				////기계공구일때
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105") {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "-" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "외주청구 등록된 행입니다. 수정할수 없습니다.2", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//			}
	//		} else if (ValidateType == "행추가03") {
	//			////행추가전 행추가가능여부검사
	//			//UPGRADE_WARNING: oMat03.Columns(LineId).Cells(oMat03Row03).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			////새로추가된 행인경우, 삭제하여도 무방하다
	//			if ((string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE))) {
	//			} else {
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP040L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//					MDC_Com.MDC_GF_Message(ref "작업일보 등록된 행 입니다. 행추가할수 없습니다.", ref "W");
	//					functionReturnValue = false;
	//					goto PS_PP030_Validate_Exit;
	//				}
	//				////삭제된행중에 외주반출등록된행
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101") {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_MM130L.U_PP030MNo = '" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "외주반출 등록된 행입니다. 행추가할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//				////삭제된행중에 외주청구등록된행
	//				//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				////기계공구일때
	//				if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105") {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40) AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "-" + Strings.Trim(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.VALUE) + "'", 0, 1)) > 0) {
	//						MDC_Com.MDC_GF_Message(ref "외주청구 등록된 행입니다. 행추가할수 없습니다.", ref "W");
	//						functionReturnValue = false;
	//						goto PS_PP030_Validate_Exit;
	//					}
	//				}
	//			}
	//		} else if (ValidateType == "취소") {
	//			////취소가능유무검사
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP030H] WHERE DocEntry = ' & oForm01.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" + oForm01.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
	//				MDC_Com.MDC_GF_Message(ref "이미취소된 문서 입니다. 취소할수 없습니다.", ref "W");
	//				functionReturnValue = false;
	//				goto PS_PP030_Validate_Exit;
	//			}

	//			//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			////몰드,기계공구
	//			if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105" | oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "106") {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = ' & oForm01.Items(DocEntry).Specific.VALUE & ' AND U_OKYN = 'Y', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + oForm01.Items.Item("DocEntry").Specific.VALUE + "' AND U_OKYN = 'Y'", 0, 1) > 0) {
	//					MDC_Com.MDC_GF_Message(ref "구매요청 결재가 승인되었습니다. 취소할수 없습니다.", ref "W");
	//					functionReturnValue = false;
	//					goto PS_PP030_Validate_Exit;
	//				}
	//			}

	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_PP040H] PS_PP040H LEFT JOIN [PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = ' & oForm01.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm01.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) > 0) {
	//				MDC_Com.MDC_GF_Message(ref "작업일보가 등록되었습니다. 취소할수 없습니다.", ref "W");
	//				functionReturnValue = false;
	//				goto PS_PP030_Validate_Exit;
	//			}

	//			////삭제된행중에 외주반출등록된행
	//			//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "101") {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'", 0, 1)) > 0) {
	//					MDC_Com.MDC_GF_Message(ref "외주반출 등록된 행입니다. 취소할수 없습니다.", ref "W");
	//					functionReturnValue = false;
	//					goto PS_PP030_Validate_Exit;
	//				}
	//			}

	//			////삭제된행중에 외주청구등록된행
	//			//UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			////기계공구일때
	//			if (oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE == "105") {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND U_PP030HNo = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE) + "'", 0, 1)) > 0) {
	//					MDC_Com.MDC_GF_Message(ref "외주청구 등록된 행입니다. 취소할수 없습니다.", ref "W");
	//					functionReturnValue = false;
	//					goto PS_PP030_Validate_Exit;
	//				}
	//			}

	//		} else if (ValidateType == "닫기") {

	//			////닫기가능유무검사
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Status FROM [PS_PP030H] WHERE DocEntry = ' & oForm01.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (MDC_PS_Common.GetValue("SELECT Status FROM [@PS_PP030H] WHERE DocEntry = '" + oForm01.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "C") {
	//				MDC_Com.MDC_GF_Message(ref "이미 닫기(종료) 처리된 문서 입니다. 닫기(종료) 처리할 수 없습니다.", ref "W");
	//				functionReturnValue = false;
	//				goto PS_PP030_Validate_Exit;
	//			}


	//			//재고가 존재하면 닫기(종료) 불가 기능 추가(2012.01.11 송명규 추가)

	//			QueryString = "                     SELECT      SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]";
	//			QueryString = QueryString + "  FROM       OINM AS A";
	//			QueryString = QueryString + "                 INNER JOIN";
	//			QueryString = QueryString + "                 OITM As B";
	//			QueryString = QueryString + "                     ON A.ItemCode = B.ItemCode";
	//			QueryString = QueryString + "  WHERE      B.U_ItmBsort IN ('105','106')";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			QueryString = QueryString + "                 AND A.ItemCode = '" + oForm01.Items.Item("ItemCode").Specific.VALUE + "'";
	//			QueryString = QueryString + "  GROUP BY  A.ItemCode";

	//			if ((string.IsNullOrEmpty((MDC_PS_Common.GetValue(QueryString, 0, 1))) ? 0 : (MDC_PS_Common.GetValue(QueryString, 0, 1))) > 0) {

	//				MDC_Com.MDC_GF_Message(ref "재고가 존재하는 작업지시입니다. 닫기(종료) 처리할 수 없습니다.", ref "W");
	//				functionReturnValue = false;
	//				goto PS_PP030_Validate_Exit;

	//			}

	//			//        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
	//			//            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND U_OKYN = 'Y'", 0, 1) > 0 Then
	//			//                MDC_Com.MDC_GF_Message "구매요청 결재가 승인되었습니다. 닫기(종료) 처리할 수 없습니다.", "W"
	//			//                PS_PP030_Validate = False
	//			//                GoTo PS_PP030_Validate_Exit
	//			//            End If
	//			//        End If

	//			//        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) > 0 Then
	//			//            MDC_Com.MDC_GF_Message "작업일보가 등록되었습니다. 닫기(종료) 처리할 수 없습니다.", "W"
	//			//            PS_PP030_Validate = False
	//			//            GoTo PS_PP030_Validate_Exit
	//			//        End If

	//			//        '//삭제된행중에 외주반출등록된행
	//			//        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
	//			//            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
	//			//                MDC_Com.MDC_GF_Message "외주반출 등록된 행입니다. 닫기(종료) 처리할 수 없습니다.", "W"
	//			//                PS_PP030_Validate = False
	//			//                GoTo PS_PP030_Validate_Exit
	//			//            End If
	//			//        End If
	//			//
	//			//        '//삭제된행중에 외주청구등록된행
	//			//        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구일때
	//			//            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
	//			//                MDC_Com.MDC_GF_Message "외주청구 등록된 행입니다. 닫기(종료) 처리할 수 없습니다.", "W"
	//			//                PS_PP030_Validate = False
	//			//                GoTo PS_PP030_Validate_Exit
	//			//            End If
	//			//        End If

	//		}
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		return functionReturnValue;
	//		PS_PP030_Validate_Exit:
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		return functionReturnValue;
	//		PS_PP030_Validate_Error:
	//		functionReturnValue = false;
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		return functionReturnValue;
	//	}

	//	private void PS_PP030_PurchaseRequest(int oDocEntry02, int oLineId02)
	//	{
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		////구매요청
	//		////조달방식이 청구이면 [@PS_MM005H] 에 추가, 구매요청의 결재(OKYN) 값이 Y로 변경된 경우 수정불가, 작지에서는 청구행에 대해 행삭제불가
	//		string query01 = null;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		string itemName = null;
	//		RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string Query02 = null;
	//		SAPbobsCOM.Recordset RecordSet02 = null;
	//		RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string DocEntry = null;

	//		query01 = "SELECT ";
	//		query01 = query01 + "'" + DocEntry + "',";
	//		query01 = query01 + "'" + DocEntry + "',";
	//		query01 = query01 + " PS_PP030L.U_ItemCode,";
	//		query01 = query01 + " PS_PP030L.U_ItemName,";
	//		query01 = query01 + " PS_PP030L.U_Weight,";
	//		query01 = query01 + " PS_PP030L.U_Weight,";
	//		query01 = query01 + " 0,";
	//		query01 = query01 + " 0,";
	//		query01 = query01 + " PS_PP030H.U_BPLId,";
	//		query01 = query01 + "'" + DocEntry + "',";
	//		query01 = query01 + " CONVERT(NVARCHAR,GETDATE(),112),";
	//		query01 = query01 + " CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),";
	//		query01 = query01 + " PS_PP030L.U_CntcCode,";
	//		query01 = query01 + " PS_PP030L.U_CntcName,";
	//		query01 = query01 + " (SELECT dept FROM [OHEM] WHERE empID = PS_PP030L.U_CntcCode),";
	//		query01 = query01 + " '',";
	//		query01 = query01 + " 'N',";
	//		query01 = query01 + " 'Y',";
	//		query01 = query01 + " '10',";
	//		query01 = query01 + " PS_PP030L.U_Comments,";
	//		query01 = query01 + " 'N',";
	//		query01 = query01 + " '',";
	//		query01 = query01 + " '10',";
	//		query01 = query01 + " '',";
	//		query01 = query01 + " 'O',";
	//		query01 = query01 + " PS_PP030H.DocEntry,";
	//		query01 = query01 + " PS_PP030L.LineId,";
	//		query01 = query01 + " CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030L.LineId),";
	//		//// 청구시 필드 추가 - 류영조
	//		query01 = query01 + " CONVERT(NVARCHAR,PS_PP030L.U_CGDate,112) As CGDate,";
	//		query01 = query01 + " PS_PP030H.U_OrdNum + '-' + PS_PP030H.U_OrdSub1 + '-' + PS_PP030H.U_OrdSub2 As OrdNum,";
	//		query01 = query01 + " PS_PP030L.U_ImportYN As ImportYN,";
	//		//수입품여부
	//		query01 = query01 + " PS_PP030L.U_EmergYN As EmergYN,";
	//		//긴급여부
	//		query01 = query01 + " PS_PP030L.U_RCode As RCode,";
	//		//재작업사유
	//		query01 = query01 + " PS_PP030L.U_RName As RName,";
	//		//재작업사유내용
	//		query01 = query01 + " PS_PP030L.U_PartNo As PartNo";
	//		//PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
	//		///'''''''''''''''''''''''''''''''''''''''
	//		query01 = query01 + " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry";
	//		query01 = query01 + " WHERE PS_PP030H.DocEntry = '" + oDocEntry02 + "'";
	//		query01 = query01 + " AND PS_PP030L.LineId = '" + oLineId02 + "'";
	//		query01 = query01 + " AND PS_PP030H.Canceled = 'N'";
	//		RecordSet01.DoQuery(query01);

	//		itemName = MDC_PS_Common.Make_ItemName(Strings.Trim(RecordSet01.Fields.Item(3).Value));

	//		//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		DocEntry = MDC_PS_Common.GetValue("SELECT CASE WHEN ISNULL(MAX(CONVERT(INT,DocEntry)),0) = 0 THEN LEFT(CONVERT(NVARCHAR,'" + RecordSet01.Fields.Item("CGDate").Value + "',112),6) + '0001' ELSE ISNULL(MAX(CONVERT(INT,DocEntry)),0)+1 END FROM [@PS_MM005H] WHERE LEFT(CONVERT(NVARCHAR,'" + RecordSet01.Fields.Item("CGDate").Value + "',112),6) = LEFT(DocEntry,6)");

	//		////구매요청이 취소되면 안되고 삭제되어야 한다.. 삭제하면서 작업지시도 동시삭제, 단 작업지시에 행이 1개만 존재한다면 삭제할수 없다.
	//		Query02 = " SELECT COUNT(*)";
	//		Query02 = Query02 + " FROM [@PS_MM005H]";
	//		Query02 = Query02 + " WHERE U_OrdType = '10' AND U_PP030HNo = '" + oDocEntry02 + "'";
	//		Query02 = Query02 + " AND U_PP030LNo = '" + oLineId02 + "'";
	//		//Query01 = Query01 & " AND Canceled = 'N'"
	//		RecordSet02.DoQuery(Query02);
	//		if (RecordSet02.Fields.Item(0).Value == 0) {
	//			query01 = "INSERT INTO [@PS_MM005H]";
	//			query01 = query01 + " (";
	//			query01 = query01 + " DocEntry,";
	//			query01 = query01 + " DocNum,";
	//			query01 = query01 + " U_ItemCode,";
	//			query01 = query01 + " U_ItemName,";
	//			query01 = query01 + " U_Qty,";
	//			query01 = query01 + " U_Weight,";
	//			//        Query01 = Query01 & " U_Price,"
	//			//        Query01 = Query01 & " U_LinTotal,"
	//			query01 = query01 + " U_BPLId,";
	//			query01 = query01 + " U_CgNum,";
	//			query01 = query01 + " U_DocDate,";
	//			query01 = query01 + " U_DueDate,";
	//			query01 = query01 + " U_CntcCode,";
	//			query01 = query01 + " U_CntcName,";
	//			query01 = query01 + " U_DeptCode,";
	//			query01 = query01 + " U_UseDept,";
	//			query01 = query01 + " U_Auto,";
	//			query01 = query01 + " U_QCYN,";
	//			//Query01 = Query01 & " U_ReType,"
	//			///'''''        Query01 = Query01 & " U_Note,"
	//			query01 = query01 + " U_OKYN,";
	//			query01 = query01 + " U_OKDate,";
	//			query01 = query01 + " U_OrdType,";
	//			query01 = query01 + " U_ProcCode,";
	//			query01 = query01 + " U_Status,";
	//			//// 청구시 필드 추가 - 류영조
	//			//        Query01 = Query01 & " U_DocDate,"
	//			//        Query01 = Query01 & " U_DueDate,"
	//			query01 = query01 + " U_Comments,";
	//			query01 = query01 + " U_OrdNum,";
	//			///''''''''''''''''''''''''''''''''
	//			query01 = query01 + " U_PP030HNo,";
	//			query01 = query01 + " U_PP030LNo,";
	//			query01 = query01 + " U_PP030DL,";
	//			query01 = query01 + " U_ImportYN,";
	//			//수입품여부(2018.09.12 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_EmergYN,";
	//			//긴급여부(2018.09.12 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_RCode,";
	//			//재청구사유(2018.09.17 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_RName,";
	//			//재청구사유내용(2018.09.17 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_PartNo,";
	//			//PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
	//			query01 = query01 + " UserSign,";
	//			//UserSign 추가(2020.04.16 송명규)
	//			query01 = query01 + " CreateDate";
	//			//생성일 추가(2014.02.24 송명규)
	//			query01 = query01 + " ) ";
	//			query01 = query01 + "VALUES(";
	//			query01 = query01 + "'" + DocEntry + "',";
	//			////DocEntry
	//			query01 = query01 + "'" + DocEntry + "',";
	//			////DocNum
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(2).Value + "',";
	//			////ItemCode
	//			query01 = query01 + "'" + itemName + "',";
	//			////ItemName
	//			query01 = query01 + "" + RecordSet01.Fields.Item(4).Value + ",";
	//			////Qty
	//			query01 = query01 + "" + RecordSet01.Fields.Item(5).Value + ",";
	//			////Weight
	//			//        Query01 = Query01 & "'" & RecordSet01.Fields(6).Value & "'," '//Price
	//			//        Query01 = Query01 & "'" & RecordSet01.Fields(7).Value & "'," '//LineTotal
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(8).Value + "',";
	//			////BPLId
	//			query01 = query01 + "'" + DocEntry + "',";
	//			////CgNum  'RecordSet01.Fields(9).Value
	//			query01 = query01 + "'" + RecordSet01.Fields.Item("CGDate").Value + "',";
	//			////DocDate
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(11).Value + "',";
	//			////DueDate
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(12).Value + "',";
	//			////CntcCode
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(13).Value + "',";
	//			////CntcName
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(14).Value + "',";
	//			////DeptCode
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(15).Value + "',";
	//			////UseDept
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(16).Value + "',";
	//			////Auto
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(17).Value + "',";
	//			////QCYN
	//			//Query01 = Query01 & "'" & RecordSet01.Fields(18).Value & "'," '//ReType
	//			///'''        Query01 = Query01 & "'" & RecordSet01.Fields(19).Value & "'," '//Note
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(20).Value + "',";
	//			//OKYN
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(21).Value + "',";
	//			//OKDate
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(22).Value + "',";
	//			//OrdType
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(23).Value + "',";
	//			//ProcCode
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(24).Value + "',";
	//			//Status
	//			//// 청구시 필드 추가 - 류영조
	//			//        Query01 = Query01 & "'" & RecordSet01.Fields("CGDate").Value & "'," 'U_DocDate
	//			//        Query01 = Query01 & "'" & RecordSet01.Fields("CGDate").Value & "'," 'U_DueDate
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(19).Value + "',";
	//			//U_Comments
	//			query01 = query01 + "'" + RecordSet01.Fields.Item("OrdNum").Value + "',";
	//			//U_OrdNum
	//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(25).Value + "',";
	//			//PP030HNo
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(26).Value + "',";
	//			//PP030LNo
	//			query01 = query01 + "'" + RecordSet01.Fields.Item(27).Value + "',";
	//			//PP030DL
	//			query01 = query01 + "'" + RecordSet01.Fields.Item("ImportYN").Value + "',";
	//			//수입품여부(2018.09.12 송명규, 김석태 과장 요청)
	//			query01 = query01 + "'" + RecordSet01.Fields.Item("EmergYN").Value + "',";
	//			//긴급여부(2018.09.12 송명규, 김석태 과장 요청)
	//			query01 = query01 + "'" + RecordSet01.Fields.Item("RCode").Value + "',";
	//			//재청구사유(2018.09.17 송명규, 김석태 과장 요청)
	//			query01 = query01 + "'" + RecordSet01.Fields.Item("RName").Value + "',";
	//			//재청구사유내용(2018.09.17 송명규, 김석태 과장 요청)
	//			query01 = query01 + "'" + RecordSet01.Fields.Item("PartNo").Value + "',";
	//			//PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
	//			query01 = query01 + "'" + SubMain.Sbo_Company.UserSignature + "',";
	//			//UserSign 추가(2020.04.16 송명규)
	//			query01 = query01 + " GETDATE()";
	//			//생성일 추가(2014.02.24 송명규)
	//			query01 = query01 + ")";
	//			RecordSet01.DoQuery(query01);
	//		} else {
	//			query01 = "UPDATE [@PS_MM005H] SET";
	//			query01 = query01 + " U_ItemCode = '" + RecordSet01.Fields.Item(2).Value + "',";
	//			query01 = query01 + " U_ItemName = '" + itemName + "',";
	//			query01 = query01 + " U_Qty = " + RecordSet01.Fields.Item(4).Value + ",";
	//			query01 = query01 + " U_Weight = " + RecordSet01.Fields.Item(5).Value + ",";
	//			//        Query01 = Query01 & " U_Price = '" & RecordSet01.Fields(6).Value & "',"
	//			//        Query01 = Query01 & " U_LinTotal = '" & RecordSet01.Fields(7).Value & "',"
	//			query01 = query01 + " U_BPLId = '" + RecordSet01.Fields.Item(8).Value + "',";
	//			query01 = query01 + " U_DocDate = '" + RecordSet01.Fields.Item("CGDate").Value + "',";
	//			query01 = query01 + " U_DueDate = '" + RecordSet01.Fields.Item(11).Value + "',";
	//			query01 = query01 + " U_CntcCode = '" + RecordSet01.Fields.Item(12).Value + "',";
	//			query01 = query01 + " U_CntcName = '" + RecordSet01.Fields.Item(13).Value + "',";
	//			query01 = query01 + " U_DeptCode = '" + RecordSet01.Fields.Item(14).Value + "',";
	//			query01 = query01 + " U_UseDept = '" + RecordSet01.Fields.Item(15).Value + "',";
	//			query01 = query01 + " U_Auto = '" + RecordSet01.Fields.Item(16).Value + "',";
	//			query01 = query01 + " U_QCYN = '" + RecordSet01.Fields.Item(17).Value + "',";
	//			//        Query01 = Query01 & " U_ReType = '" & RecordSet01.Fields(18).Value & "',"
	//			query01 = query01 + " U_Note = '" + RecordSet01.Fields.Item(19).Value + "',";
	//			//        Query01 = Query01 & " U_OKYN = '" & RecordSet01.Fields(20).Value & "',"
	//			//        Query01 = Query01 & " U_OKDate = '" & RecordSet01.Fields(21).Value & "',"
	//			query01 = query01 + " U_OrdType = '" + RecordSet01.Fields.Item(22).Value + "',";
	//			query01 = query01 + " U_ProcCode = '" + RecordSet01.Fields.Item(23).Value + "',";
	//			//// 청구시 필드 추가 - 류영조
	//			//        Query01 = Query01 & " U_DocDate = '" & RecordSet01.Fields("CGDate").Value & "',"
	//			//        Query01 = Query01 & " U_DueDate = '" & RecordSet01.Fields("CGDate").Value & "',"
	//			query01 = query01 + " U_Comments = '" + RecordSet01.Fields.Item(19).Value + "',";
	//			query01 = query01 + " U_ImportYN = '" + RecordSet01.Fields.Item("ImportYN").Value + "',";
	//			//수입품여부(2018.09.12 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_EmergYN = '" + RecordSet01.Fields.Item("EmergYN").Value + "',";
	//			//긴급여부(2018.09.12 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_RCode = '" + RecordSet01.Fields.Item("RCode").Value + "',";
	//			//재작업사유(2018.09.17 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_RName = '" + RecordSet01.Fields.Item("RName").Value + "',";
	//			//재작업사유내용(2018.09.17 송명규, 김석태 과장 요청)
	//			query01 = query01 + " U_PartNo = '" + RecordSet01.Fields.Item("PartNo").Value + "',";
	//			//PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
	//			query01 = query01 + " UserSign = '" + SubMain.Sbo_Company.UserSignature + "',";
	//			//UserSign(2020.04.16 송명규)
	//			query01 = query01 + " UpdateDate = GETDATE()";
	//			//수정일 추가(2014.02.24 송명규)
	//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	//			query01 = query01 + " WHERE U_OrdType = '10' And U_PP030HNo = '" + RecordSet01.Fields.Item(25).Value + "'";
	//			query01 = query01 + " AND U_PP030LNo = '" + RecordSet01.Fields.Item(26).Value + "'";
	//			RecordSet01.DoQuery(query01);
	//		}

	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet02 = null;
	//		return;
	//		PS_PP030_PurchaseRequest_Error:
	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet02 = null;
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_PurchaseRequest_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//	}

	//	private bool PS_PP030_AutoCreateMultiGage()
	//	{
	//		bool functionReturnValue = false;
	//		 // ERROR: Not supported in C#: OnErrorStatement

	//		functionReturnValue = true;
	//		object j = null;
	//		object i = null;
	//		object h = null;
	//		int s = 0;
	//		SAPbobsCOM.Recordset RecordSet01 = null;
	//		SAPbobsCOM.Recordset RecordSet02 = null;
	//		SAPbobsCOM.Recordset RecordSet03 = null;
	//		string query01 = null;
	//		string Query02 = null;
	//		string Query03 = null;
	//		RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
	//		RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
	//		RecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		int CurrentDocEntry = 0;

	//		if (SubMain.Sbo_Company.InTransaction == true) {
	//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
	//		}
	//		SubMain.Sbo_Company.StartTransaction();

	//		////투입자재의 수량만큼
	//		for (i = 1; i <= oMat02.VisualRowCount; i++) {
	//			query01 = "SELECT AutoKey FROM [ONNM] WHERE ObjectCode = 'PS_PP030'";
	//			RecordSet01.DoQuery(query01);
	//			////PS_PP031의 최종문서번호
	//			//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			CurrentDocEntry = RecordSet01.Fields.Item(0).Value;
	//			query01 = "UPDATE [ONNM] SET AutoKey = AutoKey +1 WHERE ObjectCode = 'PS_PP030'";
	//			RecordSet01.DoQuery(query01);
	//			////문서번호증가, 이전문서번호는 현재 유저가 선점

	//			query01 = "INSERT INTO [@PS_PP030H] (";
	//			query01 = query01 + "DocEntry,";
	//			query01 = query01 + "DocNum,";
	//			query01 = query01 + "Period,";
	//			query01 = query01 + "Instance,";
	//			query01 = query01 + "Series,";
	//			query01 = query01 + "Handwrtten,";
	//			query01 = query01 + "Canceled,";
	//			query01 = query01 + "Object,";
	//			query01 = query01 + "LogInst,";
	//			query01 = query01 + "UserSign,";
	//			query01 = query01 + "Transfered,";
	//			query01 = query01 + "Status,";
	//			query01 = query01 + "CreateDate,";
	//			query01 = query01 + "CreateTime,";
	//			query01 = query01 + "UpdateDate,";
	//			query01 = query01 + "UpdateTime,";
	//			query01 = query01 + "DataSource,";
	//			query01 = query01 + "U_BaseType,";
	//			query01 = query01 + "U_BaseNum,";
	//			query01 = query01 + "U_OrdGbn,";
	//			query01 = query01 + "U_DocDate,";
	//			query01 = query01 + "U_DueDate,";
	//			query01 = query01 + "U_ItemCode,";
	//			query01 = query01 + "U_ItemName,";
	//			query01 = query01 + "U_CntcCode,";
	//			query01 = query01 + "U_CntcName,";
	//			query01 = query01 + "U_SjNum,";
	//			query01 = query01 + "U_SjLine,";
	//			query01 = query01 + "U_OrdMgNum,";
	//			query01 = query01 + "U_OrdNum,";
	//			query01 = query01 + "U_OrdSub1,";
	//			query01 = query01 + "U_OrdSub2,";
	//			query01 = query01 + "U_JakMyung,";
	//			query01 = query01 + "U_ReqWt,";
	//			query01 = query01 + "U_SelWt,";
	//			query01 = query01 + "U_LotNo,";
	//			query01 = query01 + "U_SjPrice,";
	//			query01 = query01 + "U_MulGbn1,";
	//			query01 = query01 + "U_MulGbn2,";
	//			query01 = query01 + "U_MulGbn3,";
	//			query01 = query01 + "U_Comments,";
	//			query01 = query01 + "U_BPLId,";
	//			query01 = query01 + "U_BasicGub";
	//			query01 = query01 + ")";
	//			query01 = query01 + " VALUES(";
	//			query01 = query01 + "'" + CurrentDocEntry + "'" + ",";
	//			query01 = query01 + "'" + CurrentDocEntry + "'" + ",";
	//			query01 = query01 + "'11'" + ",";
	//			query01 = query01 + "'0'" + ",";
	//			query01 = query01 + "'-1'" + ",";
	//			query01 = query01 + "'N'" + ",";
	//			query01 = query01 + "'N'" + ",";
	//			query01 = query01 + "'PS_PP030'" + ",";
	//			query01 = query01 + "NULL" + ",";
	//			query01 = query01 + "'" + SubMain.Sbo_Company.UserSignature + "'" + ",";
	//			query01 = query01 + "'N'" + ",";
	//			query01 = query01 + "'O'" + ",";
	//			////Status
	//			query01 = query01 + "CONVERT(NVARCHAR,GETDATE(),112)" + ",";
	//			query01 = query01 + "SUBSTRING(CONVERT(NVARCHAR,GETDATE(),108),1,2) + SUBSTRING(CONVERT(NVARCHAR,GETDATE(),108),4,2)" + ",";
	//			query01 = query01 + "NULL" + ",";
	//			////UpdateDate
	//			query01 = query01 + "NULL" + ",";
	//			////UpdateTime
	//			query01 = query01 + "'I'" + ",";
	//			////DataSource
	//			query01 = query01 + "NULL" + ",";
	//			////BaseType
	//			query01 = query01 + "NULL" + ",";
	//			////BaseNum
	//			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("OrdGbn").Specific.Selected.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (string.IsNullOrEmpty(oForm01.Items.Item("DocDate").Specific.VALUE)) {
	//				query01 = query01 + "NULL" + ",";
	//			} else {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oForm01.Items.Item("DocDate").Specific.VALUE + "'" + ",";
	//			}
	//			//UPGRADE_WARNING: oForm01.Items(DueDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (string.IsNullOrEmpty(oForm01.Items.Item("DueDate").Specific.VALUE)) {
	//				query01 = query01 + "NULL" + ",";
	//			} else {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oForm01.Items.Item("DueDate").Specific.VALUE + "'" + ",";
	//			}
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("ItemCode").Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("ItemName").Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("CntcCode").Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("CntcName").Specific.VALUE + "'" + ",";
	//			query01 = query01 + "NULL" + ",";
	//			////SjNum
	//			query01 = query01 + "NULL" + ",";
	//			////SjLine
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("OrdMgNum").Specific.VALUE + "'" + ",";
	//			////신규작업지시번호를 조회
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP030_01 ' & oForm01.Items(OrdMgNum).Specific.VALUE & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("OrdMgNum").Specific.VALUE + MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" + oForm01.Items.Item("OrdMgNum").Specific.VALUE + "'") + "'" + ",";
	//			query01 = query01 + "'" + "00" + "'" + ",";
	//			query01 = query01 + "'" + "000" + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("JakMyung").Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("ReqWt").Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("Weight").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			////투입자재의 중량으로 입력되어야함
	//			query01 = query01 + "NULL" + ",";
	//			////LotNo
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("SjPrice").Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items(MulGbn1).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (oForm01.Items.Item("MulGbn1").Specific.Selected == null) {
	//				query01 = query01 + "'" + "" + "'" + ",";
	//			} else {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + Strings.Trim(oForm01.Items.Item("MulGbn1").Specific.Selected.VALUE) + "'" + ",";
	//			}
	//			//UPGRADE_WARNING: oForm01.Items(MulGbn2).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (oForm01.Items.Item("MulGbn2").Specific.Selected == null) {
	//				query01 = query01 + "'" + "" + "'" + ",";
	//			} else {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + Strings.Trim(oForm01.Items.Item("MulGbn2").Specific.Selected.VALUE) + "'" + ",";
	//			}
	//			//UPGRADE_WARNING: oForm01.Items(MulGbn3).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (oForm01.Items.Item("MulGbn3").Specific.Selected == null) {
	//				query01 = query01 + "'" + "" + "'" + ",";
	//			} else {
	//				//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + Strings.Trim(oForm01.Items.Item("MulGbn3").Specific.Selected.VALUE) + "'" + ",";
	//			}
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + Strings.Trim(oForm01.Items.Item("Comments").Specific.VALUE) + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE) + "'" + ",";
	//			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oForm01.Items.Item("BasicGub").Specific.Selected.VALUE + "'";
	//			query01 = query01 + ")";
	//			RecordSet01.DoQuery(query01);

	//			query01 = "INSERT INTO [@PS_PP030L] (";
	//			query01 = query01 + "DocEntry,";
	//			query01 = query01 + "LineId,";
	//			query01 = query01 + "VisOrder,";
	//			query01 = query01 + "Object,";
	//			query01 = query01 + "LogInst,";
	//			query01 = query01 + "U_LineNum,";
	//			query01 = query01 + "U_InputGbn,";
	//			query01 = query01 + "U_ItemCode,";
	//			query01 = query01 + "U_ItemName,";
	//			query01 = query01 + "U_ItemGpCd,";
	//			query01 = query01 + "U_Weight,";
	//			query01 = query01 + "U_DueDate,";
	//			query01 = query01 + "U_CntcCode,";
	//			query01 = query01 + "U_CntcName,";
	//			query01 = query01 + "U_ProcType,";
	//			query01 = query01 + "U_Comments,";
	//			query01 = query01 + "U_BatchNum,";
	//			query01 = query01 + "U_LineId";
	//			query01 = query01 + ")";
	//			query01 = query01 + " VALUES(";
	//			query01 = query01 + "'" + CurrentDocEntry + "'" + ",";
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + i + "'" + ",";
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + i - 1 + "'" + ",";
	//			query01 = query01 + "'PS_PP030'" + ",";
	//			query01 = query01 + "NULL" + ",";
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + i + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.Selected.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("ItemName").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("Weight").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns(DueDate).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			if (string.IsNullOrEmpty(oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.VALUE)) {
	//				query01 = query01 + "NULL" + ",";
	//			} else {
	//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			}
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("CntcCode").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("CntcName").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("Comments").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + oMat02.Columns.Item("BatchNum").Cells.Item(i).Specific.VALUE + "'" + ",";
	//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//			query01 = query01 + "'" + i + "'";
	//			query01 = query01 + ")";
	//			RecordSet01.DoQuery(query01);

	//			for (j = 1; j <= oMat03.VisualRowCount; j++) {
	//				query01 = "INSERT INTO [@PS_PP030M] (";
	//				query01 = query01 + "DocEntry,";
	//				query01 = query01 + "LineId,";
	//				query01 = query01 + "VisOrder,";
	//				query01 = query01 + "Object,";
	//				query01 = query01 + "LogInst,";
	//				query01 = query01 + "U_LineNum,";
	//				query01 = query01 + "U_Sequence,";
	//				query01 = query01 + "U_CpBCode,";
	//				query01 = query01 + "U_CpBName,";
	//				query01 = query01 + "U_CpCode,";
	//				query01 = query01 + "U_CpName,";
	//				query01 = query01 + "U_StdHour,";
	//				query01 = query01 + "U_Unit,";
	//				query01 = query01 + "U_ReDate,";
	//				query01 = query01 + "U_WorkGbn,";
	//				query01 = query01 + "U_ReWorkYN,";
	//				query01 = query01 + "U_ResultYN,";
	//				query01 = query01 + "U_ReportYN,";
	//				query01 = query01 + "U_LineId";
	//				query01 = query01 + ")";
	//				query01 = query01 + " VALUES(";
	//				query01 = query01 + "'" + CurrentDocEntry + "'" + ",";
	//				//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + j + "'" + ",";
	//				//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + j - 1 + "'" + ",";
	//				query01 = query01 + "'PS_PP030'" + ",";
	//				query01 = query01 + "NULL" + ",";
	//				//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + j + "'" + ",";
	//				//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + j + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("CpBCode").Cells.Item(j).Specific.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("CpBName").Cells.Item(j).Specific.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("CpCode").Cells.Item(j).Specific.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("CpName").Cells.Item(j).Specific.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("StdHour").Cells.Item(j).Specific.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("Unit").Cells.Item(j).Specific.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns(ReDate).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				if (string.IsNullOrEmpty(oMat03.Columns.Item("ReDate").Cells.Item(j).Specific.VALUE)) {
	//					query01 = query01 + "NULL" + ",";
	//				} else {
	//					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//					query01 = query01 + "'" + oMat03.Columns.Item("ReDate").Cells.Item(j).Specific.VALUE + "'" + ",";
	//				}
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("WorkGbn").Cells.Item(j).Specific.Selected.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("ReWorkYN").Cells.Item(j).Specific.Selected.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("ResultYN").Cells.Item(j).Specific.Selected.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + oMat03.Columns.Item("ReportYN").Cells.Item(j).Specific.Selected.VALUE + "'" + ",";
	//				//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//				query01 = query01 + "'" + j + "'";
	//				query01 = query01 + ")";
	//				RecordSet01.DoQuery(query01);
	//			}
	//		}

	//		if (SubMain.Sbo_Company.InTransaction == true) {
	//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
	//		}

	//		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet01 = null;
	//		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet02 = null;
	//		//UPGRADE_NOTE: RecordSet03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		RecordSet03 = null;
	//		return functionReturnValue;
	//		PS_PP030_AutoCreateMultiGage_Error:
	//		if (SubMain.Sbo_Company.InTransaction == true) {
	//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
	//		}
	//		functionReturnValue = false;
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_AutoCreateMultiGage_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		return functionReturnValue;
	//	}

	//	private bool PS_PP030_CheckDate()
	//	{
	//		bool functionReturnValue = false;
	//		//******************************************************************************
	//		//Function ID : PS_PP030_CheckDate()
	//		//해당모듈    : PS_PP030
	//		//기능        : 선행프로세스와 일자 비교
	//		//인수        : 없음
	//		//반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
	//		//특이사항    : 없음
	//		//******************************************************************************
	//		 // ERROR: Not supported in C#: OnErrorStatement


	//		string query01 = null;
	//		short loopCount = 0;
	//		SAPbobsCOM.Recordset oRecordSet01 = null;
	//		oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string baseEntry = null;
	//		string baseLine = null;
	//		string docType = null;
	//		string CurDocDate = null;

	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		baseEntry = Strings.Trim(oForm01.Items.Item("BaseNum").Specific.VALUE);
	//		baseLine = "";
	//		docType = "PS_PP030";
	//		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	//		CurDocDate = Strings.Trim(oForm01.Items.Item("DocDate").Specific.VALUE);

	//		query01 = "         EXEC PS_Z_CHECK_DATE '";
	//		query01 = query01 + baseEntry + "','";
	//		query01 = query01 + baseLine + "','";
	//		query01 = query01 + docType + "','";
	//		query01 = query01 + CurDocDate + "'";

	//		oRecordSet01.DoQuery(query01);

	//		if (oRecordSet01.Fields.Item("ReturnValue").Value == "False") {
	//			functionReturnValue = false;
	//		} else {
	//			functionReturnValue = true;
	//		}

	//		//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oRecordSet01 = null;
	//		return functionReturnValue;
	//		PS_PP030_CheckDate_Error:

	//		functionReturnValue = false;
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_CheckDate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		return functionReturnValue;
	//	}

	//	private bool PS_PP030_Check_DupReq(string pDocEntry, string pItemCode, string pLineID)
	//	{
	//		bool functionReturnValue = false;
	//		//******************************************************************************
	//		//Function ID : PS_PP030_Check_DupReq()
	//		//해당모듈    : PS_PP030
	//		//기능        : 중복청구 여부 조회
	//		//인수        : pDocEntry(문서번호), pItemCode(원재료품목코드), pLineID(라인번호)
	//		//반환값      : True-중복청구(O), False-중복청구(X)
	//		//특이사항    : 없음
	//		//******************************************************************************
	//		 // ERROR: Not supported in C#: OnErrorStatement


	//		string query01 = null;
	//		short loopCount = 0;
	//		SAPbobsCOM.Recordset oRecordSet01 = null;
	//		oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

	//		string DocEntry = null;
	//		string itemCode = null;
	//		string LineId = null;

	//		DocEntry = pDocEntry;
	//		//Trim(oForm01.Items("DocEntry").Specific.VALUE)
	//		itemCode = pItemCode;
	//		LineId = pLineID;

	//		query01 = "         EXEC PS_Z_Check_DupReq '";
	//		query01 = query01 + DocEntry + "','";
	//		query01 = query01 + itemCode + "','";
	//		query01 = query01 + LineId + "'";

	//		oRecordSet01.DoQuery(query01);

	//		if (oRecordSet01.Fields.Item("ReturnValue").Value == "FALSE") {
	//			functionReturnValue = false;
	//		} else {
	//			functionReturnValue = true;
	//		}

	//		//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	//		oRecordSet01 = null;
	//		return functionReturnValue;
	//		PS_PP030_CheckDate_Error:

	//		functionReturnValue = false;
	//		SubMain.Sbo_Application.SetStatusBarMessage("PS_PP030_Check_DupReq_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
	//		return functionReturnValue;
	//	}
	}
}
