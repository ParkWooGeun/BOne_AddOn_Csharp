using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 불량중분류코드등록
	/// </summary>
	internal class PS_PP002 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_PP002H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP002L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP002.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP002_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP002");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				//PS_PP002_CreateItems();
				//PS_PP002_ComboBox_Setting();
				//PS_PP002_CF_ChooseFromList();
				//PS_PP002_EnableMenus();
				//PS_PP002_SetDocument(oFormDocEntry);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}

		}

		//public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	switch (pVal.EventType) {
		//		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
		//			////1
		//			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
		//			////2
		//			Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
		//			////5
		//			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_CLICK:
		//			////6
		//			Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
		//			////7
		//			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
		//			////8
		//			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
		//			////10
		//			Raise_EVENT_VALIDATE(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
		//			////11
		//			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
		//			////18
		//			break;
		//		////et_FORM_ACTIVATE
		//		case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
		//			////19
		//			break;
		//		////et_FORM_DEACTIVATE
		//		case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
		//			////20
		//			Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
		//			////27
		//			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
		//			////3
		//			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
		//			////4
		//			break;
		//		////et_LOST_FOCUS
		//		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
		//			////17
		//			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
		//			break;
		//	}
		//	return;
		//	Raise_ItemEvent_Error:
		//	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}


		//public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////BeforeAction = True
		//	if ((pVal.BeforeAction == true)) {
		//		switch (pVal.MenuUID) {
		//			case "1284":
		//				//취소
		//				break;
		//			case "1286":
		//				//닫기
		//				break;
		//			case "1293":
		//				//행삭제
		//				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
		//				break;
		//			case "1281":
		//				//찾기
		//				break;
		//			case "1282":
		//				//추가
		//				break;
		//			case "1288":
		//			case "1289":
		//			case "1290":
		//			case "1291":
		//				//레코드이동버튼
		//				break;
		//		}
		//	////BeforeAction = False
		//	} else if ((pVal.BeforeAction == false)) {
		//		switch (pVal.MenuUID) {
		//			case "1284":
		//				//취소
		//				break;
		//			case "1286":
		//				//닫기
		//				break;
		//			case "1293":
		//				//행삭제
		//				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
		//				break;
		//			case "1281":
		//				//찾기
		//				PS_PP002_FormItemEnabled();
		//				////UDO방식
		//				break;
		//			case "1282":
		//				//추가
		//				PS_PP002_FormItemEnabled();
		//				////UDO방식
		//				PS_PP002_AddMatrixRow(0, ref true);
		//				////UDO방식
		//				break;
		//			case "1288":
		//			case "1289":
		//			case "1290":
		//			case "1291":
		//				//레코드이동버튼
		//				PS_PP002_FormItemEnabled();
		//				break;
		//		}
		//	}
		//	return;
		//	Raise_MenuEvent_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////BeforeAction = True
		//	if ((BusinessObjectInfo.BeforeAction == true)) {
		//		switch (BusinessObjectInfo.EventType) {
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
		//				////33
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
		//				////34
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
		//				////35
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
		//				////36
		//				break;
		//		}
		//	////BeforeAction = False
		//	} else if ((BusinessObjectInfo.BeforeAction == false)) {
		//		switch (BusinessObjectInfo.EventType) {
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
		//				////33
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
		//				////34
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
		//				////35
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
		//				////36
		//				break;
		//		}
		//	}
		//	return;
		//	Raise_FormDataEvent_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {
		//		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
		//		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
		//		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
		//		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
		//		//            MenuCreationParams01.uniqueID = "MenuUID"
		//		//            MenuCreationParams01.String = "메뉴명"
		//		//            MenuCreationParams01.Enabled = True
		//		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
		//		//        End If
		//	} else if (pVal.BeforeAction == false) {
		//		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
		//		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
		//		//        End If
		//	}
		//	if (pVal.ItemUID == "Mat01") {
		//		if (pVal.Row > 0) {
		//			oLastItemUID01 = pVal.ItemUID;
		//			oLastColUID01 = pVal.ColUID;
		//			oLastColRow01 = pVal.Row;
		//		}
		//	} else {
		//		oLastItemUID01 = pVal.ItemUID;
		//		oLastColUID01 = "";
		//		oLastColRow01 = 0;
		//	}
		//	return;
		//	Raise_RightClickEvent_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {
		//		if (pVal.ItemUID == "PS_PP002") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//			}
		//		}
		//		if (pVal.ItemUID == "1") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
		//				if (PS_PP002_DataValidCheck() == false) {
		//					BubbleEvent = false;
		//					return;
		//				}
		//				////해야할일 작업
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//				if (PS_PP002_DataValidCheck() == false) {
		//					BubbleEvent = false;
		//					return;
		//				}
		//				////해야할일 작업
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//			}
		//		}
		//	} else if (pVal.BeforeAction == false) {
		//		if (pVal.ItemUID == "PS_PP002") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//			}
		//		}
		//		if (pVal.ItemUID == "1") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
		//				if (pVal.ActionSuccess == true) {
		//					PS_PP002_FormItemEnabled();
		//					PS_PP002_AddMatrixRow(oMat01.RowCount, ref true);
		//					////UDO방식일때
		//				}
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//				if (pVal.ActionSuccess == true) {
		//					PS_PP002_FormItemEnabled();
		//				}
		//			}
		//		}
		//	}
		//	return;
		//	Raise_EVENT_ITEM_PRESSED_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {
		//		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "ItemCode", "") '//사용자값활성
		//		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
		//	} else if (pVal.BeforeAction == false) {

		//	}
		//	return;
		//	Raise_EVENT_KEY_DOWN_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {
		//		if (pVal.ItemChanged == true) {
		//			oForm.Freeze(true);
		//			if ((pVal.ItemUID == "BigCode")) {
		//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				oDS_PS_PP002H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
		//			}
		//			oMat01.LoadFromDataSource();
		//			oMat01.AutoResizeColumns();
		//			oForm.Update();
		//			oForm.Freeze(false);
		//		}
		//	}
		//	oForm.Freeze(false);
		//	return;
		//	Raise_EVENT_COMBO_SELECT_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {

		//	}
		//	return;
		//	Raise_EVENT_CLICK_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {

		//	}
		//	return;
		//	Raise_EVENT_DOUBLE_CLICK_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {

		//	}
		//	return;
		//	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	if (pVal.BeforeAction == true) {
		//		if (pVal.ItemChanged == true) {
		//			if ((pVal.ItemUID == "Mat01")) {
		//				if ((pVal.ColUID == "MidCode")) {
		//					////기타작업
		//					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					oDS_PS_PP002L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
		//					if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP002L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
		//						PS_PP002_AddMatrixRow((pVal.Row));
		//					}
		//				} else {
		//					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					oDS_PS_PP002L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
		//				}
		//			} else {
		//				if ((pVal.ItemUID == "DocEntry")) {
		//					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					oDS_PS_PP002H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
		//				} else {
		//					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					oDS_PS_PP002H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
		//				}
		//			}
		//			oMat01.LoadFromDataSource();
		//			oMat01.AutoResizeColumns();
		//			oForm.Update();
		//		}
		//	} else if (pVal.BeforeAction == false) {

		//	}
		//	oForm.Freeze(false);
		//	return;
		//	Raise_EVENT_VALIDATE_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {
		//		PS_PP002_FormItemEnabled();
		//		PS_PP002_AddMatrixRow(oMat01.VisualRowCount);
		//		////UDO방식
		//	}
		//	return;
		//	Raise_EVENT_MATRIX_LOAD_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_RESIZE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {

		//	}
		//	return;
		//	Raise_EVENT_RESIZE_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {
		//		//        If (pVal.ItemUID = "ItemCode") Then
		//		//            Dim oDataTable01 As SAPbouiCOM.DataTable
		//		//            Set oDataTable01 = pVal.SelectedObjects
		//		//            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
		//		//            Set oDataTable01 = Nothing
		//		//        End If
		//		//        If (pVal.ItemUID = "CardCode" Or pVal.ItemUID = "CardName") Then
		//		//            Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP002H", "U_CardCode,U_CardName")
		//		//        End If
		//	}
		//	return;
		//	Raise_EVENT_CHOOSE_FROM_LIST_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}


		//private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.ItemUID == "Mat01") {
		//		if (pVal.Row > 0) {
		//			oLastItemUID01 = pVal.ItemUID;
		//			oLastColUID01 = pVal.ColUID;
		//			oLastColRow01 = pVal.Row;
		//		}
		//	} else {
		//		oLastItemUID01 = pVal.ItemUID;
		//		oLastColUID01 = "";
		//		oLastColRow01 = 0;
		//	}
		//	return;
		//	Raise_EVENT_GOT_FOCUS_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {
		//	} else if (pVal.BeforeAction == false) {
		//		SubMain.RemoveForms(oFormUniqueID);
		//		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//		oForm = null;
		//		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//		oMat01 = null;
		//	}
		//	return;
		//	Raise_EVENT_FORM_UNLOAD_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;
		//	if ((oLastColRow01 > 0)) {
		//		if (pVal.BeforeAction == true) {
		//			////행삭제전 행삭제가능여부검사
		//		} else if (pVal.BeforeAction == false) {
		//			for (i = 1; i <= oMat01.VisualRowCount; i++) {
		//				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
		//			}
		//			oMat01.FlushToDataSource();
		//			oDS_PS_PP002L.RemoveRecord(oDS_PS_PP002L.Size - 1);
		//			oMat01.LoadFromDataSource();
		//			if (oMat01.RowCount == 0) {
		//				PS_PP002_AddMatrixRow(0);
		//			} else {
		//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP002L.GetValue("U_MidCode", oMat01.RowCount - 1)))) {
		//					PS_PP002_AddMatrixRow(oMat01.RowCount);
		//				}
		//			}
		//		}
		//	}
		//	return;
		//	Raise_EVENT_ROW_DELETE_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}


		//private bool PS_PP002_CreateItems()
		//{
		//	bool returnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	string oQuery01 = null;
		//	SAPbobsCOM.Recordset oRecordSet01 = null;
		//	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oDS_PS_PP002H = oForm.DataSources.DBDataSources("@PS_PP002H");
		//	oDS_PS_PP002L = oForm.DataSources.DBDataSources("@PS_PP002L");
		//	oMat01 = oForm.Items.Item("Mat01").Specific;
		//	oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
		//	oMat01.AutoResizeColumns();

		//	//    Call oForm.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
		//	//    Call oForm.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
		//	//    Call oForm.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
		//	//    Call oForm.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")

		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//	oForm.Freeze(false);
		//	return returnValue;
		//	PS_PP002_CreateItems_Error:
		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	return returnValue;
		//}

		//public void PS_PP002_ComboBox_Setting()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	////콤보에 기본값설정
		//	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP002", "Mat01", "ItemCode", "01", "완제품")
		//	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP002", "Mat01", "ItemCode", "02", "반제품")
		//	//    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP002", "Mat01", "ItemCode")
		//	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP002", "ItemCode", "", "01", "완제품")
		//	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP002", "ItemCode", "", "02", "반제품")
		//	//    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm.Items("Item").Specific, "PS_PP002", "ItemCode")

		//	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("BigCode").Specific), ref "SELECT U_Minor,U_CdName FROM [@PS_SY001L] WHERE Code = 'Q001' order by LineId", ref "1", ref false, ref false);
		//	//    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
		//	oForm.Freeze(false);
		//	return;
		//	PS_PP002_ComboBox_Setting_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void PS_PP002_CF_ChooseFromList()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	//    '//ChooseFromList 설정
		//	//    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
		//	//    Dim oCons               As SAPbouiCOM.Conditions
		//	//    Dim oCon                As SAPbouiCOM.Condition
		//	//    Dim oCFL                As SAPbouiCOM.ChooseFromList
		//	//    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
		//	//    Dim oEdit               As SAPbouiCOM.EditText
		//	//    Dim oColumn             As SAPbouiCOM.Column
		//	//
		//	//    Set oEdit = oForm.Items("CARDCODE").Specific
		//	//    Set oCFLs = oForm.ChooseFromLists
		//	//    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
		//	//
		//	//    oCFLCreationParams.ObjectType = lf_BusinessPartner
		//	//    oCFLCreationParams.uniqueID = "CFLCARDCODE"
		//	//    oCFLCreationParams.MultiSelection = False
		//	//'    Set oCFL = oCFLs.Add(oCFLCreationParams)

		//	//    Set oCons = oCFL.GetConditions()
		//	//    Set oCon = oCons.Add()
		//	//    oCon.Alias = "CardType"
		//	//    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
		//	//    oCon.CondVal = "C"
		//	//    oCFL.SetConditions oCons
		//	//
		//	//    oEdit.ChooseFromListUID = "CFLCARDCODE"
		//	//    oEdit.ChooseFromListAlias = "CardCode"
		//	//
		//	//    Set oEdit = oForm.Items("CARDNAME").Specific
		//	//    Set oCFLs = oForm.ChooseFromLists
		//	//    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
		//	//
		//	//    oCFLCreationParams.ObjectType = lf_BusinessPartner
		//	//    oCFLCreationParams.uniqueID = "CFLCARDNAME"
		//	//    oCFLCreationParams.MultiSelection = False
		//	//    Set oCFL = oCFLs.Add(oCFLCreationParams)

		//	//    Set oCons = oCFL.GetConditions()
		//	//    Set oCon = oCons.Add()
		//	//    oCon.Alias = "CardType"
		//	//    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
		//	//    oCon.CondVal = "C"
		//	//    oCFL.SetConditions oCons
		//	//
		//	//    oEdit.ChooseFromListUID = "CFLCARDNAME"
		//	//    oEdit.ChooseFromListAlias = "CardName"
		//	return;
		//	PS_PP002_CF_ChooseFromList_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void PS_PP002_FormItemEnabled()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
		//		////각모드에따른 아이템설정
		//		oForm.Items.Item("DocEntry").Enabled = false;
		//		oForm.Items.Item("BigCode").Enabled = true;
		//		oForm.Items.Item("Mat01").Enabled = true;
		//		PS_PP002_FormClear();
		//		////UDO방식
		//		oForm.EnableMenu("1281", true);
		//		////찾기
		//		oForm.EnableMenu("1282", false);
		//		////추가
		//	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
		//		////각모드에따른 아이템설정
		//		oForm.Items.Item("DocEntry").Enabled = true;
		//		oForm.Items.Item("BigCode").Enabled = true;
		//		oForm.Items.Item("Mat01").Enabled = false;
		//		oForm.EnableMenu("1281", false);
		//		////찾기
		//		oForm.EnableMenu("1282", true);
		//		////추가
		//	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
		//		////각모드에따른 아이템설정
		//		oForm.Items.Item("DocEntry").Enabled = false;
		//		oForm.Items.Item("BigCode").Enabled = false;
		//		oForm.Items.Item("Mat01").Enabled = true;
		//	}
		//	oForm.Freeze(false);
		//	return;
		//	PS_PP002_FormItemEnabled_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void PS_PP002_AddMatrixRow(int oRow, ref bool RowIserted = false)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	////행추가여부
		//	if (RowIserted == false) {
		//		oDS_PS_PP002L.InsertRecord((oRow));
		//	}
		//	oMat01.AddRow();
		//	oDS_PS_PP002L.Offset = oRow;
		//	oDS_PS_PP002L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
		//	oMat01.LoadFromDataSource();
		//	oForm.Freeze(false);
		//	return;
		//	PS_PP002_AddMatrixRow_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void PS_PP002_FormClear()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	string DocEntry = null;
		//	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_PP002'", ref "");
		//	if (Convert.ToDouble(DocEntry) == 0) {
		//		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oForm.Items.Item("DocEntry").Specific.Value = 1;
		//	} else {
		//		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
		//	}
		//	return;
		//	PS_PP002_FormClear_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void PS_PP002_EnableMenus()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////메뉴활성화
		//	//    Call oForm.EnableMenu("1288", True)
		//	//    Call oForm.EnableMenu("1289", True)
		//	//    Call oForm.EnableMenu("1290", True)
		//	//    Call oForm.EnableMenu("1291", True)
		//	////Call MDC_GP_EnableMenus(oForm, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
		//	MDC_Com.MDC_GP_EnableMenus(oForm, false, false, true, true, false, true, true, true, true,
		//	false, false, false, false, false, false);
		//	////메뉴설정
		//	return;
		//	PS_PP002_EnableMenus_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void PS_PP002_SetDocument(string oFormDocEntry)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if ((string.IsNullOrEmpty(oFormDocEntry))) {
		//		PS_PP002_FormItemEnabled();
		//		PS_PP002_AddMatrixRow(0, ref true);
		//		////UDO방식일때
		//	} else {
		//		//        oForm.Mode = fm_FIND_MODE
		//		//        Call PS_PP002_FormItemEnabled
		//		//        oForm.Items("DocEntry").Specific.Value = oFormDocEntry
		//		//        oForm.Items("1").Click ct_Regular
		//	}
		//	return;
		//	PS_PP002_SetDocument_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}


		//public bool PS_PP002_DataValidCheck()
		//{
		//	bool returnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	returnValue = false;
		//	object i = null;
		//	int j = 0;
		//	//UPGRADE_WARNING: oForm.Items(BigCode).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	if (oForm.Items.Item("BigCode").Specific.Selected == null) {
		//		SubMain.Sbo_Application.SetStatusBarMessage("대분류가 선택되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//		oForm.Items.Item("BigCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//		returnValue = false;
		//		return returnValue;
		//	}
		//	if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
		//		//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_PP002H] WHERE U_BigCode = ' & oForm.Items(BigCode).Specific.Selected.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP002H] WHERE U_BigCode = '" + oForm.Items.Item("BigCode").Specific.Selected.Value + "'", 0, 1) > 0) {
		//			SubMain.Sbo_Application.SetStatusBarMessage("중복된 기준분류입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//			oForm.Items.Item("BigCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//			returnValue = false;
		//			return returnValue;
		//		}
		//	}
		//	//    If oForm.Items("WhsCode").Specific.Value = "" Then
		//	//        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
		//	//        oForm.Items("WhsCode").Click ct_Regular
		//	//        PS_PP002_DataValidCheck = False
		//	//        Exit Function
		//	//    End If
		//	if (oMat01.VisualRowCount <= 1) {
		//		SubMain.Sbo_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//		returnValue = false;
		//		return returnValue;
		//	}
		//	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
		//		//UPGRADE_WARNING: oMat01.Columns(MidCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		if ((string.IsNullOrEmpty(oMat01.Columns.Item("MidCode").Cells.Item(i).Specific.Value))) {
		//			SubMain.Sbo_Application.SetStatusBarMessage("중분류코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//			oMat01.Columns.Item("MidCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//			returnValue = false;
		//			return returnValue;
		//		}
		//		//UPGRADE_WARNING: oMat01.Columns(MidName).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		if ((string.IsNullOrEmpty(oMat01.Columns.Item("MidName").Cells.Item(i).Specific.Value))) {
		//			SubMain.Sbo_Application.SetStatusBarMessage("중분류이름은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//			oMat01.Columns.Item("MidName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//			returnValue = false;
		//			return returnValue;
		//		}
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		for (j = i + 1; j <= oMat01.VisualRowCount - 1; j++) {
		//			//UPGRADE_WARNING: oMat01.Columns(MidCode).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			//UPGRADE_WARNING: oMat01.Columns(MidCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			if ((oMat01.Columns.Item("MidCode").Cells.Item(i).Specific.Value == oMat01.Columns.Item("MidCode").Cells.Item(j).Specific.Value)) {
		//				SubMain.Sbo_Application.SetStatusBarMessage("중분류코드가 중복 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//				oMat01.Columns.Item("MidCode").Cells.Item(j).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//				returnValue = false;
		//				return returnValue;
		//			}
		//		}
		//	}
		//	oDS_PS_PP002L.RemoveRecord(oDS_PS_PP002L.Size - 1);
		//	oMat01.LoadFromDataSource();
		//	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
		//		PS_PP002_FormClear();
		//	}
		//	returnValue = true;
		//	return returnValue;
		//	PS_PP002_DataValidCheck_Error:
		//	returnValue = false;
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	return returnValue;
		//}

		//private void PS_PP002_MTX01()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////메트릭스에 데이터 로드
		//	oForm.Freeze(true);
		//	int i = 0;
		//	string Query01 = null;
		//	SAPbobsCOM.Recordset RecordSet01 = null;
		//	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	string Param01 = null;
		//	string Param02 = null;
		//	string Param03 = null;
		//	string Param04 = null;
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	Param01 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	Param02 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	Param03 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	Param04 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);

		//	Query01 = "SELECT 10";
		//	RecordSet01.DoQuery(Query01);

		//	if ((RecordSet01.RecordCount == 0)) {
		//		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
		//		goto PS_PP002_MTX01_Exit;
		//	}
		//	oMat01.Clear();
		//	oMat01.FlushToDataSource();
		//	oMat01.LoadFromDataSource();

		//	SAPbouiCOM.ProgressBar ProgressBar01 = null;
		//	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

		//	for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
		//		if (i != 0) {
		//			oDS_PS_PP002L.InsertRecord((i));
		//		}
		//		oDS_PS_PP002L.Offset = i;
		//		oDS_PS_PP002L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
		//		oDS_PS_PP002L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
		//		RecordSet01.MoveNext();
		//		ProgressBar01.Value = ProgressBar01.Value + 1;
		//		ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
		//	}
		//	oMat01.LoadFromDataSource();
		//	oMat01.AutoResizeColumns();
		//	oForm.Update();

		//	ProgressBar01.Stop();
		//	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	ProgressBar01 = null;
		//	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	RecordSet01 = null;
		//	oForm.Freeze(false);
		//	return;
		//	PS_PP002_MTX01_Exit:
		//	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	RecordSet01 = null;
		//	oForm.Freeze(false);
		//	return;
		//	PS_PP002_MTX01_Error:
		//	ProgressBar01.Stop();
		//	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	ProgressBar01 = null;
		//	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	RecordSet01 = null;
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP002_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}
	}
}
