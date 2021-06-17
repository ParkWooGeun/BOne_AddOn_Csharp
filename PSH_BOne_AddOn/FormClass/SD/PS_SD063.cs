using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 수주 상세금액 조회
	/// </summary>
	internal class PS_SD063 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.DataTable oDS_PS_SD063A;
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD063.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD063_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD063");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy="DocEntry"

				oForm.Freeze(true);

				//PS_SD063_CreateItems();
				//PS_SD063_ComboBox_Setting();
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
		//				break;
		//			////Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
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
		//				break;
		//			////Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
		//			case "1281":
		//				//찾기
		//				break;
		//			////Call PS_SD063_FormItemEnabled '//UDO방식
		//			case "1282":
		//				//추가
		//				break;
		//			////Call PS_SD063_FormItemEnabled '//UDO방식
		//			case "1288":
		//			case "1289":
		//			case "1290":
		//			case "1291":
		//				//레코드이동버튼
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
		//	if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02") {
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

		//		if (pVal.ItemUID == "BtnSearch") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

		//				if (PS_SD063_DataValidCheck() == false) {

		//					BubbleEvent = false;
		//					return;

		//				} else {

		//					PS_SD063_MTX01();

		//				}

		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//			}

		//		}

		//		if (pVal.ItemUID == "BtnPrint") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

		//				if (PS_SD063_DataValidCheck() == false) {

		//					BubbleEvent = false;
		//					return;

		//				} else {

		//					PS_SD063_Print_Report01();

		//				}

		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//			}

		//		}

		//		//        If pVal.ItemUID = "1" Then
		//		//            If oForm.Mode = fm_ADD_MODE Then
		//		//                If PS_SD063_DataValidCheck = False Then
		//		//                    BubbleEvent = False
		//		//                    Exit Sub
		//		//                End If
		//		//                '//해야할일 작업
		//		//            ElseIf oForm.Mode = fm_UPDATE_MODE Then
		//		//            ElseIf oForm.Mode = fm_OK_MODE Then
		//		//            End If
		//		//        End If
		//	} else if (pVal.BeforeAction == false) {

		//		if (pVal.ItemUID == "PS_SD063") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//			}
		//		}
		//		//        If pVal.ItemUID = "1" Then
		//		//            If oForm.Mode = fm_ADD_MODE Then
		//		//                If pVal.ActionSuccess = True Then
		//		//                    Call PS_SD063_FormItemEnabled
		//		//                    Call PS_SD063_FormClear '//UDO방식일때
		//		//                    Call PS_SD063_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
		//		//                End If
		//		//            ElseIf oForm.Mode = fm_UPDATE_MODE Then
		//		//            ElseIf oForm.Mode = fm_OK_MODE Then
		//		//                If pVal.ActionSuccess = True Then
		//		//                    Call PS_SD063_FormItemEnabled
		//		//                End If
		//		//            End If
		//		//        End If
		//	}
		//	return;
		//	Raise_EVENT_ITEM_PRESSED_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.BeforeAction == true) {

		//		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
		//		//거래처
		//		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
		//		//작번

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
		//		PS_SD063_FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
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
		//		if (pVal.ItemUID == "Grid01") {
		//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
		//				if (pVal.Row > 0) {

		//				}
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
		//			}
		//		}
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
		//		if (pVal.ItemUID == "Grid01") {
		//			if (pVal.Row == -1) {
		//				//                oGrid01.Columns(pVal.ColUID).TitleObject.Sortable = True

		//			} else {
		//				if (oGrid01.Rows.SelectedRows.Count > 0) {

		//					//Call PS_SD063_GetDetail

		//					//                    Call PS_SD063_SetBaseForm '//부모폼에입력
		//					//                    If Trim(oForm.DataSources.UserDataSources("Check01").Value) = "N" Then
		//					//                        Call oForm.Close
		//					//                    End If
		//				} else {
		//					BubbleEvent = false;
		//				}
		//			}
		//		}
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

		//	} else if (pVal.BeforeAction == false) {

		//		if (pVal.ItemChanged == true) {
		//			PS_SD063_FlushToItemValue(pVal.ItemUID);
		//		}

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
		//		PS_SD063_FormItemEnabled();
		//		////Call PS_SD063_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
		//		PS_SD063_FormResize();
		//	}
		//	return;
		//	Raise_EVENT_RESIZE_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	SAPbouiCOM.DataTable oDataTable01 = null;
		//	if (pVal.BeforeAction == true) {

		//	} else if (pVal.BeforeAction == false) {
		//		//If (pVal.ItemUID = "ItemCode") Then
		//		//   Set oDataTable01 = pVal.SelectedObjects
		//		//    If oDataTable01 Is Nothing Then
		//		//    Else
		//		//  oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
		//		//     '  oForm.DataSources.UserDataSources("ItemName").Value = oDataTable01.Columns(1).Cells(0).Value
		//		//   End If
		//		// End If
		//		oForm.Update();
		//	}
		//	//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oDataTable01 = null;
		//	return;
		//	Raise_EVENT_CHOOSE_FROM_LIST_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02") {
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
		//		SubMain.RemoveForms(oFormUniqueID01);
		//		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//		oForm = null;
		//		//UPGRADE_NOTE: oGrid01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//		oGrid01 = null;
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
		//			//        For i = 1 To oMat01.VisualRowCount
		//			//            oMat01.Columns("COL01").Cells(i).Specific.Value = i
		//			//        Next i
		//			//        oMat01.FlushToDataSource
		//			//        Call oDS_PS_SD063L.RemoveRecord(oDS_PS_SD063L.Size - 1)
		//			//        oMat01.LoadFromDataSource
		//			//        If oMat01.RowCount = 0 Then
		//			//            Call PS_SD063_AddMatrixRow(0)
		//			//        Else
		//			//            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
		//			//                Call PS_SD063_AddMatrixRow(oMat01.RowCount)
		//			//            End If
		//			//        End If
		//		}
		//	}
		//	return;
		//	Raise_EVENT_ROW_DELETE_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private bool PS_SD063_CreateItems()
		//{
		//	bool functionReturnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	oForm.Freeze(true);
		//	string oQuery01 = null;
		//	//Dim C_Date   As Date
		//	//    Dim oRecordSet01 As SAPbobsCOM.Recordset
		//	//    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

		//	oGrid01 = oForm.Items.Item("Grid01").Specific;
		//	//oGrid01.SelectionMode = ms_NotSupported

		//	oForm.DataSources.DataTables.Add("PS_SD063A");

		//	oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_SD063A");

		//	oDS_PS_SD063A = oForm.DataSources.DataTables.Item("PS_SD063A");

		//	//거래처코드
		//	oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

		//	//거래처명
		//	oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

		//	//작번
		//	oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

		//	//품명
		//	oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

		//	//규격
		//	oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

		//	//수주일자(Fr)
		//	oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

		//	//수주일자(To)
		//	oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

		//	//관계사/사외
		//	oForm.DataSources.UserDataSources.Add("CardCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("CardCls").Specific.DataBind.SetBound(true, "", "CardCls");

		//	//판매완료 포함
		//	oForm.DataSources.UserDataSources.Add("SalesYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("SalesYN").Specific.DataBind.SetBound(true, "", "SalesYN");

		//	//    Set oRecordSet01 = Nothing
		//	oForm.Freeze(false);
		//	return functionReturnValue;
		//	PS_SD063_CreateItems_Error:
		//	//    Set oRecordSet01 = Nothing
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	return functionReturnValue;
		//}

		//public void PS_SD063_ComboBox_Setting()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	SAPbouiCOM.ComboBox oCombo = null;
		//	string sQry = null;
		//	SAPbobsCOM.Recordset oRecordSet01 = null;
		//	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oForm.Freeze(true);
		//	////콤보에 기본값설정

		//	//    '사업장
		//	//    Call oForm.Items("BPLId").Specific.ValidValues.Add("%", "전체")
		//	//    Call MDC_SetMod.Set_ComboList(oForm.Items("BPLId").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", "%", False, False)

		//	//관계사/사외
		//	//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("CardCls").Specific.ValidValues.Add("%", "전체");
		//	//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("CardCls").Specific.ValidValues.Add("01", "관계사");
		//	//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("CardCls").Specific.ValidValues.Add("02", "사외");
		//	//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("CardCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

		//	oForm.Freeze(false);
		//	return;
		//	PS_SD063_ComboBox_Setting_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void PS_SD063_CF_ChooseFromList()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////ChooseFromList 설정
		//	//    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
		//	//    Dim oCons               As SAPbouiCOM.Conditions
		//	//    Dim oCon                As SAPbouiCOM.Condition
		//	//    Dim oCFL                As SAPbouiCOM.ChooseFromList
		//	//    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
		//	//    Dim oEdit               As SAPbouiCOM.EditText
		//	//    Dim oColumn             As SAPbouiCOM.Column
		//	//
		//	//    Set oEdit = oForm.Items("ItemCode").Specific
		//	//    Set oCFLs = oForm.ChooseFromLists
		//	//    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
		//	//
		//	//    oCFLCreationParams.ObjectType = "4"
		//	//    oCFLCreationParams.uniqueID = "CFLITEMCD"
		//	//    oCFLCreationParams.MultiSelection = False
		//	//    Set oCFL = oCFLs.Add(oCFLCreationParams)
		//	//
		//	//'    Set oCons = oCFL.GetConditions()
		//	//'    Set oCon = oCons.Add()
		//	//'    oCon.Alias = "CardType"
		//	//'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
		//	//'    oCon.CondVal = "C"
		//	//'    oCFL.SetConditions oCons
		//	//
		//	//    oEdit.ChooseFromListUID = "CFLITEMCD"
		//	//    oEdit.ChooseFromListAlias = "ItemCode"
		//	return;
		//	PS_SD063_CF_ChooseFromList_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void PS_SD063_FormItemEnabled()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
		//		////각모드에따른 아이템설정
		//		////Call PS_SD063_FormClear '//UDO방식
		//	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
		//		////각모드에따른 아이템설정
		//	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
		//		////각모드에따른 아이템설정
		//	}
		//	oForm.Freeze(false);
		//	return;
		//	PS_SD063_FormItemEnabled_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void PS_SD063_AddMatrixRow(int oRow, ref bool RowIserted = false)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	oForm.Freeze(true);
		//	//    If RowIserted = False Then '//행추가여부
		//	//        oDS_PS_SD063L.InsertRecord (oRow)
		//	//    End If
		//	//    oMat01.AddRow
		//	//    oDS_PS_SD063L.Offset = oRow
		//	//    oDS_PS_SD063L.setValue "U_LineNum", oRow, oRow + 1
		//	//    oMat01.LoadFromDataSource
		//	oForm.Freeze(false);
		//	return;
		//	PS_SD063_AddMatrixRow_Error:
		//	oForm.Freeze(false);
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public bool PS_SD063_DataValidCheck()
		//{
		//	bool functionReturnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;

		//	functionReturnValue = true;
		//	return functionReturnValue;
		//	PS_SD063_DataValidCheck_Error:

		//	//    If oForm.Items("WorkGbn").Specific.Selected.Value = "%" Then
		//	//        Sbo_Application.SetStatusBarMessage "작업구분은 필수입니다.", bmt_Short, True
		//	//        oForm.Items("WorkGbn").Click ct_Regular
		//	//        PS_SD063_DataValidCheck = False
		//	//        Exit Function
		//	//    End If

		//	//    If oMat01.VisualRowCount = 0 Then
		//	//        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
		//	//        PS_SD063_DataValidCheck = False
		//	//        Exit Function
		//	//    End If
		//	//    For i = 1 To oMat01.VisualRowCount
		//	//        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
		//	//            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
		//	//            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
		//	//            PS_SD063_DataValidCheck = False
		//	//            Exit Function
		//	//        End If
		//	//    Next
		//	//    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
		//	//    Call oMat01.LoadFromDataSource
		//	//    Call PS_SD063_FormClear
		//	functionReturnValue = false;
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	return functionReturnValue;
		//}

		//private void PS_SD063_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	//    Dim i        As Integer
		//	//    Dim ErrNum   As Integer
		//	//    Dim sQry     As String
		//	//    Dim ItemCode As String

		//	//    Dim oRecordSet01 As SAPbobsCOM.Recordset
		//	//    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

		//	//    Dim OrdNum As String
		//	//    Dim SubNo1 As String
		//	//    Dim SubNo2 As String

		//	switch (oUID) {

		//		case "CardCode":

		//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			oForm.DataSources.UserDataSources.Item("CardName").Value = MDC_GetData.Get_ReData("CardName", "CardCode", "OCRD", "'" + Strings.Trim(oForm.Items.Item("CardCode").Specific.Value) + "'");
		//			//거래처
		//			break;

		//		case "ItemCode":

		//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			oForm.DataSources.UserDataSources.Item("ItemName").Value = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "OITM", "'" + Strings.Trim(oForm.Items.Item("ItemCode").Specific.Value) + "'");
		//			//품명
		//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			oForm.DataSources.UserDataSources.Item("ItemSpec").Value = MDC_GetData.Get_ReData("U_Size", "ItemCode", "OITM", "'" + Strings.Trim(oForm.Items.Item("ItemCode").Specific.Value) + "'");
		//			//규격
		//			break;

		//	}

		//	//    Set oRecordSet01 = Nothing

		//	return;
		//	PS_SD063_FlushToItemValue_Error:

		//	//    Set oRecordSet01 = Nothing

		//	MDC_Com.MDC_GF_Message(ref "PS_SD063_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");

		//}

		//private void PS_SD063_MTX01()
		//{
		//	//******************************************************************************
		//	//Function ID : PS_SD063_MTX01()
		//	//해당모듈    : PS_SD063
		//	//기능        : 그리드 조회
		//	//인수        : 없음
		//	//반환값      : 없음
		//	//특이사항    : 없음
		//	//******************************************************************************
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	oForm.Freeze(true);

		//	string Query01 = null;

		//	string CardCode = null;
		//	//거래처
		//	string ItemCode = null;
		//	//작번
		//	string FrDt = null;
		//	//수주일(Fr)
		//	string ToDt = null;
		//	//수주일(To)
		//	string CardCls = null;
		//	//관계사/사외
		//	string SalesYN = null;
		//	//판매완료 포함

		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
		//	//거래처
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	ItemCode = Strings.Trim(oForm.Items.Item("ItemCode").Specific.Value);
		//	//작번
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	FrDt = Strings.Trim(oForm.Items.Item("FrDt").Specific.Value);
		//	//수주일(Fr)
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	ToDt = Strings.Trim(oForm.Items.Item("ToDt").Specific.Value);
		//	//수주일(To)
		//	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	CardCls = Strings.Trim(oForm.Items.Item("CardCls").Specific.Selected.Value);
		//	//관계사/사외
		//	//UPGRADE_WARNING: oForm.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	SalesYN = (oForm.Items.Item("SalesYN").Specific.Checked == true ? "Y" : "N");
		//	//판매완료 포함

		//	SAPbouiCOM.ProgressBar ProgBar01 = null;
		//	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

		//	Query01 = "         EXEC PS_SD063_01 ";
		//	Query01 = Query01 + "'" + CardCode + "',";
		//	//거래처
		//	Query01 = Query01 + "'" + ItemCode + "',";
		//	//작번
		//	Query01 = Query01 + "'" + FrDt + "',";
		//	//수주일자(Fr)
		//	Query01 = Query01 + "'" + ToDt + "',";
		//	//수주일자(To)
		//	Query01 = Query01 + "'" + CardCls + "',";
		//	//관계사/사외
		//	Query01 = Query01 + "'" + SalesYN + "'";
		//	//판매완료 포함

		//	oGrid01.DataTable.Clear();
		//	oDS_PS_SD063A.ExecuteQuery(Query01);

		//	//    oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable")

		//	oGrid01.Columns.Item(6).RightJustified = true;
		//	oGrid01.Columns.Item(8).RightJustified = true;
		//	oGrid01.Columns.Item(9).RightJustified = true;
		//	oGrid01.Columns.Item(10).RightJustified = true;
		//	oGrid01.Columns.Item(11).RightJustified = true;
		//	oGrid01.Columns.Item(12).RightJustified = true;
		//	oGrid01.Columns.Item(13).RightJustified = true;
		//	oGrid01.Columns.Item(14).RightJustified = true;
		//	oGrid01.Columns.Item(15).RightJustified = true;
		//	oGrid01.Columns.Item(16).RightJustified = true;
		//	oGrid01.Columns.Item(17).RightJustified = true;
		//	oGrid01.Columns.Item(18).RightJustified = true;
		//	oGrid01.Columns.Item(19).RightJustified = true;
		//	oGrid01.Columns.Item(20).RightJustified = true;
		//	oGrid01.Columns.Item(21).RightJustified = true;
		//	oGrid01.Columns.Item(22).RightJustified = true;

		//	//    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
		//	//    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
		//	//    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑

		//	//    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
		//	//    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
		//	//    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
		//	//    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
		//	//    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
		//	//    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

		//	if (oGrid01.Rows.Count == 0) {
		//		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
		//		goto PS_SD063_MTX01_Exit;
		//	}

		//	oGrid01.AutoResizeColumns();
		//	oForm.Update();

		//	ProgBar01.Value = 100;
		//	ProgBar01.Stop();
		//	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	ProgBar01 = null;

		//	oForm.Freeze(false);
		//	return;
		//	PS_SD063_MTX01_Exit:
		//	oForm.Freeze(false);
		//	return;
		//	PS_SD063_MTX01_Error:

		//	oForm.Freeze(false);

		//	ProgBar01.Value = 100;
		//	ProgBar01.Stop();
		//	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	ProgBar01 = null;

		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private bool PS_SD063_DI_API()
		//{
		//	//On Error GoTo PS_SD063_DI_API_Error
		//	//    PS_SD063_DI_API = True
		//	//    Dim i, j As Long
		//	//    Dim oDIObject As SAPbobsCOM.Documents
		//	//    Dim RetVal As Long
		//	//    Dim LineNumCount As Long
		//	//    Dim ResultDocNum As Long
		//	//    If Sbo_Company.InTransaction = True Then
		//	//        Sbo_Company.EndTransaction wf_RollBack
		//	//    End If
		//	//    Sbo_Company.StartTransaction
		//	//
		//	//    ReDim ItemInformation(0)
		//	//    ItemInformationCount = 0
		//	//    For i = 1 To oMat01.VisualRowCount
		//	//        ReDim Preserve ItemInformation(ItemInformationCount)
		//	//        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.Value
		//	//        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.Value
		//	//        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.Value
		//	//        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.Value
		//	//        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.Value
		//	//        ItemInformation(ItemInformationCount).Check = False
		//	//        ItemInformationCount = ItemInformationCount + 1
		//	//    Next
		//	//
		//	//    LineNumCount = 0
		//	//    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
		//	//    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm.Items("BPLId").Specific.Selected.Value)
		//	//    oDIObject.CardCode = Trim(oForm.Items("CardCode").Specific.Value)
		//	//    oDIObject.DocDate = Format(oForm.Items("InDate").Specific.Value, "&&&&-&&-&&")
		//	//    For i = 0 To UBound(ItemInformation)
		//	//        If ItemInformation(i).Check = True Then
		//	//            GoTo Continue_First
		//	//        End If
		//	//        If i <> 0 Then
		//	//            oDIObject.Lines.Add
		//	//        End If
		//	//        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
		//	//        oDIObject.Lines.WarehouseCode = Trim(oForm.Items("WhsCode").Specific.Value)
		//	//        oDIObject.Lines.BaseType = "22"
		//	//        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
		//	//        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
		//	//        For j = i To UBound(ItemInformation)
		//	//            If ItemInformation(j).Check = True Then
		//	//                GoTo Continue_Second
		//	//            End If
		//	//            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
		//	//                GoTo Continue_Second
		//	//            End If
		//	//            '//같은것
		//	//            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
		//	//            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
		//	//            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
		//	//            oDIObject.Lines.BatchNumbers.Add
		//	//            ItemInformation(j).PDN1No = LineNumCount
		//	//            ItemInformation(j).Check = True
		//	//Continue_Second:
		//	//        Next
		//	//        LineNumCount = LineNumCount + 1
		//	//Continue_First:
		//	//    Next
		//	//    RetVal = oDIObject.Add
		//	//    If RetVal = 0 Then
		//	//        ResultDocNum = Sbo_Company.GetNewObjectKey
		//	//        For i = 0 To UBound(ItemInformation)
		//	//            Call oDS_PS_SD063L.setValue("U_OPDNNo", i, ResultDocNum)
		//	//            Call oDS_PS_SD063L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
		//	//        Next
		//	//    Else
		//	//        GoTo PS_SD063_DI_API_Error
		//	//    End If
		//	//
		//	//    If Sbo_Company.InTransaction = True Then
		//	//        Sbo_Company.EndTransaction wf_Commit
		//	//    End If
		//	//    oMat01.LoadFromDataSource
		//	//    oMat01.AutoResizeColumns
		//	//
		//	//    Set oDIObject = Nothing
		//	//    Exit Function
		//	//PS_SD063_DI_API_DI_Error:
		//	//    If Sbo_Company.InTransaction = True Then
		//	//        Sbo_Company.EndTransaction wf_RollBack
		//	//    End If
		//	//    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
		//	//    PS_SD063_DI_API = False
		//	//    Set oDIObject = Nothing
		//	//    Exit Function
		//	//PS_SD063_DI_API_Error:
		//	//    Sbo_Application.SetStatusBarMessage "PS_SD063_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
		//	//    PS_SD063_DI_API = False
		//}

		//private void PS_SD063_SetBaseForm()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	return;
		//	PS_SD063_SetBaseForm_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_SetBaseForm_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void PS_SD063_FormResize()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	//그룹박스 크기 동적 할당
		//	//    oForm.Items("GrpBox01").Height = oForm.Items("Grid01").Height + 30
		//	//    oForm.Items("GrpBox01").Width = oForm.Items("Grid01").Width + 30

		//	if (oGrid01.Columns.Count > 0) {
		//		oGrid01.AutoResizeColumns();
		//	}

		//	return;
		//	PS_SD063_FormResize_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD063_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private void PS_SD063_Print_Report01()
		//{
		//	//******************************************************************************
		//	//Function ID : PS_SD063_Print_Report01()
		//	//해당모듈    : PS_SD063
		//	//기능        : 전체 자료 출력
		//	//인수        : 없음
		//	//반환값      : 없음
		//	//특이사항    : 없음
		//	//******************************************************************************
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	string DocNum = null;
		//	string WinTitle = null;
		//	string ReportName = null;
		//	string sQry = null;

		//	short i = 0;
		//	short ErrNum = 0;
		//	string Sub_sQry = null;

		//	string CardCode = null;
		//	//거래처
		//	string ItemCode = null;
		//	//작번
		//	string FrDt = null;
		//	//수주일(Fr)
		//	string ToDt = null;
		//	//수주일(To)
		//	string CardCls = null;
		//	//관계사/사외
		//	string SalesYN = null;
		//	//판매완료 포함

		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
		//	//거래처
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	ItemCode = Strings.Trim(oForm.Items.Item("ItemCode").Specific.Value);
		//	//작번
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	FrDt = Strings.Trim(oForm.Items.Item("FrDt").Specific.Value);
		//	//수주일(Fr)
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	ToDt = Strings.Trim(oForm.Items.Item("ToDt").Specific.Value);
		//	//수주일(To)
		//	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	CardCls = Strings.Trim(oForm.Items.Item("CardCls").Specific.Selected.Value);
		//	//관계사/사외
		//	//UPGRADE_WARNING: oForm.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	SalesYN = (oForm.Items.Item("SalesYN").Specific.Checked == true ? "Y" : "N");
		//	//판매완료 포함

		//	SAPbobsCOM.Recordset oRecordSet = null;
		//	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	SAPbouiCOM.ProgressBar ProgBar01 = null;
		//	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

		//	MDC_PS_Common.ConnectODBC();

		//	/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
		//	WinTitle = "[PS_SD063] 레포트";

		//	ReportName = "PS_SD063_01.rpt";

		//	sQry = "            EXEC PS_SD063_02 ";
		//	sQry = sQry + "'" + CardCode + "',";
		//	//거래처
		//	sQry = sQry + "'" + ItemCode + "',";
		//	//작번
		//	sQry = sQry + "'" + FrDt + "',";
		//	//수주일자(Fr)
		//	sQry = sQry + "'" + ToDt + "',";
		//	//수주일자(To)
		//	sQry = sQry + "'" + CardCls + "',";
		//	//관계사/사외
		//	sQry = sQry + "'" + SalesYN + "'";
		//	MDC_Globals.gRpt_Formula = new string[3];
		//	MDC_Globals.gRpt_Formula_Value = new string[3];
		//	MDC_Globals.gRpt_SRptSqry = new string[2];
		//	MDC_Globals.gRpt_SRptName = new string[2];
		//	MDC_Globals.gRpt_SFormula = new string[2, 2];
		//	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];
		//	//판매완료 포함

		//	//// Formula 수식필드

		//	//// SubReport


		//	MDC_Globals.gRpt_SFormula[1, 1] = "";
		//	MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

		//	//    Call oRecordSet.DoQuery(sQry)
		//	//
		//	//    If oRecordSet.RecordCount = 0 Then
		//	//        ErrNum = 1
		//	//        GoTo Print_Query_Error
		//	//    End If

		//	/// Action (sub_query가 있을때는 'Y'로...)/
		//	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
		//	}

		//	ProgBar01.Value = 100;
		//	ProgBar01.Stop();
		//	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	ProgBar01 = null;

		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	return;
		//	Print_Query_Error:


		//	ProgBar01.Value = 100;
		//	ProgBar01.Stop();
		//	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	ProgBar01 = null;

		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;

		//	if (ErrNum == 1) {
		//		MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
		//	} else {
		//		MDC_Com.MDC_GF_Message(ref "PS_SD063_Print_Report01_Error:" + Err().Number + " - " + Err().Description, ref "E");
		//	}
		//}
	}
}
