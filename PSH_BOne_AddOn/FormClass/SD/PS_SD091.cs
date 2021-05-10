using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 이동요청등록
	/// </summary>
	internal class PS_SD091 : PSH_BaseClass
	{
		public string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD091H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD091L; //등록라인
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD091.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD091_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD091");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				//PS_SD091_CreateItems();
				//PS_SD091_ComboBox_Setting();
				//PS_SD091_Initial_Setting();
				//PS_SD091_CF_ChooseFromList();
				//PS_SD091_EnableMenus();
				//PS_SD091_SetDocument(oFormDocEntry);
				//PS_SD091_FormResize();
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

        #region Raise_ItemEvent
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
        #endregion

        #region Raise_MenuEvent
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
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_SD091_FormItemEnabled();
        //				////UDO방식
        //				break;
        //			case "1282":
        //				//추가
        //				PS_SD091_FormItemEnabled();
        //				////UDO방식
        //				PS_SD091_AddMatrixRow(0, ref true);
        //				////UDO방식
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_SD091_FormItemEnabled();
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
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
        #endregion

        #region Raise_RightClickEvent
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
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "PS_SD091") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_SD091_DataValidCheck() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				////해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (PS_SD091_DataValidCheck() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				////해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		} else if (pVal.ItemUID == "Btn01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				Print_StockTrans_Docu();
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "PS_SD091") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_SD091_FormItemEnabled();
        //					PS_SD091_AddMatrixRow(oMat01.RowCount, ref true);
        //					////UDO방식일때
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_SD091_FormItemEnabled();
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object TempForm01 = null;
        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01" & pVal.ColUID == "ItemCode" & pVal.CharPressed == 9) {
        //			//UPGRADE_WARNING: oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) {
        //				TempForm01 = new PS_SM020();
        //				//UPGRADE_WARNING: TempForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				TempForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
        //				PS_SD091_AddMatrixRow(0, ref true);
        //				BubbleEvent = false;
        //			}
        //		}
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OutWhCd", "");
        //		////사용자값활성
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "InWhCd", "");
        //		////사용자값활성
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemChanged == true) {
        //			oForm.Freeze(true);
        //			if ((pVal.ItemUID == "Mat01")) {
        //				//                If (pVal.ColUID = "ItemCode") Then
        //				//                    '//기타작업
        //				//                    Call oDS_PS_SD091L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value)
        //				//                    If oMat01.RowCount = pVal.Row And Trim(oDS_PS_SD091L.GetValue("U_" & pVal.ColUID, pVal.Row - 1)) <> "" Then
        //				//                        PS_SD091_AddMatrixRow (pVal.Row)
        //				//                    End If
        //				//                Else
        //				//                    Call oDS_PS_SD091L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value)
        //				//                End If
        //			} else {
        //				if ((pVal.ItemUID == "BPLId")) {
        //					//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("BPLId").Specific.Value == "1") {
        //						oDS_PS_SD091H.SetValue("U_OutWhCd", 0, "104");
        //						oDS_PS_SD091H.SetValue("U_InWhCd", 0, "101");
        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (oForm.Items.Item("BPLId").Specific.Value == "4") {
        //						oDS_PS_SD091H.SetValue("U_OutWhCd", 0, "101");
        //						oDS_PS_SD091H.SetValue("U_InWhCd", 0, "104");
        //					}
        //				}

        //				//                If (pVal.ItemUID = "CardCode") Then
        //				//                    Call oDS_PS_SD091H.setValue(pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Value)
        //				//                ElseIf (pVal.ItemUID = "CardCode") Then
        //				//                    Call oDS_PS_SD091H.setValue("U_" & pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Value)
        //				//                    Call oDS_PS_SD091H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm.Items(pVal.ItemUID).Specific.Value & "'"))
        //				//                Else
        //				//                    Call oDS_PS_SD091H.setValue("U_" & pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Value)
        //				//                End If
        //			}
        //			//            oMat01.LoadFromDataSource
        //			//            oMat01.AutoResizeColumns
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
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oMat01.SelectRow(pVal.Row, true, false);
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
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
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
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
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string sQry01 = null;
        //	string sQry02 = null;
        //	string sQry03 = null;
        //	string sQry04 = null;
        //	string sQry05 = null;
        //	string sQry06 = null;
        //	string sQry07 = null;
        //	string sQry08 = null;
        //	string sQry09 = null;
        //	string sQry10 = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	SAPbobsCOM.Recordset oRecordset02 = null;
        //	SAPbobsCOM.Recordset oRecordSet03 = null;
        //	int SumQty = 0;
        //	decimal SumWeight = default(decimal);
        //	short i = 0;

        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemChanged == true) {
        //			if ((pVal.ItemUID == "Mat01")) {
        //				// 품목 이름 Query
        //				if ((pVal.ColUID == "ItemCode")) {

        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((pVal.Row == oMat01.RowCount | oMat01.VisualRowCount == 0) & !string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))) {
        //						oMat01.FlushToDataSource();
        //						PS_SD091_AddMatrixRow(pVal.Row, ref false);
        //						oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					}

        //					oMat01.FlushToDataSource();

        //					oRecordset02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					sQry02 = "Select ItemName, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, U_Quality, U_Mark, U_CallSize, U_SbasUnit From [OITM] Where ";
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sQry02 = sQry02 + "ItemCode = '" + Strings.Trim(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value) + "'";
        //					oRecordset02.DoQuery(sQry02);
        //					oDS_PS_SD091L.SetValue("U_ItemName", pVal.Row - 1, oRecordset02.Fields.Item(0).Value);
        //					oDS_PS_SD091L.SetValue("U_Unit1", pVal.Row - 1, oRecordset02.Fields.Item(3).Value);
        //					oDS_PS_SD091L.SetValue("U_Size", pVal.Row - 1, oRecordset02.Fields.Item(4).Value);
        //					oDS_PS_SD091L.SetValue("U_CallSize", pVal.Row - 1, oRecordset02.Fields.Item(8).Value);
        //					//oMat01.Columns("ItemName").Cells(pVal.Row).Specific.Value = Trim(oRecordSet02.Fields(0).Value)
        //					//oMat01.Columns("Unit1").Cells(pVal.Row).Specific.Value = Trim(oRecordSet02.Fields(3).Value)
        //					//oMat01.Columns("Size").Cells(pVal.Row).Specific.Value = Trim(oRecordSet02.Fields(4).Value)
        //					// 품목 대분류
        //					oRecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					sQry03 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" + Strings.Trim(oRecordset02.Fields.Item(1).Value) + "'";
        //					oRecordSet03.DoQuery(sQry03);
        //					oDS_PS_SD091L.SetValue("U_ItmBsort", pVal.Row - 1, oRecordSet03.Fields.Item(1).Value);

        //					//Call oMat01.Columns("ItmBsort").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //					//UPGRADE_NOTE: oRecordSet03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet03 = null;
        //					// 품목 중분류
        //					oRecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					sQry04 = "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" + Strings.Trim(oRecordset02.Fields.Item(2).Value) + "'";
        //					oRecordSet03.DoQuery(sQry04);
        //					oDS_PS_SD091L.SetValue("U_ItmMsort", pVal.Row - 1, oRecordSet03.Fields.Item(1).Value);
        //					//Call oMat01.Columns("ItmMsort").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //					//UPGRADE_NOTE: oRecordSet03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet03 = null;
        //					// 형태타입
        //					oRecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					sQry05 = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" + Strings.Trim(oRecordset02.Fields.Item(5).Value) + "'";
        //					oRecordSet03.DoQuery(sQry05);

        //					oDS_PS_SD091L.SetValue("U_ItemType", pVal.Row - 1, oRecordSet03.Fields.Item(1).Value);
        //					//Call oMat01.Columns("ItemType").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //					//UPGRADE_NOTE: oRecordSet03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet03 = null;
        //					// 질별

        //					oRecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					sQry06 = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" + Strings.Trim(oRecordset02.Fields.Item(6).Value) + "'";
        //					oRecordSet03.DoQuery(sQry06);

        //					oDS_PS_SD091L.SetValue("U_Quality", pVal.Row - 1, oRecordSet03.Fields.Item(1).Value);
        //					//Call oMat01.Columns("Quality").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //					//UPGRADE_NOTE: oRecordSet03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet03 = null;
        //					// 인증기호
        //					oRecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					sQry07 = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" + Strings.Trim(oRecordset02.Fields.Item(7).Value) + "'";
        //					oRecordSet03.DoQuery(sQry07);

        //					oDS_PS_SD091L.SetValue("U_Mark", pVal.Row - 1, oRecordSet03.Fields.Item(1).Value);
        //					//Call oMat01.Columns("Mark").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //					//UPGRADE_NOTE: oRecordSet03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet03 = null;
        //					// 판매기준단위
        //					oRecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					sQry08 = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" + Strings.Trim(oRecordset02.Fields.Item(9).Value) + "'";
        //					oRecordSet03.DoQuery(sQry08);

        //					oDS_PS_SD091L.SetValue("U_SbasUnit", pVal.Row - 1, oRecordSet03.Fields.Item(1).Value);
        //					//Call oMat01.Columns("SbasUnit").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //					//UPGRADE_NOTE: oRecordSet03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet03 = null;

        //					//                    oDS_PS_SD091L.setValue "U_Weight", pVal.Row - 1, Val(oMat01.Columns("Weight").Cells(pVal.Row).Specific.Value)
        //					//                    oDS_PS_SD091L.setValue "U_Unweight", pVal.Row - 1, Val(oMat01.Columns("Unweight").Cells(pVal.Row).Specific.Value)
        //					//                    oMat01.FlushToDataSource
        //					//                    Call PS_SD091_AddMatrixRow(pVal.Row, False)
        //				} else if ((pVal.ColUID == "Qty")) {
        //					//UPGRADE_WARNING: oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value < 0) {
        //						oDS_PS_SD091L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(0));
        //					} else {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_SD091L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //					}
        //					//                    Call oDS_PS_SD091L.setValue("U_" & pVal.ColUID, pVal.Row - 1, Round((oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value * Val(oMat01.Columns("Unweight").Cells(pVal.Row).Specific.Value)) / 1000, 3))
        //					//                    oMat01.Columns("Weight").Cells(pVal.Row).Specific.Value = Round((Val(oMat01.Columns("Qty").Cells(pVal.Row).Specific.Value) * _
        //					//'                                                                              Val(oMat01.Columns("Unweight").Cells(pVal.Row).Specific.Value)) / 1000, 3)
        //				} else {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD091L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //				}


        //				for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //						SumQty = SumQty;
        //					} else {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //					}
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

        //				}
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumWeight").Specific.Value = SumWeight;


        //				// 단중
        //				//                If (pVal.ColUID = "Unweight") Then
        //				//                    oMat01.Columns("Weight").Cells(pVal.Row).Specific.Value = Round((Val(oMat01.Columns("Qty").Cells(pVal.Row).Specific.Value) * _
        //				//'                                                                              Val(oMat01.Columns("Unweight").Cells(pVal.Row).Specific.Value)) / 1000, 3)
        //				//                    oMat01.FlushToDataSource
        //				//'                    Call PS_SD091_AddMatrixRow(pVal.Row, False)
        //				//                End If

        //			} else {
        //				if ((pVal.ItemUID == "DocEntry")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD091H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //				} else if ((pVal.ItemUID == "CntcCode")) {
        //					oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD091H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //					//                    Call oDS_PS_SD091H.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("lastName", "firstName", "[OHEM]", "'" & oForm.Items(pVal.ItemUID).Specific.Value & "'"))
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sQry01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + Strings.Trim(oForm.Items.Item("CntcCode").Specific.Value) + "'";
        //					oRecordSet01.DoQuery(sQry01);
        //					//UPGRADE_WARNING: oForm.Items(CntcName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("CntcName").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //					//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oRecordSet01 = null;
        //				} else {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD091H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //				}
        //			}

        //			//            oMat01.FlushToDataSource
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
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	short i = 0;
        //	int SumQty = 0;
        //	decimal SumWeight = default(decimal);

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_SD091_FormItemEnabled();
        //		PS_SD091_AddMatrixRow(oMat01.VisualRowCount);
        //		////UDO방식

        //		for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //				SumQty = SumQty;
        //			} else {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //			}
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

        //		}
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pVal = null, ref bool BubbleEvent = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_SD091_FormResize();
        //	}
        //	return;
        //	Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	SAPbouiCOM.DataTable oDataTable01 = null;
        //	SAPbouiCOM.DataTable oDataTable02 = null;

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		if ((pVal.ItemUID == "CardCode")) {
        //			//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDataTable01 = pVal.SelectedObjects;
        //			if (oDataTable01 == null) {
        //			} else {
        //				oDS_PS_SD091H.SetValue("U_CardCode", 0, oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value);
        //				oDS_PS_SD091H.SetValue("U_CardName", 0, oDataTable01.Columns.Item("CardName").Cells.Item(0).Value);
        //				// // 찾기나 문서이동 버튼 클릭 시에 갱신으로 바뀌지 않음
        //				if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //					oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //			}
        //		} else if ((pVal.ItemUID == "ShipTo")) {
        //			//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDataTable02 = pVal.SelectedObjects;
        //			if (oDataTable02 == null) {
        //			} else {
        //				oDS_PS_SD091H.SetValue("U_ShipTo", 0, oDataTable02.Columns.Item("CardCode").Cells.Item(0).Value);
        //				oDS_PS_SD091H.SetValue("U_ShipNm", 0, oDataTable02.Columns.Item("CardName").Cells.Item(0).Value);
        //				if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //					oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //			}
        //		}

        //		oForm.Update();
        //	}
        //	//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDataTable01 = null;
        //	//UPGRADE_NOTE: oDataTable02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDataTable02 = null;
        //	return;
        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
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
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
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
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	int SumQty = 0;
        //	decimal SumWeight = default(decimal);

        //	if ((oLastColRow01 > 0)) {
        //		if (pVal.BeforeAction == true) {
        //			//            If (PS_SD091_Validate("행삭제") = False) Then
        //			//                BubbleEvent = False
        //			//                Exit Sub
        //			//            End If
        //			////행삭제전 행삭제가능여부검사
        //		} else if (pVal.BeforeAction == false) {
        //			for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_SD091L.RemoveRecord(oDS_PS_SD091L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0) {
        //				PS_SD091_AddMatrixRow(0);
        //			} else {
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD091L.GetValue("U_ItemCode", oMat01.RowCount - 1)))) {
        //					PS_SD091_AddMatrixRow(oMat01.RowCount);
        //				}
        //			}

        //			for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //					SumQty = SumQty;
        //				} else {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //				}
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

        //			}
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_CreateItems
        //private bool PS_SD091_CreateItems()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	string oQuery01 = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	oDS_PS_SD091H = oForm.DataSources.DBDataSources("@PS_SD091H");
        //	oDS_PS_SD091L = oForm.DataSources.DBDataSources("@PS_SD091L");
        //	oMat01 = oForm.Items.Item("Mat01").Specific;

        //	oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
        //	oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);

        //	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");
        //	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");



        //	oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //	oMat01.AutoResizeColumns();

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	oForm.Freeze(false);
        //	return functionReturnValue;
        //	PS_SD091_CreateItems_Error:
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD091_ComboBox_Setting
        //public void PS_SD091_ComboBox_Setting()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	////콤보에 기본값설정
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "Mat01", "ItemCode", "01", "완제품")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "Mat01", "ItemCode", "02", "반제품")
        //	//    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_SD091", "Mat01", "ItemCode")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "ItemCode", "", "01", "완제품")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "ItemCode", "", "02", "반제품")
        //	//    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm.Items("Item").Specific, "PS_SD091", "ItemCode")

        //	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("BPLId").Specific), ref "SELECT BPLId, BPLName FROM OBPL Where BPLId = '1' Or BPLId = '4' order by BPLId", ref "1", ref false, ref false);

        //	// 반출등록
        //	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("OutWhCd").Specific), ref "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", ref "104", ref false, ref true);
        //	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("InWhCd").Specific), ref "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", ref "101", ref false, ref true);

        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code");
        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] order by U_Code");
        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] order by Code");
        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] order by Code");
        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code, Name FROM [@PSH_MARK] order by Code");
        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("SbasUnit"), "SELECT Code, Name  FROM [@PSH_UOMORG] order by Code");

        //	//            ' 품목 대분류
        //	//                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
        //	//                    sQry03 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet02.Fields(1).Value) & "'"
        //	//                    oRecordSet03.DoQuery sQry03
        //	//                    Call oMat01.Columns("ItmBsort").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //	//                    Set oRecordSet03 = Nothing
        //	//            ' 품목 중분류
        //	//                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
        //	//                    sQry04 = "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" & Trim(oRecordSet02.Fields(2).Value) & "'"
        //	//                    oRecordSet03.DoQuery sQry04
        //	//                    Call oMat01.Columns("ItmMsort").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //	//                    Set oRecordSet03 = Nothing
        //	//            ' 형태타입
        //	//                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
        //	//                    sQry05 = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" & Trim(oRecordSet02.Fields(5).Value) & "'"
        //	//                    oRecordSet03.DoQuery sQry05
        //	//                    Call oMat01.Columns("ItemType").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //	//                    Set oRecordSet03 = Nothing
        //	//            ' 질별
        //	//                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
        //	//                    sQry06 = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" & Trim(oRecordSet02.Fields(6).Value) & "'"
        //	//                    oRecordSet03.DoQuery sQry06
        //	//                    Call oMat01.Columns("Quality").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //	//                    Set oRecordSet03 = Nothing
        //	//            ' 인증기호
        //	//                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
        //	//                    sQry07 = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" & Trim(oRecordSet02.Fields(7).Value) & "'"
        //	//                    oRecordSet03.DoQuery sQry07
        //	//                    Call oMat01.Columns("Mark").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //	//                    Set oRecordSet03 = Nothing
        //	//            ' 판매기준단위
        //	//                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
        //	//                    sQry08 = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" & Trim(oRecordSet02.Fields(9).Value) & "'"
        //	//                    oRecordSet03.DoQuery sQry08
        //	//                    Call oMat01.Columns("SbasUnit").Cells(pVal.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
        //	//                    Set oRecordSet03 = Nothing


        //	oForm.Freeze(false);
        //	return;
        //	PS_SD091_ComboBox_Setting_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_Initial_Setting
        //public void PS_SD091_Initial_Setting()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	// 사업장
        //	string lcl_User_BPLId = null;
        //	lcl_User_BPLId = MDC_PS_Common.User_BPLId();

        //	//소속사업장이 창원이나 구로일 때만 콤보박스 Select (2011.09.21 송명규 추가)
        //	if (lcl_User_BPLId == "1" | lcl_User_BPLId == "4") {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);
        //	}
        //	// 인수자
        //	//UPGRADE_WARNING: oForm.Items(CntcCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD();
        //	return;
        //	PS_SD091_Initial_Setting_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_Initial_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_CF_ChooseFromList
        //public void PS_SD091_CF_ChooseFromList()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbouiCOM.ChooseFromList oCFL01 = null;
        //	SAPbouiCOM.ChooseFromList oCFL02 = null;
        //	SAPbouiCOM.Conditions oCons = null;
        //	SAPbouiCOM.Condition oCon = null;
        //	SAPbouiCOM.ChooseFromListCollection oCFLs01 = null;
        //	SAPbouiCOM.ChooseFromListCollection oCFLs02 = null;
        //	SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams01 = null;
        //	SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams02 = null;
        //	SAPbouiCOM.EditText oEdit01 = null;
        //	SAPbouiCOM.EditText oEdit02 = null;

        //	//    Set oEdit01 = oForm.Items("CntcCode").Specific
        //	//    Set oCFLs01 = oForm.ChooseFromLists
        //	//    Set oCFLCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
        //	//
        //	//    oCFLCreationParams01.ObjectType = "171"
        //	//    oCFLCreationParams01.uniqueID = "CFLCNTCCODE"
        //	//    oCFLCreationParams01.MultiSelection = False
        //	//    Set oCFL01 = oCFLs01.Add(oCFLCreationParams01)
        //	//
        //	//    oEdit01.ChooseFromListUID = "CFLCNTCCODE"
        //	//    oEdit01.ChooseFromListAlias = "CntcCode"
        //	//
        //	//    Set oEdit01 = Nothing
        //	//    Set oCFLs01 = Nothing
        //	//    Set oCFLCreationParams01 = Nothing

        //	oEdit01 = oForm.Items.Item("CardCode").Specific;
        //	oCFLs01 = oForm.ChooseFromLists;
        //	oCFLCreationParams01 = SubMain.Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

        //	oCFLCreationParams01.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_BusinessPartner);
        //	oCFLCreationParams01.UniqueID = "CFLCARDCODE";
        //	oCFLCreationParams01.MultiSelection = false;
        //	oCFL01 = oCFLs01.Add(oCFLCreationParams01);

        //	oEdit01.ChooseFromListUID = "CFLCARDCODE";
        //	oEdit01.ChooseFromListAlias = "CardCode";

        //	// Choose from list 에 조건을 줄 경우
        //	// choosefromlist가 화면에 나오면 서식세팅으로 원하는 필드값 추가 가능
        //	oCons = oCFL01.GetConditions();
        //	oCon = oCons.Add();
        //	oCon.Alias = "CardType";
        //	// Condition Field
        //	oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
        //	// Equal
        //	oCon.CondVal = "C";
        //	// Condition Value
        //	oCFL01.SetConditions(oCons);

        //	oEdit02 = oForm.Items.Item("ShipTo").Specific;
        //	oCFLs02 = oForm.ChooseFromLists;
        //	oCFLCreationParams02 = SubMain.Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

        //	oCFLCreationParams02.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_BusinessPartner);
        //	oCFLCreationParams02.UniqueID = "CFLSHIPCODE";
        //	oCFLCreationParams02.MultiSelection = false;
        //	oCFL02 = oCFLs02.Add(oCFLCreationParams02);

        //	oEdit02.ChooseFromListUID = "CFLSHIPCODE";
        //	oEdit02.ChooseFromListAlias = "CardCode";
        //	// 정의된 오브젝트만 사용 가능
        //	return;
        //	PS_SD091_CF_ChooseFromList_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_FormItemEnabled
        //public void PS_SD091_FormItemEnabled()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		////각모드에따른 아이템설정
        //		PS_SD091_FormClear();
        //		////UDO방식
        //		oForm.EnableMenu("1281", true);
        //		////찾기
        //		oForm.EnableMenu("1282", false);
        //		////추가
        //		oForm.EnableMenu("1293", true);
        //		oForm.Items.Item("DocEntry").Enabled = false;
        //		oForm.Items.Item("Btn01").Enabled = false;
        //	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
        //		////각모드에따른 아이템설정
        //		oForm.Items.Item("DocEntry").Enabled = true;
        //		oForm.EnableMenu("1281", false);
        //		////찾기
        //		oForm.EnableMenu("1282", true);
        //		////추가
        //		oForm.EnableMenu("1293", true);
        //		oForm.Items.Item("Btn01").Enabled = false;
        //	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
        //		oForm.Items.Item("DocEntry").Enabled = false;
        //		oForm.EnableMenu("1281", true);
        //		oForm.EnableMenu("1282", true);
        //		oForm.EnableMenu("1293", false);
        //		oForm.Items.Item("Btn01").Enabled = true;
        //		////각모드에따른 아이템설정
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	PS_SD091_FormItemEnabled_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_AddMatrixRow
        //public void PS_SD091_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	////행추가여부
        //	if (RowIserted == false) {
        //		oDS_PS_SD091L.InsertRecord((oRow));
        //	}
        //	oMat01.AddRow();
        //	oDS_PS_SD091L.Offset = oRow;
        //	oDS_PS_SD091L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oMat01.LoadFromDataSource();
        //	oForm.Freeze(false);
        //	return;
        //	PS_SD091_AddMatrixRow_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_FormClear
        //public void PS_SD091_FormClear()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocEntry = null;
        //	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_SD091'", ref "");
        //	if (Convert.ToDouble(DocEntry) == 0) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.Value = 1;
        //	} else {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
        //	}
        //	return;
        //	PS_SD091_FormClear_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_EnableMenus
        //private void PS_SD091_EnableMenus()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////메뉴활성화
        //	//    Call oForm.EnableMenu("1288", True)
        //	//    Call oForm.EnableMenu("1289", True)
        //	//    Call oForm.EnableMenu("1290", True)
        //	//    Call oForm.EnableMenu("1291", True)
        //	oForm.EnableMenu("1293", true);
        //	//// 행삭제
        //	////Call MDC_GP_EnableMenus(oForm, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
        //	////Call MDC_GP_EnableMenus(oForm, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
        //	return;
        //	PS_SD091_EnableMenus_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_SetDocument
        //private void PS_SD091_SetDocument(string oFormDocEntry)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((string.IsNullOrEmpty(oFormDocEntry))) {
        //		PS_SD091_FormItemEnabled();
        //		PS_SD091_AddMatrixRow(0, ref true);
        //		////UDO방식일때
        //	} else {
        //		//        oForm.Mode = fm_FIND_MODE
        //		//        Call PS_SD091_FormItemEnabled
        //		//        oForm.Items("DocEntry").Specific.Value = oFormDocEntry
        //		//        oForm.Items("1").Click ct_Regular
        //	}
        //	return;
        //	PS_SD091_SetDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_DataValidCheck
        //public bool PS_SD091_DataValidCheck()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = false;
        //	int i = 0;
        //	//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //		//UPGRADE_WARNING: oForm.Items(CntcCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	} else if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("요청자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //		//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	} else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("요청일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //		//UPGRADE_WARNING: oForm.Items(OutWhCd).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	} else if (string.IsNullOrEmpty(oForm.Items.Item("OutWhCd").Specific.Value)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("출고창고는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm.Items.Item("OutWhCd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //		//UPGRADE_WARNING: oForm.Items(InWhCd).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	} else if (string.IsNullOrEmpty(oForm.Items.Item("InWhCd").Specific.Value)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("입고창고는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm.Items.Item("InWhCd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	if (oMat01.VisualRowCount == 0) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		//UPGRADE_WARNING: oMat01.Columns(ItemName).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value))) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("품목은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("ItemName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value) <= 0)) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("수량은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("ItemName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}

        //	}
        //	oDS_PS_SD091L.RemoveRecord(oDS_PS_SD091L.Size - 1);
        //	oMat01.LoadFromDataSource();
        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		PS_SD091_FormClear();
        //	}
        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	PS_SD091_DataValidCheck_Error:
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD091_MTX01
        //private void PS_SD091_MTX01()
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

        //	oMat01.Clear();
        //	oMat01.FlushToDataSource();
        //	oMat01.LoadFromDataSource();

        //	if ((RecordSet01.RecordCount == 0)) {
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_SD091_MTX01_Exit;
        //	}

        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //	for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //		if (i != 0) {
        //			oDS_PS_SD091L.InsertRecord((i));
        //		}
        //		oDS_PS_SD091L.Offset = i;
        //		oDS_PS_SD091L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
        //		oDS_PS_SD091L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
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
        //	PS_SD091_MTX01_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	if ((ProgressBar01 != null)) {
        //		ProgressBar01.Stop();
        //	}
        //	return;
        //	PS_SD091_MTX01_Error:
        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_FormResize
        //private void PS_SD091_FormResize()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	return;
        //	PS_SD091_FormResize_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD091_Validate
        //public bool PS_SD091_Validate(string ValidateType)
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	if (ValidateType == "수정") {
        //		//        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
        //		//        Dim Exist As Boolean
        //		//        Exist = False
        //		//        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm.Items("8").Specific.Value & "'"
        //		//        RecordSet01.DoQuery Query01
        //		//        For i = 0 To RecordSet01.RecordCount - 1
        //		//            Exist = False
        //		//            For j = 1 To oMat01.RowCount - 1
        //		//                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
        //		//                If Val(RecordSet01.Fields(1).Value) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.Value) And RecordSet01.Fields(2).Value = oMat01.Columns("1").Cells(j).Specific.Value And oMat01.Columns("U_LineNum").Cells(j).Specific.Value <> "" Then
        //		//                    Exist = True
        //		//                End If
        //		//            Next
        //		//            If (Exist = False) Then '//삭제된 행중
        //		//                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
        //		//                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할수 없습니다.", "W"
        //		//                    PS_SD091_Validate = False
        //		//                    GoTo PS_SD091_Validate_Exit
        //		//                End If
        //		//            End If
        //		//            RecordSet01.MoveNext
        //		//        Next
        //	} else if (ValidateType == "행삭제") {
        //		////행삭제전 행삭제가능여부검사
        //		//        If oForm.Mode = fm_OK_MODE Or oForm.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
        //		//            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
        //		//            Else
        //		//                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
        //		//                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
        //		//                    PS_SD091_Validate = False
        //		//                    GoTo PS_SD091_Validate_Exit
        //		//                End If
        //		//            End If
        //		//        End If
        //	} else if (ValidateType == "취소") {
        //		//        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm.Items("8").Specific.Value & "'"
        //		//        RecordSet01.DoQuery Query01
        //		//        For i = 0 To RecordSet01.RecordCount - 1
        //		//            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
        //		//                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
        //		//                PS_SD091_Validate = False
        //		//                GoTo PS_SD091_Validate_Exit
        //		//            End If
        //		//            RecordSet01.MoveNext
        //		//        Next
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD091_Validate_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD091_Validate_Error:
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD091_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region Print_StockTrans_Docu
        //private void Print_StockTrans_Docu()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;

        //	MDC_PS_Common.ConnectODBC();
        //	WinTitle = "[PS_SD045] 이동 요청서";
        //	ReportName = "PS_SD045.rpt";
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry = "EXEC PS_SD045_01 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];


        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}
        //	return;
        //	Print_StockTrans_Docu_Error:


        //}
        #endregion
    }
}
