//using System;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;
//using PSH_BOne_AddOn.Form;
//using PSH_BOne_AddOn.DataPack;
//using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 재고장
    /// </summary>
    internal class PS_SM020 : PSH_BaseClass
    {
        //		public string oFormUniqueID01;
        //		public SAPbouiCOM.Form oForm01;
        //		public SAPbouiCOM.Matrix oMat01;
        //		public SAPbouiCOM.Matrix oMat02;
        //			//등록헤더
        //		private SAPbouiCOM.DBDataSource oDS_PS_SM020H;
        //			//등록라인
        //		private SAPbouiCOM.DBDataSource oDS_PS_SM020L;

        //			////부모폼
        //		public SAPbouiCOM.Form oBaseForm01;
        //		public string oBaseItemUID01;
        //		public string oBaseColUID01;
        //		public int oBaseColRow01;
        //		public string oBaseTradeType01;

        //			//클래스에서 선택한 마지막 아이템 Uid값
        //		private string oLastItemUID01;
        //			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        //		private string oLastColUID01;
        //			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        //		private int oLastColRow01;

        //		private int oMat01Row01;
        //		private int oMat02Row02;

        public void LoadForm(SAPbouiCOM.Form oForm, string itemUID, string colUID, int colRow)
        {
        }

        //		private void LoadForm(ref SAPbouiCOM.Form oForm02 = null, string oItemUID02 = "", string oColUID02 = "", int oColRow02 = 0, string oTradeType02 = "")
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			string oInnerXml01 = null;
        //			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

        //			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_SM020.srf");
        //			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
        //			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
        //			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

        //			//매트릭스의 타이틀높이와 셀높이를 고정
        //			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
        //				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
        //				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
        //			}

        //			oFormUniqueID01 = "PS_SM020_" + GetTotalFormsCount();
        //			SubMain.AddForms(this, oFormUniqueID01);
        //			////폼추가
        //			SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));
        //			//폼 할당
        //			oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

        //			oForm01.SupportedModes = -1;
        //			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //			////oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

        //			oForm01.Freeze(true);
        //			oBaseForm01 = oForm02;
        //			oBaseItemUID01 = oItemUID02;
        //			oBaseColUID01 = oColUID02;
        //			oBaseColRow01 = oColRow02;
        //			oBaseTradeType01 = oTradeType02;

        //			PS_SM020_CreateItems();
        //			PS_SM020_ComboBox_Setting();
        //			PS_SM020_CF_ChooseFromList();
        //			PS_SM020_FormItemEnabled();
        //			PS_SM020_EnableMenus();
        //			////Call PS_SM020_FormClear '//UDO방식일때
        //			////Call PS_SM020_AddMatrixRow(0, True) '//UDO방식일때

        //			oForm01.Update();
        //			oForm01.Freeze(false);

        //			oForm01.Visible = true;
        //			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oXmlDoc01 = null;
        //			return;
        //			LoadForm_Error:
        //			oForm01.Update();
        //			oForm01.Freeze(false);
        //			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oXmlDoc01 = null;
        //			//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oForm01 = null;
        //			SubMain.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			switch (pval.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1
        //					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //					////5
        //					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
        //					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //					////8
        //					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //					////10
        //					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
        //					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //					////18
        //					break;
        //				////et_FORM_ACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //					////19
        //					break;
        //				////et_FORM_DEACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //					////20
        //					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //					////27
        //					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3
        //					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //					////4
        //					break;
        //				////et_LOST_FOCUS
        //				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //					////17
        //					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //			}
        //			return;
        //			Raise_ItemEvent_Error:
        //			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;
        //					////Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;
        //					////Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
        //					case "1281":
        //						//찾기
        //						break;
        //					////Call PS_SM020_FormItemEnabled '//UDO방식
        //					case "1282":
        //						//추가
        //						break;
        //					////Call PS_SM020_FormItemEnabled '//UDO방식
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_MenuEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //				//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //				//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //				//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //				//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //				//            MenuCreationParams01.uniqueID = "MenuUID"
        //				//            MenuCreationParams01.String = "메뉴명"
        //				//            MenuCreationParams01.Enabled = True
        //				//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //				//        End If
        //			} else if (pval.BeforeAction == false) {
        //				//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //				//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //				//        End If
        //			}
        //			if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02") {
        //				if (pval.Row > 0) {
        //					oLastItemUID01 = pval.ItemUID;
        //					oLastColUID01 = pval.ColUID;
        //					oLastColRow01 = pval.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pval.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}
        //			return;
        //			Raise_RightClickEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //				if (pval.ItemUID == "Button01") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //						PS_SM020_MTX01();
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //				if (pval.ItemUID == "Button02") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //						PS_SM020_SetBaseForm();
        //						////부모폼에입력
        //						oForm01.Close();
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //				//        If pval.ItemUID = "1" Then
        //				//            If oForm01.Mode = fm_ADD_MODE Then
        //				//                If PS_SM020_DataValidCheck = False Then
        //				//                    BubbleEvent = False
        //				//                    Exit Sub
        //				//                End If
        //				//                '//해야할일 작업
        //				//            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        //				//            ElseIf oForm01.Mode = fm_OK_MODE Then
        //				//            End If
        //				//        End If
        //			} else if (pval.BeforeAction == false) {
        //				if (pval.ItemUID == "PS_SM020") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //				//        If pval.ItemUID = "1" Then
        //				//            If oForm01.Mode = fm_ADD_MODE Then
        //				//                If pval.ActionSuccess = True Then
        //				//                    Call PS_SM020_FormItemEnabled
        //				//                    Call PS_SM020_FormClear '//UDO방식일때
        //				//                    Call PS_SM020_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
        //				//                End If
        //				//            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        //				//            ElseIf oForm01.Mode = fm_OK_MODE Then
        //				//                If pval.ActionSuccess = True Then
        //				//                    Call PS_SM020_FormItemEnabled
        //				//                End If
        //				//            End If
        //				//        End If
        //			}
        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //				//        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        //				//        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "ItemCode", "");
        //				////사용자값활성
        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "ItemName", "");
        //				////사용자값활성
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm01.Freeze(true);
        //			int i = 0;
        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				if (pval.ItemUID == "ItmBsort") {
        //					//UPGRADE_WARNING: oForm01.Items(ItmMsort).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					for (i = 0; i <= oForm01.Items.Item("ItmMsort").Specific.ValidValues.Count - 1; i++) {
        //						//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm01.Items.Item("ItmMsort").Specific.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //					}
        //					//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm01.Items.Item("ItmMsort").Specific.ValidValues.Add("선택", "선택");
        //					//UPGRADE_WARNING: oForm01.Items(ItmBsort).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItmMsort").Specific), ref "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm01.Items.Item("ItmBsort").Specific.Selected.VALUE + "' ORDER BY U_Code", ref "", ref false, ref false);
        //					//UPGRADE_WARNING: oForm01.Items(ItmMsort).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm01.Items.Item("ItmMsort").Specific.ValidValues.Count > 0) {
        //						//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm01.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //					}
        //				}
        //			}
        //			oForm01.Freeze(false);
        //			return;
        //			Raise_EVENT_COMBO_SELECT_Error:
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //				if (pval.ItemUID == "Mat01") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //						if (pval.Row > 0) {
        //							oMat01.SelectRow(pval.Row, true, false);
        //							oMat01Row01 = pval.Row;
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							////배치를 사용하는품목
        //							if (MDC_PS_Common.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE) == "Y") {
        //								PS_SM020_MTX02();
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							////배치를 사용하지 않는품목
        //							} else if (MDC_PS_Common.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE) == "N") {
        //								PS_SM020_MTX02();
        //							}
        //						}
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //				if (pval.ItemUID == "Mat02") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //						if (pval.Row > 0) {
        //							oMat02.SelectRow(pval.Row, true, false);
        //							oMat02Row02 = pval.Row;
        //						}
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //				if (pval.ItemUID == "Opt01") {
        //					oForm01.Settings.MatrixUID = "Mat01";
        //					oForm01.Settings.Enabled = true;
        //					oForm01.Settings.EnableRowFormat = true;
        //				}
        //				if (pval.ItemUID == "Opt02") {
        //					oForm01.Settings.MatrixUID = "Mat02";
        //					oForm01.Settings.Enabled = true;
        //					oForm01.Settings.EnableRowFormat = true;
        //				}
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_CLICK_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //				if (pval.ItemUID == "Mat01") {
        //					if (pval.Row == 0) {
        //						oMat01.Columns.Item(pval.ColUID).TitleObject.Sortable = true;
        //						oMat01.FlushToDataSource();
        //					}
        //				}
        //				if (pval.ItemUID == "Mat02") {
        //					if (pval.Row == 0) {
        //						oMat02.Columns.Item(pval.ColUID).TitleObject.Sortable = true;
        //						oMat02.FlushToDataSource();
        //					}
        //				}
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_DOUBLE_CLICK_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm01.Freeze(true);
        //			string ItemCode01 = null;
        //			if (pval.BeforeAction == true) {
        //				if (pval.ItemChanged == true) {
        //					if (pval.ItemUID == "Mat01") {
        //						if (pval.ColUID == "SelQty") {
        //							//UPGRADE_WARNING: oMat01.Columns(SelQty).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE <= 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE = 0;
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = 0;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE;
        //								////EA자체품
        //								if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //									//UPGRADE_WARNING: oMat01.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE);
        //								////EAUOM
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //									//UPGRADE_WARNING: oMat01.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_Unit1(ItemCode01));
        //								////KGSPEC
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //									//UPGRADE_WARNING: oMat01.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = (Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE);
        //								////KG단중
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //									//UPGRADE_WARNING: oMat01.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = System.Math.Round(Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0);
        //								////KG입력
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //								}
        //							}
        //							oForm01.Update();
        //						} else if (pval.ColUID == "SelWeight") {
        //							//UPGRADE_WARNING: oMat01.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE <= 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE = 0;
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = 0;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE;
        //								////EA자체품
        //								if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //								////EAUOM
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //								////KGSPEC
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //									//UPGRADE_WARNING: oMat01.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = (Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE);
        //								////KG단중
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //									//UPGRADE_WARNING: oMat01.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = System.Math.Round(Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0);
        //								////KG입력
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //								}
        //							}
        //							oForm01.Update();
        //						}
        //					} else if (pval.ItemUID == "Mat02") {
        //						if (pval.ColUID == "SelQty") {
        //							//                    If (MDC_PS_Common.GetItem_ManBtchNum(ItemCode01) = "Y") Then '//배치를 사용하는 품목의경우
        //							//UPGRADE_WARNING: oMat02.Columns(SelQty).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE <= 0) {
        //								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE = 0;
        //								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = 0;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE;
        //								////EA자체품
        //								if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //									//UPGRADE_WARNING: oMat02.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = Conversion.Val(oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE);
        //								////EAUOM
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //									//UPGRADE_WARNING: oMat02.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = Conversion.Val(oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_Unit1(ItemCode01));
        //								////KGSPEC
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //									//UPGRADE_WARNING: oMat02.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = (Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE);
        //								////KG단중
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //									//UPGRADE_WARNING: oMat02.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = System.Math.Round(Conversion.Val(oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0);
        //								////KG입력
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //								}
        //							}
        //							oForm01.Update();
        //						} else if (pval.ColUID == "SelWeight") {
        //							//UPGRADE_WARNING: oMat02.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE <= 0) {
        //								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE = 0;
        //								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = 0;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE;
        //								////EA자체품
        //								if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //								////EAUOM
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //								////KGSPEC
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //									//UPGRADE_WARNING: oMat02.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = (Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE);
        //								////KG단중
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //									//UPGRADE_WARNING: oMat02.Columns(SelWeight).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat02.Columns.Item("SelWeight").Cells.Item(pval.Row).Specific.VALUE = System.Math.Round(Conversion.Val(oMat02.Columns.Item("SelQty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0);
        //								////KG입력
        //								} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //								}
        //							}
        //							oForm01.Update();
        //						}
        //					}
        //				}
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			oForm01.Freeze(false);
        //			return;
        //			Raise_EVENT_VALIDATE_Error:
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				PS_SM020_FormItemEnabled();
        //				////Call PS_SM020_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LOAD_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_RESIZE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				oForm01.Items.Item("Mat01").Top = 70;
        //				oForm01.Items.Item("Mat01").Height = (oForm01.Height / 2) - 70;
        //				oForm01.Items.Item("Mat01").Left = 7;
        //				oForm01.Items.Item("Mat01").Width = oForm01.Width - 21;
        //				oForm01.Items.Item("Mat02").Top = (oForm01.Height / 2) + 10;
        //				oForm01.Items.Item("Mat02").Height = (oForm01.Height / 2) - 75;
        //				oForm01.Items.Item("Mat02").Left = 7;
        //				oForm01.Items.Item("Mat02").Width = oForm01.Width - 21;
        //			}
        //			return;
        //			Raise_EVENT_RESIZE_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			SAPbouiCOM.DataTable oDataTable01 = null;
        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				//        If (pval.ItemUID = "ItemCode") Then
        //				//            Set oDataTable01 = pval.SelectedObjects
        //				//            If oDataTable01 Is Nothing Then
        //				//            Else
        //				//                oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
        //				//                oForm01.DataSources.UserDataSources("ItemName").Value = oDataTable01.Columns(1).Cells(0).Value
        //				//            End If
        //				//        End If
        //				//        oForm01.Update
        //			}
        //			//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oDataTable01 = null;
        //			return;
        //			Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02") {
        //				if (pval.Row > 0) {
        //					oLastItemUID01 = pval.ItemUID;
        //					oLastColUID01 = pval.ColUID;
        //					oLastColRow01 = pval.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pval.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}

        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //				SubMain.RemoveForms(oFormUniqueID01);
        //				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm01 = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //			}
        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			if ((oLastColRow01 > 0)) {
        //				if (pval.BeforeAction == true) {
        //					////행삭제전 행삭제가능여부검사
        //				} else if (pval.BeforeAction == false) {
        //					//        For i = 1 To oMat01.VisualRowCount
        //					//            oMat01.Columns("COL01").Cells(i).Specific.Value = i
        //					//        Next i
        //					//        oMat01.FlushToDataSource
        //					//        Call oDS_PS_SM020L.RemoveRecord(oDS_PS_SM020L.Size - 1)
        //					//        oMat01.LoadFromDataSource
        //					//        If oMat01.RowCount = 0 Then
        //					//            Call PS_SM020_AddMatrixRow(0)
        //					//        Else
        //					//            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
        //					//                Call PS_SM020_AddMatrixRow(oMat01.RowCount)
        //					//            End If
        //					//        End If
        //				}
        //			}
        //			return;
        //			Raise_EVENT_ROW_DELETE_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private bool PS_SM020_CreateItems()
        //		{
        //			bool functionReturnValue = false;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm01.Freeze(true);
        //			string oQuery01 = null;
        //			SAPbobsCOM.Recordset oRecordSet01 = null;
        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			oDS_PS_SM020H = oForm01.DataSources.DBDataSources("@PS_USERDS01");
        //			oDS_PS_SM020L = oForm01.DataSources.DBDataSources("@PS_USERDS02");
        //			oMat01 = oForm01.Items.Item("Mat01").Specific;
        //			oMat02 = oForm01.Items.Item("Mat02").Specific;
        //			oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //			oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //			oMat01.AutoResizeColumns();
        //			oMat02.AutoResizeColumns();

        //			oForm01.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");
        //			oForm01.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

        //			oForm01.DataSources.UserDataSources.Add("StockType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("StockType").Specific.DataBind.SetBound(true, "", "StockType");
        //			oForm01.DataSources.UserDataSources.Add("TradeType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("TradeType").Specific.DataBind.SetBound(true, "", "TradeType");

        //			oForm01.DataSources.UserDataSources.Add("ItemGpCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemGpCd").Specific.DataBind.SetBound(true, "", "ItemGpCd");
        //			oForm01.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");
        //			oForm01.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");
        //			oForm01.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("Size").Specific.DataBind.SetBound(true, "", "Size");
        //			oForm01.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");
        //			oForm01.DataSources.UserDataSources.Add("Mark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("Mark").Specific.DataBind.SetBound(true, "", "Mark");

        //			oForm01.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");
        //			oForm01.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");
        //			//UPGRADE_WARNING: oForm01.Items().Specific.GroupWith 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("Opt01").Specific.GroupWith("Opt02");

        //			oForm01.Items.Item("Mat01").Enabled = false;
        //			oForm01.Items.Item("Mat02").Enabled = false;

        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;
        //			oForm01.Freeze(false);
        //			return functionReturnValue;
        //			PS_SM020_CreateItems_Error:
        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			return functionReturnValue;
        //		}

        //		private void PS_SM020_ComboBox_Setting()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm01.Freeze(true);
        //			////콤보에 기본값설정
        //			//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SM020", "Mat01", "ItemCode", "01", "완제품")
        //			//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SM020", "Mat01", "ItemCode", "02", "반제품")
        //			//    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_SM020", "Mat01", "ItemCode")
        //			MDC_PS_Common.Combo_ValidValues_Insert("PS_SM020", "StockType", "", "1", "재고있는품목");
        //			MDC_PS_Common.Combo_ValidValues_Insert("PS_SM020", "StockType", "", "2", "전체");
        //			MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm01.Items.Item("StockType").Specific), "PS_SM020", "StockType");
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("StockType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

        //			MDC_PS_Common.Combo_ValidValues_Insert("PS_SM020", "TradeType", "", "", "전체");
        //			MDC_PS_Common.Combo_ValidValues_Insert("PS_SM020", "TradeType", "", "1", "일반");
        //			MDC_PS_Common.Combo_ValidValues_Insert("PS_SM020", "TradeType", "", "2", "임가공");
        //			MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm01.Items.Item("TradeType").Specific), "PS_SM020", "TradeType");
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("TradeType").Specific.Select(oBaseTradeType01, SAPbouiCOM.BoSearchKey.psk_ByValue);

        //			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItmBsort").Specific.ValidValues.Add("선택", "선택");
        //			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItmBsort").Specific), ref "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", ref "", ref false, ref false);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItmMsort").Specific.ValidValues.Add("선택", "선택");
        //			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItmMsort").Specific), ref "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", ref "", ref false, ref false);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemType").Specific.ValidValues.Add("선택", "선택");
        //			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItemType").Specific), ref "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", ref "", ref false, ref false);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("Mark").Specific.ValidValues.Add("선택", "선택");
        //			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("Mark").Specific), ref "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", ref "", ref false, ref false);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemGpCd").Specific.ValidValues.Add("선택", "선택");
        //			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItemGpCd").Specific), ref "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", ref "", ref false, ref false);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("Mark").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm01.Items.Item("ItemGpCd").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

        //			oForm01.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm01.Items.Item("TradeType").Enabled = false;

        //			//    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
        //			oForm01.Freeze(false);
        //			return;
        //			PS_SM020_ComboBox_Setting_Error:
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void PS_SM020_CF_ChooseFromList()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////ChooseFromList 설정
        //			//    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
        //			//    Dim oCons               As SAPbouiCOM.Conditions
        //			//    Dim oCon                As SAPbouiCOM.Condition
        //			//    Dim oCFL                As SAPbouiCOM.ChooseFromList
        //			//    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
        //			//    Dim oEdit               As SAPbouiCOM.EditText
        //			//    Dim oColumn             As SAPbouiCOM.Column
        //			//
        //			//    Set oEdit = oForm01.Items("ItemCode").Specific
        //			//    Set oCFLs = oForm01.ChooseFromLists
        //			//    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
        //			//
        //			//    oCFLCreationParams.ObjectType = "4"
        //			//    oCFLCreationParams.uniqueID = "CFLITEMCD"
        //			//    oCFLCreationParams.MultiSelection = False
        //			//    Set oCFL = oCFLs.Add(oCFLCreationParams)
        //			//
        //			//'    Set oCons = oCFL.GetConditions()
        //			//'    Set oCon = oCons.Add()
        //			//'    oCon.Alias = "CardType"
        //			//'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        //			//'    oCon.CondVal = "C"
        //			//'    oCFL.SetConditions oCons
        //			//
        //			//    oEdit.ChooseFromListUID = "CFLITEMCD"
        //			//    oEdit.ChooseFromListAlias = "ItemCode"
        //			return;
        //			PS_SM020_CF_ChooseFromList_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void PS_SM020_FormItemEnabled()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm01.Freeze(true);
        //			if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //				////각모드에따른 아이템설정
        //				////Call PS_SM020_FormClear '//UDO방식
        //			} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
        //				////각모드에따른 아이템설정
        //			} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
        //				////각모드에따른 아이템설정
        //			}
        //			oForm01.Freeze(false);
        //			return;
        //			PS_SM020_FormItemEnabled_Error:
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void PS_SM020_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm01.Freeze(true);
        //			//    If RowIserted = False Then '//행추가여부
        //			//        oDS_PS_SM020L.InsertRecord (oRow)
        //			//    End If
        //			//    oMat01.AddRow
        //			//    oDS_PS_SM020L.Offset = oRow
        //			//    oDS_PS_SM020L.setValue "U_LineNum", oRow, oRow + 1
        //			//    oMat01.LoadFromDataSource
        //			oForm01.Freeze(false);
        //			return;
        //			PS_SM020_AddMatrixRow_Error:
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void PS_SM020_FormClear()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			string DocEntry = null;
        //			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_SM020'", ref "");
        //			if (Convert.ToDouble(DocEntry) == 0) {
        //				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm01.Items.Item("DocEntry").Specific.VALUE = 1;
        //			} else {
        //				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm01.Items.Item("DocEntry").Specific.VALUE = DocEntry;
        //			}
        //			return;
        //			PS_SM020_FormClear_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void PS_SM020_EnableMenus()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////메뉴활성화
        //			//    Call oForm01.EnableMenu("1288", True)
        //			//    Call oForm01.EnableMenu("1289", True)
        //			//    Call oForm01.EnableMenu("1290", True)
        //			//    Call oForm01.EnableMenu("1291", True)
        //			////Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
        //			return;
        //			PS_SM020_EnableMenus_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		public bool PS_SM020_DataValidCheck()
        //		{
        //			bool functionReturnValue = false;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			//    If oForm01.Items("WhsCode").Specific.Value = "" Then
        //			//        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
        //			//        oForm01.Items("WhsCode").Click ct_Regular
        //			//        PS_SM020_DataValidCheck = False
        //			//        Exit Function
        //			//    End If
        //			//    If oMat01.VisualRowCount = 0 Then
        //			//        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        //			//        PS_SM020_DataValidCheck = False
        //			//        Exit Function
        //			//    End If
        //			//    For i = 1 To oMat01.VisualRowCount
        //			//        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
        //			//            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
        //			//            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
        //			//            PS_SM020_DataValidCheck = False
        //			//            Exit Function
        //			//        End If
        //			//    Next
        //			//    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
        //			//    Call oMat01.LoadFromDataSource
        //			PS_SM020_FormClear();
        //			return functionReturnValue;
        //			PS_SM020_DataValidCheck_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			return functionReturnValue;
        //		}

        //		private void PS_SM020_MTX01()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////메트릭스에 데이터 로드
        //			oForm01.Freeze(true);
        //			int i = 0;
        //			string Query01 = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			string Param01 = null;
        //			string Param02 = null;
        //			string Param03 = null;
        //			string Param04 = null;
        //			string Param05 = null;
        //			string Param06 = null;
        //			string Param07 = null;
        //			string Param08 = null;
        //			string Param09 = null;
        //			string Param10 = null;
        //			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param01 = Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param02 = Strings.Trim(oForm01.Items.Item("StockType").Specific.Selected.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param03 = Strings.Trim(oForm01.Items.Item("TradeType").Specific.Selected.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param04 = Strings.Trim(oForm01.Items.Item("ItmBsort").Specific.Selected.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param05 = Strings.Trim(oForm01.Items.Item("ItmMsort").Specific.Selected.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param06 = Strings.Trim(oForm01.Items.Item("Size").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param07 = Strings.Trim(oForm01.Items.Item("ItemType").Specific.Selected.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param08 = Strings.Trim(oForm01.Items.Item("Mark").Specific.Selected.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param09 = Strings.Trim(oForm01.Items.Item("ItemName").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param10 = Strings.Trim(oForm01.Items.Item("ItemGpCd").Specific.Selected.VALUE);

        //			if (oBaseForm01 == null) {
        //				Query01 = "EXEC PS_SM020_01 '" + Param01 + "','','','" + Param02 + "','','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
        //			} else if (oBaseForm01.Type == Convert.ToDouble("149") | oBaseForm01.Type == Convert.ToDouble("139") | oBaseForm01.Type == Convert.ToDouble("140") | oBaseForm01.Type == Convert.ToDouble("180") | oBaseForm01.Type == Convert.ToDouble("133") | oBaseForm01.Type == Convert.ToDouble("179") | oBaseForm01.Type == Convert.ToDouble("60091")) {
        //				Query01 = "EXEC PS_SM020_01 '" + Param01 + "','Y','','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
        //				////판매Y,구매,재고타입(1:재고있는것만,2:전체),거래타입(1:일반,2:임가공)
        //			} else if (oBaseForm01.Type == Convert.ToDouble("142") | oBaseForm01.Type == Convert.ToDouble("143") | oBaseForm01.Type == Convert.ToDouble("182") | oBaseForm01.Type == Convert.ToDouble("141") | oBaseForm01.Type == Convert.ToDouble("181") | oBaseForm01.Type == Convert.ToDouble("60092")) {
        //				Query01 = "EXEC PS_SM020_01 '" + Param01 + "','','Y','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
        //				////판매,구매Y
        //			} else {
        //				Query01 = "EXEC PS_SM020_01 '" + Param01 + "','','','" + Param02 + "','','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
        //			}
        //			RecordSet01.DoQuery(Query01);

        //			oMat01.Clear();
        //			oMat01.FlushToDataSource();
        //			oMat01.LoadFromDataSource();

        //			oMat02.Clear();
        //			oMat02.FlushToDataSource();
        //			oMat02.LoadFromDataSource();

        //			if ((RecordSet01.RecordCount == 0)) {
        //				oForm01.Items.Item("Mat01").Enabled = false;
        //				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
        //				goto PS_SM020_MTX01_Exit;
        //			} else {
        //				oForm01.Items.Item("Mat01").Enabled = true;
        //			}

        //			SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //			ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //				if (i != 0) {
        //					oDS_PS_SM020H.InsertRecord((i));
        //				}
        //				oDS_PS_SM020H.Offset = i;
        //				oDS_PS_SM020H.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //				oDS_PS_SM020H.SetValue("U_ColReg01", i, Convert.ToString(false));
        //				oDS_PS_SM020H.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("ItemCode").Value);
        //				oDS_PS_SM020H.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("ItemName").Value);
        //				oDS_PS_SM020H.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("CallSize").Value);
        //				oDS_PS_SM020H.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("Mark").Value);
        //				oDS_PS_SM020H.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("OnHand").Value);
        //				oDS_PS_SM020H.SetValue("U_ColQty02", i, RecordSet01.Fields.Item("IsCommited").Value);
        //				oDS_PS_SM020H.SetValue("U_ColQty03", i, RecordSet01.Fields.Item("OnOrder").Value);
        //				oDS_PS_SM020H.SetValue("U_ColQty04", i, RecordSet01.Fields.Item("OnEnabled").Value);
        //				RecordSet01.MoveNext();
        //				ProgressBar01.Value = ProgressBar01.Value + 1;
        //				ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
        //			}
        //			oMat01.LoadFromDataSource();
        //			oMat01.AutoResizeColumns();
        //			oForm01.Update();

        //			ProgressBar01.Stop();
        //			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgressBar01 = null;
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			oForm01.Freeze(false);
        //			return;
        //			PS_SM020_MTX01_Exit:
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			oForm01.Freeze(false);
        //			return;
        //			PS_SM020_MTX01_Error:
        //			ProgressBar01.Stop();
        //			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgressBar01 = null;
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void PS_SM020_MTX02()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////메트릭스에 데이터 로드
        //			oForm01.Freeze(true);
        //			int i = 0;
        //			string Query01 = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			string Param01 = null;
        //			string Param02 = null;
        //			string Param03 = null;
        //			string Param04 = null;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE;
        //			//    Param02 = oMat01.Columns("WhsCode").Cells(oMat01Row01).Specific.Value
        //			//    Param03 = Trim(oForm01.Items("Param01").Specific.VALUE)
        //			//    Param04 = Trim(oForm01.Items("Param01").Specific.VALUE)

        //			Query01 = "EXEC PS_SM020_02 '" + Param01 + "'";
        //			RecordSet01.DoQuery(Query01);

        //			oMat02.Clear();
        //			oMat02.FlushToDataSource();
        //			oMat02.LoadFromDataSource();

        //			if ((RecordSet01.RecordCount == 0)) {
        //				oForm01.Items.Item("Mat02").Enabled = false;
        //				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
        //				goto PS_SM020_MTX02_Exit;
        //			} else {
        //				oForm01.Items.Item("Mat02").Enabled = true;
        //			}

        //			////품목이 멀티이면 수량,중량필드 비활성화
        //			////멀티
        //			if (MDC_PS_Common.GetItem_ItmBsort(Param01) == "104" | MDC_PS_Common.GetItem_ItmBsort(Param01) == "302") {
        //				oMat02.Columns.Item("SelQty").Editable = false;
        //				oMat02.Columns.Item("SelWeight").Editable = false;
        //			////그외품목은 수량선택가능
        //			} else {
        //				oMat02.Columns.Item("SelQty").Editable = true;
        //				oMat02.Columns.Item("SelWeight").Editable = true;
        //			}

        //			SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //			ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //				if (i != 0) {
        //					oDS_PS_SM020L.InsertRecord((i));
        //				}
        //				oDS_PS_SM020L.Offset = i;
        //				oDS_PS_SM020L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //				oDS_PS_SM020L.SetValue("U_ColReg01", i, Convert.ToString(false));
        //				oDS_PS_SM020L.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("BatchNum").Value);
        //				oDS_PS_SM020L.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("WhsCode").Value);
        //				oDS_PS_SM020L.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("WhsName").Value);
        //				oDS_PS_SM020L.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("PackNo").Value);
        //				oDS_PS_SM020L.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("Weight").Value);
        //				oDS_PS_SM020L.SetValue("U_ColNum02", i, RecordSet01.Fields.Item("SelQty").Value);
        //				oDS_PS_SM020L.SetValue("U_ColQty02", i, RecordSet01.Fields.Item("SelWeight").Value);
        //				RecordSet01.MoveNext();
        //				ProgressBar01.Value = ProgressBar01.Value + 1;
        //				ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
        //			}
        //			oMat02.LoadFromDataSource();
        //			oMat02.AutoResizeColumns();
        //			oForm01.Update();

        //			ProgressBar01.Stop();
        //			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgressBar01 = null;
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			oForm01.Freeze(false);
        //			return;
        //			PS_SM020_MTX02_Exit:
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			oForm01.Freeze(false);
        //			return;
        //			PS_SM020_MTX02_Error:
        //			ProgressBar01.Stop();
        //			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgressBar01 = null;
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void PS_SM020_SetBaseForm()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			string ItemCode01 = null;
        //			SAPbouiCOM.Matrix oBaseMat01 = null;
        //			if (oBaseForm01 == null) {
        //				////DoNothing
        //			////호출한폼타입
        //			} else if (oBaseForm01.Type == Convert.ToDouble("133") | oBaseForm01.Type == Convert.ToDouble("139") | oBaseForm01.Type == Convert.ToDouble("140") | oBaseForm01.Type == Convert.ToDouble("141") | oBaseForm01.Type == Convert.ToDouble("142") | oBaseForm01.Type == Convert.ToDouble("143") | oBaseForm01.Type == Convert.ToDouble("149") | oBaseForm01.Type == Convert.ToDouble("179") | oBaseForm01.Type == Convert.ToDouble("180") | oBaseForm01.Type == Convert.ToDouble("181") | oBaseForm01.Type == Convert.ToDouble("182") | oBaseForm01.Type == Convert.ToDouble("60091") | oBaseForm01.Type == Convert.ToDouble("60092")) {
        //				oBaseMat01 = oBaseForm01.Items.Item("38").Specific;
        //				for (i = 1; i <= oMat01.RowCount; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns(CHK).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((Conversion.Val(oMat01.Columns.Item("SelWeight").Cells.Item(i).Specific.VALUE) <= 0)) {
        //							////중량이 선택되지 않은품목
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE;
        //							//UPGRADE_WARNING: oBaseMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.VALUE = ItemCode01;
        //							////품목
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //							oBaseColRow01 = oBaseColRow01 + 1;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE;
        //							//UPGRADE_WARNING: oBaseMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.VALUE = ItemCode01;
        //							////품목
        //							//UPGRADE_WARNING: oBaseMat01.Columns(U_Qty).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("U_Qty").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.VALUE);
        //							////수량 '//수량을 변경하면 중량이 자동변경된다.
        //							//UPGRADE_WARNING: oBaseMat01.Columns(U_Unweight).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("U_Unweight").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01));
        //							////단중
        //							//UPGRADE_WARNING: oBaseMat01.Columns(11).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("11").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(oMat01.Columns.Item("SelWeight").Cells.Item(i).Specific.VALUE);
        //							////중량
        //							//UPGRADE_WARNING: oBaseMat01.Columns(14).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oBaseForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("14").Cells.Item(oBaseColRow01).Specific.VALUE = MDC_PS_Common.GetValue("EXEC PS_SBO_GETPRICE '" + oBaseForm01.Items.Item("4").Specific.VALUE + "','" + ItemCode01 + "'", 0, 1);
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //							oBaseColRow01 = oBaseColRow01 + 1;
        //						}
        //					}
        //				}
        //				////배치선택품목
        //				for (i = 1; i <= oMat02.RowCount; i++) {
        //					//UPGRADE_WARNING: oMat02.Columns(CHK).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true) {
        //						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((Conversion.Val(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.VALUE) <= 0)) {
        //							////중량이 선택되지 않은품목
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE;
        //							//UPGRADE_WARNING: oBaseMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.VALUE = ItemCode01;
        //							////품목
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //							oBaseColRow01 = oBaseColRow01 + 1;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE;
        //							//UPGRADE_WARNING: oBaseMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.VALUE = ItemCode01;
        //							////품목
        //							//UPGRADE_WARNING: oBaseMat01.Columns(U_Qty).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("U_Qty").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.VALUE);
        //							////수량 '//수량을 변경하면 중량이 자동변경된다.
        //							//UPGRADE_WARNING: oBaseMat01.Columns(U_Unweight).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("U_Unweight").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01));
        //							////단중
        //							//UPGRADE_WARNING: oBaseMat01.Columns(11).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("11").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.VALUE);
        //							////중량
        //							//UPGRADE_WARNING: oBaseMat01.Columns(14).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oBaseForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("14").Cells.Item(oBaseColRow01).Specific.VALUE = MDC_PS_Common.GetValue("EXEC PS_SBO_GETPRICE '" + oBaseForm01.Items.Item("4").Specific.VALUE + "','" + ItemCode01 + "'", 0, 1);
        //							//                    If oBaseForm01.Type = "140" Then
        //							//                        oBaseMat01.Columns("U_BatchNum").Cells(oBaseColRow01).Specific.Value = oMat02.Columns("BatchNum").Cells(i).Specific.Value '//배치번호
        //							//                    End If
        //							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //							oBaseColRow01 = oBaseColRow01 + 1;
        //						}
        //					}
        //				}
        //			} else if (oBaseForm01.TypeEx == "720") {
        //				oBaseMat01 = oBaseForm01.Items.Item("13").Specific;
        //				////매트릭스
        //				////품목선택품목
        //				for (i = 1; i <= oMat01.RowCount; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns(CHK).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true) {
        //						//UPGRADE_WARNING: oBaseMat01.Columns(1).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.VALUE = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE;
        //						////품목
        //						oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						oBaseColRow01 = oBaseColRow01 + 1;
        //					}
        //				}
        //				////배치선택품목
        //				for (i = 1; i <= oMat02.RowCount; i++) {
        //					//UPGRADE_WARNING: oMat02.Columns(CHK).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true) {
        //						//UPGRADE_WARNING: oBaseMat01.Columns(1).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.VALUE = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE;
        //						////품목
        //						oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						oBaseColRow01 = oBaseColRow01 + 1;
        //					}
        //				}
        //			} else if (oBaseForm01.TypeEx == "PS_SD091") {
        //				oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
        //				////매트릭스
        //				//        For i = 1 To oMat01.RowCount '//품목선택품목
        //				//            If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
        //				//                If (Val(oMat01.Columns("SelQty").Cells(i).Specific.Value) <= 0) Then
        //				//                    '//수량이 선택되지 않은품목
        //				//                Else
        //				//                    oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("ItemCode").Cells(i).Specific.Value '//품목
        //				//                    oBaseMat01.Columns("OutWhCd").Cells(oBaseColRow01).Specific.Value = oMat02.Columns("WhsCode").Cells(oMat01Row01).Specific.Value '//출고창고
        //				//                    oBaseMat01.Columns("Qty").Cells(oBaseColRow01).Specific.Value = Val(oMat01.Columns("SelQty").Cells(i).Specific.Value) '//수량 '//수량을 변경하면 중량이 자동변경된다.
        //				//                    oBaseMat01.Columns("Unweight").Cells(oBaseColRow01).Specific.Value = Val(MDC_PS_Common.GetItem_UnWeight(oMat01.Columns("ItemCode").Cells(i).Specific.Value)) '//단중
        //				//'                    oBaseMat01.Columns("14").Cells(oBaseColRow01).Specific.Value = MDC_PS_Common.GetValue("EXEC PS_SBO_GETPRICE '" & oBaseForm01.Items("4").Specific.Value & "','" & oMat01.Columns("ItemCode").Cells(i).Specific.Value & "'", 0, 1)
        //				//                    oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01 + 1).Click ct_Regular
        //				//                    oBaseColRow01 = oBaseColRow01 + 1
        //				//                End If
        //				//            End If
        //				//        Next
        //				////배치선택품목
        //				for (i = 1; i <= oMat01.RowCount; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns(CHK).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.VALUE) <= 0)) {
        //							////수량이 선택되지 않은품목
        //						} else {
        //							//UPGRADE_WARNING: oBaseMat01.Columns(ItemCode).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.VALUE = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE;
        //							////품목
        //							//                    oBaseForm01.Items("OutWhCd").Specific.Value = oMat02.Columns("WhsCode").Cells(i).Specific.Value
        //							//                    oBaseMat01.Columns("OutWhCd").Cells(oBaseColRow01).Specific.Value = oMat02.Columns("WhsCode").Cells(i).Specific.Value '//출고창고
        //							//                    oBaseMat01.Columns("BatchNum").Cells(oBaseColRow01).Specific.Value = oMat02.Columns("BatchNum").Cells(i).Specific.Value '//배치번호
        //							//UPGRADE_WARNING: oBaseMat01.Columns(Qty).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("Qty").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.VALUE);
        //							////수량 '//수량을 변경하면 중량이 자동변경된다.
        //							//UPGRADE_WARNING: oBaseMat01.Columns(Unweight).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("Unweight").Cells.Item(oBaseColRow01).Specific.VALUE = Conversion.Val(MDC_PS_Common.GetItem_UnWeight(oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE));
        //							////단중
        //							//UPGRADE_WARNING: oBaseMat01.Columns(Weight).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oBaseMat01.Columns.Item("Weight").Cells.Item(oBaseColRow01).Specific.VALUE = System.Math.Round((Conversion.Val(MDC_PS_Common.GetItem_UnWeight(oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE)) * Conversion.Val(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.VALUE)) / 1000, 3);
        //							////단중
        //							//                    oBaseMat01.Columns("14").Cells(oBaseColRow01).Specific.Value = MDC_PS_Common.GetValue("EXEC PS_SBO_GETPRICE '" & oBaseForm01.Items("4").Specific.Value & "','" & oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.Value & "'", 0, 1)
        //							//                    oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01 + 1).Click ct_Regular
        //							oBaseColRow01 = oBaseColRow01 + 1;
        //						}
        //					}
        //				}

        //			}
        //			return;
        //			PS_SM020_SetBaseForm_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("PS_SM020_SetBaseForm_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
    }
}
