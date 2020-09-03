using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통합재무제표용 계정 관리
	/// </summary>
	internal class PS_CO658 : PSH_BaseClass
	{
		public string oFormUniqueID;
		//public SAPbouiCOM.Form oForm;
		public SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_CO658H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO658L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string oDocType01;
		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO658.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO658_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO658");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";
				
				oForm.Freeze(true);
				//PS_CO658_CreateItems();
				//PS_CO658_ComboBox_Setting();
				//PS_CO658_CF_ChooseFromList();
				//PS_CO658_EnableMenus();
				//PS_CO658_SetDocument(oFromDocEntry01);
				//PS_CO658_FormResize();

				oForm.EnableMenu(("1283"), true); //삭제
				oForm.EnableMenu(("1287"), true); //복제
				oForm.EnableMenu(("1286"), false); //닫기
				oForm.EnableMenu(("1284"), false); //취소
				oForm.EnableMenu(("1293"), true); //행삭제
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

		/// <summary>
		/// Form Item Event
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">pVal</param>
		/// <param name="BubbleEvent">Bubble Event</param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			switch (pVal.EventType)
			{
				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
					Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;

				//case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
				//    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
				//	Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
				//    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
				//    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
				//	Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
				//    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
				//    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
				//    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
				//    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
				//    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
				//	Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
				//	break;

				//case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
				//    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
				//    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_Drag: //39
				//    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
				//    break;
			}
		}

		/// <summary>
		/// ITEM_PRESSED 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	switch (pval.EventType)
        //	{
        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //			////1
        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //			////2
        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //			////5
        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
        //			////6
        //			Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //			////7
        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //			////8
        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //			////10
        //			Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //			////11
        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
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
        //			Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //			////27
        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //			////3
        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //			////4
        //			break;
        //		////et_LOST_FOCUS
        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //			////17
        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //	}
        //	return;
        //Raise_ItemEvent_Error:

        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true))
        //	{
        //		switch (pval.MenuUID)
        //		{
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
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
        //		////BeforeAction = False
        //	}
        //	else if ((pval.BeforeAction == false))
        //	{
        //		switch (pval.MenuUID)
        //		{
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_CO658_FormItemEnabled();
        //				////UDO방식
        //				oForm01.Items.Item("Name").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			case "1282":
        //				//추가
        //				PS_CO658_FormItemEnabled();
        //				////UDO방식
        //				PS_CO658_AddMatrixRow(0, ref true);
        //				////UDO방식
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_CO658_FormItemEnabled();
        //				break;

        //			//복제(2013.03.13 송명규 추가)
        //			case "1287":

        //				oForm01.Freeze(true);
        //				PS_CO658_FormClear();
        //				//oDS_PS_CO658H.setValue "Code", 0, ""

        //				for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //				{
        //					oMat01.FlushToDataSource();
        //					oDS_PS_CO658L.SetValue("Code", i, "");
        //					oMat01.LoadFromDataSource();
        //				}

        //				oForm01.Freeze(false);
        //				break;

        //		}
        //	}
        //	return;
        //Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true))
        //	{
        //		switch (BusinessObjectInfo.EventType)
        //		{
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
        //		////BeforeAction = False
        //	}
        //	else if ((BusinessObjectInfo.BeforeAction == false))
        //	{
        //		switch (BusinessObjectInfo.EventType)
        //		{
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
        //Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{
        //		//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //		//            MenuCreationParams01.uniqueID = "MenuUID"
        //		//            MenuCreationParams01.String = "메뉴명"
        //		//            MenuCreationParams01.Enabled = True
        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //		//        End If
        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //		//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
        //	}
        //	if (pval.ItemUID == "Mat01")
        //	{
        //		if (pval.Row > 0)
        //		{
        //			oLastItemUID01 = pval.ItemUID;
        //			oLastColUID01 = pval.ColUID;
        //			oLastColRow01 = pval.Row;
        //		}
        //	}
        //	else
        //	{
        //		oLastItemUID01 = pval.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{
        //		if (pval.ItemUID == "PS_CO658")
        //		{
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //			{
        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //			{
        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //			{
        //			}
        //		}
        //		if (pval.ItemUID == "1")
        //		{
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //			{
        //				if (PS_CO658_DataValidCheck() == false)
        //				{
        //					BubbleEvent = false;
        //					return;
        //				}

        //				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocEntry01 = oForm01.Items.Item("Code").Specific.VALUE;
        //				oFormMode01 = oForm01.Mode;

        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //			{
        //				if (PS_CO658_DataValidCheck() == false)
        //				{
        //					BubbleEvent = false;
        //					return;
        //				}

        //				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocEntry01 = oForm01.Items.Item("Code").Specific.VALUE;
        //				oFormMode01 = oForm01.Mode;

        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //			{
        //			}
        //		}
        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //		if (pval.ItemUID == "PS_CO658")
        //		{
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //			{
        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //			{
        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //			{
        //			}
        //		}
        //		if (pval.ItemUID == "1")
        //		{
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //			{
        //				if (pval.ActionSuccess == true)
        //				{
        //					PS_CO658_FormItemEnabled();
        //					PS_CO658_AddMatrixRow(0, ref true);
        //					////UDO방식일때
        //				}
        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //			{
        //			}
        //			else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //			{
        //				if (pval.ActionSuccess == true)
        //				{
        //					PS_CO658_FormItemEnabled();
        //				}
        //			}
        //		}

        //	}
        //	return;
        //Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	if (pval.BeforeAction == true)
        //	{

        //		if (pval.ItemUID == "Mat01")
        //		{

        //			if (pval.ColUID == "AcctCode")
        //			{

        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "Mat01", "AcctCode");
        //				//계정

        //			}

        //			//구분
        //		}
        //		else if (pval.ItemUID == "Code")
        //		{

        //			//Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode", "") '계정 포맷서치 설정

        //		}

        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}

        //	return;
        //Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	oForm01.Freeze(true);
        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //		if (pval.ItemChanged == true)
        //		{

        //		}
        //	}
        //	oForm01.Freeze(false);
        //	return;
        //Raise_EVENT_COMBO_SELECT_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{
        //		//        If pval.ItemUID = "Mat01" Then
        //		//            If pval.Row > 0 Then
        //		//                Call oMat01.SelectRow(pval.Row, True, False)
        //		//            End If
        //		//        End If
        //		if (pval.ItemUID == "Mat01")
        //		{
        //			if (pval.Row > 0)
        //			{
        //				oLastItemUID01 = pval.ItemUID;
        //				oLastColUID01 = pval.ColUID;
        //				oLastColRow01 = pval.Row;

        //				oMat01.SelectRow(pval.Row, true, false);
        //			}
        //		}
        //		else
        //		{
        //			oLastItemUID01 = pval.ItemUID;
        //			oLastColUID01 = "";
        //			oLastColRow01 = 0;
        //		}
        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	object oTempClass = null;
        //	if (pval.BeforeAction == true)
        //	{
        //		if (pval.ItemUID == "Mat01")
        //		{

        //		}
        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;

        //	oForm01.Freeze(true);
        //	if (pval.BeforeAction == true)
        //	{
        //		if (pval.ItemChanged == true)
        //		{

        //		}
        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //		if (pval.ItemChanged == true)
        //		{

        //			PS_CO658_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);

        //		}

        //	}
        //	oForm01.Freeze(false);
        //	return;
        //Raise_EVENT_VALIDATE_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //		PS_CO658_FormItemEnabled();
        //		PS_CO658_AddMatrixRow(oMat01.VisualRowCount);
        //		////UDO방식
        //	}
        //	return;
        //Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //		PS_CO658_FormResize();
        //	}
        //	return;
        //Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //	}
        //	return;
        //Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.ItemUID == "Mat01")
        //	{
        //		if (pval.Row > 0)
        //		{
        //			oLastItemUID01 = pval.ItemUID;
        //			oLastColUID01 = pval.ColUID;
        //			oLastColRow01 = pval.Row;
        //		}
        //	}
        //	else
        //	{
        //		oLastItemUID01 = pval.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{
        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //		SubMain.RemoveForms(oFormUniqueID01);
        //		//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm01 = null;
        //		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat01 = null;
        //	}
        //	return;
        //Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0))
        //	{
        //		if (pval.BeforeAction == true)
        //		{
        //			////행삭제전 행삭제가능여부검사
        //		}
        //		else if (pval.BeforeAction == false)
        //		{
        //			for (i = 1; i <= oMat01.VisualRowCount; i++)
        //			{
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_CO658L.RemoveRecord(oDS_PS_CO658L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0)
        //			{
        //				PS_CO658_AddMatrixRow(0);
        //			}
        //			else
        //			{
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_CO658L.GetValue("U_AcctCode", oMat01.RowCount - 1))))
        //				{
        //					PS_CO658_AddMatrixRow(oMat01.RowCount);
        //				}
        //			}
        //		}
        //	}
        //	return;
        //Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_CreateItems
        //private bool PS_CO658_CreateItems()
        //{
        //	bool functionReturnValue = false;
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	oDS_PS_CO658H = oForm01.DataSources.DBDataSources("@PS_CO658H");
        //	oDS_PS_CO658L = oForm01.DataSources.DBDataSources("@PS_CO658L");
        //	oMat01 = oForm01.Items.Item("Mat01").Specific;

        //	oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //	return functionReturnValue;
        //PS_CO658_CreateItems_Error:
        //	//    oMat01.AutoResizeColumns

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_CO658_ComboBox_Setting
        //public void PS_CO658_ComboBox_Setting()
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	SAPbouiCOM.ComboBox oCombo = null;
        //	string sQry = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;

        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_CO658", "Mat01", "UseYN", "Y", "Y");
        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_CO658", "Mat01", "UseYN", "N", "N");
        //	MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("UseYN"), "PS_CO658", "Mat01", "UseYN");

        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "0", "계산");
        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "1", "차변잔액");
        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "2", "차변합계");
        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "3", "대변잔액");
        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "4", "대변합계");
        //	MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("CRStd"), "PS_CO658", "Mat01", "CRStd");

        //	//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oCombo = null;
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;

        //	return;
        //PS_CO658_ComboBox_Setting_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_CF_ChooseFromList
        //public void PS_CO658_CF_ChooseFromList()
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	return;
        //PS_CO658_CF_ChooseFromList_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_FormItemEnabled
        //public void PS_CO658_FormItemEnabled()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO658_FormItemEnabled()
        //	//해당모듈    : PS_CO658
        //	//기능        : 모드에 따른 아이템 설정
        //	//인수        : 없음
        //	//반환값      : 없음
        //	//특이사항    : 없음
        //	//******************************************************************************
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);
        //	if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //	{

        //		oForm01.Items.Item("Code").Enabled = true;
        //		oForm01.Items.Item("Mat01").Enabled = true;
        //		PS_CO658_FormClear();
        //		////UDO방식

        //		oForm01.EnableMenu("1281", true);
        //		////찾기
        //		oForm01.EnableMenu("1282", false);
        //		////추가

        //	}
        //	else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
        //	{

        //		//        oForm01.Items("Code").Specific.VALUE = ""
        //		oForm01.Items.Item("Code").Enabled = true;
        //		oForm01.Items.Item("Mat01").Enabled = false;

        //		oForm01.EnableMenu("1281", false);
        //		////찾기
        //		oForm01.EnableMenu("1282", true);
        //		////추가

        //	}
        //	else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
        //	{

        //		oForm01.Items.Item("Code").Enabled = false;
        //		oForm01.Items.Item("Mat01").Enabled = true;

        //		oForm01.EnableMenu("1281", true);
        //		////찾기
        //		oForm01.EnableMenu("1282", true);
        //		////추가

        //	}

        //	oMat01.AutoResizeColumns();

        //	oForm01.Freeze(false);
        //	return;
        //PS_CO658_FormItemEnabled_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_AddMatrixRow
        //public void PS_CO658_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	oForm01.Freeze(true);
        //	////행추가여부
        //	if (RowIserted == false)
        //	{
        //		oDS_PS_CO658L.InsertRecord((oRow));
        //	}
        //	oMat01.AddRow();
        //	oDS_PS_CO658L.Offset = oRow;
        //	oDS_PS_CO658L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oMat01.LoadFromDataSource();
        //	oForm01.Freeze(false);
        //	return;
        //PS_CO658_AddMatrixRow_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_FormClear
        //public void PS_CO658_FormClear()
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	string DocEntry = null;
        //	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_CO658'", ref "");
        //	if (string.IsNullOrEmpty(DocEntry) | DocEntry == "0")
        //	{
        //		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm01.Items.Item("DocEntry").Specific.VALUE = 1;
        //		//oForm01.Items("Code").Specific.VALUE = 1
        //	}
        //	else
        //	{
        //		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm01.Items.Item("DocEntry").Specific.VALUE = DocEntry;
        //		//oForm01.Items("Code").Specific.VALUE = DocEntry
        //	}
        //	return;
        //PS_CO658_FormClear_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_EnableMenus
        //private void PS_CO658_EnableMenus()
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	////메뉴활성화
        //	//    Call oForm01.EnableMenu("1288", True)
        //	//    Call oForm01.EnableMenu("1289", True)
        //	//    Call oForm01.EnableMenu("1290", True)
        //	//    Call oForm01.EnableMenu("1291", True)
        //	////Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
        //	MDC_Com.MDC_GP_EnableMenus(oForm01, false, false, true, true, false, true, true, true, true,
        //	false, false, false, false, false, false);
        //	////메뉴설정
        //	return;
        //PS_CO658_EnableMenus_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_SetDocument
        //private void PS_CO658_SetDocument(string oFromDocEntry01)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if ((string.IsNullOrEmpty(oFromDocEntry01)))
        //	{
        //		PS_CO658_FormItemEnabled();
        //		PS_CO658_AddMatrixRow(0, ref true);
        //		////UDO방식일때
        //		oForm01.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //	}
        //	else
        //	{
        //		//        oForm01.Mode = fm_FIND_MODE
        //		//        Call PS_CO658_FormItemEnabled
        //		//        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        //		//        oForm01.Items("1").Click ct_Regular
        //	}
        //	return;
        //PS_CO658_SetDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_DataValidCheck
        //public bool PS_CO658_DataValidCheck()
        //{
        //	bool functionReturnValue = false;
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = false;
        //	int i = 0;
        //	if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //	{
        //		PS_CO658_FormClear();
        //	}

        //	//UPGRADE_WARNING: oForm01.Items(Code).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm01.Items.Item("Code").Specific.VALUE))
        //	{
        //		SubMain.Sbo_Application.SetStatusBarMessage("구분코드가 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	//UPGRADE_WARNING: oForm01.Items(Name).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm01.Items.Item("Name").Specific.VALUE))
        //	{
        //		SubMain.Sbo_Application.SetStatusBarMessage("구분명이 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	//라인정보 미입력 시
        //	if (oMat01.VisualRowCount == 1)
        //	{
        //		SubMain.Sbo_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
        //	{
        //		//UPGRADE_WARNING: oMat01.Columns(AcctCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.VALUE)))
        //		{
        //			SubMain.Sbo_Application.SetStatusBarMessage("계정코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("AcctCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}

        //		//        If (oMat01.Columns("AcctName").Cells(i).Specific.VALUE = "") Then
        //		//            Sbo_Application.SetStatusBarMessage "계정명은 필수입니다.", bmt_Short, True
        //		//            oMat01.Columns("AcctName").Cells(i).Click ct_Regular
        //		//            PS_CO658_DataValidCheck = False
        //		//            Exit Function
        //		//        End If

        //		//UPGRADE_WARNING: oMat01.Columns(Contents).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("Contents").Cells.Item(i).Specific.VALUE)))
        //		{
        //			SubMain.Sbo_Application.SetStatusBarMessage("목차제목은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("Contents").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        //	}

        //	oMat01.FlushToDataSource();
        //	oDS_PS_CO658L.RemoveRecord(oDS_PS_CO658L.Size - 1);
        //	oMat01.LoadFromDataSource();

        //	if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //	{
        //		PS_CO658_FormClear();
        //	}

        //	functionReturnValue = true;
        //	return functionReturnValue;
        //PS_CO658_DataValidCheck_Error:
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_CO658_FlushToItemValue
        //private void PS_CO658_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;
        //	short ErrNum = 0;
        //	string sQry = null;
        //	string ItemCode = null;

        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	switch (oUID)
        //	{

        //		case "Mat01":

        //			if (oCol == "AcctCode")
        //			{

        //				oMat01.FlushToDataSource();
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDS_PS_CO658L.SetValue("U_" + oCol, oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE);
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDS_PS_CO658L.SetValue("U_AcctName", oRow - 1, MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE + "'"));
        //				if (oMat01.RowCount == oRow & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_CO658L.GetValue("U_" + oCol, oRow - 1))))
        //				{
        //					PS_CO658_AddMatrixRow(oRow);
        //				}
        //				oMat01.LoadFromDataSource();

        //			}

        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oMat01.Columns.Item("UseYN").Cells.Item(oRow).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //			//기본으로 'Y' 세팅
        //			oMat01.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			//다음 컬럼 클릭
        //			oMat01.AutoResizeColumns();
        //			break;

        //		case "Code":
        //			break;

        //			//oForm01.Items("Name").Specific.VALUE = MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" & Trim(oForm01.Items("Code").Specific.VALUE) & "'") '계정명

        //	}

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;

        //	return;
        //PS_CO658_FlushToItemValue_Error:


        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;

        //	MDC_Com.MDC_GF_Message(ref "PS_CO658_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");

        //}
        #endregion

        #region PS_CO658_MTX01
        //private void PS_CO658_MTX01()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO658_MTX01()
        //	//해당모듈    : PS_CO658
        //	//기능        : Matrix 데이터 로드
        //	//인수        : 없음
        //	//반환값      : 없음
        //	//특이사항    : 사용안함
        //	//******************************************************************************
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string Param01 = null;
        //	string Param02 = null;
        //	string Param03 = null;
        //	string Param04 = null;
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param01 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param02 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param03 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param04 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);

        //	Query01 = "SELECT 10";
        //	RecordSet01.DoQuery(Query01);

        //	oMat01.Clear();
        //	oMat01.FlushToDataSource();
        //	oMat01.LoadFromDataSource();

        //	if ((RecordSet01.RecordCount == 0))
        //	{
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_CO658_MTX01_Exit;
        //	}

        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //	for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
        //	{
        //		if (i != 0)
        //		{
        //			oDS_PS_CO658L.InsertRecord((i));
        //		}
        //		oDS_PS_CO658L.Offset = i;
        //		oDS_PS_CO658L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
        //		oDS_PS_CO658L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
        //		RecordSet01.MoveNext();
        //		ProgressBar01.Value = ProgressBar01.Value + 1;
        //		ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	oForm01.Update();

        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //PS_CO658_MTX01_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	if ((ProgressBar01 != null))
        //	{
        //		ProgressBar01.Stop();
        //	}
        //	return;
        //PS_CO658_MTX01_Error:
        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_FormResize
        //private void PS_CO658_FormResize()
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	oMat01.AutoResizeColumns();

        //	return;
        //PS_CO658_FormResize_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_Validate
        //public bool PS_CO658_Validate(string ValidateType)
        //{
        //	bool functionReturnValue = false;
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	if (ValidateType == "수정")
        //	{
        //		////삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
        //	}
        //	else if (ValidateType == "행삭제")
        //	{
        //		////행삭제전 행삭제가능여부검사
        //	}
        //	else if (ValidateType == "취소")
        //	{
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //PS_CO658_Validate_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //PS_CO658_Validate_Error:
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO658_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion
    }
}
