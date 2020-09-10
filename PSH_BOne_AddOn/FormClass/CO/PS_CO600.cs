using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통합재무제표
	/// </summary>
	internal class PS_CO600 : PSH_BaseClass
	{

		public string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;

		public SAPbouiCOM.Grid oGrid01;
		public SAPbouiCOM.Grid oGrid02;
		public SAPbouiCOM.Grid oGrid03;
		public SAPbouiCOM.Grid oGrid04;

		public SAPbouiCOM.DataTable oDS_PS_CO600A;
		public SAPbouiCOM.DataTable oDS_PS_CO600B;
		public SAPbouiCOM.DataTable oDS_PS_CO600C;
		public SAPbouiCOM.DataTable oDS_PS_CO600D;

		//public SAPbouiCOM.Form oBaseForm01; //부모폼
		//public string oBaseItemUID01;
		//public string oBaseColUID01;
		//public int oBaseColRow01;
		//public string oBaseTradeType01;
		//public string oBaseItmBsort01;

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO600.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO600_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO600");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

				oForm.Freeze(true);
                PS_CO600_CreateItems();
                //PS_CO600_ComboBox_Setting();

                oForm.Items.Item("FrDt01").Specific.Value = DateTime.Now.ToString("yyyy0101"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY0101");
				oForm.Items.Item("ToDt01").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
				
				oForm.Items.Item("Folder01").Specific.Select();				
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

        private void PS_CO600_CreateItems()
        {
            try
            {
                //oForm.Freeze(true);

                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oGrid02 = oForm.Items.Item("Grid02").Specific;
                oGrid03 = oForm.Items.Item("Grid03").Specific;
                oGrid04 = oForm.Items.Item("Grid04").Specific;

                oForm.DataSources.DataTables.Add("PS_CO600A");
                oForm.DataSources.DataTables.Add("PS_CO600B");
                oForm.DataSources.DataTables.Add("PS_CO600C");
                oForm.DataSources.DataTables.Add("PS_CO600D");

                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_CO600A");
                oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_CO600B");
                oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_CO600C");
                oGrid04.DataTable = oForm.DataSources.DataTables.Item("PS_CO600D");

                oDS_PS_CO600A = oForm.DataSources.DataTables.Item("PS_CO600A");
                oDS_PS_CO600B = oForm.DataSources.DataTables.Item("PS_CO600B");
                oDS_PS_CO600C = oForm.DataSources.DataTables.Item("PS_CO600C");
                oDS_PS_CO600D = oForm.DataSources.DataTables.Item("PS_CO600D");

                //조회기간(시작)
                oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");

                //조회기간(종료)
                oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");

                //출력구분
                oForm.DataSources.UserDataSources.Add("Ctgr01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Ctgr01").Specific.DataBind.SetBound(true, "", "Ctgr01");

            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                //oForm.Freeze(false);
            }
        }



















        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	switch (pval.EventType) {
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
        //	Raise_ItemEvent_Error:

        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true)) {
        //		switch (pval.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			////Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
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
        //	} else if ((pval.BeforeAction == false)) {
        //		switch (pval.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			////Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
        //			case "1281":
        //				//찾기
        //				break;
        //			////Call PS_CO600_FormItemEnabled '//UDO방식
        //			case "1282":
        //				//추가
        //				break;
        //			////Call PS_CO600_FormItemEnabled '//UDO방식
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
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //		//            MenuCreationParams01.uniqueID = "MenuUID"
        //		//            MenuCreationParams01.String = "메뉴명"
        //		//            MenuCreationParams01.Enabled = True
        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //		//        End If
        //	} else if (pval.BeforeAction == false) {
        //		//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
        //	}
        //	if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02") {
        //		if (pval.Row > 0) {
        //			oLastItemUID01 = pval.ItemUID;
        //			oLastColUID01 = pval.ColUID;
        //			oLastColRow01 = pval.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pval.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pval.BeforeAction == true) {

        //		if (pval.ItemUID == "BtnSrch01") {

        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_CO600_MTX01();
        //				//대차대조표(재무상태표)
        //				PS_CO600_MTX02();
        //				//제조원가명세서
        //				PS_CO600_MTX03();
        //				//매출원가명세서
        //				PS_CO600_MTX04();
        //				//손익계산서(포괄손익계산서)
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		//대차대조표(재무상태표)
        //		} else if (pval.ItemUID == "BtnPrt01") {

        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_CO600_Print_Report01();
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		//제품원가명세서
        //		} else if (pval.ItemUID == "BtnPrt02") {

        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_CO600_Print_Report02();
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		//매출원가명세서
        //		} else if (pval.ItemUID == "BtnPrt03") {

        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_CO600_Print_Report03();
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		//손익계산서(포괄손익계산서)
        //		} else if (pval.ItemUID == "BtnPrt04") {

        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_CO600_Print_Report04();
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		}

        //	} else if (pval.BeforeAction == false) {

        //		//폴더를 사용할 때는 필수 소스_S
        //		//Folder01이 선택되었을 때
        //		if (pval.ItemUID == "Folder01") {

        //			oForm01.PaneLevel = 1;

        //		}

        //		//Folder02가 선택되었을 때
        //		if (pval.ItemUID == "Folder02") {

        //			oForm01.PaneLevel = 2;

        //		}

        //		//Folder03가 선택되었을 때
        //		if (pval.ItemUID == "Folder03") {

        //			oForm01.PaneLevel = 3;

        //		}

        //		//Folder04가 선택되었을 때
        //		if (pval.ItemUID == "Folder04") {

        //			oForm01.PaneLevel = 4;

        //		}
        //		//폴더를 사용할 때는 필수 소스_E

        //		if (pval.ItemUID == "PS_CO600") {
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		//        If pval.ItemUID = "1" Then
        //		//            If oForm01.Mode = fm_ADD_MODE Then
        //		//                If pval.ActionSuccess = True Then
        //		//                    Call PS_CO600_FormItemEnabled
        //		//                    Call PS_CO600_FormClear '//UDO방식일때
        //		//                    Call PS_CO600_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
        //		//                End If
        //		//            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        //		//            ElseIf oForm01.Mode = fm_OK_MODE Then
        //		//                If pval.ActionSuccess = True Then
        //		//                    Call PS_CO600_FormItemEnabled
        //		//                End If
        //		//            End If
        //		//        End If
        //	}
        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pval.BeforeAction == true) {

        //		// Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode01", "")

        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_CO600_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
        //	}

        //	oForm01.Freeze(false);

        //	return;
        //	Raise_EVENT_COMBO_SELECT_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "Grid01") {
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pval.Row > 0) {

        //				}
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "Grid01") {
        //			if (pval.Row == -1) {
        //				//                oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True

        //			} else {
        //				if (oGrid01.Rows.SelectedRows.Count > 0) {

        //					//                    Call PS_CO600_SetBaseForm '//부모폼에입력
        //					//                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
        //					//                        Call oForm01.Close
        //					//                    End If
        //				} else {
        //					BubbleEvent = false;
        //				}
        //			}
        //		}
        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm01.Freeze(true);
        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_CO600_FlushToItemValue(pval.ItemUID);
        //	}
        //	oForm01.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_CO600_FormItemEnabled();
        //		////Call PS_CO600_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_CO600_FormResize();
        //	}
        //	return;
        //	Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		//If (pval.ItemUID = "ItemCode") Then
        //		//   Set oDataTable01 = pval.SelectedObjects
        //		//    If oDataTable01 Is Nothing Then
        //		//    Else
        //		//  oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
        //		//     '  oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
        //		//   End If
        //		// End If
        //		oForm01.Update();
        //	}

        //	return;
        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02") {
        //		if (pval.Row > 0) {
        //			oLastItemUID01 = pval.ItemUID;
        //			oLastColUID01 = pval.ColUID;
        //			oLastColRow01 = pval.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pval.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}

        //	return;
        //	Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //	} else if (pval.BeforeAction == false) {
        //		SubMain.RemoveForms(oFormUniqueID01);
        //		//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm01 = null;
        //		//UPGRADE_NOTE: oGrid01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oGrid01 = null;
        //		//UPGRADE_NOTE: oGrid03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oGrid03 = null;
        //		//UPGRADE_NOTE: oGrid03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oGrid03 = null;
        //		//UPGRADE_NOTE: oGrid04 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oGrid04 = null;
        //	}
        //	return;
        //	Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0)) {
        //		if (pval.BeforeAction == true) {
        //			////행삭제전 행삭제가능여부검사
        //		} else if (pval.BeforeAction == false) {
        //			//        For i = 1 To oMat01.VisualRowCount
        //			//            oMat01.Columns("COL01").Cells(i).Specific.Value = i
        //			//        Next i
        //			//        oMat01.FlushToDataSource
        //			//        Call oDS_PS_CO600L.RemoveRecord(oDS_PS_CO600L.Size - 1)
        //			//        oMat01.LoadFromDataSource
        //			//        If oMat01.RowCount = 0 Then
        //			//            Call PS_CO600_AddMatrixRow(0)
        //			//        Else
        //			//            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
        //			//                Call PS_CO600_AddMatrixRow(oMat01.RowCount)
        //			//            End If
        //			//        End If
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion



        #region PS_CO600_ComboBox_Setting
        //public void PS_CO600_ComboBox_Setting()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	SAPbouiCOM.ComboBox oCombo = null;
        //	string sQry = null;

        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string BPLId = null;
        //	BPLId = MDC_PS_Common.User_BPLId();

        //	oForm01.Freeze(true);
        //	////콤보에 기본값설정

        //	//출력구분
        //	//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm01.Items.Item("Ctgr01").Specific.ValidValues.Add("10", "K-GAAP");
        //	//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm01.Items.Item("Ctgr01").Specific.ValidValues.Add("20", "K-IFRS");
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm01.Items.Item("Ctgr01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_ComboBox_Setting_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_CF_ChooseFromList
        //public void PS_CO600_CF_ChooseFromList()
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
        //	//    Set oEdit = oForm01.Items("ItemCode").Specific
        //	//    Set oCFLs = oForm01.ChooseFromLists
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
        //	PS_CO600_CF_ChooseFromList_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_FormItemEnabled
        //public void PS_CO600_FormItemEnabled()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm01.Freeze(true);
        //	if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		if (string.IsNullOrEmpty(oBaseItmBsort01)) {

        //		} else {

        //		}
        //		////각모드에따른 아이템설정
        //		////Call PS_CO600_FormClear '//UDO방식
        //	} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
        //		////각모드에따른 아이템설정
        //	} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
        //		////각모드에따른 아이템설정
        //	}
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_FormItemEnabled_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_AddMatrixRow
        //public void PS_CO600_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm01.Freeze(true);
        //	//    If RowIserted = False Then '//행추가여부
        //	//        oDS_PS_CO600L.InsertRecord (oRow)
        //	//    End If
        //	//    oMat01.AddRow
        //	//    oDS_PS_CO600L.Offset = oRow
        //	//    oDS_PS_CO600L.setValue "U_LineNum", oRow, oRow + 1
        //	//    oMat01.LoadFromDataSource
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_AddMatrixRow_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_DataValidCheck
        //public bool PS_CO600_DataValidCheck()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	PS_CO600_DataValidCheck_Error:
        //	//    If oForm01.Items("WhsCode").Specific.Value = "" Then
        //	//        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
        //	//        oForm01.Items("WhsCode").Click ct_Regular
        //	//        PS_CO600_DataValidCheck = False
        //	//        Exit Function
        //	//    End If
        //	//    If oMat01.VisualRowCount = 0 Then
        //	//        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        //	//        PS_CO600_DataValidCheck = False
        //	//        Exit Function
        //	//    End If
        //	//    For i = 1 To oMat01.VisualRowCount
        //	//        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
        //	//            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
        //	//            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
        //	//            PS_CO600_DataValidCheck = False
        //	//            Exit Function
        //	//        End If
        //	//    Next
        //	//    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
        //	//    Call oMat01.LoadFromDataSource
        //	//    Call PS_CO600_FormClear
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_CO600_FlushToItemValue
        //private void PS_CO600_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;
        //	short ErrNum = 0;
        //	string sQry = null;
        //	string ItemCode = null;

        //	//Dim oRecordSet01 As SAPbobsCOM.Recordset
        //	//Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

        //	switch (oUID) {

        //		// Case "CntcCode01"

        //		//     If Trim(oForm01.Items("CntcCode01").Specific.VALUE) = "9999999" Then
        //		//         oForm01.Items("CntcName01").Specific.VALUE = "공용" '성명
        //		//     Else
        //		//         oForm01.Items("CntcName01").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode01").Specific.VALUE) & "'") '성명
        //		//     End If

        //		// Case "TeamCode01"

        //		//     If Trim(oForm01.Items("TeamCode01").Specific.VALUE) = oForm01.Items("BPLID01").Specific.Selected.VALUE & "999" Then
        //		//         oForm01.Items("TeamName01").Specific.VALUE = "전체공용"
        //		//     ElseIf Trim(oForm01.Items("TeamCode01").Specific.VALUE) = "Z" & oForm01.Items("BPLID01").Specific.Selected.VALUE & "99" Then
        //		//         oForm01.Items("TeamName01").Specific.VALUE = "사용부서없음"
        //		//     Else
        //		//         oForm01.Items("TeamName01").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & Trim(oForm01.Items("TeamCode01").Specific.VALUE) & "'", " AND Code = '1'") '팀
        //		//     End If

        //		// Case "CntcCode02"

        //		//     If Trim(oForm01.Items("CntcCode02").Specific.VALUE) = "9999999" Then
        //		//         oForm01.Items("CntcName02").Specific.VALUE = "공용" '성명
        //		//     Else
        //		//         oForm01.Items("CntcName02").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode02").Specific.VALUE) & "'") '성명
        //		//     End If

        //		//  Case "CntcCode03"

        //		//     If Trim(oForm01.Items("CntcCode03").Specific.VALUE) = "9999999" Then
        //		//         oForm01.Items("CntcName03").Specific.VALUE = "공용" '성명
        //		//     Else
        //		//         oForm01.Items("CntcName03").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode03").Specific.VALUE) & "'") '성명
        //		//     End If

        //		// Case "CntcCode05"

        //		//     If Trim(oForm01.Items("CntcCode05").Specific.VALUE) = "9999999" Then
        //		//         oForm01.Items("CntcName05").Specific.VALUE = "공용" '성명
        //		//     Else
        //		//         oForm01.Items("CntcName05").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode05").Specific.VALUE) & "'") '성명
        //		//     End If

        //	}

        //	//Set oRecordSet01 = Nothing

        //	return;
        //	PS_CO600_FlushToItemValue_Error:

        //	//Set oRecordSet01 = Nothing

        //	MDC_Com.MDC_GF_Message(ref "PS_CO600_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");

        //}
        #endregion

        #region PS_CO600_MTX01
        //private void PS_CO600_MTX01()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_MTX01()
        //	//해당모듈    : PS_CO600
        //	//기능        : 대차대조표(재무상태표)조회
        //	//인수        : 없음
        //	//반환값      : 없음
        //	//특이사항    : 없음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "G";
        //	//그리드출력

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		Query01 = "         EXEC PS_CO600_01 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		Query01 = "         EXEC PS_CO600_21 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	}

        //	oGrid01.DataTable.Clear();
        //	oDS_PS_CO600A.ExecuteQuery(Query01);
        //	//    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

        //	oGrid01.Columns.Item(2).RightJustified = true;
        //	oGrid01.Columns.Item(3).RightJustified = true;
        //	oGrid01.Columns.Item(4).RightJustified = true;
        //	oGrid01.Columns.Item(5).RightJustified = true;
        //	oGrid01.Columns.Item(6).RightJustified = true;
        //	//    oGrid01.Columns(19).RightJustified = True

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
        //		goto PS_CO600_MTX01_Exit;
        //	}

        //	oGrid01.AutoResizeColumns();
        //	oForm01.Update();

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX01_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX01_Error:

        //	oForm01.Freeze(false);

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_MTX02
        //private void PS_CO600_MTX02()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_MTX02()
        //	//해당모듈    : PS_CO600
        //	//기능        : 제조원가명세서 조회
        //	//인수        : 없음
        //	//반환값      : 없음
        //	//특이사항    : 없음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "G";
        //	//그리드출력

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		Query01 = "         EXEC PS_CO600_02 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		Query01 = "         EXEC PS_CO600_22 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	}

        //	oGrid02.DataTable.Clear();
        //	oDS_PS_CO600B.ExecuteQuery(Query01);
        //	//    oGrid03.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

        //	oGrid02.Columns.Item(2).RightJustified = true;
        //	oGrid02.Columns.Item(3).RightJustified = true;
        //	oGrid02.Columns.Item(4).RightJustified = true;
        //	oGrid02.Columns.Item(5).RightJustified = true;
        //	oGrid02.Columns.Item(6).RightJustified = true;
        //	//    oGrid02.Columns(18).RightJustified = True
        //	//
        //	//    oGrid02.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
        //	//    oGrid02.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
        //	//    oGrid02.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑

        //	//    oGrid02.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
        //	//    oGrid02.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
        //	//    oGrid02.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
        //	//    oGrid02.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
        //	//    oGrid02.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
        //	//    oGrid02.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

        //	if (oGrid02.Rows.Count == 0) {
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_CO600_MTX02_Exit;
        //	}

        //	oGrid02.AutoResizeColumns();
        //	oForm01.Update();

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX02_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX02_Error:

        //	oForm01.Freeze(false);

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_MTX03
        //private void PS_CO600_MTX03()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_MTX03()
        //	//해당모듈    : PS_CO600
        //	//기능        : 매출원가명세서 조회
        //	//인수        : 없음
        //	//반환값      : 없음
        //	//특이사항    : 없음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "G";
        //	//그리드출력

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		Query01 = "         EXEC PS_CO600_03 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		Query01 = "         EXEC PS_CO600_23 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	}

        //	oGrid03.DataTable.Clear();
        //	oDS_PS_CO600C.ExecuteQuery(Query01);
        //	//    oGrid03.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

        //	oGrid03.Columns.Item(2).RightJustified = true;
        //	oGrid03.Columns.Item(3).RightJustified = true;
        //	oGrid03.Columns.Item(4).RightJustified = true;
        //	oGrid03.Columns.Item(5).RightJustified = true;
        //	oGrid03.Columns.Item(6).RightJustified = true;
        //	//    oGrid03.Columns(18).RightJustified = True

        //	//    oGrid03.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
        //	//    oGrid03.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
        //	//    oGrid03.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑

        //	//    oGrid03.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
        //	//    oGrid03.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
        //	//    oGrid03.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
        //	//    oGrid03.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
        //	//    oGrid03.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
        //	//    oGrid03.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

        //	if (oGrid03.Rows.Count == 0) {
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_CO600_MTX03_Exit;
        //	}

        //	oGrid03.AutoResizeColumns();
        //	oForm01.Update();

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX03_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX03_Error:

        //	oForm01.Freeze(false);

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_MTX03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_MTX04
        //private void PS_CO600_MTX04()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_MTX04()
        //	//해당모듈    : PS_CO600
        //	//기능        : 손익계산서(포괄손익계산서)
        //	//인수        : 없음
        //	//반환값      : 없음
        //	//특이사항    : 없음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "G";
        //	//그리드출력

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		Query01 = "         EXEC PS_CO600_04 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		Query01 = "         EXEC PS_CO600_24 '";
        //		Query01 = Query01 + FrDt + "','";
        //		Query01 = Query01 + ToDt + "','";
        //		Query01 = Query01 + PrtCls + "'";

        //	}

        //	oGrid04.DataTable.Clear();
        //	oDS_PS_CO600D.ExecuteQuery(Query01);
        //	//    oGrid05.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

        //	oGrid04.Columns.Item(2).RightJustified = true;
        //	oGrid04.Columns.Item(3).RightJustified = true;
        //	oGrid04.Columns.Item(4).RightJustified = true;
        //	oGrid04.Columns.Item(5).RightJustified = true;
        //	oGrid04.Columns.Item(6).RightJustified = true;

        //	//    oGrid04.Columns(14).RightJustified = True
        //	//
        //	//    oGrid04.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
        //	//    oGrid04.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
        //	//    oGrid04.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑

        //	//    oGrid04.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
        //	//    oGrid04.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
        //	//    oGrid04.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
        //	//    oGrid04.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
        //	//    oGrid04.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
        //	//    oGrid04.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

        //	if (oGrid04.Rows.Count == 0) {
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_CO600_MTX04_Exit;
        //	}

        //	oGrid04.AutoResizeColumns();
        //	oForm01.Update();

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX04_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO600_MTX04_Error:

        //	oForm01.Freeze(false);

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_MTX04_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_DI_API
        //private bool PS_CO600_DI_API()
        //{
        //	//On Error GoTo PS_CO600_DI_API_Error
        //	//    PS_CO600_DI_API = True
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
        //	//    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
        //	//    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
        //	//    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "&&&&-&&-&&")
        //	//    For i = 0 To UBound(ItemInformation)
        //	//        If ItemInformation(i).Check = True Then
        //	//            GoTo Continue_First
        //	//        End If
        //	//        If i <> 0 Then
        //	//            oDIObject.Lines.Add
        //	//        End If
        //	//        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        //	//        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
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
        //	//            Call oDS_PS_CO600L.setValue("U_OPDNNo", i, ResultDocNum)
        //	//            Call oDS_PS_CO600L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        //	//        Next
        //	//    Else
        //	//        GoTo PS_CO600_DI_API_Error
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
        //	//PS_CO600_DI_API_DI_Error:
        //	//    If Sbo_Company.InTransaction = True Then
        //	//        Sbo_Company.EndTransaction wf_RollBack
        //	//    End If
        //	//    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
        //	//    PS_CO600_DI_API = False
        //	//    Set oDIObject = Nothing
        //	//    Exit Function
        //	//PS_CO600_DI_API_Error:
        //	//    Sbo_Application.SetStatusBarMessage "PS_CO600_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
        //	//    PS_CO600_DI_API = False
        //}
        #endregion

        #region PS_CO600_SetBaseForm
        //private void PS_CO600_SetBaseForm()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	string ItemCode01 = null;
        //	SAPbouiCOM.Matrix oBaseMat01 = null;
        //	if (oBaseForm01 == null) {
        //		////DoNothing
        //	} else {

        //	}
        //	return;
        //	PS_CO600_SetBaseForm_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_SetBaseForm_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_FormResize
        //private void PS_CO600_FormResize()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	//그룹박스 크기 동적 할당
        //	oForm01.Items.Item("GrpBox01").Height = oForm01.Items.Item("Grid01").Height + 60;
        //	oForm01.Items.Item("GrpBox01").Width = oForm01.Items.Item("Grid01").Width + 30;

        //	if (oGrid01.Columns.Count > 0) {
        //		oGrid01.AutoResizeColumns();
        //	}

        //	if (oGrid03.Columns.Count > 0) {
        //		oGrid03.AutoResizeColumns();
        //	}

        //	if (oGrid03.Columns.Count > 0) {
        //		oGrid03.AutoResizeColumns();
        //	}

        //	if (oGrid04.Columns.Count > 0) {
        //		oGrid04.AutoResizeColumns();
        //	}

        //	return;
        //	PS_CO600_FormResize_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO600_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO600_Print_Report01
        //private void PS_CO600_Print_Report01()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_Print_Report01()
        //	//해당모듈    : PS_CO600
        //	//기능        : 대차대조표(재무상태표) 출력
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

        //	SAPbobsCOM.Recordset oRecordSet = null;
        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "R";
        //	//리포트출력

        //	/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //	//쿼리
        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		ReportName = "PS_CO600_51.rpt";
        //		WinTitle = "[PS_CO600] 대차대조표";

        //		sQry = "      EXEC PS_CO600_01 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		ReportName = "PS_CO600_61.rpt";
        //		WinTitle = "[PS_CO600] 재무상태표";

        //		sQry = "      EXEC PS_CO600_21 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	}
        //	MDC_Globals.gRpt_Formula = new string[3];
        //	MDC_Globals.gRpt_Formula_Value = new string[3];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

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
        //		MDC_Com.MDC_GF_Message(ref "PS_CO600_Print_Report01_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion

        #region PS_CO600_Print_Report02
        //private void PS_CO600_Print_Report02()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_Print_Report02()
        //	//해당모듈    : PS_CO600
        //	//기능        : 제조원가명세서 출력
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

        //	SAPbobsCOM.Recordset oRecordSet = null;
        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "R";
        //	//리포트출력

        //	/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //	//쿼리
        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		ReportName = "PS_CO600_52.rpt";
        //		WinTitle = "[PS_CO600] 제조원가명세서";

        //		sQry = "      EXEC PS_CO600_02 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		ReportName = "PS_CO600_62.rpt";
        //		WinTitle = "[PS_CO600] 제조원가명세서";

        //		sQry = "      EXEC PS_CO600_22 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	}
        //	MDC_Globals.gRpt_Formula = new string[3];
        //	MDC_Globals.gRpt_Formula_Value = new string[3];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

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
        //		MDC_Com.MDC_GF_Message(ref "PS_CO600_Print_Report02_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion

        #region PS_CO600_Print_Report03
        //private void PS_CO600_Print_Report03()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_Print_Report03()
        //	//해당모듈    : PS_CO600
        //	//기능        : 매출원가명세서 출력
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

        //	SAPbobsCOM.Recordset oRecordSet = null;
        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "R";
        //	//리포트출력

        //	/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //	//쿼리
        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		ReportName = "PS_CO600_53.rpt";
        //		WinTitle = "[PS_CO600] 매출원가명세서";

        //		sQry = "      EXEC PS_CO600_03 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		ReportName = "PS_CO600_63.rpt";
        //		WinTitle = "[PS_CO600] 매출원가명세서";

        //		sQry = "      EXEC PS_CO600_23 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	}
        //	MDC_Globals.gRpt_Formula = new string[3];
        //	MDC_Globals.gRpt_Formula_Value = new string[3];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

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
        //		MDC_Com.MDC_GF_Message(ref "PS_CO600_Print_Report03_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion

        #region PS_CO600_Print_Report04
        //private void PS_CO600_Print_Report04()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO600_Print_Report04()
        //	//해당모듈    : PS_CO600
        //	//기능        : 손익계산서 출력
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

        //	SAPbobsCOM.Recordset oRecordSet = null;
        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();

        //	string FrDt = null;
        //	//조회기간(시작)
        //	string ToDt = null;
        //	//조회기간(종료)
        //	string Ctgr = null;
        //	//출력구분
        //	string PrtCls = null;
        //	//그리드, 리포트 출력구분

        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm01.Items.Item("FrDt01").Specific.VALUE);
        //	//조회기간(시작)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm01.Items.Item("ToDt01").Specific.VALUE);
        //	//조회기간(종료)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Ctgr = Strings.Trim(oForm01.Items.Item("Ctgr01").Specific.Selected.VALUE);
        //	//출력구분
        //	PrtCls = "R";
        //	//리포트출력

        //	/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //	//쿼리
        //	//K-GAAP
        //	if (Ctgr == "10") {

        //		ReportName = "PS_CO600_54.rpt";
        //		WinTitle = "[PS_CO600] 손익계산서";

        //		sQry = "      EXEC PS_CO600_04 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	//K-IFRS
        //	} else {

        //		ReportName = "PS_CO600_64.rpt";
        //		WinTitle = "[PS_CO600] 포괄손익계산서";

        //		sQry = "      EXEC PS_CO600_24 '";
        //		sQry = sQry + FrDt + "','";
        //		sQry = sQry + ToDt + "','";
        //		sQry = sQry + PrtCls + "'";

        //	}
        //	MDC_Globals.gRpt_Formula = new string[3];
        //	MDC_Globals.gRpt_Formula_Value = new string[3];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

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
        //		MDC_Com.MDC_GF_Message(ref "PS_CO600_Print_Report04_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion
    }
}
