using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작업일보등록(분말)
	/// </summary>
	internal class PS_PP049 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Matrix oMat03;
		private SAPbouiCOM.DBDataSource oDS_PS_PP049H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP049L; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP049M; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP049N; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMat01Row01;
		private int oMat02Row02;
		private int oMat03Row03;
		private string oDocType01;
		private string oDocEntry01;
		private string oOrdGbn;
		private string oSequence;
		private string oDocdate;
		private SAPbouiCOM.BoFormMode oFormMode01;

		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//private static extern int Beep_Renamed(int dwFreq, int dwDuration);
		//[DllImport("uio32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//private static extern int usb_io_init(int pID);
		//[DllImport("uio32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)] // VB6.0의 버그로 인해 Long 선언
		//private static extern bool usb_io_output(int pID, byte cmd, short io1, short io2, short io3, short io4);

		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
		
			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP049.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP049_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP049");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				//PS_PP049_CreateItems();
				//PS_PP049_SetComboBox();
				//PS_PP049_CF_ChooseFromList();
				//PS_PP049_EnableMenus();
				//PS_PP049_SetDocument(oFormDocEntry);
				//PS_PP049_ResizeForm();
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

        //			// 작업시간 합계 추가 S
        //			//            Dim i&
        //			//            Dim Total As Currency
        //			//
        //			//
        //			//                For i = 0 To oMat01.VisualRowCount - 1
        //			//
        //			//                    Total = Total + Val(oMat01.Columns("WorkTime").Cells(i + 1).Specific.VALUE)
        //			//'                 oMat01.Columns("Total").Cells.Specific.VALUE = Total
        //			//                Next i
        //			//                oForm.Items("Total").Specific.VALUE = Total
        //			PS_PP049_SumWorkTime();
        //			break;
        //		// 작업시간 합계 추가 E

        //		//            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)

        //		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //			////18
        //			usb_io_output(609, 0, -1, -2, 0, 0);
        //			// 부져음 및 경광등 끄기
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
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
        //				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					if ((PS_PP049_Validate("취소") == false)) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //					if (SubMain.Sbo_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1")) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					// 분말 첫번째 공정 투입시 원자재 불출로직 추가(황영수 20181101)
        //					if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "111" | Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "601") {
        //						if (Add_oInventoryGenEntry(ref 2) == false) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //					}
        //				} else {
        //					MDC_Com.MDC_GF_Message(ref "현재 모드에서는 취소할수 없습니다.", ref "W");
        //					BubbleEvent = false;
        //					return;
        //				}
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
        //				Raise_EVENT_RECORD_MOVE(ref FormUID, ref pval, ref BubbleEvent);
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
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_PP049_EnableFormItem();
        //				////UDO방식
        //				oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			case "1282":
        //				//추가
        //				PS_PP049_EnableFormItem();
        //				////UDO방식
        //				PS_PP049_AddMatrixRow01(0, ref true);
        //				////UDO방식
        //				PS_PP049_AddMatrixRow02(0, ref true);
        //				////UDO방식
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				Raise_EVENT_RECORD_MOVE(ref FormUID, ref pval, ref BubbleEvent);
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
        //				if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //					if ((PS_PP049_FindValidateDocument("@PS_PP040H") == false)) {
        //						////찾기메뉴 활성화일때 수행
        //						if (SubMain.Sbo_Application.Menus.Item("1281").Enabled == true) {
        //							SubMain.Sbo_Application.ActivateMenuItem(("1281"));
        //						} else {
        //							SubMain.Sbo_Application.SetStatusBarMessage("관리자에게 문의바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //						}
        //						BubbleEvent = false;
        //						return;
        //					}
        //				}
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
        //	if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02" | pval.ItemUID == "Mat03") {
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
        //	if (pval.ItemUID == "Mat01") {
        //		if (pval.Row > 0) {
        //			oMat01Row01 = pval.Row;
        //		}
        //	} else if (pval.ItemUID == "Mat02") {
        //		if (pval.Row > 0) {
        //			oMat02Row02 = pval.Row;
        //		}
        //	} else if (pval.ItemUID == "Mat03") {
        //		if (pval.Row > 0) {
        //			oMat03Row03 = pval.Row;
        //		}
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

        //	string vReturnValue = null;

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("처리 중..", 100, false);

        //	string DocEntry = null;
        //	string LineNum = null;
        //	int i = 0;
        //	int ErrNum = 0;
        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string[] oText = new string[2];
        //	string sQry = null;
        //	string sQryS = null;
        //	string sQry1 = null;
        //	string WorkName = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "PS_PP049") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pval.ItemUID == "1") {

        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_PP049_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}

        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				// 분말 첫번째 공정 투입시 원자재 불출로직 추가(황영수 20181101)
        //				if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "111" | Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "601") {
        //					if (Add_oInventoryGenExit(ref 2) == false) {
        //						BubbleEvent = false;
        //						return;
        //					} else {
        //					}
        //					// End If
        //				}
        //				usb_io_output(609, 0, -1, -2, 0, 0);
        //				// 부져음 및 경광등 끄기

        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocEntry01 = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oOrdGbn = Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE);
        //				////작업구분
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oSequence = oMat01.Columns.Item("Sequence").Cells.Item(1).Specific.VALUE;
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocdate = Strings.Trim(oForm.Items.Item("DocDate").Specific.VALUE);
        //				oFormMode01 = oForm.Mode;
        //				////해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (PS_PP049_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocEntry01 = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
        //				oFormMode01 = oForm.Mode;
        //				////해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}

        //		////취소버튼 누를시 저장할 자료가 있으면 메시지 표시
        //		if (pval.ItemUID == "2") {
        //			usb_io_output(609, 0, -1, -2, 0, 0);
        //			// 부져음 및 경광등 끄기
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (oMat01.VisualRowCount > 1) {
        //					vReturnValue = Convert.ToString(SubMain.Sbo_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소"));
        //					switch (vReturnValue) {
        //						case Convert.ToString(1):
        //							break;
        //						case Convert.ToString(2):
        //							BubbleEvent = false;
        //							return;

        //							break;
        //					}
        //				}
        //			}
        //		}

        //		if (pval.ItemUID == "Button01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_PP049_LoadOrderInfo();
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				PS_PP049_LoadOrderInfo();
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pval.ItemUID == "Button02") {

        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			MDC_PS_Common.ConnectODBC();

        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
        //			for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //				if (oMat01.IsRowSelected(i + 1) == true) {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					LineNum = oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.VALUE;
        //				}
        //			}

        //			WinTitle = " 공정카드 [PS_PP040]";
        //			ReportName = "PS_PP049_01.rpt";

        //			sQry1 = "Select U_WorkName From [@PS_PP040M] Where DocEntry = '" + DocEntry + "' And IsNull(U_WorkName, '') <> ''";
        //			oRecordSet01.DoQuery(sQry1);

        //			while (!(oRecordSet01.EoF)) {
        //				WorkName = WorkName + "     " + oRecordSet01.Fields.Item(0).Value;
        //				oRecordSet01.MoveNext();
        //			}
        //			MDC_Globals.gRpt_Formula = new string[2];
        //			MDC_Globals.gRpt_Formula_Value = new string[2];

        //			////Formula 수식필드

        //			oText[1] = WorkName;

        //			for (i = 1; i <= 1; i++) {
        //				if (Strings.Len("" + i + "") == 1) {
        //					MDC_Globals.gRpt_Formula[i] = "F0" + i + "";
        //				} else {
        //					MDC_Globals.gRpt_Formula[i] = "F" + i + "";
        //				}
        //				MDC_Globals.gRpt_Formula_Value[i] = oText[i];
        //			}
        //			MDC_Globals.gRpt_SRptSqry = new string[2];
        //			MDC_Globals.gRpt_SRptName = new string[2];
        //			MDC_Globals.gRpt_SFormula = new string[2, 2];
        //			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //			////SubReport

        //			MDC_Globals.gRpt_SFormula[1, 1] = "";
        //			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

        //			sQryS = "EXEC [PS_PP049_06] '" + DocEntry + "', '" + LineNum + "', 'S'";

        //			MDC_Globals.gRpt_SRptSqry[1] = sQryS;
        //			MDC_Globals.gRpt_SRptName[1] = "PS_PP049_S1";

        //			////조회조건문
        //			sQry = "EXEC [PS_PP049_06] '" + DocEntry + "', '" + LineNum + "', 'M'";
        //			oRecordSet01.DoQuery(sQry);
        //			if (oRecordSet01.RecordCount == 0) {
        //				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다.확인해 주세요.", ref "E");
        //				return;
        //			}

        //			////CR Action
        //			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "N", "V") == false) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			}
        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;
        //		}
        //	} else if (pval.BeforeAction == false) {
        //		if (pval.ItemUID == "PS_PP049") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pval.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pval.ActionSuccess == true) {
        //					if (oOrdGbn == "101" & oSequence == "1") {
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //						PS_PP049_EnableFormItem();
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("DocEntry").Specific.VALUE = oDocEntry01;
        //						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					} else {
        //						PS_PP049_EnableFormItem();
        //						PS_PP049_AddMatrixRow01(0, ref true);
        //						////UDO방식일때
        //						PS_PP049_AddMatrixRow02(0, ref true);
        //						////UDO방식일때
        //					}
        //					//
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				if (pval.ActionSuccess == true) {
        //					if ((oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)) {
        //						oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //						PS_PP049_EnableFormItem();
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("DocEntry").Specific.VALUE = oDocEntry01;
        //						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					}
        //					PS_PP049_EnableFormItem();
        //				}
        //			}
        //		}
        //		if (pval.ItemUID == "Button01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:


        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "OrdMgNum") {
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////작업타입이 일반,조정일때
        //			if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "60") {
        //				MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pval, ref BubbleEvent, "OrdMgNum", "");
        //				////사용자값활성
        //			}
        //		}





        //		if (pval.ItemUID == "Mat01") {
        //			if (pval.ColUID == "OrdMgNum") {
        //				//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////일반,조정, 설계
        //				if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "60" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "70") {
        //					//UPGRADE_WARNING: oForm.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("OrdGbn").Specific.Selected.VALUE == "선택") {
        //						MDC_Com.MDC_GF_Message(ref "작업구분이 선택되지 않았습니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (oForm.Items.Item("BPLId").Specific.Selected.VALUE == "선택") {
        //						MDC_Com.MDC_GF_Message(ref "사업장이 선택되지 않았습니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //						//UPGRADE_WARNING: oForm.Items(ItemCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.VALUE)) {
        //						MDC_Com.MDC_GF_Message(ref "품목코드가 선택되지 않았습니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //						//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.VALUE)) {
        //						MDC_Com.MDC_GF_Message(ref "작지번호가 선택되지 않았습니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //						//UPGRADE_WARNING: oForm.Items(PP030HNo).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("PP030HNo").Specific.VALUE)) {
        //						MDC_Com.MDC_GF_Message(ref "작지문서번호가 선택되지 않았습니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //					} else {
        //						MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pval, ref BubbleEvent, "Mat01", "OrdMgNum");
        //						////사용자값활성
        //					}
        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////지원
        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {
        //					//UPGRADE_WARNING: oForm.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("OrdGbn").Specific.Selected.VALUE == "선택") {
        //						MDC_Com.MDC_GF_Message(ref "작업구분이 선택되지 않았습니다.", ref "W");
        //						oForm.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						BubbleEvent = false;
        //						return;
        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (oForm.Items.Item("BPLId").Specific.Selected.VALUE == "선택") {
        //						MDC_Com.MDC_GF_Message(ref "사업장이 선택되지 않았습니다.", ref "W");
        //						oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						BubbleEvent = false;
        //						return;
        //						//                    ElseIf oForm.Items("ItemCode").Specific.Value = "" Then
        //						//                        Call MDC_Com.MDC_GF_Message("품목코드가 선택되지 않았습니다.", "W")
        //						//                        oForm.Items("ItemCode").Click ct_Regular
        //						//                        BubbleEvent = False
        //						//                        Exit Sub
        //						//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.VALUE)) {
        //						MDC_Com.MDC_GF_Message(ref "작지번호가 선택되지 않았습니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //					} else {
        //						MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pval, ref BubbleEvent, "Mat01", "OrdMgNum");
        //						////사용자값활성
        //					}
        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////외주
        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {

        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////실적
        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {

        //				}

        //			}
        //		}
        //		if (pval.ItemUID == "Mat02") {
        //			if (pval.ColUID == "WorkCode") {
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE) == 0) {
        //					MDC_Com.MDC_GF_Message(ref "기준시간을 입력하지 않았습니다.", ref "W");
        //					oForm.Items.Item("BaseTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //		}
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "Mat02", "WorkCode");
        //		//사용자값활성
        //		MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pval, ref BubbleEvent, "Mat02", "NCode");
        //		//사용자값활성
        //		MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pval, ref BubbleEvent, "Mat03", "FailCode");
        //		//사용자값활성

        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "Mat01", "MachCode");
        //		//설비코드 사용자값활성
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "SubLot") 'sub작지번호 사용자값활성
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "Mat01", "CItemCod");
        //		//원재료코드 사용자값활성
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "Mat01", "SCpCode");
        //		//지원공정추가(2018.05.30 송명규)
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "UseMCode", "");
        //		//작업장비 사용자값활성
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "ItemCode") '사용자값활성
        //	} else if (pval.BeforeAction == false) {
        //		//// 화살표 이동 강제 코딩 - 류영조
        //		if (pval.ItemUID == "Mat01") {
        //			////위쪽 화살표
        //			if (pval.CharPressed == 38) {
        //				if (pval.Row > 1 & pval.Row <= oMat01.VisualRowCount) {
        //					oForm.Freeze(true);
        //					oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					oForm.Freeze(false);
        //				}
        //			////아래 화살표
        //			} else if (pval.CharPressed == 40) {
        //				if (pval.Row > 0 & pval.Row < oMat01.VisualRowCount) {
        //					oForm.Freeze(true);
        //					oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					oForm.Freeze(false);
        //				}
        //			}

        //			//작업시간 입력 시마다 합계 계산(2011.09.26 송명규 추가)
        //			if (pval.ColUID == "WorkTime" & pval.Row != Convert.ToDouble("0")) {

        //				PS_PP049_SumWorkTime();

        //			}

        //		} else if (pval.ItemUID == "BaseTime") {

        //			//탭 키 Press
        //			if (pval.CharPressed == 9) {

        //				oMat02.Columns.Item("WorkCode").Cells.Item(1).Click();

        //			}

        //		}
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

        //	oForm.Freeze(true);
        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		if (pval.ItemChanged == true) {
        //			oForm.Freeze(true);
        //			if ((pval.ItemUID == "Mat01")) {
        //				if ((pval.ColUID == "특정컬럼")) {
        //					////기타작업
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.VALUE);
        //					if (oMat01.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP049L.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
        //						//PS_PP049_AddMatrixRow (pval.Row)
        //					}
        //				} else {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.VALUE);
        //				}
        //			} else if ((pval.ItemUID == "Mat02")) {
        //				if ((pval.ColUID == "특정컬럼")) {
        //					////기타작업
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.VALUE);
        //					if (oMat02.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP049M.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
        //						//PS_PP049_AddMatrixRow (pval.Row)
        //					}
        //				} else {
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.VALUE);
        //				}
        //			} else if ((pval.ItemUID == "Mat03")) {
        //				if ((pval.ColUID == "특정컬럼")) {
        //				} else {
        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049N.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.VALUE);
        //				}
        //			} else {
        //				if ((pval.ItemUID == "OrdType")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE);
        //					//UPGRADE_WARNING: oForm.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					////일반,조정,설계
        //					if (oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "10" | oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "50" | oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "60" | oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "70") {
        //						//창원은 품목구분 선택하도록 수정 '2015/04/09
        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (oForm.Items.Item("BPLId").Specific.VALUE == "1") {
        //							oForm.Items.Item("OrdGbn").Enabled = true;
        //						} else {
        //							oForm.Items.Item("OrdGbn").Enabled = false;
        //						}
        //						oForm.Items.Item("BPLId").Enabled = false;
        //						oForm.Items.Item("ItemCode").Enabled = false;
        //						//UPGRADE_WARNING: oForm.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "20") {
        //						oForm.Items.Item("OrdGbn").Enabled = true;
        //						oForm.Items.Item("BPLId").Enabled = true;
        //						oForm.Items.Item("ItemCode").Enabled = true;
        //						//UPGRADE_WARNING: oForm.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "30") {
        //						oForm.Items.Item("OrdGbn").Enabled = false;
        //						oForm.Items.Item("BPLId").Enabled = false;
        //						oForm.Items.Item("ItemCode").Enabled = false;
        //						//UPGRADE_WARNING: oForm.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE == "40") {
        //						oForm.Items.Item("OrdGbn").Enabled = false;
        //						oForm.Items.Item("BPLId").Enabled = false;
        //						oForm.Items.Item("ItemCode").Enabled = false;
        //					}

        //					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("OrdMgNum").Specific.VALUE = "";
        //					//Call oForm.Items("BPLId").Specific.Select(0, psk_Index)
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("ItemCode").Specific.VALUE = "";
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("ItemName").Specific.VALUE = "";
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("OrdNum").Specific.VALUE = "";
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("OrdSub1").Specific.VALUE = "";
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("OrdSub2").Specific.VALUE = "";
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("PP030HNo").Specific.VALUE = "";
        //					oMat01.Clear();
        //					oMat01.FlushToDataSource();
        //					oMat01.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow01(0, ref true);
        //					oMat02.Clear();
        //					oMat02.FlushToDataSource();
        //					oMat02.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow02(0, ref true);
        //					oMat03.Clear();
        //					oMat03.FlushToDataSource();
        //					oMat03.LoadFromDataSource();
        //				} else if ((pval.ItemUID == "OrdGbn")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE);
        //					oMat01.Clear();
        //					oMat01.FlushToDataSource();
        //					oMat01.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow01(0, ref true);
        //					oMat02.Clear();
        //					oMat02.FlushToDataSource();
        //					oMat02.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow02(0, ref true);
        //					oMat03.Clear();
        //					oMat03.FlushToDataSource();
        //					oMat03.LoadFromDataSource();
        //				} else if ((pval.ItemUID == "BPLId")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE);
        //					oMat01.Clear();
        //					oMat01.FlushToDataSource();
        //					oMat01.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow01(0, ref true);
        //					oMat02.Clear();
        //					oMat02.FlushToDataSource();
        //					oMat02.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow02(0, ref true);
        //					oMat03.Clear();
        //					oMat03.FlushToDataSource();
        //					oMat03.LoadFromDataSource();
        //				} else {
        //					//거래처구분이 아닐 경우만 실행(2012.02.02 송명규 추가)
        //					if (pval.ItemUID != "CardType") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.Selected.VALUE);
        //					}
        //				}
        //			}
        //			oMat01.LoadFromDataSource();
        //			oMat01.AutoResizeColumns();
        //			oMat02.LoadFromDataSource();
        //			oMat02.AutoResizeColumns();
        //			oMat03.LoadFromDataSource();
        //			oMat03.AutoResizeColumns();
        //			oForm.Update();
        //			if (pval.ItemUID == "Mat01") {
        //				oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //			} else if (pval.ItemUID == "Mat02") {
        //				oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //			} else if (pval.ItemUID == "Mat03") {
        //				oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //			} else {

        //			}
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
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object TempForm01 = null;

        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "Opt01") {
        //			oForm.Freeze(true);
        //			oForm.Settings.MatrixUID = "Mat02";
        //			oForm.Settings.EnableRowFormat = true;
        //			oForm.Settings.Enabled = true;
        //			oMat01.AutoResizeColumns();
        //			oMat02.AutoResizeColumns();
        //			oMat03.AutoResizeColumns();
        //			oForm.Freeze(false);
        //		}
        //		if (pval.ItemUID == "Opt02") {
        //			oForm.Freeze(true);
        //			oForm.Settings.MatrixUID = "Mat03";
        //			oForm.Settings.EnableRowFormat = true;
        //			oForm.Settings.Enabled = true;
        //			oMat01.AutoResizeColumns();
        //			oMat02.AutoResizeColumns();
        //			oMat03.AutoResizeColumns();
        //			oForm.Freeze(false);
        //		}
        //		if (pval.ItemUID == "Opt03") {
        //			oForm.Freeze(true);
        //			oForm.Settings.MatrixUID = "Mat01";
        //			oForm.Settings.EnableRowFormat = true;
        //			oForm.Settings.Enabled = true;
        //			oMat01.AutoResizeColumns();
        //			oMat02.AutoResizeColumns();
        //			oMat03.AutoResizeColumns();
        //			oForm.Freeze(false);
        //		}
        //		//        If pval.ItemUID = "Mat01" Then
        //		//            If pval.Row > 0 Then
        //		//                Call oMat01.SelectRow(pval.Row, True, False)
        //		//            End If
        //		//        End If
        //		//        If pval.ItemUID = "Mat01" Then
        //		//            If pval.Row > 0 Then
        //		//                If pval.ColUID = "BatchNum" Then
        //		//                    If oMat01.Columns("BatchNum").Cells(pval.Row).Specific.VALUE <> "" And oMat01.Columns("CItemCod").Cells(pval.Row).Specific.VALUE <> "" Then
        //		//                        oMat01.Columns("BatchNum").Cells(pval.Row).Click ct_Regular
        //		//                    End If
        //		//                End If
        //		//
        //		//                Call oMat01.SelectRow(pval.Row, True, False)
        //		//                oMat01Row01 = pval.Row
        //		//            End If
        //		//        End If
        //		if (pval.ItemUID == "Mat02") {
        //			if (pval.Row > 0) {
        //				oMat02.SelectRow(pval.Row, true, false);
        //				oMat02Row02 = pval.Row;
        //			}
        //		}
        //		if (pval.ItemUID == "Mat03") {
        //			if (pval.Row > 0) {
        //				oMat03.SelectRow(pval.Row, true, false);
        //				oMat03Row03 = pval.Row;
        //			}
        //		}
        //	} else if (pval.BeforeAction == false) {
        //		//// 작업지시번호 링크 번튼 - 류영조
        //		if (pval.ItemUID == "LBtn01") {
        //			TempForm01 = new PS_PP030();
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: TempForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TempForm01.LoadForm(oForm.Items.Item("PP030HNo").Specific.VALUE);
        //			//UPGRADE_NOTE: TempForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			TempForm01 = null;
        //		}
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
        //		if (pval.ItemUID == "Mat01") {
        //			if (pval.Row > 0) {
        //				//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////작업타입이 일반,조정인경우
        //				if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "60") {
        //					//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.VALUE)) {

        //					} else {
        //						if (oMat03.VisualRowCount == 0) {
        //							PS_PP049_AddMatrixRow03(0, ref true);
        //						} else {
        //							PS_PP049_AddMatrixRow03(oMat03.VisualRowCount);
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pval.Row).Specific.VALUE);
        //						oDS_PS_PP049N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pval.Row));
        //						oMat03.LoadFromDataSource();
        //						oMat03.AutoResizeColumns();
        //						//                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
        //						//                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
        //						oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
        //						oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
        //						oMat03.FlushToDataSource();
        //					}
        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////작업타입이 PSMT지원인경우
        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {
        //					//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.VALUE)) {

        //					} else {
        //						if (oMat03.VisualRowCount == 0) {
        //							PS_PP049_AddMatrixRow03(0, ref true);
        //						} else {
        //							PS_PP049_AddMatrixRow03(oMat03.VisualRowCount);
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pval.Row).Specific.VALUE);
        //						oDS_PS_PP049N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pval.Row));
        //						oMat03.LoadFromDataSource();
        //						oMat03.AutoResizeColumns();
        //						//                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
        //						//                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
        //						oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
        //						oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
        //						oMat03.FlushToDataSource();
        //					}
        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////작업타입이 외주인경우
        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {
        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////작업타입이 실적인경우
        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {
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

        //	object oTempClass = null;
        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "Mat01") {
        //			if (pval.ColUID == "OrdMgNum") {
        //				oTempClass = new PS_PP030();
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oTempClass.LoadForm(Strings.Mid(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE, 1, Strings.InStr(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE, "-") - 1));
        //			}
        //			if (pval.ColUID == "PP030HNo") {
        //				oTempClass = new PS_PP030();
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oTempClass.LoadForm(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //			}
        //		}
        //		if (pval.ItemUID == "Mat03") {
        //			if (pval.ColUID == "OrdMgNum") {
        //				oTempClass = new PS_PP030();
        //				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oTempClass.LoadForm(Strings.Mid(oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE, 1, Strings.InStr(oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE, "-") - 1));
        //			}
        //		}
        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE(
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	double Weight = 0;

        //	double Time = 0;
        //	//UPGRADE_NOTE: Hour이(가) Hour_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //	int Hour_Renamed = 0;
        //	//UPGRADE_NOTE: Minute이(가) Minute_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //	int Minute_Renamed = 0;
        //	string ReturnValue = null;
        //	string BatchNumErr = null;
        //	bool result = false;

        //	oForm.Freeze(true);
        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemChanged == true) {
        //			if ((pval.ItemUID == "Mat01")) {
        //				if ((PS_PP049_Validate("수정01") == false)) {
        //					oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, Strings.Trim(oDS_PS_PP049L.GetValue("U_" + pval.ColUID, pval.Row - 1)));
        //				} else {
        //					if ((pval.ColUID == "OrdMgNum")) {
        //						RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //						ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("실행 중...", 100, false);
        //						//
        //						//                        If oForm.Items("OrdNum").Specific.VALUE = "" Then '//작지번호에 값이 없으면 작업지시가 불러오기전
        //						//                            Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //						//                        Else '//작업지시가 선택된상태
        //						//                            If oForm.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm.Items("OrdType").Specific.Selected.VALUE = "60" Or oForm.Items("OrdType").Specific.Selected.VALUE = "70" Then '//작업타입이 일반,조정, 설계
        //						//                                '//작지문서헤더번호가 일치하지 않으면
        //						//                                If oForm.Items("PP030HNo").Specific.VALUE <> Mid(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1) Then
        //						//                                    Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //						//                                Else '//작지문서번호가 일치하면
        //						//                                    If oForm.Items("BPLId").Specific.Selected.VALUE <> "1" Then
        //						//                                        '//신동사업부를 제외한 사업부만 체크
        //						//                                        For i = 1 To oMat01.RowCount
        //						//                                            '//현재 입력한 값이 이미 입력되어 있는경우
        //						//                                            If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE And i <> pval.Row Then
        //						//                                                Call MDC_Com.MDC_GF_Message("이미 입력한 공정입니다.", "W")
        //						//                                                Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //						//                                                GoTo Continue
        //						//                                            End If
        //						//    '                                        '//공정라인의 공정순서가 앞공정보다 높으면
        //						//    '                                        If Val(oMat01.Columns("Sequence").Cells(i).Specific.Value) >= MDC_PS_Common.GetValue("SELECT PS_PP030M.U_Sequence FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE CONVERT(NVARCHAR,PS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" & oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.Value & "'") Then
        //						//    '                                            Call MDC_Com.MDC_GF_Message("공정순서가 올바르지 않습니다.", "W")
        //						//    '                                            Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //						//    '                                            GoTo Continue
        //						//    '                                        End If
        //						//                                        Next
        //						//
        //						//
        //						//
        //						//
        //						//
        //						//                                    End If
        //						//
        //						//                                    Query01 = "EXEC PS_PP049_01 '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "', '" & oForm.Items("OrdType").Specific.Selected.VALUE & "'"
        //						//                                    RecordSet01.DoQuery Query01
        //						//                                    If RecordSet01.RecordCount = 0 Then
        //						//                                        Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //						//                                    Else
        //						//'                                        Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, RecordSet01.Fields("OrdMgNum").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_Sequence", pval.Row - 1, RecordSet01.Fields("Sequence").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_CpCode", pval.Row - 1, RecordSet01.Fields("CpCode").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_CpName", pval.Row - 1, RecordSet01.Fields("CpName").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_OrdGbn", pval.Row - 1, RecordSet01.Fields("OrdGbn").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_BPLId", pval.Row - 1, RecordSet01.Fields("BPLId").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_OrdNum", pval.Row - 1, RecordSet01.Fields("OrdNum").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_OrdSub1", pval.Row - 1, RecordSet01.Fields("OrdSub1").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_OrdSub2", pval.Row - 1, RecordSet01.Fields("OrdSub2").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_PP030HNo", pval.Row - 1, RecordSet01.Fields("PP030HNo").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_PP030MNo", pval.Row - 1, RecordSet01.Fields("PP030MNo").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_SelWt", pval.Row - 1, RecordSet01.Fields("SelWt").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_PSum", pval.Row - 1, RecordSet01.Fields("PSum").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_BQty", pval.Row - 1, RecordSet01.Fields("BQty").VALUE)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_PQty", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_PWeight", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_YQty", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_YWeight", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_NQty", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_NWeight", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_ScrapWt", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_WorkTime", pval.Row - 1, 0)
        //						//'                                        Call oDS_PS_PP049L.setValue("U_LineId", pval.Row - 1, "")
        //						//
        //						//                                        '//설비코드,명 Reset
        //						//                                        Call oDS_PS_PP049L.setValue("U_MachCode", pval.Row - 1, "")
        //						//                                        Call oDS_PS_PP049L.setValue("U_MachName", pval.Row - 1, "")
        //						//                                        '//불량코드테이블
        //						//                                        If oMat03.VisualRowCount = 0 Then
        //						//                                            Call PS_PP049_AddMatrixRow03(0, True)
        //						//                                        Else
        //						//                                            Call PS_PP049_AddMatrixRow03(oMat03.VisualRowCount)
        //						//                                        End If
        //						//
        //						//                                        Call oDS_PS_PP049N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields("OrdMgNum").VALUE)
        //						//                                        Call oDS_PS_PP049N.setValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields("CpCode").VALUE)
        //						//                                        Call oDS_PS_PP049N.setValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields("CpName").VALUE)
        //						//                                        Call oDS_PS_PP049N.setValue("U_OLineNum", oMat03.VisualRowCount - 1, pval.Row)
        //						//
        //						//
        //						//
        //						//                                        '// 류영조
        //						//                                        If oForm.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm.Items("OrdType").Specific.Selected.VALUE = "60" Then
        //						//                                            Call oDS_PS_PP049H.setValue("U_BaseTime", 0, "1")
        //						//                                            oMat02.Columns("WorkCode").Cells(1).Specific.VALUE = "9999999"
        //						//'                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = "조정"
        //						//'                                            Call oDS_PS_PP049M.setValue("U_WorkCode", 0, "9999999")
        //						//                                            Call oDS_PS_PP049M.setValue("U_WorkName", 0, "조정")
        //						//                                            oMat02.LoadFromDataSource
        //						//                                        Else
        //						//'                                            Call oDS_PS_PP049H.setValue("U_BaseTime", 0, "")
        //						//'                                            oMat02.Columns("WorkCode").Cells(1).Specific.Value = ""
        //						//'                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = ""
        //						//                    '                        Call oDS_PS_PP049M.setValue("U_WorkCode", 0, "")
        //						//                    '                        Call oDS_PS_PP049M.setValue("U_WorkName", 0, "")
        //						//                                        End If
        //						//                                    End If
        //						//                                End If
        //						//                            ElseIf oForm.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원
        //						//                                '//올바른 공정코드인지 검사
        //						//                                If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'") = 0 Then
        //						//                                    Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //						//                                Else
        //						//                                    For i = 1 To oMat01.RowCount
        //						//                                        '//현재 입력한 값이 이미 입력되어 있는경우
        //						//                                        If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE And i <> pval.Row Then
        //						//                                            Call MDC_Com.MDC_GF_Message("이미 입력한 공정입니다.", "W")
        //						//                                            Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //						//                                            GoTo Continue
        //						//                                        End If
        //						//                                    Next
        //						//                                    Call oDS_PS_PP049L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
        //						//                                    Call oDS_PS_PP049L.setValue("U_CpCode", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
        //						//                                    Call oDS_PS_PP049L.setValue("U_CpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
        //						//                                    Call oDS_PS_PP049L.setValue("U_OrdGbn", pval.Row - 1, oForm.Items("OrdGbn").Specific.Selected.VALUE)
        //						//                                    Call oDS_PS_PP049L.setValue("U_BPLId", pval.Row - 1, oForm.Items("BPLId").Specific.Selected.VALUE)
        //						//                                    Call oDS_PS_PP049L.setValue("U_ItemCode", pval.Row - 1, "")
        //						//                                    Call oDS_PS_PP049L.setValue("U_ItemName", pval.Row - 1, "")
        //						//                                    '//PSMT지원은 품목코드 필요없음
        //						//'                                    Call oDS_PS_PP049L.setValue("U_ItemCode", pval.Row - 1, oForm.Items("ItemCode").Specific.Value)
        //						//'                                    Call oDS_PS_PP049L.setValue("U_ItemName", pval.Row - 1, oForm.Items("ItemName").Specific.Value)
        //						//                                    Call oDS_PS_PP049L.setValue("U_OrdNum", pval.Row - 1, oForm.Items("OrdNum").Specific.VALUE)
        //						//                                    Call oDS_PS_PP049L.setValue("U_OrdSub1", pval.Row - 1, oForm.Items("OrdSub1").Specific.VALUE)
        //						//                                    Call oDS_PS_PP049L.setValue("U_OrdSub2", pval.Row - 1, oForm.Items("OrdSub2").Specific.VALUE)
        //						//                                    Call oDS_PS_PP049L.setValue("U_PP030HNo", pval.Row - 1, "")
        //						//                                    Call oDS_PS_PP049L.setValue("U_PP030MNo", pval.Row - 1, "")
        //						//                                    Call oDS_PS_PP049L.setValue("U_PSum", pval.Row - 1, 0)
        //						//                                    Call oDS_PS_PP049L.setValue("U_PQty", pval.Row - 1, 0)
        //						//                                    Call oDS_PS_PP049L.setValue("U_PWeight", pval.Row - 1, 0)
        //						//                                    Call oDS_PS_PP049L.setValue("U_YQty", pval.Row - 1, 0)
        //						//                                    Call oDS_PS_PP049L.setValue("U_YWeight", pval.Row - 1, 0)
        //						//                                    Call oDS_PS_PP049L.setValue("U_NQty", pval.Row - 1, 0)
        //						//                                    Call oDS_PS_PP049L.setValue("U_NWeight", pval.Row - 1, 0)
        //						//                                    Call oDS_PS_PP049L.setValue("U_ScrapWt", pval.Row - 1, 0)
        //						//                                    '//불량코드테이블
        //						//                                    If oMat03.VisualRowCount = 0 Then
        //						//                                        Call PS_PP049_AddMatrixRow03(0, True)
        //						//                                    Else
        //						//                                        If oDS_PS_PP049L.GetValue("U_OrdMgNum", pval.Row - 1) = oDS_PS_PP049N.GetValue("U_OrdMgNum", oMat03.VisualRowCount - 1) Then
        //						//                                        Else
        //						//                                            Call PS_PP049_AddMatrixRow03(oMat03.VisualRowCount)
        //						//                                        End If
        //						//                                    End If
        //						//                                    Call oDS_PS_PP049N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
        //						//                                    Call oDS_PS_PP049N.setValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
        //						//                                    Call oDS_PS_PP049N.setValue("U_CpName", oMat03.VisualRowCount - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
        //						//                                End If
        //						//                            ElseIf oForm.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주
        //						//
        //						//                            ElseIf oForm.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적
        //						//
        //						//                            End If
        //						//Continue:
        //						//                            If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP049L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
        //						//                                Call PS_PP049_AddMatrixRow01(pval.Row)
        //						//                            End If
        //						//                        End If
        //						//
        //						//                        ProgBar01.VALUE = 100
        //						//                        ProgBar01.Stop
        //						//                        Set ProgBar01 = Nothing
        //						//
        //						//                        Set RecordSet01 = Nothing
        //					} else if (pval.ColUID == "BatchNum") {
        //						usb_io_output(609, 0, -1, -2, 0, 0);
        //						// 부져음 및 경광등 끄기
        //						RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //						Query01 = "EXEC PS_PP049_99 '";
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						Query01 = Query01 + oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row).Specific.VALUE + "','";
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						Query01 = Query01 + oForm.Items.Item("ItemCode").Specific.VALUE + "'";


        //						RecordSet01.DoQuery(Query01);

        //						//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						ReturnValue = RecordSet01.Fields.Item("ReturnValue").Value;
        //						//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						BatchNumErr = RecordSet01.Fields.Item("BatchNum").Value;

        //						if (Strings.Trim(ReturnValue) == "O") {
        //							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_OrdMgNum", pval.Row - 1, Convert.ToString(oForm.Items.Item("PP030HNo").Specific.VALUE + "-1"));
        //							oDS_PS_PP049L.SetValue("U_CItemCod", pval.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
        //							oDS_PS_PP049L.SetValue("U_CItemNam", pval.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
        //							oDS_PS_PP049L.SetValue("U_OnHandWt", pval.Row - 1, RecordSet01.Fields.Item("OnHandWt").Value);
        //						} else if (ReturnValue == "E1") {
        //							usb_io_output(609, 34, 1, 2, 0, 0);
        //							goto Raise_EVENT_VALIDATE_Error;
        //						} else if (ReturnValue == "E2") {
        //							usb_io_output(609, 34, 1, 2, 0, 0);
        //							goto Raise_EVENT_VALIDATE_Error;
        //						} else if (ReturnValue == "E3") {
        //							usb_io_output(609, 34, 1, 2, 0, 0);
        //							goto Raise_EVENT_VALIDATE_Error;
        //						} else if (ReturnValue == "E4") {
        //							usb_io_output(609, 34, 1, 2, 0, 0);
        //							goto Raise_EVENT_VALIDATE_Error;
        //						}

        //						//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						Query01 = "EXEC PS_PP049_01 '" + Strings.Trim(Convert.ToString(oForm.Items.Item("PP030HNo").Specific.VALUE + "-1")) + "', '" + oForm.Items.Item("OrdType").Specific.Selected.VALUE + "'";
        //						RecordSet01.DoQuery(Query01);
        //						if (RecordSet01.RecordCount == 0) {
        //							oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row).Specific.VALUE);
        //							oDS_PS_PP049L.SetValue("U_OrdMgNum", pval.Row - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
        //							oDS_PS_PP049L.SetValue("U_Sequence", pval.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
        //							oDS_PS_PP049L.SetValue("U_CpCode", pval.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
        //							oDS_PS_PP049L.SetValue("U_CpName", pval.Row - 1, RecordSet01.Fields.Item("CpName").Value);
        //							oDS_PS_PP049L.SetValue("U_OrdGbn", pval.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
        //							oDS_PS_PP049L.SetValue("U_BPLId", pval.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
        //							oDS_PS_PP049L.SetValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
        //							oDS_PS_PP049L.SetValue("U_ItemName", pval.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
        //							oDS_PS_PP049L.SetValue("U_OrdNum", pval.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
        //							oDS_PS_PP049L.SetValue("U_OrdSub1", pval.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
        //							oDS_PS_PP049L.SetValue("U_OrdSub2", pval.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
        //							oDS_PS_PP049L.SetValue("U_PP030HNo", pval.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
        //							oDS_PS_PP049L.SetValue("U_PP030MNo", pval.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
        //							oDS_PS_PP049L.SetValue("U_SelWt", pval.Row - 1, RecordSet01.Fields.Item("SelWt").Value);
        //							oDS_PS_PP049L.SetValue("U_PSum", pval.Row - 1, RecordSet01.Fields.Item("PSum").Value);
        //							oDS_PS_PP049L.SetValue("U_BQty", pval.Row - 1, RecordSet01.Fields.Item("BQty").Value);
        //							oDS_PS_PP049L.SetValue("U_PQty", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_YQty", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_NQty", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_ScrapWt", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_WorkTime", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_LineId", pval.Row - 1, "");

        //							////설비코드,명 Reset
        //							oDS_PS_PP049L.SetValue("U_MachCode", pval.Row - 1, "");
        //							oDS_PS_PP049L.SetValue("U_MachName", pval.Row - 1, "");
        //							////불량코드테이블
        //							if (oMat03.VisualRowCount == 0) {
        //								PS_PP049_AddMatrixRow03(0, ref true);
        //							} else {
        //								PS_PP049_AddMatrixRow03(oMat03.VisualRowCount);
        //							}

        //							oDS_PS_PP049N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
        //							oDS_PS_PP049N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpCode").Value);
        //							oDS_PS_PP049N.SetValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpName").Value);
        //							oDS_PS_PP049N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pval.Row));
        //							if (oMat01.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP049L.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
        //								PS_PP049_AddMatrixRow01(pval.Row);
        //							}
        //						}
        //					} else if (pval.ColUID == "PQty") {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE) <= 0) {
        //							if (Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "60") {
        //								goto Skip_PQty;
        //							} else {
        //								oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oDS_PS_PP049L.GetValue("U_" + pval.ColUID, pval.Row - 1));
        //							}
        //						} else {
        //							Skip_PQty:
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_YQty", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //							//UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							Weight = Conversion.Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1)) / 1000;
        //							if (Weight == 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //							}
        //							oDS_PS_PP049L.SetValue("U_NQty", pval.Row - 1, Convert.ToString(0));
        //							oDS_PS_PP049L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(0));
        //						}
        //					} else if (pval.ColUID == "NQty") {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE) <= 0) {
        //							if (Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "60") {
        //								goto skip_Nqty;
        //							} else {
        //								oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oDS_PS_PP049L.GetValue("U_" + pval.ColUID, pval.Row - 1));
        //							}
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						} else if (Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE) > Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.VALUE)) {
        //							if (Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "60") {
        //								goto skip_Nqty;
        //							} else {
        //								oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oDS_PS_PP049L.GetValue("U_" + pval.ColUID, pval.Row - 1));
        //							}
        //						} else {
        //							skip_Nqty:
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_YQty", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.VALUE) - Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //							//UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							Weight = Conversion.Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1)) / 1000;
        //							if (Weight == 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.VALUE) - Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_PP049L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Weight * (Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.VALUE) - Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE))));
        //							}
        //						}

        //					//작업시간(공수)을 입력할 때
        //					} else if (pval.ColUID == "WorkTime") {

        //						RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (oForm.Items.Item("BPLId").Specific.Selected.VALUE != "1") {


        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));

        //						} else {

        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));

        //						}

        //						//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						RecordSet01 = null;

        //					} else if (pval.ColUID == "CItemCod") {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049L.SetValue("U_CItemNam", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.VALUE + "' and U_ItemCod2 = '" + oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
        //					} else {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //					}
        //				}
        //			} else if ((pval.ItemUID == "Mat02")) {
        //				if ((pval.ColUID == "WorkCode")) {
        //					////기타작업
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_WorkName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
        //					if (oMat02.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP049M.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
        //						PS_PP049_AddMatrixRow02(pval.Row);
        //					}
        //				} else if (pval.ColUID == "NStart") {
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE) == 0 | Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE) == 0) {
        //						oDS_PS_PP049M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(0));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE)));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE)));
        //					} else {
        //						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE) <= Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE)) {
        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							Time = Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE) - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE);
        //						} else {
        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							Time = (2400 - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE)) + Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE);
        //						}
        //						Hour_Renamed = Conversion.Fix(Time / 100);
        //						//UPGRADE_WARNING: Mod에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        //						Minute_Renamed = Time % 100;
        //						Time = Hour_Renamed;
        //						if (Minute_Renamed > 0) {
        //							Time = Time + 0.5;
        //						}
        //						oDS_PS_PP049M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(Time));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE) - Time));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE) - Time));
        //					}
        //				} else if (pval.ColUID == "NEnd") {
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE) == 0 | Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE) == 0) {
        //						oDS_PS_PP049M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(0));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE)));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE)));
        //					} else {
        //						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE) <= Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE)) {
        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							Time = Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE) - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE);
        //						} else {
        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							Time = (2400 - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.VALUE)) + Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.VALUE);
        //						}
        //						Hour_Renamed = Conversion.Fix(Time / 100);
        //						//UPGRADE_WARNING: Mod에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        //						Minute_Renamed = Time % 100;
        //						Time = Hour_Renamed;
        //						if (Minute_Renamed > 0) {
        //							Time = Time + 0.5;
        //						}
        //						oDS_PS_PP049M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(Time));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE) - Time));
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.VALUE) - Time));
        //					}
        //				} else if (pval.ColUID == "YTime") {
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE)));
        //				} else {
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //				}
        //			} else if ((pval.ItemUID == "Mat03")) {
        //				if ((pval.ColUID == "FailCode")) {
        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049N.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049N.SetValue("U_FailName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE + "'", 0, 1));
        //				} else {
        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049N.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //				}
        //			} else {
        //				if ((pval.ItemUID == "DocEntry")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue(pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.VALUE);
        //				} else if ((pval.ItemUID == "BaseTime")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, Convert.ToString(Conversion.Val(oForm.Items.Item(pval.ItemUID).Specific.VALUE)));
        //				} else if ((pval.ItemUID == "OrdMgNum")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.VALUE);
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //						PS_PP049_LoadOrderInfo();
        //					}
        //				} else if ((pval.ItemUID == "ItemCode")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.VALUE);
        //					oMat01.Clear();
        //					oMat01.FlushToDataSource();
        //					oMat01.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow01(0, ref true);
        //					oMat02.Clear();
        //					oMat02.FlushToDataSource();
        //					oMat02.LoadFromDataSource();
        //					PS_PP049_AddMatrixRow02(0, ref true);
        //					oMat03.Clear();
        //					oMat03.FlushToDataSource();
        //					oMat03.LoadFromDataSource();

        //				} else if ((pval.ItemUID == "UseMCode")) {

        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					Query01 = "EXEC PS_PP049_98 '" + oForm.Items.Item("UseMCode").Specific.VALUE;

        //					RecordSet01.DoQuery(Query01);

        //					//UPGRADE_WARNING: oForm.Items(UseMName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("UseMName").Specific.VALUE = Strings.Trim(RecordSet01.Fields.Item(0).Value);

        //				} else {
        //					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_PP049H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.VALUE);
        //				}
        //			}
        //			oMat01.LoadFromDataSource();
        //			oMat01.AutoResizeColumns();
        //			oMat02.LoadFromDataSource();
        //			oMat02.AutoResizeColumns();
        //			oMat03.LoadFromDataSource();
        //			oMat03.AutoResizeColumns();
        //			oForm.Update();
        //			if (pval.ItemUID == "Mat01") {
        //				oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (pval.ItemUID == "Mat02") {
        //				oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (pval.ItemUID == "Mat03") {
        //				oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else {
        //				oForm.Items.Item(pval.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			}
        //		}
        //		//oMat01.Columns("BatchNum").Cells(pval.Row + 1).Click ct_Regular
        //	} else if (pval.BeforeAction == false) {
        //		//        If pval.ItemUID = "Mat01" Then
        //		//            If pval.ColUID = "BatchNum" Then
        //		//                If oMat01.Columns("CItemCod").Cells(pval.Row).Specific.VALUE = "" Then
        //		//                    oMat01.Columns("BatchNum").Cells(pval.Row).Specific.VALUE = ""
        //		//                Else
        //		//
        //		//                End If
        //		//            End If
        //		//        End If
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	oForm.Freeze(false);
        //	//    ProgBar01.VALUE = 100
        //	//    ProgBar01.Stop
        //	//    Set ProgBar01 = Nothing
        //	if (ReturnValue == "E1") {
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		SubMain.Sbo_Application.SetStatusBarMessage("배치번호 재고없음 : " + oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row).Specific.VALUE, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	} else if (ReturnValue == "E3") {
        //		SubMain.Sbo_Application.SetStatusBarMessage("배치번호 선입선출 오류 : " + BatchNumErr + " 선입고된 배치번호입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	} else if (ReturnValue == "E4") {
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		SubMain.Sbo_Application.SetStatusBarMessage("원재료 매칭 오류 : " + oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row).Specific.VALUE + " - " + "재고없음", SAPbouiCOM.BoMessageTime.bmt_Short, true);

        //	} else {
        //		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

        //	}
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_PP049_EnableFormItem();
        //		if (pval.ItemUID == "Mat01") {
        //			PS_PP049_AddMatrixRow01(oMat01.VisualRowCount);
        //			////UDO방식
        //		} else if (pval.ItemUID == "Mat02") {
        //			PS_PP049_AddMatrixRow02(oMat02.VisualRowCount);
        //			////UDO방식
        //		}
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_PP049_ResizeForm();
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

        //	SAPbouiCOM.DataTable oDataTable01 = null;
        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		if ((pval.ItemUID == "ItemCode")) {
        //			//UPGRADE_WARNING: pval.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (pval.SelectedObjects == null) {
        //			} else {
        //				MDC_Com.MDC_GP_CF_DBDatasourceReturn(pval, (pval.FormUID), "@PS_PP040H", "U_ItemCode,U_ItemName");
        //				oMat01.Clear();
        //				oMat01.FlushToDataSource();
        //				oMat01.LoadFromDataSource();
        //				PS_PP049_AddMatrixRow01(0, ref true);
        //				oMat02.Clear();
        //				oMat02.FlushToDataSource();
        //				oMat02.LoadFromDataSource();
        //				PS_PP049_AddMatrixRow02(0, ref true);
        //				oMat03.Clear();
        //				oMat03.FlushToDataSource();
        //				oMat03.LoadFromDataSource();
        //			}
        //		}
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

        //	if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02" | pval.ItemUID == "Mat03") {
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
        //	if (pval.ItemUID == "Mat01") {
        //		if (pval.ColUID == "CItemCod") {
        //			//UPGRADE_WARNING: oMat01.Columns(CItemCod).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat01.Columns(BatchNum).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (!string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row).Specific.VALUE) & string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(pval.Row).Specific.VALUE)) {
        //				oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			}

        //			//UPGRADE_WARNING: oMat01.Columns(CItemCod).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat01.Columns(BatchNum).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (!string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row).Specific.VALUE) & !string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(pval.Row).Specific.VALUE)) {
        //				oMat01.Columns.Item("BatchNum").Cells.Item(pval.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			}
        //		}

        //		if (pval.Row > 0) {
        //			oMat01Row01 = pval.Row;
        //		}


        //	} else if (pval.ItemUID == "Mat02") {
        //		if (pval.Row > 0) {
        //			oMat02Row02 = pval.Row;
        //		}
        //	} else if (pval.ItemUID == "Mat03") {
        //		if (pval.Row > 0) {
        //			oMat03Row03 = pval.Row;
        //		}
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
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object i = null;
        //	int j = 0;
        //	bool Exist = false;
        //	if ((oLastColRow01 > 0)) {
        //		if (pval.BeforeAction == true) {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//분말 첫번째 공정일 경우 오류
        //			if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "111" & (Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.VALUE) == "CP80111" | Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.VALUE) == "CP80101")) {
        //				MDC_Com.MDC_GF_Message(ref "첫공정은 행삭제 할수 없습니다.", ref "E");
        //				BubbleEvent = false;
        //				return;
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//분말 첫번째 공정일 경우 오류
        //			} else if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "601" & (Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.VALUE) == "CP80111" | Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.VALUE) == "CP80101")) {
        //				MDC_Com.MDC_GF_Message(ref "첫공정은 행삭제 할수 없습니다.", ref "E");
        //				BubbleEvent = false;
        //				return;
        //			}
        //			//추가 End
        //			if (oLastItemUID01 == "Mat01") {
        //				if ((PS_PP049_Validate("행삭제01") == false)) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				Continue_Renamed:
        //				for (i = 1; i <= oMat03.RowCount; i++) {
        //					//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(oLastColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(oLastColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.VALUE == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE & oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.VALUE == oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.VALUE) {
        //						//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.RemoveRecord((i - 1));
        //						oMat03.DeleteRow((i));
        //						oMat03.FlushToDataSource();
        //						goto Continue_Renamed;
        //					}
        //				}
        //			}
        //			////행삭제전 행삭제가능여부검사
        //		} else if (pval.BeforeAction == false) {
        //			if (oLastItemUID01 == "Mat01") {
        //				for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //				}

        //				for (i = 1; i <= oMat03.VisualRowCount; i++) {
        //					//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.VALUE != 1) {
        //						//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.VALUE = oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.VALUE - 1;
        //						////i
        //					}
        //				}

        //				oMat01.FlushToDataSource();
        //				oDS_PS_PP049L.RemoveRecord(oDS_PS_PP049L.Size - 1);
        //				oMat01.LoadFromDataSource();
        //				if (oMat01.RowCount == 0) {
        //					PS_PP049_AddMatrixRow01(0);
        //				} else {
        //					if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP049L.GetValue("U_OrdMgNum", oMat01.RowCount - 1)))) {
        //						PS_PP049_AddMatrixRow01(oMat01.RowCount);
        //					}
        //				}
        //			} else if (oLastItemUID01 == "Mat02") {
        //				for (i = 1; i <= oMat02.VisualRowCount; i++) {
        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //				}
        //				oMat02.FlushToDataSource();
        //				oDS_PS_PP049M.RemoveRecord(oDS_PS_PP049M.Size - 1);
        //				oMat02.LoadFromDataSource();
        //				if (oMat02.RowCount == 0) {
        //					PS_PP049_AddMatrixRow02(0);
        //				} else {
        //					if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP049M.GetValue("U_WorkCode", oMat02.RowCount - 1)))) {
        //						PS_PP049_AddMatrixRow02(oMat02.RowCount);
        //					}
        //				}
        //			} else if (oLastItemUID01 == "Mat03") {
        //				for (i = 1; i <= oMat03.VisualRowCount; i++) {
        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //				}
        //				oMat03.FlushToDataSource();
        //				////사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
        //				if (oDS_PS_PP049N.Size == 1) {
        //				} else {
        //					oDS_PS_PP049N.RemoveRecord(oDS_PS_PP049N.Size - 1);
        //				}
        //				oMat03.LoadFromDataSource();

        //				////공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
        //				for (i = 1; i <= oMat01.RowCount - 1; i++) {
        //					Exist = false;
        //					for (j = 1; j <= oMat03.RowCount; j++) {
        //						//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.VALUE & oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.VALUE) {
        //							////If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(j).Specific.VALUE Then
        //							Exist = true;
        //						}
        //					}
        //					////불량코드테이블에 값이 존재하지 않으면
        //					if (Exist == false) {
        //						if (oMat03.VisualRowCount == 0) {
        //							PS_PP049_AddMatrixRow03(0, ref true);
        //						} else {
        //							PS_PP049_AddMatrixRow03(oMat03.VisualRowCount);
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.VALUE);
        //						//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_PP049N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, i);
        //						oMat03.LoadFromDataSource();
        //						oMat03.AutoResizeColumns();
        //						oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
        //						oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
        //						//                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
        //						//                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
        //						oMat03.FlushToDataSource();
        //					}
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RECORD_MOVE
        //private void Raise_EVENT_RECORD_MOVE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string DocEntry = null;
        //	string DocEntryNext = null;
        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
        //	////원본문서
        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntryNext = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
        //	////다음문서

        //	////다음
        //	if (pval.MenuUID == "1288") {
        //		if (pval.BeforeAction == true) {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //				BubbleEvent = false;
        //				return;
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.VALUE))) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //			if (PS_PP049_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////이전
        //	} else if (pval.MenuUID == "1289") {
        //		if (pval.BeforeAction == true) {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //				BubbleEvent = false;
        //				return;
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.VALUE))) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //			if (PS_PP049_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////첫번째레코드로이동
        //	} else if (pval.MenuUID == "1290") {
        //		if (pval.BeforeAction == true) {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry DESC";
        //			////가장마지막행을 부여
        //			RecordSet01.DoQuery(Query01);
        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////원본문서
        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////다음문서
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			if (PS_PP049_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////마지막문서로이동
        //	} else if (pval.MenuUID == "1291") {
        //		if (pval.BeforeAction == true) {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry ASC";
        //			////가장 첫행을 부여
        //			RecordSet01.DoQuery(Query01);
        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////원본문서
        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////다음문서
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			if (PS_PP049_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	Raise_EVENT_RECORD_MOVE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RECORD_MOVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_CreateItems
        //private bool PS_PP049_CreateItems()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	string oQuery01 = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	oDS_PS_PP049H = oForm.DataSources.DBDataSources("@PS_PP040H");
        //	oDS_PS_PP049L = oForm.DataSources.DBDataSources("@PS_PP040L");
        //	oDS_PS_PP049M = oForm.DataSources.DBDataSources("@PS_PP040M");
        //	oDS_PS_PP049N = oForm.DataSources.DBDataSources("@PS_PP040N");

        //	oMat01 = oForm.Items.Item("Mat01").Specific;
        //	oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //	oMat01.AutoResizeColumns();

        //	oMat02 = oForm.Items.Item("Mat02").Specific;
        //	oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //	oMat02.AutoResizeColumns();

        //	oMat03 = oForm.Items.Item("Mat03").Specific;
        //	oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //	oMat03.AutoResizeColumns();

        //	oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //	oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //	oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");
        //	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");
        //	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");
        //	//UPGRADE_WARNING: oForm.Items().Specific.GroupWith 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
        //	//UPGRADE_WARNING: oForm.Items().Specific.GroupWith 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");

        //	oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

        //	oForm.DataSources.UserDataSources.Add("EmpChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("EmpChk").Specific.DataBind.SetBound(true, "", "EmpChk");


        //	oDocType01 = "작업일보등록(작지)";
        //	if ((oDocType01 == "작업일보등록(작지)")) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //	} else if ((oDocType01 == "작업일보등록(공정)")) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //	}

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	oForm.Freeze(false);
        //	return functionReturnValue;
        //	PS_PP049_CreateItems_Error:
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_PP049_SetComboBox
        //public void PS_PP049_SetComboBox()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	short loopCount = 0;

        //	oForm.Freeze(true);

        //	string sQry = null;

        //	//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
        //	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("BPLId").Specific), ref "SELECT BPLId, BPLName FROM OBPL order by BPLId", ref "", ref false, ref false);

        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "OrdType", "", "10", "일반");
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "OrdType", "", "20", "PSMT지원")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "OrdType", "", "30", "외주")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "OrdType", "", "40", "실적")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "OrdType", "", "50", "일반조정")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "OrdType", "", "60", "외주조정")
        //	//    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "OrdType", "", "70", "설계시간")
        //	MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm.Items.Item("OrdType").Specific), "PS_PP049", "OrdType");

        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "DocType", "", "10", "작지기준");
        //	MDC_PS_Common.Combo_ValidValues_Insert("PS_PP049", "DocType", "", "20", "공정기준");
        //	MDC_PS_Common.Combo_ValidValues_SetValueItem((oForm.Items.Item("DocType").Specific), "PS_PP049", "DocType");



        //	//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
        //	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("OrdGbn").Specific), ref "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE IN('111','601') order by Code", ref "", ref false, ref false);



        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId");
        //	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code");
        //	//    Call MDC_SetMod.Set_ComboList(oForm.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
        //	//    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")

        //	//거래처구분 콤보(2012.02.02 송명규 추가)
        //	//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "선택");
        //	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("CardType").Specific), ref "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", ref "", ref false, ref false);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

        //	//작업구분코드(2014.04.15 송명규 수정)
        //	sQry = "           SELECT      U_Minor,";
        //	sQry = sQry + "                U_CdName";
        //	sQry = sQry + " FROM       [@PS_SY001L]";
        //	sQry = sQry + " WHERE      Code = 'P203'";
        //	sQry = sQry + "                AND U_UseYN = 'Y'";
        //	sQry = sQry + " ORDER BY  U_Seq";
        //	if (oMat01.Columns.Item("WorkCls").ValidValues.Count > 0) {

        //		for (loopCount = 0; loopCount <= oMat01.Columns.Item("WorkCls").ValidValues.Count - 1; loopCount++) {
        //			oMat01.Columns.Item("WorkCls").ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //		}

        //		MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCls"), sQry);
        //	} else {
        //		MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCls"), sQry);
        //	}

        //	oForm.Freeze(false);
        //	return;
        //	PS_PP049_SetComboBox_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_SetComboBox_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_CF_ChooseFromList
        //public void PS_PP049_CF_ChooseFromList()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////ChooseFromList 설정
        //	SAPbouiCOM.ChooseFromListCollection oCFLs = null;
        //	SAPbouiCOM.Conditions oCons = null;
        //	SAPbouiCOM.Condition oCon = null;
        //	SAPbouiCOM.ChooseFromList oCFL = null;
        //	SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
        //	SAPbouiCOM.EditText oEdit = null;
        //	SAPbouiCOM.Column oColumn = null;

        //	oEdit = oForm.Items.Item("ItemCode").Specific;
        //	oCFLs = oForm.ChooseFromLists;
        //	oCFLCreationParams = SubMain.Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

        //	oCFLCreationParams.ObjectType = "4";
        //	oCFLCreationParams.UniqueID = "CFLITEMCODE";
        //	oCFLCreationParams.MultiSelection = false;
        //	oCFL = oCFLs.Add(oCFLCreationParams);

        //	oCons = oCFL.GetConditions();
        //	oCon = oCons.Add();
        //	oCon.Alias = "ItmsGrpCod";
        //	oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
        //	oCon.CondVal = "102";
        //	oCFL.SetConditions(oCons);

        //	oEdit.ChooseFromListUID = "CFLITEMCODE";
        //	oEdit.ChooseFromListAlias = "ItemCode";

        //	return;
        //	PS_PP049_CF_ChooseFromList_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_EnableFormItem
        //public void PS_PP049_EnableFormItem()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		////각모드에따른 아이템설정
        //		oForm.EnableMenu("1281", true);
        //		////찾기
        //		oForm.EnableMenu("1282", false);
        //		////추가
        //		oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		oForm.Items.Item("DocEntry").Enabled = false;
        //		oForm.Items.Item("OrdType").Enabled = true;
        //		oForm.Items.Item("OrdMgNum").Enabled = true;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oForm.Items.Item("Button01").Enabled = true;
        //		oForm.Items.Item("1").Enabled = true;
        //		oForm.Items.Item("Mat01").Enabled = true;
        //		oForm.Items.Item("Mat02").Enabled = true;
        //		oForm.Items.Item("Mat03").Enabled = true;
        //		oMat02.Columns.Item("NTime").Editable = true;
        //		////비가동시간만 사용
        //		//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_Index);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //		if (string.IsNullOrEmpty(Strings.Trim(oOrdGbn))) {
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //		} else {
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("OrdGbn").Specific.Select(oOrdGbn, SAPbouiCOM.BoSearchKey.psk_ByValue);
        //		}
        //		//Call oForm.Items("BPLId").Specific.Select(0, psk_Index)

        //		PS_PP049_FormClear();
        //		////UDO방식
        //		if ((oDocType01 == "작업일보등록(작지)")) {
        //			oDS_PS_PP049H.SetValue("U_DocType", 0, "10");
        //			//            Call oForm.Items("DocType").Specific.Select("10", psk_ByValue)
        //		} else if ((oDocType01 == "작업일보등록(공정)")) {
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //		}
        //		if (string.IsNullOrEmpty(Strings.Trim(oDocdate))) {
        //			//UPGRADE_WARNING: oForm.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("DocDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(System.Date.FromOADate(DateAndTime.Now.ToOADate() - 1), "YYYYMMDD");
        //		} else {
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("DocDate").Specific.VALUE = oDocdate;
        //		}
        //	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
        //		////각모드에따른 아이템설정
        //		oForm.EnableMenu("1281", false);
        //		////찾기
        //		oForm.EnableMenu("1282", true);
        //		////추가
        //		oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		oForm.Items.Item("DocEntry").Enabled = true;
        //		oForm.Items.Item("OrdType").Enabled = true;
        //		oForm.Items.Item("OrdMgNum").Enabled = true;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oForm.Items.Item("Button01").Enabled = true;
        //		oForm.Items.Item("1").Enabled = true;
        //		oForm.Items.Item("Mat01").Enabled = false;
        //		oForm.Items.Item("Mat02").Enabled = false;
        //		oForm.Items.Item("Mat03").Enabled = false;
        //	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
        //		oForm.EnableMenu("1281", true);
        //		////찾기
        //		oForm.EnableMenu("1282", true);
        //		////추가
        //		////각모드에따른 아이템설정
        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP040H] WHERE DocEntry = ' & Trim(oDS_PS_PP049H.GetValue(DocEntry, 0)) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + Strings.Trim(oDS_PS_PP049H.GetValue("DocEntry", 0)) + "'", 0, 1) == "Y") {
        //			oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm.Items.Item("DocEntry").Enabled = false;
        //			oForm.Items.Item("OrdType").Enabled = false;
        //			oForm.Items.Item("OrdMgNum").Enabled = false;
        //			oForm.Items.Item("DocDate").Enabled = false;
        //			oForm.Items.Item("Button01").Enabled = false;
        //			oForm.Items.Item("1").Enabled = false;
        //			oForm.Items.Item("Mat01").Enabled = false;
        //			oForm.Items.Item("Mat02").Enabled = false;
        //			oForm.Items.Item("Mat03").Enabled = false;
        //		} else {
        //			////조정, 설계
        //			if (Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "10" | Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "60" | Strings.Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) == "70") {
        //				oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("DocEntry").Enabled = false;
        //				oForm.Items.Item("OrdType").Enabled = false;
        //				oForm.Items.Item("OrdMgNum").Enabled = true;
        //				oForm.Items.Item("DocDate").Enabled = true;
        //				oForm.Items.Item("Button01").Enabled = true;
        //				oForm.Items.Item("1").Enabled = true;
        //				oForm.Items.Item("Mat01").Enabled = true;
        //				oForm.Items.Item("Mat02").Enabled = true;
        //				oForm.Items.Item("Mat03").Enabled = true;
        //				//            ElseIf Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) = "20" Then '//PSMT
        //				//                oForm.Items("Focus").Click ct_Regular
        //				//                oForm.Items("DocEntry").Enabled = False
        //				//                oForm.Items("OrdType").Enabled = False
        //				//                oForm.Items("OrdMgNum").Enabled = True
        //				//                oForm.Items("DocDate").Enabled = True
        //				//                oForm.Items("Button01").Enabled = True
        //				//                oForm.Items("1").Enabled = True
        //				//                oForm.Items("Mat01").Enabled = True
        //				//                oForm.Items("Mat02").Enabled = True
        //				//                oForm.Items("Mat03").Enabled = True
        //				//            ElseIf Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) = "30" Then '//외주
        //				//                oForm.Items("Focus").Click ct_Regular
        //				//                oForm.Items("DocEntry").Enabled = False
        //				//                oForm.Items("OrdType").Enabled = False
        //				//                oForm.Items("OrdMgNum").Enabled = False
        //				//                oForm.Items("DocDate").Enabled = False
        //				//                oForm.Items("Button01").Enabled = False
        //				//                oForm.Items("1").Enabled = False
        //				//                oForm.Items("Mat01").Enabled = False
        //				//                oForm.Items("Mat02").Enabled = False
        //				//                oForm.Items("Mat03").Enabled = False
        //				//            ElseIf Trim(oDS_PS_PP049H.GetValue("U_OrdType", 0)) = "40" Then '//실적
        //				//                oForm.Items("Focus").Click ct_Regular
        //				//                oForm.Items("DocEntry").Enabled = False
        //				//                oForm.Items("OrdType").Enabled = False
        //				//                oForm.Items("OrdMgNum").Enabled = False
        //				//                oForm.Items("DocDate").Enabled = False
        //				//                oForm.Items("Button01").Enabled = False
        //				//                oForm.Items("1").Enabled = False
        //				//                oForm.Items("Mat01").Enabled = False
        //				//                oForm.Items("Mat02").Enabled = False
        //				//                oForm.Items("Mat03").Enabled = False
        //			}
        //		}
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	PS_PP049_EnableFormItem_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_EnableFormItem_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_AddMatrixRow01
        //public void PS_PP049_AddMatrixRow01(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	////행추가여부
        //	if (RowIserted == false) {
        //		oDS_PS_PP049L.InsertRecord((oRow));
        //	}
        //	oMat01.AddRow();
        //	oDS_PS_PP049L.Offset = oRow;
        //	oDS_PS_PP049L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oDS_PS_PP049L.SetValue("U_WorkCls", oRow, "A");
        //	//작업구분을 기본으로 선택(2014.04.15 송명규 추가)
        //	oMat01.LoadFromDataSource();
        //	oForm.Freeze(false);
        //	return;
        //	PS_PP049_AddMatrixRow01_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_AddMatrixRow01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_AddMatrixRow02
        //public void PS_PP049_AddMatrixRow02(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	////행추가여부
        //	if (RowIserted == false) {
        //		oDS_PS_PP049M.InsertRecord((oRow));
        //	}
        //	oMat02.AddRow();
        //	oDS_PS_PP049M.Offset = oRow;
        //	oDS_PS_PP049M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oMat02.LoadFromDataSource();
        //	oForm.Freeze(false);
        //	return;
        //	PS_PP049_AddMatrixRow02_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_AddMatrixRow02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_AddMatrixRow03
        //public void PS_PP049_AddMatrixRow03(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	////행추가여부
        //	if (RowIserted == false) {
        //		oDS_PS_PP049N.InsertRecord((oRow));
        //	}
        //	oMat03.AddRow();
        //	oDS_PS_PP049N.Offset = oRow;
        //	oDS_PS_PP049N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oMat03.LoadFromDataSource();
        //	oForm.Freeze(false);
        //	return;
        //	PS_PP049_AddMatrixRow03_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_AddMatrixRow03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_FormClear
        //public void PS_PP049_FormClear()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocEntry = null;
        //	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_PP040'", ref "");
        //	if (Convert.ToDouble(DocEntry) == 0) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.VALUE = 1;
        //	} else {
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
        //	}
        //	return;
        //	PS_PP049_FormClear_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_EnableMenus
        //private void PS_PP049_EnableMenus()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////메뉴활성화
        //	//    Call oForm.EnableMenu("1288", True)
        //	//    Call oForm.EnableMenu("1289", True)
        //	//    Call oForm.EnableMenu("1290", True)
        //	//    Call oForm.EnableMenu("1291", True)
        //	////Call MDC_GP_EnableMenus(oForm, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
        //	MDC_Com.MDC_GP_EnableMenus(ref oForm, false, false, true, true, false, true, true, true, true,
        //	true, false, false, false, false, false, false);
        //	////메뉴설정
        //	return;
        //	PS_PP049_EnableMenus_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_SetDocument
        //private void PS_PP049_SetDocument(string oFormDocEntry)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((string.IsNullOrEmpty(oFormDocEntry))) {
        //		PS_PP049_EnableFormItem();
        //		PS_PP049_AddMatrixRow01(0, ref true);
        //		////UDO방식일때
        //		PS_PP049_AddMatrixRow02(0, ref true);
        //		////UDO방식일때
        //	} else {
        //		oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //		PS_PP049_EnableFormItem();
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.VALUE = oFormDocEntry;
        //		oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //	}
        //	return;
        //	PS_PP049_SetDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_CheckDataValid
        //public bool PS_PP049_CheckDataValid()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = false;
        //	object i = null;
        //	int j = 0;
        //	double FailQty = 0;



        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	//UPGRADE_WARNING: MDC_PS_Common.GetValue(select Count(*) from OFPR Where ' & oForm.Items(DocDate).Specific.VALUE & ' between F_RefDate and T_RefDate And PeriodStat = 'Y') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (MDC_PS_Common.GetValue("select Count(*) from OFPR Where '" + oForm.Items.Item("DocDate").Specific.VALUE + "' between F_RefDate and T_RefDate And PeriodStat = 'Y'") > 0) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("해당일자는 전기기간이 잠겼습니다. 일자를 확인바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		PS_PP049_FormClear();
        //	}
        //	//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("OrdType").Specific.Selected.VALUE != "10" & oForm.Items.Item("OrdType").Specific.Selected.VALUE != "20" & oForm.Items.Item("OrdType").Specific.Selected.VALUE != "50" & oForm.Items.Item("OrdType").Specific.Selected.VALUE != "60" & oForm.Items.Item("OrdType").Specific.Selected.VALUE != "70") {
        //		SubMain.Sbo_Application.SetStatusBarMessage("작업타입이 일반, PSMT지원, 조정, 설계가 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.VALUE)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("작지번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm.Items.Item("OrdNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	if (oMat01.VisualRowCount == 1) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("공정정보 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	if (oMat02.VisualRowCount == 1) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("작업자정보 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	//    '마감상태 체크_S(2017.11.23 송명규 추가)
        //	//    If MDC_PS_Common.Check_Finish_Status(Trim(oForm.Items("BPLId").Specific.VALUE), oForm.Items("DocDate").Specific.VALUE, oForm.TypeEx) = False Then
        //	//        Call Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작업일보일자를 확인하고, 회계부서로 문의하세요.", bmt_Short, True)
        //	//        PS_PP049_CheckDataValid = False
        //	//        Exit Function
        //	//    End If
        //	//    '마감상태 체크_E(2017.11.23 송명규 추가)

        //	// 작업자 1명이상 가능토록 수정 (이병각)
        //	//    If oMat02.VisualRowCount > 2 Then '//한명이상 입력했을경우
        //	//        If oForm.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드
        //	//            Sbo_Application.SetStatusBarMessage "작업자정보 한명만 입력할수 있습니다.", bmt_Short, True
        //	//            PS_PP049_CheckDataValid = False
        //	//            Exit Function
        //	//        Else
        //	//            '//휘팅,부품은 여러명 입력할수 있다.
        //	//        End If
        //	//    End If

        //	if (oMat03.VisualRowCount == 0) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("불량정보 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE))) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("작지문서번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (Strings.Trim(oForm.Items.Item("OrdType").Specific.VALUE) != "50" & Strings.Trim(oForm.Items.Item("OrdType").Specific.VALUE) != "60") {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if ((Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.VALUE) <= 0)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("생산수량은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat01.Columns.Item("PQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}
        //		}
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (Strings.Trim(oForm.Items.Item("OrdType").Specific.VALUE) != "50" & Strings.Trim(oForm.Items.Item("OrdType").Specific.VALUE) != "60" & Strings.Trim(oForm.Items.Item("OrdType").Specific.VALUE) != "70") {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if ((Conversion.Val(oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.VALUE) <= 0)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("실동시간은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat01.Columns.Item("WorkTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}
        //		}

        //		//작업완료여부(2012.02.02. 송명규 추가)
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		//기계공구, 몰드일 경우만 작업완료여부 필수 체크
        //		if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "105" | Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "106") {

        //			//UPGRADE_WARNING: oMat01.Columns(CompltYN).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if ((oMat01.Columns.Item("CompltYN").Cells.Item(i).Specific.VALUE == "%")) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("작업구분이 기계공구, 몰드일경우는 작업완료여부가 필수입니다. 확인하십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat01.Columns.Item("CompltYN").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}

        //		}

        //		////불량수량 검사
        //		FailQty = 0;
        //		for (j = 1; j <= oMat03.VisualRowCount; j++) {
        //			////불량코드를 입력했는지 check
        //			//UPGRADE_WARNING: oMat03.Columns(FailCode).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (Conversion.Val(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.VALUE) != 0 & string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.VALUE)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("불량수량이 입력되었을 때는 불량코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}

        //			//UPGRADE_WARNING: oMat03.Columns(FailCode).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (Conversion.Val(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.VALUE) == 0 & !string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.VALUE)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("불량코드를 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}

        //			//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.VALUE) & (oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.VALUE)) {
        //				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				FailQty = FailQty + Conversion.Val(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.VALUE);
        //			}
        //		}
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (Strings.Trim(oForm.Items.Item("OrdType").Specific.VALUE) != "50" & Strings.Trim(oForm.Items.Item("OrdType").Specific.VALUE) != "60") {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (Conversion.Val(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.VALUE) != FailQty) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}
        //		}

        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "601" | Strings.Trim(oForm.Items.Item("OrdGbn").Specific.VALUE) == "111") {
        //			//If oMat01.Columns("CpCode").Cells(i).Specific.VALUE = "CP80101" And Trim(oMat01.Columns("CItemCod").Cells(i).Specific.VALUE) = "" Then
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat01.Columns(Sequence).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.VALUE == 1 & string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.VALUE))) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("공정 사용 원재료코드가 없습니다. 사용 원재료를 선택해 주세요", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}
        //		}
        //	}

        //	//비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_S
        //	for (i = 1; i <= oMat02.VisualRowCount - 1; i++) {

        //		//UPGRADE_WARNING: oMat02.Columns(NCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((!string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.VALUE))) {

        //			//UPGRADE_WARNING: oMat02.Columns(NTime).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.VALUE)) {

        //				SubMain.Sbo_Application.SetStatusBarMessage("비가동코드가 입력되었을 때는 비가동시간은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat02.Columns.Item("NTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;

        //			}

        //		}

        //		//UPGRADE_WARNING: oMat02.Columns(NTime).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((!string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.VALUE))) {

        //			//UPGRADE_WARNING: oMat02.Columns(NCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.VALUE)) {

        //				SubMain.Sbo_Application.SetStatusBarMessage("비가동시간이 입력되었을 때는 비가동코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat02.Columns.Item("NCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;

        //			}

        //		}

        //	}
        //	//비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_E

        //	if ((PS_PP049_Validate("검사01") == false)) {
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	oDS_PS_PP049L.RemoveRecord(oDS_PS_PP049L.Size - 1);
        //	oMat01.LoadFromDataSource();
        //	oDS_PS_PP049M.RemoveRecord(oDS_PS_PP049M.Size - 1);
        //	oMat02.LoadFromDataSource();

        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		PS_PP049_FormClear();
        //	}
        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	PS_PP049_CheckDataValid_Error:

        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_CheckDataValid_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_PP049_MTX01
        //private void PS_PP049_MTX01()
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
        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param01 = Strings.Trim(oForm.Items.Item("Param01").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param02 = Strings.Trim(oForm.Items.Item("Param01").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param03 = Strings.Trim(oForm.Items.Item("Param01").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param04 = Strings.Trim(oForm.Items.Item("Param01").Specific.VALUE);

        //	Query01 = "SELECT 10";
        //	RecordSet01.DoQuery(Query01);

        //	oMat01.Clear();
        //	oMat01.FlushToDataSource();
        //	oMat01.LoadFromDataSource();

        //	if ((RecordSet01.RecordCount == 0)) {
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_PP049_MTX01_Exit;
        //	}

        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //	for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //		if (i != 0) {
        //			oDS_PS_PP049L.InsertRecord((i));
        //		}
        //		oDS_PS_PP049L.Offset = i;
        //		oDS_PS_PP049L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
        //		oDS_PS_PP049L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
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
        //	PS_PP049_MTX01_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	if ((ProgressBar01 != null)) {
        //		ProgressBar01.Stop();
        //	}
        //	return;
        //	PS_PP049_MTX01_Error:
        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_SumWorkTime
        //private void PS_PP049_SumWorkTime()
        //{
        //	//******************************************************************************
        //	//Function ID    : PS_PP049_SumWorkTime()
        //	//해 당 모 듈    : 생산관리
        //	//기        능    : 근무시간의 총합을 구함
        //	//인        수    : 없음
        //	//반   환   값   : 없음
        //	//특 이 사 항    : 없음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	short loopCount = 0;
        //	double Total = 0;

        //	for (loopCount = 0; loopCount <= oMat01.RowCount - 2; loopCount++) {
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		Total = Total + Convert.ToDouble((string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.VALUE)) ? 0 : Strings.Trim(oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.VALUE)));
        //	}

        //	//UPGRADE_WARNING: oForm.Items(Total).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oForm.Items.Item("Total").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Total, "##0.#0");

        //	return;
        //	PS_PP049_SumWorkTime_Error:

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_SumWorkTime_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_ResizeForm
        //private void PS_PP049_ResizeForm()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Items.Item("Mat02").Top = 170;
        //	oForm.Items.Item("Mat02").Left = 7;
        //	oForm.Items.Item("Mat02").Height = ((oForm.Height - 170) / 3 * 1) - 20;
        //	oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 14;

        //	oForm.Items.Item("Mat03").Top = 170;
        //	oForm.Items.Item("Mat03").Left = oForm.Width / 2;
        //	oForm.Items.Item("Mat03").Height = ((oForm.Height - 170) / 3 * 1) - 20;
        //	oForm.Items.Item("Mat03").Width = oForm.Width / 2 - 14;

        //	oForm.Items.Item("Mat01").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 40;
        //	oForm.Items.Item("Mat01").Left = 7;
        //	oForm.Items.Item("Mat01").Height = ((oForm.Height - 170) / 3 * 2) - 80;
        //	oForm.Items.Item("Mat01").Width = oForm.Width - 21;

        //	oForm.Items.Item("Opt01").Left = 10;
        //	oForm.Items.Item("Opt02").Left = oForm.Width / 2;
        //	oForm.Items.Item("Opt03").Left = 10;
        //	oForm.Items.Item("Opt03").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 20;
        //	return;
        //	PS_PP049_ResizeForm_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_ResizeForm_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_Validate
        //public bool PS_PP049_Validate(string ValidateType)
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	int PrevDBCpQty = 0;
        //	int PrevMATRIXCpQty = 0;
        //	int CurrentDBCpQty = 0;
        //	int CurrentMATRIXCpQty = 0;
        //	int NextDBCpQty = 0;
        //	int NextMATRIXCpQty = 0;
        //	string PrevCpInfo = null;
        //	string CurrentCpInfo = null;
        //	string NextCpInfo = null;

        //	string OrdMgNum = null;
        //	bool Exist = false;
        //	string LineNum = null;
        //	string DocEntry = null;

        //	if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP040H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
        //			MDC_Com.MDC_GF_Message(ref "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", ref "W");
        //			functionReturnValue = false;
        //			goto PS_PP049_Validate_Exit;
        //		}
        //	}

        //	//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	////작업타입이 일반,조정인경우
        //	if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "60") {
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	////작업타입이 PSMT지원인경우
        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	////작업타입이 외주인경우
        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {
        //		MDC_Com.MDC_GF_Message(ref "해당작업타입은 변경이 불가능합니다.", ref "W");
        //		functionReturnValue = false;
        //		goto PS_PP049_Validate_Exit;
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	////작업타입이 실적인경우
        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {
        //		MDC_Com.MDC_GF_Message(ref "해당작업타입은 변경이 불가능합니다.", ref "W");
        //		functionReturnValue = false;
        //		goto PS_PP049_Validate_Exit;
        //	}

        //	string QueryString = null;
        //	if (ValidateType == "검사01") {
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 일반인경우
        //		if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10") {
        //			////입력된 행에 대해
        //			for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_PP030H] PS_PP030H LEFT JOIN [PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = ' & oMat01.Columns(OrdMgNum).Cells(i).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE + "'", 0, 1) <= 0) {
        //					MDC_Com.MDC_GF_Message(ref "작업지시문서가 존재하지 않습니다.", ref "W");
        //					functionReturnValue = false;
        //					goto PS_PP049_Validate_Exit;
        //				}
        //			}

        //			if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //				////삭제된 행에 대한처리
        //				Query01 = "SELECT ";
        //				Query01 = Query01 + " PS_PP040H.DocEntry,";
        //				Query01 = Query01 + " PS_PP040L.LineId,";
        //				Query01 = Query01 + " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
        //				Query01 = Query01 + " PS_PP040L.U_OrdGbn AS OrdGbn,";
        //				Query01 = Query01 + " PS_PP040L.U_PP030HNo AS PP030HNo,";
        //				Query01 = Query01 + " PS_PP040L.U_PP030MNo AS PP030MNo,";
        //				Query01 = Query01 + " PS_PP040L.U_OrdMgNum AS OrdMgNum ";
        //				Query01 = Query01 + " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				Query01 = Query01 + " WHERE PS_PP040L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'";
        //				RecordSet01.DoQuery(Query01);
        //				for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //					Exist = false;
        //					////기존에 있는 행에대한처리
        //					for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
        //						//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						////새로추가된 행인경우, 검사할필요없다
        //						if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.VALUE))) {
        //						} else {
        //							////라인번호가 같고, 문서번호가 같으면 존재하는행
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (Conversion.Val(RecordSet01.Fields.Item(0).Value) == Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) & Conversion.Val(RecordSet01.Fields.Item(1).Value) == Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.VALUE)) {
        //								Exist = true;
        //							}
        //						}
        //					}
        //					////삭제된 행중 수량관계를 알아본다.
        //					if (Exist == false) {
        //						////휘팅이면서
        //						if (RecordSet01.Fields.Item("OrdGbn").Value == "101") {
        //							////현재 공정이 실적공정이면..
        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP049_05 ' & RecordSet01.Fields(OrdMgNum).VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (MDC_PS_Common.GetValue("EXEC PS_PP049_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y") {
        //								////휘팅벌크포장
        //								//                            PP040_CurrentPQty = 0
        //								//                            PP040_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP040L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)
        //								//                            PP070_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)
        //								//                            PP080_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)

        //								if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
        //									MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
        //									functionReturnValue = false;
        //									goto PS_PP049_Validate_Exit;
        //								}
        //								////휘팅실적
        //								if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
        //									MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
        //									functionReturnValue = false;
        //									goto PS_PP049_Validate_Exit;
        //								}
        //							}
        //						}

        //						////기계공구,몰드
        //						if (RecordSet01.Fields.Item("OrdGbn").Value == "105" | RecordSet01.Fields.Item("OrdGbn").Value == "106") {
        //							////그냥 입력가능
        //						////휘팅,부품
        //						} else if (RecordSet01.Fields.Item("OrdGbn").Value == "101" | RecordSet01.Fields.Item("OrdGbn").Value == "102") {
        //							////삭제된 행에 대한 검사..
        //							//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							OrdMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
        //							//// DocEntry + '-' + LineId
        //							CurrentCpInfo = OrdMgNum;

        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP049_02 '" + OrdMgNum + "'");
        //							if (string.IsNullOrEmpty(PrevCpInfo)) {
        //								////해당공정이 첫공정이면 입력되어도 상관없다.
        //							} else {
        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'");
        //								////재공이동 수량
        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");
        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");

        //								PrevMATRIXCpQty = 0;
        //								for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
        //									//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.VALUE == PrevCpInfo)) {
        //										//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //										PrevMATRIXCpQty = PrevMATRIXCpQty + Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.VALUE);
        //									}
        //								}
        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'");
        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");
        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");

        //								CurrentMATRIXCpQty = 0;
        //								for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
        //									//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.VALUE == CurrentCpInfo)) {
        //										//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //										CurrentMATRIXCpQty = CurrentMATRIXCpQty + Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.VALUE);
        //									}
        //								}
        //								if (((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))) {
        //									SubMain.Sbo_Application.SetStatusBarMessage("삭제된 공정의 선행공정의 생산수량이 삭제된 공정의 생산수량을 미달합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //									functionReturnValue = false;
        //									goto PS_PP049_Validate_Exit;
        //								}
        //							}
        //							//                        If oForm.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
        //							//                            NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP049_03 '" & OrdMgNum & "'")
        //							//                            If NextCpInfo = "" Then
        //							//                                '//해당공정이 마지막공정이면 삭제되어도 상관없다.
        //							//                            Else
        //							//                                NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
        //							//                                NextMATRIXCpQty = 0
        //							//                                For j = 1 To oMat01.VisualRowCount - 1
        //							//                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
        //							//                                        NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
        //							//                                    End If
        //							//                                Next
        //							//                                CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
        //							//                                CurrentMATRIXCpQty = 0
        //							//                                For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
        //							//                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
        //							//                                        CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
        //							//                                    End If
        //							//                                Next
        //							//                                If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
        //							//                                    Sbo_Application.SetStatusBarMessage "삭제된 공정의 후행공정의 생산수량이 삭제된 공정의 생산수량을 초과합니다.", bmt_Short, True
        //							//                                    PS_PP049_Validate = False
        //							//                                    GoTo PS_PP049_Validate_Exit
        //							//                                End If
        //							//                            End If
        //							//                        End If
        //						}
        //					}
        //					RecordSet01.MoveNext();
        //				}
        //			}

        //			if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //				for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					////새로추가된 행인경우, 검사할필요없다
        //					if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.VALUE))) {
        //					} else {
        //						//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						////휘팅이면서
        //						if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.VALUE == "101") {
        //							////현재공정이 실적공정이면
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP049_05 ' & Val(oForm.Items(DocEntry).Specific.VALUE) & - & Val(oMat01.Columns(LineId).Cells(i).Specific.VALUE) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							////현재 공정이 바렐 앞공정이면..
        //							if (MDC_PS_Common.GetValue("EXEC PS_PP049_05 '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "-" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'", 0, 1) == "Y") {
        //								//                            '//휘팅벌크포장,휘팅실적
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP070L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'", 0, 1)) > 0 | (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP080L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'", 0, 1)) > 0) {
        //									////작업일보등록된문서중에 수정이 된문서를 구함
        //									Query01 = "SELECT ";
        //									Query01 = Query01 + " PS_PP040L.U_OrdMgNum,";
        //									Query01 = Query01 + " PS_PP040L.U_Sequence,";
        //									Query01 = Query01 + " PS_PP040L.U_CpCode,";
        //									Query01 = Query01 + " PS_PP040L.U_ItemCode,";
        //									Query01 = Query01 + " PS_PP040L.U_PP030HNo,";
        //									Query01 = Query01 + " PS_PP040L.U_PP030MNo,";
        //									Query01 = Query01 + " PS_PP040L.U_PQty,";
        //									Query01 = Query01 + " PS_PP040L.U_NQty,";
        //									Query01 = Query01 + " PS_PP040L.U_ScrapWt,";
        //									Query01 = Query01 + " PS_PP040L.U_WorkTime";
        //									Query01 = Query01 + " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry";
        //									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									Query01 = Query01 + " WHERE PS_PP040H.DocEntry = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "'";
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									Query01 = Query01 + " AND PS_PP040L.LineId = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.VALUE) + "'";
        //									Query01 = Query01 + " AND PS_PP040H.Canceled = 'N'";
        //									RecordSet01.DoQuery(Query01);
        //									//UPGRADE_WARNING: oMat01.Columns(WorkTime).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(ScrapWt).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(NQty).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(PQty).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(PP030MNo).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(PP030HNo).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(ItemCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(Sequence).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if ((RecordSet01.Fields.Item(0).Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(1).Value == oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(2).Value == oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(3).Value == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(4).Value == oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(5).Value == oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(6).Value == oMat01.Columns.Item("PQty").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(7).Value == oMat01.Columns.Item("NQty").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(8).Value == oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.VALUE & RecordSet01.Fields.Item(9).Value == oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.VALUE)) {
        //									////값이 변경된 행의경우
        //									} else {
        //										MDC_Com.MDC_GF_Message(ref "생산실적이 등록된 행은 수정할수 없습니다.", ref "W");
        //										functionReturnValue = false;
        //										goto PS_PP049_Validate_Exit;
        //									}
        //								}
        //							}
        //						}
        //					}
        //				}
        //			}

        //			////입력된 모든행에 대해 입력가능성 검사
        //			for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////기계공구,몰드
        //				if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.VALUE == "105" | oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.VALUE == "106") {
        //					////그냥 입력가능
        //					//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////휘팅,부품
        //				} else if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.VALUE == "101" | oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.VALUE == "102") {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					OrdMgNum = oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.VALUE;
        //					CurrentCpInfo = OrdMgNum;

        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP049_02 '" + OrdMgNum + "'");
        //					if (string.IsNullOrEmpty(PrevCpInfo)) {
        //						////해당공정이 첫공정이면 입력되어도 상관없다.
        //					} else {

        //						//PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & PrevCpInfo & "' AND PS_PP040H.DocEntry <> '" & oForm.Items("DocEntry").Specific.VALUE & "' AND PS_PP040H.Canceled = 'N'")
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						PrevDBCpQty = MDC_PS_Common.GetValue("EXEC PS_PP049_07 '" + PrevCpInfo + "', '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'");
        //						////재공 이동수량 반영
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");

        //						PrevMATRIXCpQty = 0;
        //						for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
        //							//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.VALUE == PrevCpInfo)) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								PrevMATRIXCpQty = PrevMATRIXCpQty + Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.VALUE);
        //							}
        //						}
        //						//CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & oForm.Items("DocEntry").Specific.VALUE & "' AND PS_PP040H.Canceled = 'N'")
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						CurrentDBCpQty = MDC_PS_Common.GetValue("EXEC PS_PP049_07 '" + CurrentCpInfo + "', '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'");
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");

        //						CurrentMATRIXCpQty = 0;
        //						for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
        //							//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(j).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.VALUE == CurrentCpInfo)) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								CurrentMATRIXCpQty = CurrentMATRIXCpQty + Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.VALUE);
        //							}
        //						}
        //						//// 노대리님 요청 주석
        //						if (((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))) {
        //							SubMain.Sbo_Application.SetStatusBarMessage("선행공정의 생산수량이 현공정의 생산수량에 미달 합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //							//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oMat01.SelectRow(i, true, false);
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}

        //					}
        //					//                    If oForm.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
        //					//                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP049_03 '" & OrdMgNum & "'")
        //					//                        If NextCpInfo = "" Then
        //					//                            '//해당공정이 마지막공정이면 삭제되어도 상관없다.
        //					//                        Else
        //					//                            NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & oForm.Items("DocEntry").Specific.Value & "' AND PS_PP040H.Canceled = 'N'")
        //					//                            NextMATRIXCpQty = 0
        //					//                            For j = 1 To oMat01.VisualRowCount - 1
        //					//                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
        //					//                                    NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
        //					//                                End If
        //					//                            Next
        //					//                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & oForm.Items("DocEntry").Specific.Value & "' AND PS_PP040H.Canceled = 'N'")
        //					//                            CurrentMATRIXCpQty = 0
        //					//                            For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
        //					//                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
        //					//                                    CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
        //					//                                End If
        //					//                            Next
        //					//                            If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
        //					//                                Sbo_Application.SetStatusBarMessage "후행공정의 생산수량이 현공정의 생산수량을 초과 합니다.", bmt_Short, True
        //					//                                PS_PP049_Validate = False
        //					//                                GoTo PS_PP049_Validate_Exit
        //					//                            End If
        //					//                        End If
        //					//                    End If
        //				}
        //			}
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 PSMT지원인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {
        //			////현재는 특별한 조건이 필요치 않음
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 외주인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {
        //			////현재는 특별한 조건이 필요치 않음
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 실적인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {
        //			////현재는 특별한 조건이 필요치 않음
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 조정인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50") {
        //			////현재는 특별한 조건이 필요치 않음
        //		}
        //	} else if (ValidateType == "행삭제01") {
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 일반인경우
        //		if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10") {
        //			////행삭제전 행삭제가능여부검사
        //			//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(oMat01Row01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////새로추가된 행인경우, 삭제하여도 무방하다
        //			if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.VALUE))) {
        //			} else {
        //				//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(oMat01Row01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////휘팅이면서
        //				if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.VALUE == "101") {
        //					////현재공정이 실적공정이면
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP049_05 ' & Val(oMat01.Columns(PP030HNo).Cells(oMat01Row01).Specific.VALUE) & - & Val(oMat01.Columns(PP030MNo).Cells(oMat01Row01).Specific.VALUE) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					////현재 공정이 바렐 앞공정이면..
        //					if (MDC_PS_Common.GetValue("EXEC PS_PP049_05 '" + Conversion.Val(oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.VALUE) + "-" + Conversion.Val(oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.VALUE) + "'", 0, 1) == "Y") {
        //						////휘팅벌크포장
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP070L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.VALUE) + "'", 0, 1)) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //						////휘팅실적
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP080L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.VALUE) + "'", 0, 1)) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //					}

        //					//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(oMat01Row01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////기계공구,몰드
        //				} else if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.VALUE == "105" | oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.VALUE == "106") {

        //					//재고가 존재하면 행삭제 불가 기능 추가(2011.12.15 송명규 추가)

        //					QueryString = "                     SELECT      SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]";
        //					QueryString = QueryString + "  FROM       OINM AS A";
        //					QueryString = QueryString + "                 INNER JOIN";
        //					QueryString = QueryString + "                 OITM As B";
        //					QueryString = QueryString + "                     ON A.ItemCode = B.ItemCode";
        //					QueryString = QueryString + "  WHERE      B.U_ItmBsort IN ('105','106')";
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					QueryString = QueryString + "                 AND A.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.VALUE + "'";
        //					QueryString = QueryString + "  GROUP BY  A.ItemCode";

        //					if ((string.IsNullOrEmpty((MDC_PS_Common.GetValue(QueryString, 0, 1))) ? 0 : (MDC_PS_Common.GetValue(QueryString, 0, 1))) > 0) {

        //						MDC_Com.MDC_GF_Message(ref "재고가 존재하는 작번입니다. 삭제할 수 없습니다.", ref "W");
        //						functionReturnValue = false;
        //						goto PS_PP049_Validate_Exit;

        //					}

        //				}

        //			}
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 PSMT인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 외주인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 실적인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 조정인경
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50") {

        //		}
        //	} else if (ValidateType == "수정01") {
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 일반인경우
        //		if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10") {
        //			////수정전 수정가능여부검사
        //			//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(oMat01Row01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////새로추가된 행인경우, 수정하여도 무방하다
        //			if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.VALUE))) {
        //			} else {
        //				//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(oMat01Row01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////분말
        //				if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.VALUE == "111" | oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.VALUE == "601") {
        //					//UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(oMat01Row01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oMat01.Columns.Item("CpCode").Cells.Item(oMat01Row01).Specific.VALUE == "CP80111" | oMat01.Columns.Item("CpCode").Cells.Item(oMat01Row01).Specific.VALUE == "CP80101") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						LineNum = oMat01.Columns.Item("LineNum").Cells.Item(oMat01Row01).Specific.VALUE;

        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue(select U_pqty from [PS_PP040L] where DocEntry =' & DocEntry & ' and u_linenum =' & LineNum & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (Strings.Trim(oMat01.Columns.Item("PQty").Cells.Item(oMat01Row01).Specific.VALUE) != MDC_PS_Common.GetValue("select U_pqty from [@PS_PP040L] where DocEntry ='" + DocEntry + "' and u_linenum ='" + LineNum + "'")) {
        //							SubMain.Sbo_Application.MessageBox("원자재 불출이 진행된 행은 생산수량을 수정할 수 없습니다.");
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //					}
        //				}
        //				//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(oMat01Row01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////휘팅이면서
        //				if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.VALUE == "101") {
        //					////현재공정이 실적공정이면
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP049_05 ' & Val(oMat01.Columns(PP030HNo).Cells(oMat01Row01).Specific.VALUE) & - & Val(oMat01.Columns(PP030MNo).Cells(oMat01Row01).Specific.VALUE) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					////현재 공정이 바렐 앞공정이면..
        //					if (MDC_PS_Common.GetValue("EXEC PS_PP049_05 '" + Conversion.Val(oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.VALUE) + "-" + Conversion.Val(oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.VALUE) + "'", 0, 1) == "Y") {
        //						////휘팅벌크포장
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP070L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.VALUE) + "'", 0, 1)) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "수정된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //						////휘팅실적
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.VALUE) + "' AND PS_PP080L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.VALUE) + "'", 0, 1)) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "수정된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //					}
        //				}
        //			}
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 PSMT인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 외주인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 실적인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 조정인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50") {
        //		}
        //	} else if (ValidateType == "취소") {
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP040H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
        //			MDC_Com.MDC_GF_Message(ref "이미취소된 문서 입니다. 취소할수 없습니다.", ref "W");
        //			functionReturnValue = false;
        //			goto PS_PP049_Validate_Exit;
        //		}
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 일반인경우
        //		if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10") {
        //			////삭제된 행에 대한처리
        //			Query01 = "SELECT ";
        //			Query01 = Query01 + " PS_PP040H.DocEntry,";
        //			Query01 = Query01 + " PS_PP040L.LineId,";
        //			Query01 = Query01 + " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
        //			Query01 = Query01 + " PS_PP040L.U_OrdGbn AS OrdGbn,";
        //			Query01 = Query01 + " PS_PP040L.U_PP030HNo AS PP030HNo,";
        //			Query01 = Query01 + " PS_PP040L.U_PP030MNo AS PP030MNo,";
        //			Query01 = Query01 + " PS_PP040L.U_OrdMgNum AS OrdMgNum ";
        //			Query01 = Query01 + " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Query01 = Query01 + " WHERE PS_PP040L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'";
        //			RecordSet01.DoQuery(Query01);
        //			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //				////휘팅이면서
        //				if (RecordSet01.Fields.Item("OrdGbn").Value == "101") {
        //					////현재공정이 실적포인트이면
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP049_05 ' & RecordSet01.Fields(OrdMgNum).VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (MDC_PS_Common.GetValue("EXEC PS_PP049_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y") {
        //						////휘팅벌크포장
        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "생산실적 등록된 문서입니다. 적용할수 없습니다.", ref "W");
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //						////휘팅실적
        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "생산실적 등록된 문서입니다. 적용할수 없습니다.", ref "W");
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //					}
        //				}

        //				////기계공구,몰드
        //				if (RecordSet01.Fields.Item("OrdGbn").Value == "105" | RecordSet01.Fields.Item("OrdGbn").Value == "106") {
        //					////그냥 입력가능
        //				////휘팅,부품
        //				} else if (RecordSet01.Fields.Item("OrdGbn").Value == "101" | RecordSet01.Fields.Item("OrdGbn").Value == "102") {
        //					////삭제된 행에 대한 검사..
        //					//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					OrdMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
        //					//// DocEntry + '-' + LineId
        //					CurrentCpInfo = OrdMgNum;

        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP049_02 '" + OrdMgNum + "'");
        //					if (string.IsNullOrEmpty(PrevCpInfo)) {
        //						////해당공정이 첫공정이면 입력되어도 상관없다.
        //					} else {
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'");
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");

        //						PrevMATRIXCpQty = 0;
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'");
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");
        //						CurrentMATRIXCpQty = 0;
        //						if (((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))) {
        //							SubMain.Sbo_Application.SetStatusBarMessage("취소문서의 선행공정의 생산수량이 취소문서의 생산수량을 미달합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //							functionReturnValue = false;
        //							goto PS_PP049_Validate_Exit;
        //						}
        //					}

        //					//                    If oForm.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
        //					//                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP049_03 '" & OrdMgNum & "'")
        //					//                        If NextCpInfo = "" Then
        //					//                            '//해당공정이 마지막공정이면 삭제되어도 상관없다.
        //					//                        Else
        //					//                            NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
        //					//                            NextMATRIXCpQty = 0
        //					//                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
        //					//                            CurrentMATRIXCpQty = 0
        //					//                            If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
        //					//                                Sbo_Application.SetStatusBarMessage "취소문서의 후행공정의 생산수량이 취소문서의 생산수량을 초과합니다.", bmt_Short, True
        //					//                                PS_PP049_Validate = False
        //					//                                GoTo PS_PP049_Validate_Exit
        //					//                            End If
        //					//                        End If
        //					//                    End If
        //				}
        //				RecordSet01.MoveNext();
        //			}
        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 PSMT인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {

        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 외주인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {

        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 실적인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {

        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////작업타입이 조정인경우
        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50") {

        //		}
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_PP049_Validate_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_PP049_Validate_Error:
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_PP049_LoadOrderInfo
        //private void PS_PP049_LoadOrderInfo()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	////일반,조정, 설계
        //	if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "10" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "50" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "60" | oForm.Items.Item("OrdType").Specific.Selected.VALUE == "70") {
        //		//UPGRADE_WARNING: oForm.Items(OrdMgNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.VALUE)) {
        //			MDC_Com.MDC_GF_Message(ref "작업지시 관리번호를 입력하지 않습니다.", ref "W");
        //			goto PS_PP049_LoadOrderInfo_Exit;
        //		} else {
        //			Query01 = "SELECT ";
        //			Query01 = Query01 + "U_OrdGbn,";
        //			Query01 = Query01 + "U_BPLId,";
        //			Query01 = Query01 + "U_ItemCode,";
        //			Query01 = Query01 + "U_ItemName,";
        //			Query01 = Query01 + "U_OrdNum,";
        //			Query01 = Query01 + "U_OrdSub1,";
        //			Query01 = Query01 + "U_OrdSub2,";
        //			Query01 = Query01 + "DocEntry";
        //			Query01 = Query01 + " FROM [@PS_PP030H]";
        //			Query01 = Query01 + " WHERE ";
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Query01 = Query01 + " U_OrdNum + U_OrdSub1 + U_OrdSub2 = '" + oForm.Items.Item("OrdMgNum").Specific.VALUE + "'";
        //			Query01 = Query01 + " AND U_OrdGbn NOT IN('104','107') ";
        //			Query01 = Query01 + " AND Canceled = 'N'";
        //			RecordSet01.DoQuery(Query01);
        //			if (RecordSet01.RecordCount == 0) {
        //				MDC_Com.MDC_GF_Message(ref "작업지시 정보가 존재하지 않습니다.", ref "W");
        //				goto PS_PP049_LoadOrderInfo_Exit;
        //			} else {
        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("OrdGbn").Specific.Select(RecordSet01.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("BPLId").Specific.Select(RecordSet01.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
        //				//UPGRADE_WARNING: oForm.Items(ItemCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("ItemCode").Specific.VALUE = RecordSet01.Fields.Item(2).Value;
        //				//UPGRADE_WARNING: oForm.Items(ItemName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("ItemName").Specific.VALUE = RecordSet01.Fields.Item(3).Value;
        //				//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("OrdNum").Specific.VALUE = RecordSet01.Fields.Item(4).Value;
        //				//UPGRADE_WARNING: oForm.Items(OrdSub1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("OrdSub1").Specific.VALUE = RecordSet01.Fields.Item(5).Value;
        //				//UPGRADE_WARNING: oForm.Items(OrdSub2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("OrdSub2").Specific.VALUE = RecordSet01.Fields.Item(6).Value;
        //				//UPGRADE_WARNING: oForm.Items(PP030HNo).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("PP030HNo").Specific.VALUE = RecordSet01.Fields.Item(7).Value;
        //				//                '//매트릭스삭제
        //				//                oMat01.Clear
        //				//                oMat01.FlushToDataSource
        //				//                oMat01.LoadFromDataSource
        //				//                Call PS_PP049_AddMatrixRow01(0, True)
        //				//                oMat02.Clear
        //				//                oMat02.FlushToDataSource
        //				//                oMat02.LoadFromDataSource
        //				//                Call PS_PP049_AddMatrixRow02(0, True)
        //				//                oMat03.Clear
        //				//                oMat03.FlushToDataSource
        //				//                oMat03.LoadFromDataSource
        //				oForm.Update();
        //			}
        //		}
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	////PSMT
        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "20") {
        //		//UPGRADE_WARNING: oForm.Items(OrdMgNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.VALUE)) {
        //			MDC_Com.MDC_GF_Message(ref "작업지시 관리번호를 입력하지 않습니다.", ref "W");
        //			goto PS_PP049_LoadOrderInfo_Exit;
        //		} else {
        //			//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("OrdNum").Specific.VALUE = oForm.Items.Item("OrdMgNum").Specific.VALUE;
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("OrdSub1").Specific.VALUE = "000";
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("OrdSub2").Specific.VALUE = "00";
        //			////매트릭스삭제
        //			oMat01.Clear();
        //			oMat01.FlushToDataSource();
        //			oMat01.LoadFromDataSource();
        //			PS_PP049_AddMatrixRow01(0, ref true);
        //			oMat02.Clear();
        //			oMat02.FlushToDataSource();
        //			oMat02.LoadFromDataSource();
        //			PS_PP049_AddMatrixRow02(0, ref true);
        //			oMat03.Clear();
        //			oMat03.FlushToDataSource();
        //			oMat03.LoadFromDataSource();
        //			oForm.Update();
        //		}
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "30") {
        //		MDC_Com.MDC_GF_Message(ref "외주은 입력할수 없습니다.", ref "W");
        //		goto PS_PP049_LoadOrderInfo_Exit;
        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.VALUE == "40") {
        //		MDC_Com.MDC_GF_Message(ref "실적은 입력할수 없습니다.", ref "W");
        //		goto PS_PP049_LoadOrderInfo_Exit;
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	PS_PP049_LoadOrderInfo_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	PS_PP049_LoadOrderInfo_Error:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP049_LoadOrderInfo_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP049_FindValidateDocument
        //public bool PS_PP049_FindValidateDocument(string ObjectType)
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string Query02 = null;
        //	SAPbobsCOM.Recordset RecordSet02 = null;

        //	int i = 0;
        //	string DocEntry = null;
        //	//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
        //	////원본문서

        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	Query01 = " SELECT DocEntry";
        //	Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry = ";
        //	Query01 = Query01 + DocEntry;
        //	if ((oDocType01 == "작업일보등록(작지)")) {
        //		Query01 = Query01 + " AND U_DocType = '10'";
        //	} else if ((oDocType01 == "작업일보등록(공정)")) {
        //		Query01 = Query01 + " AND U_DocType = '20'";
        //	}
        //	RecordSet01.DoQuery(Query01);
        //	if ((RecordSet01.RecordCount == 0)) {
        //		if ((oDocType01 == "작업일보등록(작지)")) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		} else if ((oDocType01 == "작업일보등록(공정)")) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("작업일보등록(작지)문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        //		functionReturnValue = false;
        //		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet01 = null;
        //		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet02 = null;
        //		return functionReturnValue;
        //	}

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return functionReturnValue;
        //	PS_PP049_FindValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	functionReturnValue = false;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_PP049_DirectionValidateDocument
        //public bool PS_PP049_DirectionValidateDocument(string DocEntry, string DocEntryNext, string Direction, string ObjectType)
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string Query02 = null;
        //	SAPbobsCOM.Recordset RecordSet02 = null;

        //	int i = 0;
        //	string MaxDocEntry = null;
        //	string MinDocEntry = null;
        //	bool DoNext = false;
        //	bool IsFirst = false;
        //	////시작유무
        //	DoNext = true;
        //	IsFirst = true;

        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	while ((DoNext == true)) {
        //		if ((IsFirst != true)) {
        //			////문서전체를 경유하고도 유효값을 찾지못했다면
        //			if ((DocEntry == DocEntryNext)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				functionReturnValue = false;
        //				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet01 = null;
        //				//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet02 = null;
        //				return functionReturnValue;
        //			}
        //		}
        //		if ((Direction == "Next")) {
        //			Query01 = " SELECT TOP 1 DocEntry";
        //			Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
        //			Query01 = Query01 + DocEntryNext;
        //			if ((oDocType01 == "작업일보등록(작지)")) {
        //				Query01 = Query01 + " AND U_DocType = '10'";
        //			} else if ((oDocType01 == "작업일보등록(공정)")) {
        //				Query01 = Query01 + " AND U_DocType = '20'";
        //			}
        //			Query01 = Query01 + " ORDER BY DocEntry ASC";
        //		} else if ((Direction == "Prev")) {
        //			Query01 = " SELECT TOP 1 DocEntry";
        //			Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //			Query01 = Query01 + DocEntryNext;
        //			if ((oDocType01 == "작업일보등록(작지)")) {
        //				Query01 = Query01 + " AND U_DocType = '10'";
        //			} else if ((oDocType01 == "작업일보등록(공정)")) {
        //				Query01 = Query01 + " AND U_DocType = '20'";
        //			}
        //			Query01 = Query01 + " ORDER BY DocEntry DESC";
        //		}
        //		RecordSet01.DoQuery(Query01);
        //		////해당문서가 마지막문서라면
        //		if ((RecordSet01.Fields.Item(0).Value == 0)) {
        //			if ((Direction == "Next")) {
        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //				if ((oDocType01 == "작업일보등록(작지)")) {
        //					Query02 = Query02 + " WHERE U_DocType = '10'";
        //				} else if ((oDocType01 == "작업일보등록(공정)")) {
        //					Query02 = Query02 + " WHERE U_DocType = '20'";
        //				}
        //				Query02 = Query02 + " ORDER BY DocEntry ASC";
        //			} else if ((Direction == "Prev")) {
        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //				if ((oDocType01 == "작업일보등록(작지)")) {
        //					Query02 = Query02 + " WHERE U_DocType = '10'";
        //				} else if ((oDocType01 == "작업일보등록(공정)")) {
        //					Query02 = Query02 + " WHERE U_DocType = '20'";
        //				}
        //				Query02 = Query02 + " ORDER BY DocEntry DESC";
        //			}
        //			RecordSet02.DoQuery(Query02);
        //			////문서가 아예 존재하지 않는다면
        //			if ((RecordSet02.RecordCount == 0)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet01 = null;
        //				//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet02 = null;
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			} else {
        //				if ((Direction == "Next")) {
        //					DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) - 1);
        //					Query01 = " SELECT TOP 1 DocEntry";
        //					Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
        //					Query01 = Query01 + DocEntryNext;
        //					if ((oDocType01 == "작업일보등록(작지)")) {
        //						Query01 = Query01 + " AND U_DocType = '10'";
        //					} else if ((oDocType01 == "작업일보등록(공정)")) {
        //						Query01 = Query01 + " AND U_DocType = '20'";
        //					}
        //					Query01 = Query01 + " ORDER BY DocEntry ASC";
        //					RecordSet01.DoQuery(Query01);
        //				} else if ((Direction == "Prev")) {
        //					DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) + 1);
        //					Query01 = " SELECT TOP 1 DocNum";
        //					Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //					Query01 = Query01 + DocEntryNext;
        //					if ((oDocType01 == "작업일보등록(작지)")) {
        //						Query01 = Query01 + " AND U_DocType = '10'";
        //					} else if ((oDocType01 == "작업일보등록(공정)")) {
        //						Query01 = Query01 + " AND U_DocType = '20'";
        //					}
        //					Query01 = Query01 + " ORDER BY DocEntry DESC";
        //					RecordSet01.DoQuery(Query01);
        //				}
        //			}
        //		}
        //		if ((oDocType01 == "작업일보등록(작지)")) {
        //			DoNext = false;
        //			if ((Direction == "Next")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
        //			} else if ((Direction == "Prev")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
        //			}
        //		} else if ((oDocType01 == "작업일보등록(공정)")) {
        //			DoNext = false;
        //			if ((Direction == "Next")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
        //			} else if ((Direction == "Prev")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
        //			}
        //		}
        //		IsFirst = false;
        //	}
        //	////다음문서가 유효하다면 그냥 넘어가고
        //	if ((DocEntry == DocEntryNext)) {
        //		PS_PP049_EnableFormItem();
        //		////UDO방식
        //	////다음문서가 유효하지 않다면
        //	} else {
        //		oForm.Freeze(true);
        //		oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //		PS_PP049_EnableFormItem();
        //		////UDO방식
        //		////문서번호 필드가 입력이 가능하다면
        //		if (oForm.Items.Item("DocEntry").Enabled == true) {
        //			if ((Direction == "Next")) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.VALUE = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) + 1));
        //			} else if ((Direction == "Prev")) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.VALUE = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) - 1));
        //			}
        //			oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		}
        //		oForm.Freeze(false);
        //		functionReturnValue = false;
        //		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet01 = null;
        //		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet02 = null;
        //		return functionReturnValue;
        //	}
        //	functionReturnValue = true;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return functionReturnValue;
        //	PS_PP049_DirectionValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return functionReturnValue;
        //}
        #endregion

        #region Add_oInventoryGenExit
        //private bool Add_oInventoryGenExit(ref short ChkType)
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbobsCOM.Documents DI_oInventoryGenExit = null;
        //	////재고출고 문서 객체

        //	int j = 0;
        //	int i = 0;
        //	int Cnt = 0;
        //	short ErrNum = 0;
        //	int ErrCode = 0;
        //	string ErrMsg = null;
        //	int RetVal = 0;

        //	string BatchNum = null;
        //	string CItemCod = null;
        //	string DocDate = null;
        //	string DocNum = null;
        //	string CpCode = null;
        //	string WhsCode = null;
        //	int IssueQty = 0;
        //	decimal IssueWt = default(decimal);
        //	string sDocEntry = null;
        //	SAPbobsCOM.Recordset oRecordSet = null;
        //	string sQry = null;

        //	Cnt = 0;

        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	oMat01.FlushToDataSource();

        //	DocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_PS_PP049H.GetValue("U_DocDate", 0), "0000-00-00");
        //	//[If ChkType = 2 Then Call FormClear
        //	DocNum = Strings.Trim(oDS_PS_PP049H.GetValue("DocEntry", 0));


        //	//UPGRADE_WARNING: oMat01.Columns(OutDoc).Cells(1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oMat01.Columns.Item("OutDoc").Cells.Item(1).Specific.VALUE)) {
        //		SubMain.Sbo_Company.StartTransaction();
        //		//UPGRADE_NOTE: DI_oInventoryGenExit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		DI_oInventoryGenExit = null;
        //		DI_oInventoryGenExit = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

        //		var _with1 = DI_oInventoryGenExit;
        //		_with1.DocDate = Convert.ToDateTime(DocDate);
        //		_with1.TaxDate = Convert.ToDateTime(DocDate);
        //		_with1.Comments = "원재료 불출 등록(" + DocNum + ") 출고";

        //		j = 0;
        //		for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			IssueQty = oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			IssueWt = oMat01.Columns.Item("PWeight").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i + 1).Specific.VALUE;

        //			WhsCode = "101";

        //			if ((CpCode == "CP80101" | CpCode == "CP80111") & !string.IsNullOrEmpty(CItemCod) & IssueQty >= 0 & IssueWt != 0 & !string.IsNullOrEmpty(WhsCode)) {
        //				//If (CpCode = "CP80101" Or CpCode = "CP80111" Or CpCode = "CP80104" Or CpCode = "CP80105") And CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
        //				if (j > 0)
        //					_with1.Lines.Add();
        //				_with1.Lines.SetCurrentLine(j);
        //				_with1.Lines.ItemCode = CItemCod;
        //				_with1.Lines.WarehouseCode = WhsCode;
        //				_with1.Lines.BatchNumbers.BatchNumber = BatchNum;
        //				_with1.Lines.BatchNumbers.Quantity = IssueWt;
        //				_with1.Lines.Quantity = IssueWt;
        //				_with1.Lines.UserFields.Fields.Item("U_Qty").VALUE = IssueQty;
        //				Cnt = Cnt + 1;
        //				j = j + 1;
        //			}
        //		}

        //		//// 완료
        //		if (Cnt > 0) {
        //			RetVal = DI_oInventoryGenExit.Add();
        //			if ((0 != RetVal)) {
        //				SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
        //				ErrNum = 1;
        //				goto Add_oInventoryGenExit_Error;
        //			}
        //		}

        //		if (ChkType == 1) {
        //			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //		} else if (ChkType == 2) {
        //			SubMain.Sbo_Company.GetNewObjectCode(out sDocEntry);
        //			Cnt = 1;
        //			for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.VALUE;
        //				if (CpCode == "CP80101" | CpCode == "CP80111") {
        //					//If CpCode = "CP80101" Or CpCode = "CP80111" Or CpCode = "CP80104" Or CpCode = "CP80105" Then
        //					oDS_PS_PP049L.SetValue("U_OutDoc", i, sDocEntry);
        //					oDS_PS_PP049L.SetValue("U_OutLin", i, Convert.ToString(Cnt));
        //					Cnt = Cnt + 1;
        //				}
        //			}
        //			oMat01.LoadFromDataSource();
        //			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //		}
        //	}

        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	//UPGRADE_NOTE: DI_oInventoryGenExit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	DI_oInventoryGenExit = null;
        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	Add_oInventoryGenExit_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	//UPGRADE_NOTE: DI_oInventoryGenExit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	DI_oInventoryGenExit = null;
        //	if (SubMain.Sbo_Company.InTransaction)
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	functionReturnValue = false;
        //	if (ErrNum == 1) {
        //		MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenExit_Error:" + ErrCode + " - " + ErrMsg, ref "E");
        //	} else {
        //		MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenExit_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //	return functionReturnValue;
        //}
        #endregion

        #region Add_oInventoryGenEntry
        //private bool Add_oInventoryGenEntry(ref short ChkType)
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbobsCOM.Documents DI_oInventoryGenEntry = null;
        //	////재고출고 문서 객체

        //	int j = 0;
        //	int i = 0;
        //	int Cnt = 0;
        //	short ErrNum = 0;
        //	int ErrCode = 0;
        //	string ErrMsg = null;
        //	int RetVal = 0;

        //	string BatchNum = null;
        //	string CItemCod = null;
        //	string DocDate = null;
        //	string DocNum = null;
        //	string CpCode = null;
        //	string WhsCode = null;
        //	int IssueQty = 0;
        //	decimal IssueWt = default(decimal);
        //	string sDocEntry = null;
        //	SAPbobsCOM.Recordset oRecordSet = null;
        //	string sQry = null;
        //	string OIGEDoc = null;

        //	Cnt = 0;

        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	oMat01.FlushToDataSource();

        //	DocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_PS_PP049H.GetValue("U_DocDate", 0), "0000-00-00");
        //	DocNum = Strings.Trim(oDS_PS_PP049H.GetValue("DocEntry", 0));
        //	//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	OIGEDoc = Strings.Trim(oMat01.Columns.Item("OutDoc").Cells.Item(1).Specific.VALUE);

        //	//UPGRADE_WARNING: oMat01.Columns(OutDocC).Cells(1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oMat01.Columns.Item("OutDocC").Cells.Item(1).Specific.VALUE)) {
        //		SubMain.Sbo_Company.StartTransaction();
        //		//UPGRADE_NOTE: DI_oInventoryGenEntry 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		DI_oInventoryGenEntry = null;
        //		DI_oInventoryGenEntry = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

        //		var _with2 = DI_oInventoryGenEntry;
        //		_with2.DocDate = Convert.ToDateTime(DocDate);
        //		_with2.TaxDate = Convert.ToDateTime(DocDate);
        //		_with2.Comments = "원재료 불출 등록 출고 취소 (" + DocNum + ") 입고";

        //		_with2.UserFields.Fields.Item("U_CancDoc").VALUE = OIGEDoc;
        //		////입고취소 문서번호

        //		j = 0;
        //		for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			IssueQty = oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			IssueWt = oMat01.Columns.Item("PWeight").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.VALUE;
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i + 1).Specific.VALUE;

        //			WhsCode = "101";

        //			if ((CpCode == "CP80101" | CpCode == "CP80111") & !string.IsNullOrEmpty(CItemCod) & IssueQty >= 0 & IssueWt != 0 & !string.IsNullOrEmpty(WhsCode)) {
        //				//If (CpCode = "CP80101" Or CpCode = "CP80111" Or CpCode = "CP80104" Or CpCode = "CP80105") And CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
        //				if (j > 0)
        //					_with2.Lines.Add();
        //				_with2.Lines.SetCurrentLine(j);
        //				_with2.Lines.ItemCode = CItemCod;
        //				_with2.Lines.WarehouseCode = WhsCode;
        //				_with2.Lines.BatchNumbers.BatchNumber = BatchNum;
        //				_with2.Lines.BatchNumbers.Quantity = IssueWt;
        //				_with2.Lines.Quantity = IssueWt;
        //				_with2.Lines.UserFields.Fields.Item("U_Qty").VALUE = IssueQty;
        //				Cnt = Cnt + 1;

        //				j = j + 1;
        //			}
        //		}

        //		//// 완료
        //		if (Cnt > 0) {
        //			RetVal = DI_oInventoryGenEntry.Add();
        //			if ((0 != RetVal)) {
        //				SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
        //				ErrNum = 1;
        //				goto Add_oInventoryGenEntry_Error;
        //			}
        //		}


        //		if (ChkType == 1) {
        //			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //		} else if (ChkType == 2) {
        //			SubMain.Sbo_Company.GetNewObjectCode(out sDocEntry);
        //			//cnt = 1
        //			//For i = 0 To oMat01.VisualRowCount - 1
        //			//     CpCode = oMat01.Columns("CpCode").Cells(i + 1).Specific.VALUE
        //			//     If CpCode = "CP80101" Or CpCode = "CP80111" Then
        //			//         oDS_PS_PP049L.setValue "U_OutDocC", i, sDocEntry
        //			//         oDS_PS_PP049L.setValue "U_OutLinC", i, cnt
        //			//         cnt = cnt + 1
        //		}
        //		// Next i
        //		oMat01.LoadFromDataSource();
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

        //		sQry = "Update [@PS_PP040L] set U_OutDocC = '" + sDocEntry + "', U_OutLinC = U_OutLin";
        //		//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sQry = sQry + " From [@PS_PP040L] where 1=1 and u_cpcode in ('CP80101','CP80111') and docentry = '" + Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE) + "' ";
        //		oRecordSet.DoQuery(sQry);
        //		//End If
        //	}

        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	//UPGRADE_NOTE: DI_oInventoryGenEntry 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	DI_oInventoryGenEntry = null;
        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	Add_oInventoryGenEntry_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	//UPGRADE_NOTE: DI_oInventoryGenEntry 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	DI_oInventoryGenEntry = null;
        //	if (SubMain.Sbo_Company.InTransaction)
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	functionReturnValue = false;
        //	if (ErrNum == 1) {
        //		MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenEntry_Error:" + ErrCode + " - " + ErrMsg, ref "E");
        //	} else {
        //		MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenEntry_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //	return functionReturnValue;
        //}
        #endregion
    }
}
