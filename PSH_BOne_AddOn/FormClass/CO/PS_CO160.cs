using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 재공 원가 이동등록
	/// </summary>
	internal class PS_CO160 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
			
		private SAPbouiCOM.DBDataSource oDS_PS_CO160H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO160L; //등록라인

		//private string oDocType01;
			
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		
		private string oDocEntry;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{

			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO160.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO160_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO160");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);
                //oDocType01 = oFromDocType01;
                PS_CO160_CreateItems();
                PS_CO160_ComboBox_Setting();
                //PS_CO160_EnableMenus();
                //PS_CO160_SetDocument(oFromDocEntry01);
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

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_CO160_CreateItems()
        {
            try
            {
                //oForm.Freeze(true);

                oDS_PS_CO160H = oForm.DataSources.DBDataSources.Item("@PS_CO160H");
                oDS_PS_CO160L = oForm.DataSources.DBDataSources.Item("@PS_CO160L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
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

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_CO160_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //oForm.Freeze(true);

                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

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
        //				if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {

        //					if (SubMain.Sbo_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1")) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //				} else {
        //					MDC_Com.MDC_GF_Message(ref "현재 모드에서는 취소할수 없습니다.", ref "W");
        //					BubbleEvent = false;
        //					return;
        //				}
        //				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocEntry = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE);
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
        //	////BeforeAction = False
        //	} else if ((pval.BeforeAction == false)) {
        //		switch (pval.MenuUID) {
        //			case "1284":
        //				//취소
        //				MDC_PS_Common.DoQuery("EXEC PS_CO160_03 '" + oDocEntry + "', 'D'");
        //				//삭제
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
        //				PS_CO160_FormItemEnabled();
        //				////UDO방식
        //				oForm01.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			case "1282":
        //				//추가
        //				PS_CO160_FormItemEnabled();
        //				////UDO방식
        //				PS_CO160_AddMatrixRow(0, ref true);
        //				////UDO방식
        //				break;
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
        //				if ((oForm01.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //					if ((PS_CO160_FindValidateDocument("@PS_CO160H") == false)) {
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
        //		if (pval.ItemUID == "Mat01") {
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

        //	} else if (pval.BeforeAction == false) {
        //		if (pval.ItemUID == "Mat01") {
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

        //	int DocEntry = 0;
        //	int i = 0;
        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "1") {
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_CO160_DataValidCheck() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocEntry = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE);
        //				////해야할일 작업
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (PS_CO160_DataValidCheck() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDocEntry = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE);
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}


        //	} else if (pval.BeforeAction == false) {
        //		if (pval.ItemUID == "1") {
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pval.ActionSuccess == true) {
        //					MDC_PS_Common.DoQuery("EXEC PS_CO160_03 '" + oDocEntry + "', 'I'");
        //					//입력
        //					PS_CO160_FormItemEnabled();
        //					PS_CO160_AddMatrixRow(oMat01.RowCount, ref true);
        //					////UDO방식일때
        //				}
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (pval.ActionSuccess == true) {
        //					MDC_PS_Common.DoQuery("EXEC PS_CO160_03 '" + oDocEntry + "', 'U'");
        //					//갱신
        //				}
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				if (pval.ActionSuccess == true) {
        //					PS_CO160_FormItemEnabled();
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
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		if (pval.CharPressed == 9) {
        //			if (pval.ItemUID == "ItemCode") {
        //				//UPGRADE_WARNING: oForm01.Items(ItemCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm01.Items.Item("ItemCode").Specific.VALUE)) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //					BubbleEvent = false;
        //				}
        //			}
        //			if (pval.ItemUID == "MItemCod") {
        //				//UPGRADE_WARNING: oForm01.Items(MItemCod).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm01.Items.Item("MItemCod").Specific.VALUE)) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //					BubbleEvent = false;
        //				}
        //			}
        //			if (pval.ItemUID == "Mat01") {
        //				if (pval.ColUID == "PO") {
        //					//UPGRADE_WARNING: oMat01.Columns(PO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE)) {
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
        //					}
        //				}
        //				if (pval.ColUID == "MPO") {
        //					//UPGRADE_WARNING: oMat01.Columns(MPO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE)) {
        //						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //						BubbleEvent = false;
        //					}
        //				}
        //			}

        //		}
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        //		//        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrderNum") '//사용자값활성
        //		//Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "WhsCode") '//사용자값활성
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
        //		if (pval.ItemUID == "Mat01") {
        //			if (pval.Row > 0) {
        //				oMat01.SelectRow(pval.Row, true, false);
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
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string ItemCode01 = null;
        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemChanged == true) {
        //			if ((pval.ItemUID == "Mat01")) {

        //				if (pval.ColUID == "PO") {
        //					//UPGRADE_WARNING: oMat01.Columns(PO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE)) {
        //						goto Raise_EVENT_VALIDATE_Exit;
        //					}
        //					for (i = 1; i <= oMat01.RowCount; i++) {
        //						////현재 선택되어있는 행이 아니면
        //						if (pval.Row != i) {
        //							//UPGRADE_WARNING: oMat01.Columns(PO).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns(PO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if ((oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE == oMat01.Columns.Item("PO").Cells.Item(i).Specific.VALUE)) {
        //								MDC_Com.MDC_GF_Message(ref "동일한 항목이 존재합니다.", ref "W");
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE = "";
        //								goto Raise_EVENT_VALIDATE_Exit;
        //							}
        //							//                            If (Mid(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, "-") - 1) <> _
        //							//'                            Mid(oMat01.Columns("OrderNum").Cells(i).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(i).Specific.Value, "-") - 1)) Then
        //							//                                Call MDC_Com.MDC_GF_Message("동일하지않은 수주문서가 존재합니다.", "W")
        //							//                                oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value = ""
        //							//                                GoTo Raise_EVENT_VALIDATE_Exit
        //							//                            End If
        //						}
        //					}



        //					RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					//UPGRADE_WARNING: oMat01.Columns(PO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm01.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					Query01 = "EXEC PS_CO160_01 '" + oForm01.Items.Item("BPLId").Specific.VALUE + "','" + oForm01.Items.Item("YM").Specific.VALUE + "','" + oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE + "'";
        //					RecordSet01.DoQuery(Query01);
        //					for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //						oDS_PS_CO160L.SetValue("U_PO", pval.Row - 1, RecordSet01.Fields.Item("PO").Value);
        //						oDS_PS_CO160L.SetValue("U_POEntry", pval.Row - 1, RecordSet01.Fields.Item("POEntry").Value);
        //						oDS_PS_CO160L.SetValue("U_POLine", pval.Row - 1, RecordSet01.Fields.Item("POLine").Value);
        //						oDS_PS_CO160L.SetValue("U_Sequence", pval.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
        //						oDS_PS_CO160L.SetValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
        //						oDS_PS_CO160L.SetValue("U_ItemName", pval.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
        //						oDS_PS_CO160L.SetValue("U_CpCode", pval.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
        //						oDS_PS_CO160L.SetValue("U_CpName", pval.Row - 1, RecordSet01.Fields.Item("CpName").Value);
        //						oDS_PS_CO160L.SetValue("U_StcQty", pval.Row - 1, RecordSet01.Fields.Item("StcQty").Value);
        //						oDS_PS_CO160L.SetValue("U_StcAmt", pval.Row - 1, RecordSet01.Fields.Item("StcAmt").Value);
        //						RecordSet01.MoveNext();
        //					}
        //					if (oMat01.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_CO160L.GetValue("U_PO", pval.Row - 1)))) {
        //						PS_CO160_AddMatrixRow((pval.Row));
        //					}
        //					oMat01.LoadFromDataSource();
        //					oMat01.AutoResizeColumns();


        //					oForm01.Update();
        //					//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					RecordSet01 = null;
        //				} else if (pval.ColUID == "MPO") {
        //					//UPGRADE_WARNING: oMat01.Columns(MPO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE)) {
        //						goto Raise_EVENT_VALIDATE_Exit;
        //					}
        //					for (i = 1; i <= oMat01.RowCount; i++) {
        //						////현재 선택되어있는 행이 아니면
        //						if (pval.Row != i) {
        //							//UPGRADE_WARNING: oMat01.Columns(MPO).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns(MPO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if ((oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE == oMat01.Columns.Item("MPO").Cells.Item(i).Specific.VALUE)) {
        //								MDC_Com.MDC_GF_Message(ref "동일한 항목이 존재합니다.", ref "W");
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE = "";
        //								goto Raise_EVENT_VALIDATE_Exit;
        //							}
        //							//                            If (Mid(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, "-") - 1) <> _
        //							//'                            Mid(oMat01.Columns("OrderNum").Cells(i).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(i).Specific.Value, "-") - 1)) Then
        //							//                                Call MDC_Com.MDC_GF_Message("동일하지않은 수주문서가 존재합니다.", "W")
        //							//                                oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value = ""
        //							//                                GoTo Raise_EVENT_VALIDATE_Exit
        //							//                            End If
        //						}
        //					}
        //					RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					//UPGRADE_WARNING: oMat01.Columns(MPO).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm01.Items(MItemCod).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm01.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					Query01 = "EXEC PS_CO160_02 '" + oForm01.Items.Item("BPLId").Specific.VALUE + "','" + oForm01.Items.Item("YM").Specific.VALUE + "','" + oForm01.Items.Item("MItemCod").Specific.VALUE + "','" + oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE + "'";
        //					RecordSet01.DoQuery(Query01);
        //					for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //						oDS_PS_CO160L.SetValue("U_MPO", pval.Row - 1, RecordSet01.Fields.Item("MPO").Value);
        //						oDS_PS_CO160L.SetValue("U_MPOEntry", pval.Row - 1, RecordSet01.Fields.Item("MPOEntry").Value);
        //						oDS_PS_CO160L.SetValue("U_MPOLine", pval.Row - 1, RecordSet01.Fields.Item("MPOLine").Value);
        //						oDS_PS_CO160L.SetValue("U_MSequenc", pval.Row - 1, RecordSet01.Fields.Item("MSequenc").Value);
        //						oDS_PS_CO160L.SetValue("U_MItemCod", pval.Row - 1, RecordSet01.Fields.Item("MItemCod").Value);
        //						oDS_PS_CO160L.SetValue("U_MItemNam", pval.Row - 1, RecordSet01.Fields.Item("MItemNam").Value);
        //						oDS_PS_CO160L.SetValue("U_MCpCode", pval.Row - 1, RecordSet01.Fields.Item("MCpCode").Value);
        //						oDS_PS_CO160L.SetValue("U_MCpName", pval.Row - 1, RecordSet01.Fields.Item("MCpName").Value);
        //						RecordSet01.MoveNext();
        //					}
        //					//                        If oMat01.RowCount = pval.Row And Trim(oDS_PS_CO160L.GetValue("U_MPO", pval.Row - 1)) <> "" Then
        //					//                            PS_CO160_AddMatrixRow (pval.Row)
        //					//                        End If
        //					oMat01.LoadFromDataSource();
        //					oMat01.AutoResizeColumns();


        //					oForm01.Update();
        //					//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					RecordSet01 = null;
        //				} else if (pval.ColUID == "Qty") {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE) <= 0)) {
        //						oDS_PS_CO160L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(0));
        //						oDS_PS_CO160L.SetValue("U_Weight", pval.Row - 1, Convert.ToString(0));
        //						oDS_PS_CO160L.SetValue("U_LinTotal", pval.Row - 1, Convert.ToString(0));
        //					} else {
        //						ItemCode01 = Strings.Trim(oDS_PS_CO160L.GetValue("U_ItemCode", pval.Row - 1));
        //						////EA자체품
        //						if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_CO160L.SetValue("U_Weight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pval.Row).Specific.VALUE)));
        //						////EAUOM
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_CO160L.SetValue("U_Weight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))));
        //						////KGSPEC
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_CO160L.SetValue("U_Weight", pval.Row - 1, Convert.ToString((Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pval.Row).Specific.VALUE)));
        //						////KG단중
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_CO160L.SetValue("U_Weight", pval.Row - 1, Convert.ToString(System.Math.Round(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pval.Row).Specific.VALUE) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
        //						////KG선택
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //						}
        //						oDS_PS_CO160L.SetValue("U_LinTotal", pval.Row - 1, Convert.ToString(Convert.ToDouble(Strings.Trim(oDS_PS_CO160L.GetValue("U_Weight", pval.Row - 1))) * Convert.ToDouble(Strings.Trim(oDS_PS_CO160L.GetValue("U_Price", pval.Row - 1)))));
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_CO160L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //					}

        //				} else {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_CO160L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
        //				}

        //				oMat01.LoadFromDataSource();
        //				oMat01.AutoResizeColumns();
        //				oForm01.Update();
        //				oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else {
        //				if ((pval.ItemUID == "DocEntry")) {
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_CO160H.SetValue(pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
        //				} else if ((pval.ItemUID == "ItemCode")) {
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_CO160H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_CO160H.SetValue("U_ItemName", 0, MDC_PS_Common.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'", 0, 1));
        //				} else if ((pval.ItemUID == "MItemCod")) {
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_CO160H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_CO160H.SetValue("U_MItemNam", 0, MDC_PS_Common.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'", 0, 1));
        //				} else {
        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_CO160H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.VALUE);
        //				}
        //			}
        //		}
        //	} else if (pval.BeforeAction == false) {

        //	}
        //	oForm01.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Exit:
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
        //		PS_CO160_FormItemEnabled();
        //		PS_CO160_AddMatrixRow(oMat01.VisualRowCount);
        //		////UDO방식
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
        //		if ((pval.ItemUID == "ItemCode" | pval.ItemUID == "ItemName")) {
        //			MDC_Com.MDC_GP_CF_DBDatasourceReturn(pval, (pval.FormUID), "@PS_CO160H", "U_ItemCode,U_ItemName");
        //		}
        //		if ((pval.ItemUID == "MItemCod" | pval.ItemUID == "MItemNam")) {
        //			MDC_Com.MDC_GP_CF_DBDatasourceReturn(pval, (pval.FormUID), "@PS_CO160H", "U_MItemCod,U_MItemNam");
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

        //	if (pval.ItemUID == "Mat01") {
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

        //	int i = 0;
        //	if ((oLastColRow01 > 0)) {
        //		if (pval.BeforeAction == true) {

        //		} else if (pval.BeforeAction == false) {
        //			for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_CO160L.RemoveRecord(oDS_PS_CO160L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0) {
        //				PS_CO160_AddMatrixRow(0);
        //			} else {
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_CO160L.GetValue("U_PO", oMat01.RowCount - 1)))) {
        //					PS_CO160_AddMatrixRow(oMat01.RowCount);
        //				}
        //			}

        //			oForm01.Update();
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion




        #region PS_CO160_FormItemEnabled
        //public void PS_CO160_FormItemEnabled()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	oForm01.Freeze(true);

        //	if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		////각모드에따른 아이템설정
        //		oForm01.Items.Item("DocEntry").Enabled = false;
        //		oForm01.Items.Item("BPLId").Enabled = true;
        //		oForm01.Items.Item("YM").Enabled = true;
        //		oForm01.Items.Item("ItemCode").Enabled = true;
        //		oForm01.Items.Item("MItemCod").Enabled = true;
        //		oForm01.Items.Item("Mat01").Enabled = true;
        //		oMat01.AutoResizeColumns();
        //		PS_CO160_FormClear();
        //		////UDO방식
        //		//        Call oForm01.Items("BPLId").Specific.Select("1", psk_ByValue)
        //		//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm01.Items.Item("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);
        //		////2010.12.06 추가



        //		oForm01.EnableMenu("1281", true);
        //		////찾기
        //		oForm01.EnableMenu("1282", false);
        //		////추가

        //	} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
        //		oForm01.Items.Item("DocEntry").Enabled = true;
        //		oForm01.Items.Item("BPLId").Enabled = true;
        //		oForm01.Items.Item("YM").Enabled = true;
        //		oForm01.Items.Item("ItemCode").Enabled = true;
        //		oForm01.Items.Item("MItemCod").Enabled = true;
        //		oForm01.Items.Item("Comment").Enabled = true;
        //		oForm01.Items.Item("Mat01").Enabled = false;
        //		oMat01.AutoResizeColumns();
        //		oForm01.EnableMenu("1281", false);
        //		oForm01.EnableMenu("1282", true);


        //		////각모드에따른 아이템설정
        //	} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
        //		oForm01.EnableMenu("1281", true);
        //		////찾기
        //		oForm01.EnableMenu("1282", true);
        //		////추가

        //		oForm01.Items.Item("DocEntry").Enabled = false;
        //		oForm01.Items.Item("BPLId").Enabled = false;
        //		oForm01.Items.Item("YM").Enabled = false;
        //		oForm01.Items.Item("ItemCode").Enabled = false;
        //		oForm01.Items.Item("MItemCod").Enabled = false;

        //		oMat01.AutoResizeColumns();
        //		oForm01.EnableMenu("1281", true);
        //		oForm01.EnableMenu("1282", false);

        //	}
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO160_FormItemEnabled_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO160_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO160_AddMatrixRow
        //public void PS_CO160_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm01.Freeze(true);
        //	////행추가여부
        //	if (RowIserted == false) {
        //		oDS_PS_CO160L.InsertRecord((oRow));
        //	}
        //	oMat01.AddRow();
        //	oDS_PS_CO160L.Offset = oRow;
        //	oDS_PS_CO160L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oMat01.LoadFromDataSource();
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO160_AddMatrixRow_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO160_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO160_FormClear
        //public void PS_CO160_FormClear()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocEntry = null;
        //	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_CO160'", ref "");
        //	if (Convert.ToDouble(DocEntry) == 0) {
        //		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm01.Items.Item("DocEntry").Specific.VALUE = 1;
        //	} else {
        //		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm01.Items.Item("DocEntry").Specific.VALUE = DocEntry;
        //	}
        //	return;
        //	PS_CO160_FormClear_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO160_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO160_EnableMenus
        //private void PS_CO160_EnableMenus()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////메뉴활성화
        //	//    Call oForm01.EnableMenu("1293", True)
        //	//    Call oForm01.EnableMenu("1288", True)
        //	//    Call oForm01.EnableMenu("1289", True)
        //	//    Call oForm01.EnableMenu("1290", True)
        //	//    Call oForm01.EnableMenu("1291", True)
        //	MDC_Com.MDC_GP_EnableMenus(ref oForm01, false, false, true, true, true, true, true, true, true,
        //	true, false, false, false, false, true, false);
        //	////메뉴설정
        //	return;
        //	PS_CO160_EnableMenus_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO160_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO160_SetDocument
        //private void PS_CO160_SetDocument(string oFromDocEntry01)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((string.IsNullOrEmpty(oFromDocEntry01))) {
        //		PS_CO160_FormItemEnabled();
        //		PS_CO160_AddMatrixRow(0, ref true);
        //		////UDO방식일때
        //	} else {
        //		oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //		PS_CO160_FormItemEnabled();
        //		//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm01.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
        //		oForm01.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //	}
        //	return;
        //	PS_CO160_SetDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO160_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO160_DataValidCheck
        //public bool PS_CO160_DataValidCheck()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	//UPGRADE_WARNING: oForm01.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm01.Items.Item("YM").Specific.VALUE)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm01.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	//UPGRADE_WARNING: oForm01.Items(ItemCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm01.Items.Item("ItemCode").Specific.VALUE)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("품목코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm01.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	//UPGRADE_WARNING: oForm01.Items(MItemCod).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm01.Items.Item("MItemCod").Specific.VALUE)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("이동품목코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm01.Items.Item("MItemCod").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	if (oMat01.VisualRowCount == 1) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		//UPGRADE_WARNING: oMat01.Columns(PO).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(i).Specific.VALUE))) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("작지문서라인은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("PO").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((Conversion.Val(oMat01.Columns.Item("MPO").Cells.Item(i).Specific.VALUE) <= 0)) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("이동작지문서라인은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("MPO").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}

        //	}

        //	oDS_PS_CO160L.RemoveRecord(oDS_PS_CO160L.Size - 1);
        //	oMat01.LoadFromDataSource();
        //	if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //		PS_CO160_FormClear();
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_CO160_DataValidCheck_Error:
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO160_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_CO160_FindValidateDocument
        //public bool PS_CO160_FindValidateDocument(string ObjectType)
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
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.VALUE);
        //	////원본문서

        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	Query01 = " SELECT DocEntry";
        //	Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry = ";
        //	Query01 = Query01 + DocEntry;
        //	if ((oDocType01 == "출하요청")) {
        //		Query01 = Query01 + " AND U_DocType = '1'";
        //	} else if ((oDocType01 == "선출요청")) {
        //		Query01 = Query01 + " AND U_DocType = '2'";
        //	}
        //	RecordSet01.DoQuery(Query01);
        //	if ((RecordSet01.RecordCount == 0)) {
        //		if ((oDocType01 == "출하요청")) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("선출요청문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		} else if ((oDocType01 == "선출요청")) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("출하요청문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        //	PS_CO160_FindValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	functionReturnValue = false;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_CO160_DirectionValidateDocument
        //public bool PS_CO160_DirectionValidateDocument(string DocEntry, string DocEntryNext, string Direction, string ObjectType)
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
        //			if ((oDocType01 == "출하요청")) {
        //				Query01 = Query01 + " AND U_DocType = '1'";
        //			} else if ((oDocType01 == "선출요청")) {
        //				Query01 = Query01 + " AND U_DocType = '2'";
        //			}
        //			Query01 = Query01 + " ORDER BY DocEntry ASC";
        //		} else if ((Direction == "Prev")) {
        //			Query01 = " SELECT TOP 1 DocEntry";
        //			Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //			Query01 = Query01 + DocEntryNext;
        //			if ((oDocType01 == "출하요청")) {
        //				Query01 = Query01 + " AND U_DocType = '1'";
        //			} else if ((oDocType01 == "선출요청")) {
        //				Query01 = Query01 + " AND U_DocType = '2'";
        //			}
        //			Query01 = Query01 + " ORDER BY DocEntry DESC";
        //		}
        //		RecordSet01.DoQuery(Query01);
        //		////해당문서가 마지막문서라면
        //		if ((RecordSet01.Fields.Item(0).Value == 0)) {
        //			if ((Direction == "Next")) {
        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //				if ((oDocType01 == "출하요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '1'";
        //				} else if ((oDocType01 == "선출요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '2'";
        //				}
        //				Query02 = Query02 + " ORDER BY DocEntry ASC";
        //			} else if ((Direction == "Prev")) {
        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //				if ((oDocType01 == "출하요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '1'";
        //				} else if ((oDocType01 == "선출요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '2'";
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
        //					if ((oDocType01 == "출하요청")) {
        //						Query01 = Query01 + " AND U_DocType = '1'";
        //					} else if ((oDocType01 == "선출요청")) {
        //						Query01 = Query01 + " AND U_DocType = '2'";
        //					}
        //					Query01 = Query01 + " ORDER BY DocEntry ASC";
        //					RecordSet01.DoQuery(Query01);
        //				} else if ((Direction == "Prev")) {
        //					DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) + 1);
        //					Query01 = " SELECT TOP 1 DocNum";
        //					Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //					Query01 = Query01 + DocEntryNext;
        //					if ((oDocType01 == "출하요청")) {
        //						Query01 = Query01 + " AND U_DocType = '1'";
        //					} else if ((oDocType01 == "선출요청")) {
        //						Query01 = Query01 + " AND U_DocType = '2'";
        //					}
        //					Query01 = Query01 + " ORDER BY DocEntry DESC";
        //					RecordSet01.DoQuery(Query01);
        //				}
        //			}
        //		}
        //		if ((oDocType01 == "출하요청")) {
        //			DoNext = false;
        //			if ((Direction == "Next")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
        //			} else if ((Direction == "Prev")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
        //			}
        //		} else if ((oDocType01 == "선출요청")) {
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
        //		PS_CO160_FormItemEnabled();
        //		////UDO방식
        //	////다음문서가 유효하지 않다면
        //	} else {
        //		oForm01.Freeze(true);
        //		oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //		PS_CO160_FormItemEnabled();
        //		////UDO방식
        //		////문서번호 필드가 입력이 가능하다면
        //		if (oForm01.Items.Item("DocEntry").Enabled == true) {
        //			if ((Direction == "Next")) {
        //				//UPGRADE_WARNING: oForm01.Items(DocEntry).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm01.Items.Item("DocEntry").Specific.VALUE = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) + 1));
        //			} else if ((Direction == "Prev")) {
        //				//UPGRADE_WARNING: oForm01.Items(DocEntry).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm01.Items.Item("DocEntry").Specific.VALUE = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) - 1));
        //			}
        //			oForm01.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		}
        //		oForm01.Freeze(false);
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
        //	PS_CO160_DirectionValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return functionReturnValue;
        //}
        #endregion
    }
}
