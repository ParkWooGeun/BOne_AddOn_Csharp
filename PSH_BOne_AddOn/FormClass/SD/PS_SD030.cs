using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using SAP.Middleware.Connector;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 출하(선출)요청등록
	/// </summary>
	internal class PS_SD030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD030H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD030L; //등록라인
		private string oDocType01;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		public class ItemInformation
		{
			public string ItemCode; //수량
			public int Qty; //중량
			public double Weight; //통화
			public string Currency; //단가
			public double Price; //총계
			public double LineTotal; //창고
			public string WhsCode; //판매오더문서
			public int ORDRNum; //판매오더라인
			public int RDR1Num;
			public bool Check; //납품문서
			public int ODLNNum; //납품라인
			public int DLN1Num; //반품문서
			public int ORDNNum; //반품라인
			public int RDN1Num; //출하(선출)문서
			public int SD030HNum; //출하(선출)라인
			public int SD030LNum;
		}

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				oDocType01 = "출하요청";
                PS_SD030_CreateItems();
                PS_SD030_SetComboBox();
                PS_SD030_CF_ChooseFromList();
                PS_SD030_EnableMenus();
                //PS_SD030_SetDocument(oFormDocEntry);
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

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_SD030_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PS_SD030H = oForm.DataSources.DBDataSources.Item("@PS_SD030H");
                oDS_PS_SD030L = oForm.DataSources.DBDataSources.Item("@PS_SD030L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                if (oDocType01 == "출하요청")
                {
                    oForm.Title = "출하요청[PS_SD030]";
                    oForm.Items.Item("DocType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if ((oDocType01 == "선출요청"))
                {
                    oForm.Title = "선출요청[PS_SD031]";
                    oForm.Items.Item("DocType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                
                oDS_PS_SD030H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD()); //담당자
                oDS_PS_SD030H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value + "'", 0, 1));
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_SD030_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "DocType", "", "1", "출하요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "DocType", "", "2", "선출요청");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PS_SD030", "DocType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "TrType", "", "1", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "TrType", "", "2", "임가공");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("TrType").Specific, "PS_PS_SD030", "TrType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "1", "출하요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "2", "선출요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "3", "납품");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "4", "반품");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("ProgStat").Specific, "PS_PS_SD030", "ProgStat", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "Status", "O", "미결");
                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "Status", "C", "완료");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("Status"), "PS_SD030", "Mat01", "Status", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "TrType", "1", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "TrType", "2", "군납");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("TrType"), "PS_SD030", "Mat01", "TrType", false);

                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemGpCd"), "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code,Name FROM [@PSH_ITMBSORT]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code,Name FROM [@PSH_SHAPE]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code,Name FROM [@PSH_QUALITY]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code,Name FROM [@PSH_MARK]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("SbasUnit"), "SELECT Code,Name FROM [@PSH_UOMORG]", "", "");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_SD030_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromListCollection oCFLs01 = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Conditions oCons01 = null;
            SAPbouiCOM.Condition oCon = null;
            SAPbouiCOM.Condition oCon01 = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromList oCFL01 = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams01 = null;
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.EditText oEdit01 = null;
            SAPbouiCOM.Column oColumn = null;

            try
            {
                oEdit = oForm.Items.Item("DCardCod").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                oCFLCreationParams.UniqueID = "CFLCARDCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oEdit.ChooseFromListUID = "CFLCARDCODE";
                oEdit.ChooseFromListAlias = "CardCode";

                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";
                oCFL.SetConditions(oCons);

                oEdit01 = oForm.Items.Item("CardCode").Specific;
                oCFLs01 = oForm.ChooseFromLists;
                oCFLCreationParams01 = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams01.ObjectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                oCFLCreationParams01.UniqueID = "CFLCARD2CODE";
                oCFLCreationParams01.MultiSelection = false;
                oCFL01 = oCFLs01.Add(oCFLCreationParams01);

                oEdit01.ChooseFromListUID = "CFLCARD2CODE";
                oEdit01.ChooseFromListAlias = "CardCode";

                oCons01 = oCFL01.GetConditions();
                oCon01 = oCons01.Add();
                oCon01.Alias = "CardType";
                oCon01.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon01.CondVal = "C";
                oCFL01.SetConditions(oCons01);

                oColumn = oMat01.Columns.Item("WhsCode");
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "64"; //SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                oCFLCreationParams.UniqueID = "CFLWAREHOUSES";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oColumn.ChooseFromListUID = "CFLWAREHOUSES";
                oColumn.ChooseFromListAlias = "WhsCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }

                if (oCFLs01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs01);
                }

                if (oCons != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
                }

                if (oCons01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons01);
                }

                if (oCon != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
                }

                if (oCon01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon01);
                }

                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }

                if (oCFL01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL01);
                }

                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }

                if (oCFLCreationParams01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams01);
                }

                if (oEdit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                }

                if (oEdit01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit01);
                }

                if (oColumn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                }
            }

        }

        /// <summary>
        /// 메뉴설정
        /// </summary>
        private void PS_SD030_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, true, true, true, true, true, true, false, false, false, false, true, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        #region PS_SD030_SetDocument
        //private void PS_SD030_SetDocument(string oFormDocEntry)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((string.IsNullOrEmpty(oFormDocEntry))) {
        //		PS_SD030_FormItemEnabled();
        //		PS_SD030_AddMatrixRow(0, ref true);
        //		////UDO방식일때
        //	} else {
        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_DocType FROM [PS_SD030H] WHERE DocEntry = ' & oFormDocEntry & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + oFormDocEntry + "'", 0, 1) == "1")) {
        //			oForm.Title = "출하요청[PS_SD030]";
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("DocType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_DocType FROM [PS_SD030H] WHERE DocEntry = ' & oFormDocEntry & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		} else if ((MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + oFormDocEntry + "'", 0, 1) == "2")) {
        //			oForm.Title = "선출요청[PS_SD031]";
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("DocType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //		}
        //		oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //		PS_SD030_FormItemEnabled();
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
        //		oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //	}
        //	return;
        //	PS_SD030_SetDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion






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
        //				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					if ((PS_SD030_Validate("취소") == false)) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //					if (SubMain.Sbo_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1")) {
        //						BubbleEvent = false;
        //						return;
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
        //				Raise_EVENT_RECORD_MOVE(ref FormUID, ref pVal, ref BubbleEvent);
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
        //				PS_SD030_FormItemEnabled();
        //				////UDO방식
        //				oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			case "1282":
        //				//추가
        //				PS_SD030_FormItemEnabled();
        //				////UDO방식
        //				PS_SD030_AddMatrixRow(0, ref true);
        //				////UDO방식
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				Raise_EVENT_RECORD_MOVE(ref FormUID, ref pVal, ref BubbleEvent);
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
        //					if ((PS_SD030_FindValidateDocument("@PS_SD030H") == false)) {
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
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}
        //		} else {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = "";
        //			oLastColRow01 = 0;
        //		}
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
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}
        //		} else {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = "";
        //			oLastColRow01 = 0;
        //		}
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
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

        //	int DocEntry = 0;
        //	int i = 0;
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_SD030_DataValidCheck() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				////해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (PS_SD030_DataValidCheck() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "Button01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				////작지완료여부검사후 DI '//납품전기가능상황에서 납품(실적이 등록된것)
        //				if (PS_SD030_ValidateDelivery() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				if (PS_SD030_DI_API_01() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				SubMain.Sbo_Application.ActivateMenuItem("1281");
        //				//                DocEntry = Trim(oForm.Items("DocEntry").Specific.Value)
        //				//                oForm.Mode = fm_FIND_MODE
        //				//                oForm.Items("DocEntry").Specific.Value = DocEntry
        //				//                oForm.Items("1").Click ct_Regular
        //			}
        //		}
        //		if (pVal.ItemUID == "Button04") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(ProgStat).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////납품상태이면
        //				if (oForm.Items.Item("ProgStat").Specific.Selected.Value == "3") {
        //					////AR송장처리된 문서존재유무검사
        //					for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(DLN1Num).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = ' & oMat01.Columns(ODLNNum).Cells(i).Specific.Value & ' AND BaseLine = ' & oMat01.Columns(DLN1Num).Cells(i).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "AR송장처리된 문서가 존재합니다.", ref "W");
        //							BubbleEvent = false;
        //							return;
        //						}
        //					}
        //					////반품처리된 문서존재유무검사
        //					for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(DLN1Num).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = ' & oMat01.Columns(ODLNNum).Cells(i).Specific.Value & ' AND BaseLine = ' & oMat01.Columns(DLN1Num).Cells(i).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "반품처리된 문서가 존재합니다.", ref "W");
        //							BubbleEvent = false;
        //							return;
        //						}
        //					}
        //					if (PS_SD030_DI_API_02() == false) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //					SubMain.Sbo_Application.ActivateMenuItem("1281");
        //				}
        //				//                DocEntry = Trim(oForm.Items("DocEntry").Specific.Value)
        //				//                oForm.Mode = fm_FIND_MODE
        //				//                oForm.Items("DocEntry").Specific.Value = DocEntry
        //				//                oForm.Items("1").Click ct_Regular
        //			}
        //		}
        //		if (pVal.ItemUID == "Button02") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				PS_SD030_Print_Report01();
        //			}
        //		}
        //		if (pVal.ItemUID == "Button03") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				PS_SD030_Print_Report02();
        //			}
        //		}
        //		if (pVal.ItemUID == "Button05") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				PS_SD030_Print_Report03();
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_SD030_FormItemEnabled();
        //					PS_SD030_AddMatrixRow(oMat01.RowCount, ref true);
        //					////UDO방식일때
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_SD030_FormItemEnabled();
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

        //	if (pVal.BeforeAction == true) {
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
        //		////사용자값활성
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "ItemCode", "") '//사용자값활성
        //		MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrderNum");
        //		////사용자값활성
        //		//Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm, pVal, BubbleEvent, "Mat01", "WhsCode") '//사용자값활성
        //	} else if (pVal.BeforeAction == false) {

        //	}
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

        //	oForm.Freeze(true);
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string ItemCode01 = null;
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemChanged == true) {
        //			if ((pVal.ItemUID == "Mat01")) {
        //				if ((PS_SD030_Validate("수정") == false)) {
        //					oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Strings.Trim(oDS_PS_SD030L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)));
        //				} else {
        //					if (pVal.ColUID == "OrderNum") {
        //						//UPGRADE_WARNING: oMat01.Columns(OrderNum).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value)) {
        //							goto Raise_EVENT_VALIDATE_Exit;
        //						}
        //						for (i = 1; i <= oMat01.RowCount; i++) {
        //							////현재 선택되어있는 행이 아니면
        //							if (pVal.Row != i) {
        //								//UPGRADE_WARNING: oMat01.Columns(OrderNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oMat01.Columns(OrderNum).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if ((oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value)) {
        //									MDC_Com.MDC_GF_Message(ref "동일한 수주가 존재합니다.", ref "W");
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value = "";
        //									goto Raise_EVENT_VALIDATE_Exit;
        //								}
        //								//                            If (Mid(oMat01.Columns("OrderNum").Cells(pVal.Row).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(pVal.Row).Specific.Value, "-") - 1) <> _
        //								//'                            Mid(oMat01.Columns("OrderNum").Cells(i).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(i).Specific.Value, "-") - 1)) Then
        //								//                                Call MDC_Com.MDC_GF_Message("동일하지않은 수주문서가 존재합니다.", "W")
        //								//                                oMat01.Columns("OrderNum").Cells(pVal.Row).Specific.Value = ""
        //								//                                GoTo Raise_EVENT_VALIDATE_Exit
        //								//                            End If
        //							}
        //						}
        //						RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //						//UPGRADE_WARNING: oForm.Items(DocType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						Query01 = "EXEC PS_SD030_01 '" + oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value + "','" + oForm.Items.Item("DocType").Specific.Selected.Value + "'";
        //						RecordSet01.DoQuery(Query01);
        //						for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //							oDS_PS_SD030L.SetValue("U_OrderNum", pVal.Row - 1, RecordSet01.Fields.Item("OrderNum").Value);
        //							oDS_PS_SD030L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
        //							oDS_PS_SD030L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
        //							oDS_PS_SD030L.SetValue("U_ItemGpCd", pVal.Row - 1, RecordSet01.Fields.Item("ItemGpCd").Value);
        //							oDS_PS_SD030L.SetValue("U_ItmBsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmBsort").Value);
        //							oDS_PS_SD030L.SetValue("U_ItmMsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmMsort").Value);
        //							oDS_PS_SD030L.SetValue("U_Unit1", pVal.Row - 1, RecordSet01.Fields.Item("Unit1").Value);
        //							oDS_PS_SD030L.SetValue("U_Size", pVal.Row - 1, RecordSet01.Fields.Item("Size").Value);
        //							oDS_PS_SD030L.SetValue("U_ItemType", pVal.Row - 1, RecordSet01.Fields.Item("ItemType").Value);
        //							oDS_PS_SD030L.SetValue("U_Quality", pVal.Row - 1, RecordSet01.Fields.Item("Quality").Value);
        //							oDS_PS_SD030L.SetValue("U_Mark", pVal.Row - 1, RecordSet01.Fields.Item("Mark").Value);
        //							oDS_PS_SD030L.SetValue("U_SbasUnit", pVal.Row - 1, RecordSet01.Fields.Item("SbasUnit").Value);
        //							oDS_PS_SD030L.SetValue("U_SjQty", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("SjQty").Value)));
        //							oDS_PS_SD030L.SetValue("U_SjWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("SjWeight").Value)));
        //							oDS_PS_SD030L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Qty").Value)));
        //							oDS_PS_SD030L.SetValue("U_UnWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("UnWeight").Value)));
        //							oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Weight").Value)));
        //							oDS_PS_SD030L.SetValue("U_Currency", pVal.Row - 1, RecordSet01.Fields.Item("Currency").Value);
        //							oDS_PS_SD030L.SetValue("U_Price", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Price").Value)));
        //							oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("LinTotal").Value)));
        //							oDS_PS_SD030L.SetValue("U_WhsCode", pVal.Row - 1, RecordSet01.Fields.Item("WhsCode").Value);
        //							oDS_PS_SD030L.SetValue("U_WhsName", pVal.Row - 1, RecordSet01.Fields.Item("WhsName").Value);
        //							oDS_PS_SD030L.SetValue("U_Comments", pVal.Row - 1, RecordSet01.Fields.Item("Comments").Value);
        //							oDS_PS_SD030L.SetValue("U_TrType", pVal.Row - 1, RecordSet01.Fields.Item("TrType").Value);
        //							oDS_PS_SD030L.SetValue("U_ORDRNum", pVal.Row - 1, RecordSet01.Fields.Item("ORDRNum").Value);
        //							oDS_PS_SD030L.SetValue("U_RDR1Num", pVal.Row - 1, RecordSet01.Fields.Item("RDR1Num").Value);
        //							oDS_PS_SD030L.SetValue("U_Status", pVal.Row - 1, RecordSet01.Fields.Item("Status").Value);
        //							oDS_PS_SD030L.SetValue("U_LineId", pVal.Row - 1, RecordSet01.Fields.Item("LineId").Value);
        //							RecordSet01.MoveNext();
        //						}
        //						if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD030L.GetValue("U_OrderNum", pVal.Row - 1)))) {
        //							PS_SD030_AddMatrixRow((pVal.Row));
        //						}
        //						oMat01.LoadFromDataSource();
        //						oMat01.AutoResizeColumns();

        //						if ((oMat01.RowCount > 1)) {
        //							oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //							oForm.Items.Item("CardCode").Enabled = false;
        //							oForm.Items.Item("BPLId").Enabled = false;
        //							oForm.Items.Item("TrType").Enabled = false;
        //						} else {
        //							oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //							oForm.Items.Item("CardCode").Enabled = true;
        //							oForm.Items.Item("BPLId").Enabled = true;
        //							oForm.Items.Item("TrType").Enabled = true;
        //						}
        //						oForm.Update();
        //						//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						RecordSet01 = null;
        //					} else if (pVal.ColUID == "Qty") {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)) {
        //							oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(0));
        //							oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(0));
        //							oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(0));
        //						} else {
        //							ItemCode01 = Strings.Trim(oDS_PS_SD030L.GetValue("U_ItemCode", pVal.Row - 1));
        //							////EA자체품
        //							if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
        //							////EAUOM
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))));
        //							////KGSPEC
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
        //							////KG단중
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
        //							////KG선택
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //							}
        //							oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(Strings.Trim(oDS_PS_SD030L.GetValue("U_Weight", pVal.Row - 1))) * Convert.ToDouble(Strings.Trim(oDS_PS_SD030L.GetValue("U_Price", pVal.Row - 1)))));
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //						}
        //					} else if (pVal.ColUID == "Weight") {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if ((Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)) {
        //							oDS_PS_SD030L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(0));
        //							oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(0));
        //							oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(0));
        //						} else {
        //							ItemCode01 = Strings.Trim(oDS_PS_SD030L.GetValue("U_ItemCode", pVal.Row - 1));
        //							if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //							////EAUOM
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //							////KGSPEC
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
        //							////KG단중
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
        //							////KG선택
        //							} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //							}
        //							oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(Strings.Trim(oDS_PS_SD030L.GetValue("U_Weight", pVal.Row - 1))) * Convert.ToDouble(Strings.Trim(oDS_PS_SD030L.GetValue("U_Price", pVal.Row - 1)))));
        //						}
        //						//                    ElseIf pVal.ColUID = "WhsCode" Then
        //						//                        Call oDS_PS_SD030L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value)
        //						//                        Call oDS_PS_SD030L.setValue("U_WhsName", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" & oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value & "'", 0, 1))
        //					} else {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //					}
        //				}
        //				oMat01.LoadFromDataSource();
        //				oMat01.AutoResizeColumns();
        //				oForm.Update();
        //				oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else {
        //				if ((pVal.ItemUID == "DocEntry")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD030H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //				} else if ((pVal.ItemUID == "CntcCode")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD030H.SetValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
        //				} else {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //				}
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Exit:
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

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_SD030_FormItemEnabled();
        //		PS_SD030_AddMatrixRow(oMat01.VisualRowCount);
        //		////UDO방식
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
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
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbouiCOM.DataTable oDataTable01 = null;
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		if ((pVal.ItemUID == "CardCode" | pVal.ItemUID == "CardName")) {
        //			MDC_Com.MDC_GP_CF_DBDatasourceReturn(pVal, (pVal.FormUID), "@PS_SD030H", "U_CardCode,U_CardName");
        //		}
        //		if ((pVal.ItemUID == "DCardCod" | pVal.ItemUID == "DCardNam")) {
        //			MDC_Com.MDC_GP_CF_DBDatasourceReturn(pVal, (pVal.FormUID), "@PS_SD030H", "U_DCardCod,U_DCardNam");
        //		}
        //		if ((pVal.ItemUID == "Mat01")) {
        //			if ((pVal.ColUID == "WhsCode")) {
        //				//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (pVal.SelectedObjects == null) {
        //				} else {
        //					//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDataTable01 = pVal.SelectedObjects;
        //					oDS_PS_SD030L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
        //					oDS_PS_SD030L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
        //					//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oDataTable01 = null;
        //					//Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
        //					oMat01.LoadFromDataSource();
        //					oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				}
        //			}
        //		}
        //		if ((pVal.ItemUID == "DCardCod" | pVal.ItemUID == "DCardNam")) {
        //			MDC_Com.MDC_GP_CF_DBDatasourceReturn(pVal, (pVal.FormUID), "@PS_SD030H", "U_DCardCod,U_DCardNam");
        //		}
        //	}
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
        //	if ((oLastColRow01 > 0)) {
        //		if (pVal.BeforeAction == true) {
        //			if ((PS_SD030_Validate("행삭제") == false)) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		} else if (pVal.BeforeAction == false) {
        //			for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_SD030L.RemoveRecord(oDS_PS_SD030L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0) {
        //				PS_SD030_AddMatrixRow(0);
        //			} else {
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD030L.GetValue("U_OrderNum", oMat01.RowCount - 1)))) {
        //					PS_SD030_AddMatrixRow(oMat01.RowCount);
        //				}
        //			}
        //			if ((oMat01.RowCount > 1)) {
        //				oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("CardCode").Enabled = false;
        //				oForm.Items.Item("BPLId").Enabled = false;
        //				oForm.Items.Item("TrType").Enabled = false;
        //			} else {
        //				oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("CardCode").Enabled = true;
        //				oForm.Items.Item("BPLId").Enabled = true;
        //				oForm.Items.Item("TrType").Enabled = true;
        //			}
        //			oForm.Update();
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RECORD_MOVE
        //private void Raise_EVENT_RECORD_MOVE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string DocEntry = null;
        //	string DocEntryNext = null;
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
        //	////원본문서
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntryNext = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
        //	////다음문서

        //	////다음
        //	if (pVal.MenuUID == "1288") {
        //		if (pVal.BeforeAction == true) {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //				BubbleEvent = false;
        //				return;
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_SD030H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////이전
        //	} else if (pVal.MenuUID == "1289") {
        //		if (pVal.BeforeAction == true) {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //				BubbleEvent = false;
        //				return;
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_SD030H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////첫번째레코드로이동
        //	} else if (pVal.MenuUID == "1290") {
        //		if (pVal.BeforeAction == true) {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_SD030H] ORDER BY DocEntry DESC";
        //			////가장마지막행을 부여
        //			RecordSet01.DoQuery(Query01);
        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////원본문서
        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////다음문서
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_SD030H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////마지막문서로이동
        //	} else if (pVal.MenuUID == "1291") {
        //		if (pVal.BeforeAction == true) {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_SD030H] ORDER BY DocEntry ASC";
        //			////가장 첫행을 부여
        //			RecordSet01.DoQuery(Query01);
        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////원본문서
        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////다음문서
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_SD030H") == false) {
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






        #region PS_SD030_FormItemEnabled
        //public void PS_SD030_FormItemEnabled()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	oForm.Freeze(true);
        //	oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //	bool Enabled = false;
        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //		////각모드에따른 아이템설정
        //		oForm.Items.Item("DocEntry").Enabled = false;
        //		oForm.Items.Item("CardCode").Enabled = true;
        //		oForm.Items.Item("BPLId").Enabled = true;
        //		oForm.Items.Item("CntcCode").Enabled = true;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oForm.Items.Item("DueDate").Enabled = true;
        //		oForm.Items.Item("TranCard").Enabled = true;
        //		oForm.Items.Item("TranCode").Enabled = true;
        //		oForm.Items.Item("Destin").Enabled = true;
        //		oForm.Items.Item("TranCost").Enabled = true;
        //		oForm.Items.Item("Comments").Enabled = true;
        //		oForm.Items.Item("TrType").Enabled = true;
        //		oForm.Items.Item("Mat01").Enabled = true;
        //		oMat01.AutoResizeColumns();
        //		PS_SD030_FormClear();
        //		////UDO방식
        //		//        Call oForm.Items("BPLId").Specific.Select("1", psk_ByValue)
        //		//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);
        //		////2010.12.06 추가
        //		if ((oDocType01 == "출하요청")) {
        //			oForm.Items.Item("DocType").Enabled = true;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("DocType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //			oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm.Items.Item("DocType").Enabled = false;
        //			oForm.Items.Item("Button01").Visible = false;
        //			oForm.Items.Item("Button04").Visible = false;
        //			oForm.Items.Item("ProgStat").Enabled = true;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("ProgStat").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //			oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm.Items.Item("ProgStat").Enabled = false;
        //			oMat01.Columns.Item("ODLNNum").Visible = false;
        //			oMat01.Columns.Item("DLN1Num").Visible = false;
        //			oForm.Update();
        //		} else if ((oDocType01 == "선출요청")) {
        //			oForm.Items.Item("DocType").Enabled = true;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("DocType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //			oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm.Items.Item("DocType").Enabled = false;
        //			oForm.Items.Item("BPLId").Enabled = true;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("BPLId").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //			oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm.Items.Item("BPLId").Enabled = false;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //			oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm.Items.Item("TrType").Enabled = false;
        //			oForm.Items.Item("Button01").Visible = true;
        //			oForm.Items.Item("Button04").Visible = true;
        //			oForm.Items.Item("ProgStat").Enabled = true;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.Items.Item("ProgStat").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //			oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			oForm.Items.Item("ProgStat").Enabled = false;
        //			oMat01.Columns.Item("ODLNNum").Visible = true;
        //			oMat01.Columns.Item("DLN1Num").Visible = true;
        //			oForm.Items.Item("Button03").Visible = false;
        //		}
        //		//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //		////2010.12.06 추가
        //		//담당자
        //		oDS_PS_SD030H.SetValue("U_CntcCode", 0, MDC_PS_Common.User_MSTCOD());
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDS_PS_SD030H.SetValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value + "'", 0, 1));
        //		oForm.EnableMenu("1281", true);
        //		////찾기
        //		oForm.EnableMenu("1282", false);
        //		////추가
        //		oForm.Items.Item("Button01").Enabled = false;
        //		oForm.Items.Item("Button04").Enabled = false;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oForm.Items.Item("DueDate").Enabled = true;
        //		//UPGRADE_WARNING: oForm.Items(DocDate).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocDate").Specific.String = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd");
        //		//UPGRADE_WARNING: oForm.Items(DueDate).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DueDate").Specific.String = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd");

        //		oForm.Items.Item("14").Visible = false;
        //		oForm.Items.Item("16").Visible = false;
        //		oForm.Items.Item("18").Visible = false;
        //		oForm.Items.Item("20").Visible = false;
        //		oForm.Items.Item("TranCard").Visible = false;
        //		oForm.Items.Item("TranCode").Visible = false;
        //		oForm.Items.Item("Destin").Visible = false;
        //		oForm.Items.Item("TranCost").Visible = false;
        //	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
        //		oForm.Items.Item("DocEntry").Enabled = true;
        //		oForm.Items.Item("CardCode").Enabled = true;
        //		oForm.Items.Item("BPLId").Enabled = true;
        //		oForm.Items.Item("CntcCode").Enabled = true;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oForm.Items.Item("DueDate").Enabled = true;
        //		oForm.Items.Item("TranCard").Enabled = true;
        //		oForm.Items.Item("TranCode").Enabled = true;
        //		oForm.Items.Item("Destin").Enabled = true;
        //		oForm.Items.Item("TranCost").Enabled = true;
        //		oForm.Items.Item("Comments").Enabled = true;
        //		oForm.Items.Item("TrType").Enabled = true;
        //		oForm.Items.Item("Mat01").Enabled = false;
        //		oMat01.AutoResizeColumns();
        //		oForm.EnableMenu("1281", false);
        //		oForm.EnableMenu("1282", true);
        //		oForm.Items.Item("Button01").Enabled = false;
        //		oForm.Items.Item("Button04").Enabled = false;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oForm.Items.Item("DueDate").Enabled = true;

        //		oForm.Items.Item("14").Visible = false;
        //		oForm.Items.Item("16").Visible = false;
        //		oForm.Items.Item("18").Visible = false;
        //		oForm.Items.Item("20").Visible = false;
        //		oForm.Items.Item("TranCard").Visible = false;
        //		oForm.Items.Item("TranCode").Visible = false;
        //		oForm.Items.Item("Destin").Visible = false;
        //		oForm.Items.Item("TranCost").Visible = false;
        //		////각모드에따른 아이템설정
        //	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
        //		oForm.EnableMenu("1281", true);
        //		////찾기
        //		oForm.EnableMenu("1282", true);
        //		////추가
        //		////출하요청일때
        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_DocType FROM [PS_SD030H] WHERE DocEntry = ' & Trim(oDS_PS_SD030H.GetValue(DocEntry, 0)) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + Strings.Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) + "'", 0, 1) == "1") {
        //			Enabled = false;
        //			for (i = 0; i <= oDS_PS_SD030L.Size - 1; i++) {
        //				////매트릭스의 중량과 납품문서의 중량을 비교
        //				//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (Conversion.Val(Strings.Trim(oDS_PS_SD030L.GetValue("U_Weight", i))) > Conversion.Val(MDC_PS_Common.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD040L] WHERE U_SD030H = '" + Strings.Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) + "' AND U_SD030L = '" + Strings.Trim(oDS_PS_SD030L.GetValue("U_LineId", i)) + "'", 0, 1))) {
        //					Enabled = true;
        //				}
        //			}
        //			////문서가 수정불가능한경우
        //			if (Enabled == false) {
        //				oForm.Items.Item("DocEntry").Enabled = false;
        //				oForm.Items.Item("CardCode").Enabled = false;
        //				oForm.Items.Item("BPLId").Enabled = false;
        //				oForm.Items.Item("CntcCode").Enabled = false;
        //				oForm.Items.Item("DocDate").Enabled = false;
        //				oForm.Items.Item("DueDate").Enabled = false;
        //				oForm.Items.Item("TranCard").Enabled = false;
        //				oForm.Items.Item("TranCode").Enabled = false;
        //				oForm.Items.Item("Destin").Enabled = false;
        //				oForm.Items.Item("TranCost").Enabled = false;
        //				oForm.Items.Item("Comments").Enabled = true;
        //				oForm.Items.Item("Mat01").Enabled = false;
        //				oForm.Items.Item("TrType").Enabled = false;
        //				oMat01.AutoResizeColumns();
        //				oForm.EnableMenu("1281", true);
        //				oForm.EnableMenu("1282", false);
        //				oForm.Items.Item("Button01").Enabled = false;
        //				oForm.Items.Item("Button04").Enabled = false;
        //				oForm.Items.Item("DocDate").Enabled = false;
        //				oForm.Items.Item("DueDate").Enabled = false;
        //			////문서가 수정가능한경우
        //			} else {
        //				oForm.Items.Item("DocEntry").Enabled = false;
        //				oForm.Items.Item("CardCode").Enabled = false;
        //				oForm.Items.Item("BPLId").Enabled = false;
        //				oForm.Items.Item("CntcCode").Enabled = true;
        //				oForm.Items.Item("DocDate").Enabled = true;
        //				oForm.Items.Item("DueDate").Enabled = true;
        //				oForm.Items.Item("TranCard").Enabled = true;
        //				oForm.Items.Item("TranCode").Enabled = true;
        //				oForm.Items.Item("Destin").Enabled = true;
        //				oForm.Items.Item("TranCost").Enabled = true;
        //				oForm.Items.Item("Comments").Enabled = true;
        //				oForm.Items.Item("TrType").Enabled = false;
        //				oForm.Items.Item("Mat01").Enabled = true;
        //				oMat01.AutoResizeColumns();
        //				oForm.EnableMenu("1281", true);
        //				oForm.EnableMenu("1282", false);
        //				oForm.Items.Item("Button01").Enabled = false;
        //				oForm.Items.Item("Button04").Enabled = false;
        //				oForm.Items.Item("DocDate").Enabled = false;
        //				oForm.Items.Item("DueDate").Enabled = false;
        //			}
        //			oForm.Items.Item("14").Visible = false;
        //			oForm.Items.Item("16").Visible = false;
        //			oForm.Items.Item("18").Visible = false;
        //			oForm.Items.Item("20").Visible = false;
        //			oForm.Items.Item("TranCard").Visible = false;
        //			oForm.Items.Item("TranCode").Visible = false;
        //			oForm.Items.Item("Destin").Visible = false;
        //			oForm.Items.Item("TranCost").Visible = false;
        //			////선출요청일때
        //			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_DocType FROM [PS_SD030H] WHERE DocEntry = ' & Trim(oDS_PS_SD030H.GetValue(DocEntry, 0)) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		} else if (MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + Strings.Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) + "'", 0, 1) == "2") {
        //			////납품일때
        //			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_ProgStat FROM [PS_SD030H] WHERE DocEntry = ' & Trim(oDS_PS_SD030H.GetValue(DocEntry, 0)) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + Strings.Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) + "'", 0, 1) == "3") {
        //				oForm.Items.Item("DocEntry").Enabled = false;
        //				oForm.Items.Item("CardCode").Enabled = false;
        //				oForm.Items.Item("BPLId").Enabled = false;
        //				oForm.Items.Item("CntcCode").Enabled = false;
        //				oForm.Items.Item("DocDate").Enabled = false;
        //				oForm.Items.Item("DueDate").Enabled = false;
        //				oForm.Items.Item("TranCard").Enabled = false;
        //				oForm.Items.Item("TranCode").Enabled = false;
        //				oForm.Items.Item("Destin").Enabled = false;
        //				oForm.Items.Item("TranCost").Enabled = false;
        //				oForm.Items.Item("Comments").Enabled = true;
        //				oForm.Items.Item("Mat01").Enabled = false;
        //				oForm.Items.Item("TrType").Enabled = false;
        //				oMat01.AutoResizeColumns();
        //				oForm.EnableMenu("1281", true);
        //				oForm.EnableMenu("1282", false);
        //				oForm.Items.Item("Button01").Enabled = false;
        //				oForm.Items.Item("Button04").Enabled = true;
        //				oForm.Items.Item("DocDate").Enabled = false;
        //				oForm.Items.Item("DueDate").Enabled = false;

        //				oForm.Items.Item("14").Visible = true;
        //				oForm.Items.Item("16").Visible = true;
        //				oForm.Items.Item("18").Visible = true;
        //				oForm.Items.Item("20").Visible = true;
        //				oForm.Items.Item("TranCard").Visible = true;
        //				oForm.Items.Item("TranCode").Visible = true;
        //				oForm.Items.Item("Destin").Visible = true;
        //				oForm.Items.Item("TranCost").Visible = true;
        //			////납품이 아닐때
        //			} else {
        //				oForm.Items.Item("DocEntry").Enabled = false;
        //				oForm.Items.Item("CardCode").Enabled = false;
        //				oForm.Items.Item("BPLId").Enabled = false;
        //				oForm.Items.Item("CntcCode").Enabled = true;
        //				oForm.Items.Item("DocDate").Enabled = true;
        //				oForm.Items.Item("DueDate").Enabled = true;
        //				oForm.Items.Item("TranCard").Enabled = true;
        //				oForm.Items.Item("TranCode").Enabled = true;
        //				oForm.Items.Item("Destin").Enabled = true;
        //				oForm.Items.Item("TranCost").Enabled = true;
        //				oForm.Items.Item("Comments").Enabled = true;
        //				oForm.Items.Item("TrType").Enabled = false;
        //				oForm.Items.Item("Mat01").Enabled = true;
        //				oMat01.AutoResizeColumns();
        //				oForm.EnableMenu("1281", true);
        //				oForm.EnableMenu("1282", false);
        //				oForm.Items.Item("Button01").Enabled = true;
        //				oForm.Items.Item("Button04").Enabled = false;
        //				oForm.Items.Item("DocDate").Enabled = false;
        //				oForm.Items.Item("DueDate").Enabled = false;

        //				oForm.Items.Item("14").Visible = false;
        //				oForm.Items.Item("16").Visible = false;
        //				oForm.Items.Item("18").Visible = false;
        //				oForm.Items.Item("20").Visible = false;
        //				oForm.Items.Item("TranCard").Visible = false;
        //				oForm.Items.Item("TranCode").Visible = false;
        //				oForm.Items.Item("Destin").Visible = false;
        //				oForm.Items.Item("TranCost").Visible = false;
        //			}
        //		}
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	PS_SD030_FormItemEnabled_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD030_AddMatrixRow
        //public void PS_SD030_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	////행추가여부
        //	if (RowIserted == false) {
        //		oDS_PS_SD030L.InsertRecord((oRow));
        //	}
        //	oMat01.AddRow();
        //	oDS_PS_SD030L.Offset = oRow;
        //	oDS_PS_SD030L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oMat01.LoadFromDataSource();
        //	oForm.Freeze(false);
        //	return;
        //	PS_SD030_AddMatrixRow_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD030_FormClear
        //public void PS_SD030_FormClear()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocEntry = null;
        //	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_SD030'", ref "");
        //	if (Convert.ToDouble(DocEntry) == 0) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.Value = 1;
        //	} else {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
        //	}
        //	return;
        //	PS_SD030_FormClear_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion



        #region PS_SD030_DataValidCheck
        //public bool PS_SD030_DataValidCheck()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("전기일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	if (oMat01.VisualRowCount == 1) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		//UPGRADE_WARNING: oMat01.Columns(OrderNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value))) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("수주는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("OrderNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if ((Conversion.Val(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value) <= 0)) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("중량(수량)은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			oMat01.Columns.Item("Weight").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		for (j = i + 1; j <= oMat01.VisualRowCount - 1; j++) {
        //			//UPGRADE_WARNING: oMat01.Columns(OrderNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat01.Columns(OrderNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if ((oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrderNum").Cells.Item(j).Specific.Value)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("동일한 수주가 존재합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				oMat01.Columns.Item("OrderNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}
        //			//            If (Mid(oMat01.Columns("OrderNum").Cells(i).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(i).Specific.Value, "-") - 1) <> _
        //			//'            Mid(oMat01.Columns("OrderNum").Cells(j).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(j).Specific.Value, "-") - 1)) Then
        //			//                Sbo_Application.SetStatusBarMessage "동일하지않은 수주문서가 존재합니다.", bmt_Short, True
        //			//                oMat01.Columns("OrderNum").Cells(i).Click ct_Regular
        //			//                PS_SD030_DataValidCheck = False
        //			//                Exit Function
        //			//            End If
        //		}
        //	}

        //	if ((PS_SD030_Validate("검사") == false)) {
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}

        //	oDS_PS_SD030L.RemoveRecord(oDS_PS_SD030L.Size - 1);
        //	oMat01.LoadFromDataSource();
        //	if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //		PS_SD030_FormClear();
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD030_DataValidCheck_Error:
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD030_MTX01
        //private void PS_SD030_MTX01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////메트릭스에 데이터 로드
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
        //		goto PS_SD030_MTX01_Exit;
        //	}
        //	oMat01.Clear();
        //	oMat01.FlushToDataSource();
        //	oMat01.LoadFromDataSource();

        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //	for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //		if (i != 0) {
        //			oDS_PS_SD030L.InsertRecord((i));
        //		}
        //		oDS_PS_SD030L.Offset = i;
        //		oDS_PS_SD030L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
        //		oDS_PS_SD030L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
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
        //	return;
        //	PS_SD030_MTX01_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	PS_SD030_MTX01_Error:
        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD030_DI_API_01
        //private bool PS_SD030_DI_API_01()
        //{
        //	bool functionReturnValue = false;
        //	////납품
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Qty = oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Weight = oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Currency = oMat01.Columns.Item("Currency").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Price = oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].LineTotal = oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ORDRNum = Conversion.Val(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].RDR1Num = Conversion.Val(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030HNum = Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030LNum = Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value);
        //		////ItemInformation(ItemInformationCount).Check = False
        //		ItemInformationCount = ItemInformationCount + 1;
        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.Value));
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardCod").Value = Strings.Trim(oForm.Items.Item("DCardCod").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardNam").Value = Strings.Trim(oForm.Items.Item("DCardNam").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_TradeType").Value = Strings.Trim(oForm.Items.Item("TrType").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //	}
        //	//UPGRADE_WARNING: oForm.Items(DueDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDueDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DueDate").Specific.Value, "&&&&-&&-&&"));
        //	}

        //	for (i = 0; i <= ItemInformationCount - 1; i++) {
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (i != 0) {
        //			oDIObject.Lines.Add();
        //		}

        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.WarehouseCode = ItemInformation[i].WhsCode;
        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD030";
        //		////별로 의미가 없을듯..
        //		oDIObject.Lines.BaseType = 17;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseEntry = ItemInformation[i].ORDRNum;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseLine = ItemInformation[i].RDR1Num;

        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = ItemInformation[i].Qty;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Quantity = ItemInformation[i].Weight;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Currency = ItemInformation[i].Currency;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Price = ItemInformation[i].Price;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.LineTotal = ItemInformation[i].LineTotal;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[i].DLN1Num = LineNumCount;
        //		LineNumCount = LineNumCount + 1;
        //	}
        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //		////문서상태 납품으로 변경
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030H] SET U_ProgStat = '3' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'"));
        //		////납품,반품문서번호 업데이트
        //		for (i = 0; i <= ItemInformationCount - 1; i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030L] SET U_ODLNNum = '" + ResultDocNum + "', U_DLN1Num = '" + ItemInformation[i].DLN1Num + "', U_ORDNNum = '', U_RDN1Num = '' WHERE DocEntry = '" + ItemInformation[i].SD030HNum + "' AND LineId = '" + ItemInformation[i].SD030LNum + "'"));
        //		}
        //	} else {
        //		goto PS_SD030_DI_API_01_DI_Error;
        //	}

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	oForm.Update();
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //	PS_SD030_DI_API_01_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //	PS_SD030_DI_API_01_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_DI_API_01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD030_DI_API_02
        //private bool PS_SD030_DI_API_02()
        //{
        //	bool functionReturnValue = false;
        //	////반품
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Qty = oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Weight = oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Currency = oMat01.Columns.Item("Currency").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Price = oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].LineTotal = oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ORDRNum = Conversion.Val(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].RDR1Num = Conversion.Val(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030HNum = Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030LNum = Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ODLNNum = Conversion.Val(oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value);
        //		////납품
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].DLN1Num = Conversion.Val(oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value);
        //		////납품
        //		////ItemInformation(ItemInformationCount).Check = False
        //		ItemInformationCount = ItemInformationCount + 1;
        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.Value));
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardCod").Value = Strings.Trim(oForm.Items.Item("DCardCod").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardNam").Value = Strings.Trim(oForm.Items.Item("DCardNam").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_TradeType").Value = Strings.Trim(oForm.Items.Item("TrType").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //	}
        //	//UPGRADE_WARNING: oForm.Items(DueDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDueDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DueDate").Specific.Value, "&&&&-&&-&&"));
        //	}

        //	for (i = 0; i <= ItemInformationCount - 1; i++) {
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (i != 0) {
        //			oDIObject.Lines.Add();
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.WarehouseCode = ItemInformation[i].WhsCode;
        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD030";
        //		////별로 의미가 없을듯..
        //		oDIObject.Lines.BaseType = 15;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseEntry = ItemInformation[i].ODLNNum;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseLine = ItemInformation[i].DLN1Num;

        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = ItemInformation[i].Qty;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Quantity = ItemInformation[i].Weight;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Currency = ItemInformation[i].Currency;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Price = ItemInformation[i].Price;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.LineTotal = ItemInformation[i].LineTotal;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[i].RDN1Num = LineNumCount;
        //		LineNumCount = LineNumCount + 1;
        //	}
        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //		////문서상태 반품으로 변경
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030H] SET U_ProgStat = '4' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'"));
        //		////납품,반품문서번호 업데이트
        //		for (i = 0; i <= ItemInformationCount - 1; i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030L] SET U_ORDNNum = '" + ResultDocNum + "', U_RDN1Num = '" + ItemInformation[i].RDN1Num + "' WHERE DocEntry = '" + ItemInformation[i].SD030HNum + "' AND LineId = '" + ItemInformation[i].SD030LNum + "'"));
        //		}
        //	} else {
        //		goto PS_SD030_DI_API_02_DI_Error;
        //	}

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	oForm.Update();
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //	PS_SD030_DI_API_02_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //	PS_SD030_DI_API_02_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_DI_API_02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD030_FindValidateDocument
        //public bool PS_SD030_FindValidateDocument(string ObjectType)
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
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
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
        //	PS_SD030_FindValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	functionReturnValue = false;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD030_DirectionValidateDocument
        //public bool PS_SD030_DirectionValidateDocument(string DocEntry, string DocEntryNext, string Direction, string ObjectType)
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
        //		PS_SD030_FormItemEnabled();
        //		////UDO방식
        //	////다음문서가 유효하지 않다면
        //	} else {
        //		oForm.Freeze(true);
        //		oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //		PS_SD030_FormItemEnabled();
        //		////UDO방식
        //		////문서번호 필드가 입력이 가능하다면
        //		if (oForm.Items.Item("DocEntry").Enabled == true) {
        //			if ((Direction == "Next")) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) + 1));
        //			} else if ((Direction == "Prev")) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) - 1));
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
        //	PS_SD030_DirectionValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD030_Validate
        //public bool PS_SD030_Validate(string ValidateType)
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object i = null;
        //	int j = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_SD030H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
        //		MDC_Com.MDC_GF_Message(ref "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", ref "W");
        //		functionReturnValue = false;
        //		goto PS_SD030_Validate_Exit;
        //	}

        //	bool Exist = false;
        //	if (ValidateType == "검사") {
        //		////입력된 행에 대해
        //		for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [ORDR] ORDR LEFT JOIN [RDR1] RDR1 ON ORDR.DocEntry = RDR1.DocEntry WHERE CONVERT(NVARCHAR,ORDR.DocEntry) + '-' + CONVERT(NVARCHAR,RDR1.LineNum) = ' & oMat01.Columns(OrderNum).Cells(i).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [ORDR] ORDR LEFT JOIN [RDR1] RDR1 ON ORDR.DocEntry = RDR1.DocEntry WHERE CONVERT(NVARCHAR,ORDR.DocEntry) + '-' + CONVERT(NVARCHAR,RDR1.LineNum) = '" + oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value + "'", 0, 1) <= 0) {
        //				MDC_Com.MDC_GF_Message(ref "판매오더문서가 존재하지 않습니다.", ref "W");
        //				functionReturnValue = false;
        //				goto PS_SD030_Validate_Exit;
        //			}
        //		}
        //		////삭제된 행을 찾아서 삭제가능성 검사
        //		Exist = false;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		Query01 = "SELECT DocEntry,LineId FROM [@PS_SD030L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //		RecordSet01.DoQuery(Query01);
        //		for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //			Exist = false;
        //			for (j = 1; j <= oMat01.RowCount - 1; j++) {
        //				////라인번호가 같고, 품목코드가 같으면 존재하는행, LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
        //				//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (Conversion.Val(RecordSet01.Fields.Item(1).Value) == Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value) & !string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value)) {
        //					Exist = true;
        //				}
        //			}
        //			////삭제된 행중
        //			if ((Exist == false)) {
        //				//UPGRADE_WARNING: oForm.Items(DocType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////출하요청
        //				if ((oForm.Items.Item("DocType").Specific.Value == "1")) {
        //					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + Conversion.Val(RecordSet01.Fields.Item(0).Value) + "' AND PS_SD040L.U_SD030L = '" + Conversion.Val(RecordSet01.Fields.Item(1).Value) + "'", 0, 1)) > 0) {
        //						MDC_Com.MDC_GF_Message(ref "삭제된행이 다른사용자에 의해 납품되었습니다. 적용할수 없습니다.", ref "W");
        //						functionReturnValue = false;
        //						goto PS_SD030_Validate_Exit;
        //					}
        //					//UPGRADE_WARNING: oForm.Items(DocType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////선출요청
        //				} else if ((oForm.Items.Item("DocType").Specific.Value == "2")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_ProgStat FROM [PS_SD030H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "3")) {
        //						MDC_Com.MDC_GF_Message(ref "이미납품된 문서입니다. 삭제할수 없습니다.", ref "W");
        //						functionReturnValue = false;
        //						goto PS_SD030_Validate_Exit;
        //					}
        //				}
        //			}
        //			RecordSet01.MoveNext();
        //		}
        //		////수량가능성검사.
        //		for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////새로추가된 행인경우, 검사할필요없다
        //			if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value))) {
        //			} else {
        //				////매트릭스에 입력된 수량과 DB상에 존재하는 수량의 값비교
        //				//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (Conversion.Val(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value) < Conversion.Val(MDC_PS_Common.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_SD040L.U_SD030L = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1))) {
        //					MDC_Com.MDC_GF_Message(ref "출하요청,선출요청 수량보다 작습니다.", ref "W");
        //					functionReturnValue = false;
        //					goto PS_SD030_Validate_Exit;
        //				}
        //				////납품된 행이 있으면 값이 수정되어서는 안된다..
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "' AND PS_SD040L.U_SD030L = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value) + "'", 0, 1)) > 0) {
        //					Query01 = "SELECT ";
        //					////품목코드는 변경되면 안된다.
        //					Query01 = Query01 + " U_OrderNum, ";
        //					Query01 = Query01 + " U_ItemCode, ";
        //					Query01 = Query01 + " U_WhsCode ";
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					Query01 = Query01 + " FROM [@PS_SD030L] PS_SD030L WHERE PS_SD030L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_SD030L.LineId = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value) + "'";
        //					RecordSet01.DoQuery(Query01);
        //					//UPGRADE_WARNING: oMat01.Columns(WhsCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oMat01.Columns(ItemCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oMat01.Columns(OrderNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (RecordSet01.Fields.Item(0).Value == oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(1).Value == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(2).Value == oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value) {
        //					} else {
        //						MDC_Com.MDC_GF_Message(ref "이미납품된 행입니다. 수정할수 없습니다.", ref "W");
        //						functionReturnValue = false;
        //						goto PS_SD030_Validate_Exit;
        //					}
        //				}
        //			}
        //		}
        //		////
        //	} else if (ValidateType == "수정") {
        //		////수정전 수정가능여부검사
        //		//        If (oMat01.Columns("LineId").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 수정하여도 무방하다
        //		//        Else
        //		//            If oForm.Mode = fm_OK_MODE Or oForm.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
        //		//                If (oForm.Items("DocType").Specific.Value = "1") Then '//출하요청
        //		//                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" & oForm.Items("DocEntry").Specific.Value & "' AND PS_SD040L.U_SD030L = '" & Val(oMat01.Columns("LineId").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
        //		//                        MDC_Com.MDC_GF_Message "납품된 행입니다. 수정할수 없습니다.", "W"
        //		//                        PS_SD030_Validate = False
        //		//                        GoTo PS_SD030_Validate_Exit
        //		//                    End If
        //		//                ElseIf (oForm.Items("DocType").Specific.Value = "2") Then '//선출요청
        //		//                    If (MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "3") Then
        //		//                        MDC_Com.MDC_GF_Message "납품된 행입니다. 수정할수 없습니다.", "W"
        //		//                        PS_SD030_Validate = False
        //		//                        GoTo PS_SD030_Validate_Exit
        //		//                    End If
        //		//                End If
        //		//            End If
        //		//        End If
        //	} else if (ValidateType == "행삭제") {
        //		////행삭제전 행삭제가능여부검사
        //		//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(oLastColRow01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		////새로추가된 행인경우, 삭제하여도 무방하다
        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oLastColRow01).Specific.Value))) {
        //		} else {
        //			////추가,수정모드일때행삭제가능검사
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(DocType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////출하요청
        //				if ((oForm.Items.Item("DocType").Specific.Value == "1")) {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_SD040L.U_SD030L = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oLastColRow01).Specific.Value) + "'", 0, 1)) > 0) {
        //						MDC_Com.MDC_GF_Message(ref "납품된 행입니다. 삭제할수 없습니다.", ref "W");
        //						functionReturnValue = false;
        //						goto PS_SD030_Validate_Exit;
        //					}
        //					//UPGRADE_WARNING: oForm.Items(DocType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				////선출요청
        //				} else if ((oForm.Items.Item("DocType").Specific.Value == "2")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_ProgStat FROM [PS_SD030H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "3")) {
        //						MDC_Com.MDC_GF_Message(ref "납품된 행입니다. 삭제할수 없습니다.", ref "W");
        //						functionReturnValue = false;
        //						goto PS_SD030_Validate_Exit;
        //					}
        //				}
        //			}
        //		}
        //	} else if (ValidateType == "취소") {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		Query01 = "SELECT DocEntry,LineId,U_ItemCode FROM [@PS_SD030L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //		RecordSet01.DoQuery(Query01);
        //		for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //			//UPGRADE_WARNING: oForm.Items(DocType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////출하요청
        //			if ((oForm.Items.Item("DocType").Specific.Value == "1")) {
        //				if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + Conversion.Val(RecordSet01.Fields.Item(0).Value) + "' AND PS_SD040L.U_SD030L = '" + Conversion.Val(RecordSet01.Fields.Item(1).Value) + "'", 0, 1)) > 0) {
        //					MDC_Com.MDC_GF_Message(ref "납품된 문서입니다. 적용할수 없습니다.", ref "W");
        //					functionReturnValue = false;
        //					goto PS_SD030_Validate_Exit;
        //				}
        //				//UPGRADE_WARNING: oForm.Items(DocType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////선출요청
        //			} else if ((oForm.Items.Item("DocType").Specific.Value == "2")) {
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT U_ProgStat FROM [PS_SD030H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "3")) {
        //					MDC_Com.MDC_GF_Message(ref "납품된 문서입니다. 삭제할수 없습니다.", ref "W");
        //					functionReturnValue = false;
        //					goto PS_SD030_Validate_Exit;
        //				}
        //			}
        //			RecordSet01.MoveNext();
        //		}
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD030_Validate_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD030_Validate_Error:
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD030_Print_Report01
        //private void PS_SD030_Print_Report01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	int i = 0;

        //	//UPGRADE_WARNING: oForm.Items(ProgStat).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3" & oDocType01 != "선출요청") {
        //		MDC_Com.MDC_GF_Message(ref "문서상태가 납품이 아닙니다.", ref "W");
        //		return;
        //	}

        //	//UPGRADE_WARNING: oForm.Items(DocType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("DocType").Specific.Selected.Value != "2") {
        //		MDC_Com.MDC_GF_Message(ref "선출요청이 아닙니다.", ref "W");
        //		return;
        //	}
        //	//UPGRADE_WARNING: oForm.Items(ProgStat).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3" & oDocType01 != "선출요청") {
        //		MDC_Com.MDC_GF_Message(ref "납품이 되지 않았습니다.", ref "W");
        //		return;
        //	}
        //	//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("BPLId").Specific.Selected.Value != "2") {
        //		MDC_Com.MDC_GF_Message(ref "사업장이 동래가 아닙니다.", ref "W");
        //		return;
        //	}
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		//UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (oMat01.Columns.Item("ItmBsort").Cells.Item(i).Specific.Selected.Value != "105" & oMat01.Columns.Item("ItmBsort").Cells.Item(i).Specific.Selected.Value != "106") {
        //			MDC_Com.MDC_GF_Message(ref "품목이 기계공구,몰드가 아닙니다.", ref "W");
        //			return;
        //		}
        //	}

        //	MDC_PS_Common.ConnectODBC();
        //	WinTitle = "[PS_PP540_10] 출하요청서";
        //	ReportName = "PS_PP540_10.rpt";
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry = "EXEC PS_PP540_10 '선출','" + oForm.Items.Item("DocEntry").Specific.Value + "'";
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
        //	PS_SD030_Print_Report01_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD030_Print_Report02
        //private void PS_SD030_Print_Report02()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	int i = 0;
        //	string sQry01 = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	////출하요청서
        //	////여신한도체크
        //	if (PS_SD030_ValidateCreditLine() == false) {
        //		return;
        //	}


        //	//    If oForm.Items("BPLId").Specific.Selected.Value <> "1" And oForm.Items("BPLId").Specific.Selected.Value <> "4" Then
        //	//        Call MDC_Com.MDC_GF_Message("사업장이 창원,서울이 아닙니다.", "W")
        //	//        Exit Sub
        //	//    End If

        //	//// 한도초과체크후 정상프린트가 되면 요청프린트유무에 'Y' Update
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry01 = "Update [@PS_SD030H] Set U_PrtYn = 'Y' Where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	oRecordSet01.DoQuery(sQry01);

        //	//    For i = 1 To oMat01.VisualRowCount - 1
        //	//        If oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.Value <> "101" Then
        //	//            Call MDC_Com.MDC_GF_Message("품목이 휘팅이 아닙니다.", "W")
        //	//            Exit Sub
        //	//        End If
        //	//    Next

        //	MDC_PS_Common.ConnectODBC();
        //	WinTitle = "[PS_SD030_10] 레포트";
        //	ReportName = "PS_SD030_10.rpt";
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry = "EXEC PS_SD030_10 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];


        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry01 = "Select BPLName FROM [OBPL] WHERE BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
        //	oRecordSet01.DoQuery(sQry01);
        //	MDC_Globals.gRpt_Formula[1] = "BPLName";
        //	MDC_Globals.gRpt_Formula_Value[1] = Strings.Trim(oRecordSet01.Fields.Item(0).Value);

        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	PS_SD030_Print_Report02_Error:
        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_Print_Report02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD030_Print_Report03
        //////임시 출하요청서 출력
        //private void PS_SD030_Print_Report03()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	int i = 0;
        //	string sQry01 = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	//    If oForm.Items("BPLId").Specific.Selected.Value <> "1" And oForm.Items("BPLId").Specific.Selected.Value <> "4" Then
        //	//        Call MDC_Com.MDC_GF_Message("사업장이 창원,서울이 아닙니다.", "W")
        //	//        Exit Sub
        //	//    End If

        //	MDC_PS_Common.ConnectODBC();
        //	WinTitle = "[PS_SD030_20] 레포트";
        //	ReportName = "PS_SD030_20.rpt";
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry = "EXEC PS_SD030_10 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];


        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry01 = "Select BPLName FROM [OBPL] WHERE BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
        //	oRecordSet01.DoQuery(sQry01);
        //	MDC_Globals.gRpt_Formula[1] = "BPLName";
        //	MDC_Globals.gRpt_Formula_Value[1] = Strings.Trim(oRecordSet01.Fields.Item(0).Value);

        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	PS_SD030_Print_Report03_Error:
        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_Print_Report03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD030_ValidateDelivery
        //private bool PS_SD030_ValidateDelivery()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	if (PS_SD030_ValidateCreditLine() == false) {
        //		functionReturnValue = false;
        //		return functionReturnValue;
        //	}
        //	return functionReturnValue;
        //	PS_SD030_ValidateDelivery_Error:

        //	////생산완료 되었는지 확인
        //	//    Dim i As Long
        //	//
        //	//    For i = 1 To oMat01.VisualRowCount - 1
        //	//        mdc_ps_common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H
        //	//    Next
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_ValidateDelivery_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD030_ValidateCreditLine
        //private bool PS_SD030_ValidateCreditLine()
        //{
        //	bool functionReturnValue = false;
        //	////여신한도체크
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;

        //	int i = 0;
        //	decimal OCRDCreditLine = default(decimal);
        //	////고객여신한도
        //	decimal SD080CreditLine = default(decimal);
        //	////추가여신한도
        //	decimal OCRDBalance = default(decimal);
        //	////계정잔액
        //	decimal OCRDDNotesBal = default(decimal);
        //	////납품액
        //	decimal CurrentLineSum = default(decimal);
        //	////현재문서총계
        //	decimal OutPreP = default(decimal);
        //	//출고예정금액
        //	////If oMat01.Columns("WhsCode").Cells(1).Specific.Value = "104" And oMat01.Columns("ItmBsort").Cells(1).Specific.Value = "101" Then
        //	////휘팅 서울출고
        //	//UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	//UPGRADE_WARNING: oMat01.Columns(WhsCode).Cells(1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oMat01.Columns.Item("WhsCode").Cells.Item(1).Specific.Value == "101" & oMat01.Columns.Item("ItmBsort").Cells.Item(1).Specific.Value == "111") {
        //		////창원 분말
        //		if (oDS_PS_SD030H.GetValue("U_PrtYn", 0) != "Y") {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Query01 = "EXEC PS_SD081_hando '" + Strings.Trim(oDS_PS_SD030H.GetValue("U_BPLId", 0)) + "', '" + oForm.Items.Item("CardCode").Specific.Value + "','" + oForm.Items.Item("DocDate").Specific.Value + "'";
        //			RecordSet01.DoQuery(Query01);

        //			if (RecordSet01.RecordCount > 0) {
        //				////여신한도금액
        //				if (RecordSet01.Fields.Item("OverAmt").Value > 0) {
        //					MDC_Com.MDC_GF_Message(ref "여신한도로 초과했습니다.", ref "W");
        //					functionReturnValue = false;
        //					return functionReturnValue;
        //				}
        //			}
        //		}


        //		//EXEC [PS_SD081_hando] '4','12494','20110316'

        //		//        OCRDCreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT CreditLine FROM [OCRD] WHERE CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//        SD080CreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Sum(PS_SD080L.U_RequestP) FROM [@PS_SD080H] PS_SD080H LEFT JOIN [@PS_SD080L] PS_SD080L ON PS_SD080H.DocEntry = PS_SD080L.DocEntry WHERE PS_SD080H.Canceled = 'N' And PS_SD080H.U_OkYN = 'Y' AND PS_SD080H.U_DocDate = '" & oForm.Items("DocDate").Specific.Value & "' AND PS_SD080L.U_CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//'        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Balance FROM [OCRD] WHERE CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((Select Sum(Debit - Credit) from JDT1 where ShortName = '" & oForm.Items("CardCode").Specific.Value & "' And Account = '11104010'),0)", 0, 1)
        //		//        OCRDDNotesBal = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT DNotesBal FROM [OCRD] WHERE CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//'        OutPreP = MDC_PS_Common.GetValue("Select IsNull((Select Sum(LineTotal) From  [ORDR] a Inner Join [RDR1] b On a.DocEntry = b.DocEntry Where  b.LineStatus = 'O' And a.DocStatus = 'O' And a.CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//        CurrentLineSum = 0
        //		//        For i = 1 To oMat01.VisualRowCount - 1
        //		//            CurrentLineSum = CurrentLineSum + Val(oMat01.Columns("LinTotal").Cells(i).Specific.Value)
        //		//        Next
        //		//'        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal) < CurrentLineSum) Then
        //		//        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal + OutPreP) < CurrentLineSum) Then
        //		//            Call MDC_Com.MDC_GF_Message("여신한도가 부족합니다.", "W")
        //		//            PS_SD030_ValidateCreditLine = False
        //		//            Exit Function
        //		//        End If
        //	}
        //	return functionReturnValue;
        //	PS_SD030_ValidateCreditLine_Error:
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_ValidateCreditLine_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion
    }
}
