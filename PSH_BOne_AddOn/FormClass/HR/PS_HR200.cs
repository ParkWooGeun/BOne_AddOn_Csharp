using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
namespace MDC_PS_Addon
{
	internal class PS_HR200
	{
//****************************************************************************************************************
////  File              : PS_HR200.cls
////  Module         : 인사관리 > 인사급여코드등록
////  Desc             :
////  FormType       : PS_HR200
////  Create Date    : 2012.2.9
////  Modified Date  :
////  Creator           : N.G.Y
////  Company         : Poongsan Holdings
//****************************************************************************************************************

		public string oFormUniqueID01;
		public SAPbouiCOM.Form oForm01;
		public SAPbouiCOM.Matrix oMat01;
			//등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_HR200H;
			//등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_HR200L;

			//클래스에서 선택한 마지막 아이템 Uid값
		private string oLastItemUID01;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private string oLastColUID01;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLastColRow01;

////사용자구조체
		private struct ItemInformations
		{
			public string ItemCode;
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

		private string oDocType01;
		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

//*******************************************************************
// .srf 파일로부터 폼을 로드한다.
//*******************************************************************
		public void LoadForm(string oFromDocEntry01 = "")
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string oInnerXml01 = null;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_HR200.srf");
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

			//매트릭스의 타이틀높이와 셀높이를 고정
			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
			}

			oFormUniqueID01 = "PS_HR200_" + GetTotalFormsCount();
			SubMain.AddForms(this, oFormUniqueID01);
			////폼추가
			SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));
			//폼 할당
			oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

			oForm01.SupportedModes = -1;
			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
			oForm01.DataBrowser.BrowseBy = "Code";
			////UDO방식일때

			oForm01.Freeze(true);
			PS_HR200_CreateItems();
			PS_HR200_ComboBox_Setting();
			PS_HR200_CF_ChooseFromList();
			PS_HR200_EnableMenus();
			PS_HR200_SetDocument(oFromDocEntry01);
			PS_HR200_FormResize();

			oForm01.EnableMenu(("1283"), true);
			//// 삭제
			oForm01.EnableMenu(("1287"), true);
			//// 복제
			oForm01.EnableMenu(("1286"), false);
			//// 닫기
			oForm01.EnableMenu(("1284"), false);
			//// 취소
			oForm01.EnableMenu(("1293"), true);
			//// 행삭제

			oForm01.Update();
			oForm01.Freeze(false);

			oForm01.Visible = true;
			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			return;
			LoadForm_Error:
			oForm01.Update();
			oForm01.Freeze(false);
			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oForm01 = null;
			SubMain.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			switch (pval.EventType) {
				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
					////1
					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
					////2
					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
					////5
					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK:
					////6
					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
					////7
					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
					////8
					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
					////10
					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
					////11
					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
					////18
					break;
				////et_FORM_ACTIVATE
				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
					////19
					break;
				////et_FORM_DEACTIVATE
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
					////20
					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
					////27
					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
					////3
					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
					////4
					break;
				////et_LOST_FOCUS
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
					////17
					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
					break;
			}
			return;
			Raise_ItemEvent_Error:
			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}


		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement


			short i = 0;

			////BeforeAction = True
			if ((pval.BeforeAction == true)) {
				switch (pval.MenuUID) {
					case "1284":
						//취소
						break;
					case "1286":
						//닫기
						break;
					case "1293":
						//행삭제
						Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
						break;
					case "1281":
						//찾기
						break;
					case "1282":
						//추가
						break;
					case "1288":
					case "1289":
					case "1290":
					case "1291":
						//레코드이동버튼
						break;
				}
			////BeforeAction = False
			} else if ((pval.BeforeAction == false)) {
				switch (pval.MenuUID) {
					case "1284":
						//취소
						break;
					case "1286":
						//닫기
						break;
					case "1293":
						//행삭제
						Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
						break;
					case "1281":
						//찾기
						PS_HR200_FormItemEnabled();
						////UDO방식
						oForm01.Items.Item("Name").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						break;
					case "1282":
						//추가
						PS_HR200_FormItemEnabled();
						////UDO방식
						PS_HR200_AddMatrixRow(0, ref true);
						////UDO방식
						break;
					case "1288":
					case "1289":
					case "1290":
					case "1291":
						//레코드이동버튼
						PS_HR200_FormItemEnabled();
						break;

					//복제(2013.03.13 송명규 추가)
					case "1287":

						oForm01.Freeze(true);
						PS_HR200_FormClear();
						//oDS_PS_HR200H.setValue "Code", 0, ""

						for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
							oMat01.FlushToDataSource();
							oDS_PS_HR200L.SetValue("Code", i, "");
							oMat01.LoadFromDataSource();
						}

						oForm01.Freeze(false);
						break;

				}
			}
			return;
			Raise_MenuEvent_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////BeforeAction = True
			if ((BusinessObjectInfo.BeforeAction == true)) {
				switch (BusinessObjectInfo.EventType) {
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
						////33
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
						////34
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
						////35
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
						////36
						break;
				}
			////BeforeAction = False
			} else if ((BusinessObjectInfo.BeforeAction == false)) {
				switch (BusinessObjectInfo.EventType) {
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
						////33
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
						////34
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
						////35
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
						////36
						break;
				}
			}
			return;
			Raise_FormDataEvent_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {
				//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
				//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
				//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
				//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
				//            MenuCreationParams01.uniqueID = "MenuUID"
				//            MenuCreationParams01.String = "메뉴명"
				//            MenuCreationParams01.Enabled = True
				//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
				//        End If
			} else if (pval.BeforeAction == false) {
				//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
				//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
				//        End If
			}
			if (pval.ItemUID == "Mat01") {
				if (pval.Row > 0) {
					oLastItemUID01 = pval.ItemUID;
					oLastColUID01 = pval.ColUID;
					oLastColRow01 = pval.Row;
				}
			} else {
				oLastItemUID01 = pval.ItemUID;
				oLastColUID01 = "";
				oLastColRow01 = 0;
			}
			return;
			Raise_RightClickEvent_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {
				if (pval.ItemUID == "PS_HR200") {
					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
					}
				}
				if (pval.ItemUID == "1") {
					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
						if (PS_HR200_DataValidCheck() == false) {
							BubbleEvent = false;
							return;
						}

						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						oDocEntry01 = oForm01.Items.Item("Code").Specific.VALUE;
						oFormMode01 = oForm01.Mode;

					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
						if (PS_HR200_DataValidCheck() == false) {
							BubbleEvent = false;
							return;
						}

						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						oDocEntry01 = oForm01.Items.Item("Code").Specific.VALUE;
						oFormMode01 = oForm01.Mode;

					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
					}
				}
			} else if (pval.BeforeAction == false) {
				if (pval.ItemUID == "PS_HR200") {
					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
					}
				}
				if (pval.ItemUID == "1") {
					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
						if (pval.ActionSuccess == true) {
							PS_HR200_FormItemEnabled();
							PS_HR200_AddMatrixRow(0, ref true);
							////UDO방식일때
						}
					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
						if (pval.ActionSuccess == true) {
							PS_HR200_FormItemEnabled();
						}
					}
				}

			}
			return;
			Raise_EVENT_ITEM_PRESSED_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {


			} else if (pval.BeforeAction == false) {

			}
			return;
			Raise_EVENT_KEY_DOWN_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			oForm01.Freeze(true);
			if (pval.BeforeAction == true) {

			} else if (pval.BeforeAction == false) {

				if (pval.ItemChanged == true) {

				}
			}
			oForm01.Freeze(false);
			return;
			Raise_EVENT_COMBO_SELECT_Error:
			oForm01.Freeze(false);
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {
				//        If pval.ItemUID = "Mat01" Then
				//            If pval.Row > 0 Then
				//                Call oMat01.SelectRow(pval.Row, True, False)
				//            End If
				//        End If
				if (pval.ItemUID == "Mat01") {
					if (pval.Row > 0) {
						oLastItemUID01 = pval.ItemUID;
						oLastColUID01 = pval.ColUID;
						oLastColRow01 = pval.Row;

						oMat01.SelectRow(pval.Row, true, false);
					}
				} else {
					oLastItemUID01 = pval.ItemUID;
					oLastColUID01 = "";
					oLastColRow01 = 0;
				}
			} else if (pval.BeforeAction == false) {

			}
			return;
			Raise_EVENT_CLICK_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {

			} else if (pval.BeforeAction == false) {

			}
			return;
			Raise_EVENT_DOUBLE_CLICK_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			object oTempClass = null;
			if (pval.BeforeAction == true) {
				if (pval.ItemUID == "Mat01") {

				}
			} else if (pval.BeforeAction == false) {

			}
			return;
			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement


			string Query01 = null;
			SAPbobsCOM.Recordset RecordSet01 = null;

			oForm01.Freeze(true);
			if (pval.BeforeAction == true) {
				if (pval.ItemChanged == true) {


					if ((pval.ItemUID == "Mat01")) {
						if (pval.ColUID == "Code") {
							oMat01.FlushToDataSource();

							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							oDS_PS_HR200L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE);
							oDS_PS_HR200L.SetValue("U_Seq", pval.Row - 1, Convert.ToString(pval.Row));
							if (oMat01.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_HR200L.GetValue("U_" + pval.ColUID, pval.Row - 1)))) {
								PS_HR200_AddMatrixRow((pval.Row));
							}
							oMat01.LoadFromDataSource();
						}
						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						oMat01.Columns.Item("UseYN").Cells.Item(pval.Row).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						//기본으로 'Y' 세팅

						//                Call oMat01.Columns(pval.ColUID).Cells(pval.Row).Click(ct_Regular)
						oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					} else {

					}
				}

			} else if (pval.BeforeAction == false) {

			}
			oForm01.Freeze(false);
			return;
			Raise_EVENT_VALIDATE_Error:
			oForm01.Freeze(false);
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {

			} else if (pval.BeforeAction == false) {
				PS_HR200_FormItemEnabled();
				PS_HR200_AddMatrixRow(oMat01.VisualRowCount);
				////UDO방식
			}
			return;
			Raise_EVENT_MATRIX_LOAD_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {

			} else if (pval.BeforeAction == false) {
				PS_HR200_FormResize();
			}
			return;
			Raise_EVENT_RESIZE_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {

			} else if (pval.BeforeAction == false) {
			}
			return;
			Raise_EVENT_CHOOSE_FROM_LIST_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}


		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.ItemUID == "Mat01") {
				if (pval.Row > 0) {
					oLastItemUID01 = pval.ItemUID;
					oLastColUID01 = pval.ColUID;
					oLastColRow01 = pval.Row;
				}
			} else {
				oLastItemUID01 = pval.ItemUID;
				oLastColUID01 = "";
				oLastColRow01 = 0;
			}
			return;
			Raise_EVENT_GOT_FOCUS_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if (pval.BeforeAction == true) {
			} else if (pval.BeforeAction == false) {
				SubMain.RemoveForms(oFormUniqueID01);
				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oForm01 = null;
				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oMat01 = null;
			}
			return;
			Raise_EVENT_FORM_UNLOAD_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			if ((oLastColRow01 > 0)) {
				if (pval.BeforeAction == true) {
					////행삭제전 행삭제가능여부검사
				} else if (pval.BeforeAction == false) {
					for (i = 1; i <= oMat01.VisualRowCount; i++) {
						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
					}
					oMat01.FlushToDataSource();
					oDS_PS_HR200L.RemoveRecord(oDS_PS_HR200L.Size - 1);
					oMat01.LoadFromDataSource();
					if (oMat01.RowCount == 0) {
						PS_HR200_AddMatrixRow(0);
					} else {
						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_HR200L.GetValue("U_Code", oMat01.RowCount - 1)))) {
							PS_HR200_AddMatrixRow(oMat01.RowCount);
						}
					}
				}
			}
			return;
			Raise_EVENT_ROW_DELETE_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}


		private bool PS_HR200_CreateItems()
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement


			oDS_PS_HR200H = oForm01.DataSources.DBDataSources("@PS_HR200H");
			oDS_PS_HR200L = oForm01.DataSources.DBDataSources("@PS_HR200L");
			oMat01 = oForm01.Items.Item("Mat01").Specific;

			oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
			return functionReturnValue;
			PS_HR200_CreateItems_Error:
			//    oMat01.AutoResizeColumns

			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
			return functionReturnValue;
		}

		public void PS_HR200_ComboBox_Setting()
		{
			 // ERROR: Not supported in C#: OnErrorStatement


			////콤보에 기본값설정
			SAPbouiCOM.ComboBox oCombo = null;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;

			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			MDC_PS_Common.Combo_ValidValues_Insert("PS_HR200", "Mat01", "UseYN", "Y", "Y");
			MDC_PS_Common.Combo_ValidValues_Insert("PS_HR200", "Mat01", "UseYN", "N", "N");
			MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("UseYN"), "PS_HR200", "Mat01", "UseYN");

			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;

			return;
			PS_HR200_ComboBox_Setting_Error:
			oForm01.Freeze(false);
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void PS_HR200_CF_ChooseFromList()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			return;
			PS_HR200_CF_ChooseFromList_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void PS_HR200_FormItemEnabled()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			oForm01.Freeze(true);
			if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
				////각모드에따른 아이템설정
				//
				oForm01.Items.Item("Code").Enabled = true;
				oForm01.Items.Item("Mat01").Enabled = true;
				PS_HR200_FormClear();
				////UDO방식
				oForm01.EnableMenu("1281", true);
				////찾기
				oForm01.EnableMenu("1282", false);
				////추가



			} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
				////각모드에따른 아이템설정
				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oForm01.Items.Item("Code").Specific.VALUE = "";
				oForm01.Items.Item("Code").Enabled = true;
				oForm01.Items.Item("Mat01").Enabled = false;
				oForm01.EnableMenu("1281", false);
				////찾기
				oForm01.EnableMenu("1282", true);
				////추가

			} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
				////각모드에따른 아이템설정

				oForm01.Items.Item("Code").Enabled = false;
				oForm01.Items.Item("Mat01").Enabled = true;

			}
			oForm01.Freeze(false);
			return;
			PS_HR200_FormItemEnabled_Error:
			oForm01.Freeze(false);
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void PS_HR200_AddMatrixRow(int oRow, ref bool RowIserted = false)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			oForm01.Freeze(true);
			////행추가여부
			if (RowIserted == false) {
				oDS_PS_HR200L.InsertRecord((oRow));
			}
			oMat01.AddRow();
			oDS_PS_HR200L.Offset = oRow;
			oDS_PS_HR200L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
			oMat01.LoadFromDataSource();
			oForm01.Freeze(false);
			return;
			PS_HR200_AddMatrixRow_Error:
			oForm01.Freeze(false);
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void PS_HR200_FormClear()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			string DocEntry = null;
			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_HR200'", ref "");
			if (string.IsNullOrEmpty(DocEntry) | DocEntry == "0") {
				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oForm01.Items.Item("DocEntry").Specific.VALUE = 1;
				//        oForm01.Items("Code").Specific.VALUE = 1
			} else {
				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oForm01.Items.Item("DocEntry").Specific.VALUE = DocEntry;
				//        oForm01.Items("Code").Specific.VALUE = DocEntry
			}
			return;
			PS_HR200_FormClear_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void PS_HR200_EnableMenus()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////메뉴활성화
			//    Call oForm01.EnableMenu("1288", True)
			//    Call oForm01.EnableMenu("1289", True)
			//    Call oForm01.EnableMenu("1290", True)
			//    Call oForm01.EnableMenu("1291", True)
			////Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
			MDC_Com.MDC_GP_EnableMenus(oForm01, false, false, true, true, false, true, true, true, true,
			false, false, false, false, false, false);
			////메뉴설정
			return;
			PS_HR200_EnableMenus_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void PS_HR200_SetDocument(string oFromDocEntry01)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
				PS_HR200_FormItemEnabled();
				PS_HR200_AddMatrixRow(0, ref true);
				////UDO방식일때
			} else {
				//        oForm01.Mode = fm_FIND_MODE
				//        Call PS_HR200_FormItemEnabled
				//        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
				//        oForm01.Items("1").Click ct_Regular
			}
			return;
			PS_HR200_SetDocument_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}


		public bool PS_HR200_DataValidCheck()
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement

			functionReturnValue = false;
			int i = 0;
			if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
				PS_HR200_FormClear();
			}

			//사업장 미입력 시
			//UPGRADE_WARNING: oForm01.Items(Name).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (string.IsNullOrEmpty(oForm01.Items.Item("Name").Specific.VALUE)) {
				SubMain.Sbo_Application.SetStatusBarMessage("코드명이 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
				functionReturnValue = false;
				return functionReturnValue;
			}

			//라인정보 미입력 시
			if (oMat01.VisualRowCount == 1) {
				SubMain.Sbo_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
				functionReturnValue = false;
				return functionReturnValue;
			}

			for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
				//UPGRADE_WARNING: oMat01.Columns(Code).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				if ((string.IsNullOrEmpty(oMat01.Columns.Item("Code").Cells.Item(i).Specific.VALUE))) {
					SubMain.Sbo_Application.SetStatusBarMessage("코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
					oMat01.Columns.Item("Code").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					functionReturnValue = false;
					return functionReturnValue;
				}

				//UPGRADE_WARNING: oMat01.Columns(CodeNm).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				if ((string.IsNullOrEmpty(oMat01.Columns.Item("CodeNm").Cells.Item(i).Specific.VALUE))) {
					SubMain.Sbo_Application.SetStatusBarMessage("코드명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
					oMat01.Columns.Item("CodeNm").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					functionReturnValue = false;
					return functionReturnValue;
				}
			}



			oMat01.FlushToDataSource();
			oDS_PS_HR200L.RemoveRecord(oDS_PS_HR200L.Size - 1);
			oMat01.LoadFromDataSource();

			if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
				PS_HR200_FormClear();
			}

			functionReturnValue = true;
			return functionReturnValue;
			PS_HR200_DataValidCheck_Error:
			functionReturnValue = false;
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
			return functionReturnValue;
		}


		private void PS_HR200_MTX01()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////메트릭스에 데이터 로드
			oForm01.Freeze(true);
			int i = 0;
			string Query01 = null;
			SAPbobsCOM.Recordset RecordSet01 = null;
			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string Param01 = null;
			string Param02 = null;
			string Param03 = null;
			string Param04 = null;
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param01 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param02 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param03 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param04 = Strings.Trim(oForm01.Items.Item("Param01").Specific.VALUE);

			Query01 = "SELECT 10";
			RecordSet01.DoQuery(Query01);

			oMat01.Clear();
			oMat01.FlushToDataSource();
			oMat01.LoadFromDataSource();

			if ((RecordSet01.RecordCount == 0)) {
				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
				goto PS_HR200_MTX01_Exit;
			}

			SAPbouiCOM.ProgressBar ProgressBar01 = null;
			ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
				if (i != 0) {
					oDS_PS_HR200L.InsertRecord((i));
				}
				oDS_PS_HR200L.Offset = i;
				oDS_PS_HR200L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
				oDS_PS_HR200L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
				RecordSet01.MoveNext();
				ProgressBar01.Value = ProgressBar01.Value + 1;
				ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
			}
			oMat01.LoadFromDataSource();
			oMat01.AutoResizeColumns();
			oForm01.Update();

			ProgressBar01.Stop();
			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgressBar01 = null;
			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			oForm01.Freeze(false);
			return;
			PS_HR200_MTX01_Exit:
			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			oForm01.Freeze(false);
			if ((ProgressBar01 != null)) {
				ProgressBar01.Stop();
			}
			return;
			PS_HR200_MTX01_Error:
			ProgressBar01.Stop();
			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgressBar01 = null;
			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			oForm01.Freeze(false);
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}


		private void PS_HR200_FormResize()
		{
			 // ERROR: Not supported in C#: OnErrorStatement


			return;
			PS_HR200_FormResize_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public bool PS_HR200_Validate(string ValidateType)
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement

			functionReturnValue = true;
			object i = null;
			int j = 0;
			string Query01 = null;
			SAPbobsCOM.Recordset RecordSet01 = null;
			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			if (ValidateType == "수정") {
				////삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
			} else if (ValidateType == "행삭제") {
				////행삭제전 행삭제가능여부검사
			} else if (ValidateType == "취소") {
			}
			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			return functionReturnValue;
			PS_HR200_Validate_Exit:
			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			return functionReturnValue;
			PS_HR200_Validate_Error:
			functionReturnValue = false;
			SubMain.Sbo_Application.SetStatusBarMessage("PS_HR200_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
			return functionReturnValue;
		}
	}
}
