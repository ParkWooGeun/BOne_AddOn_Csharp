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
	internal class PS_PP071
	{
//****************************************************************************************************************
////  File           : PS_PP071.cls
////  Module         : PP
////  Description    : 작지조회
////  FormType       : PS_PP071
////  Create Date    : 2010.12.23
////  Modified Date  :
////  Creator        : Youn Je Hyung
////  Company        : Poongsan Holdings
//****************************************************************************************************************

		public string oFormUniqueID01;
		public SAPbouiCOM.Form oForm01;
		public SAPbouiCOM.Matrix oMat01;
		public SAPbouiCOM.Grid oGrid01;
			//등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP071H;
			//등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP071L;

			////부모폼
		public SAPbouiCOM.Form oBaseForm01;
		public string oBaseItemUID01;
		public string oBaseColUID01;
		public int oBaseColRow01;
		public string oBaseBPLId01;
		public string oBaseOrdGbn01;

			//클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Item_UID;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private string oLast_Col_UID;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Col_Row;

		private int oLast_Mode;

//****************************************************************************************************************
// .srf 파일로부터 폼을 로드한다.
//****************************************************************************************************************
		public void LoadForm(ref SAPbouiCOM.Form oForm02 = null, string oItemUID02 = "", string oColUID02 = "", int oColRow02 = 0, string oBPLId02 = "", string oOrdGbn02 = "")
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string oInnerXml01 = null;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_PP071.srf");
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

			//매트릭스의 타이틀높이와 셀높이를 고정
			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
			}

			oFormUniqueID01 = "PS_PP071_" + GetTotalFormsCount();
			SubMain.AddForms(this, oFormUniqueID01);
			////폼추가
			SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));

			//폼 할당
			oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

			oForm01.SupportedModes = -1;
			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

			//////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//************************************************************************************************************
			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
			//    oForm01.DataBrowser.BrowseBy = "DocNum"
			//************************************************************************************************************
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////

			oForm01.Freeze(true);

			oBaseForm01 = oForm02;
			oBaseItemUID01 = oItemUID02;
			oBaseColUID01 = oColUID02;
			oBaseColRow01 = oColRow02;
			oBaseBPLId01 = oBPLId02;
			oBaseOrdGbn01 = oOrdGbn02;

			CreateItems();
			ComboBox_Setting();
			//    FormItemEnabled

			oForm01.EnableMenu(("1283"), false);
			//// 삭제
			oForm01.EnableMenu(("1286"), false);
			//// 닫기
			oForm01.EnableMenu(("1287"), false);
			//// 복제
			oForm01.EnableMenu(("1284"), false);
			//// 취소
			oForm01.EnableMenu(("1293"), false);
			//// 행삭제

			oForm01.Update();
			oForm01.Freeze(false);
			oForm01.Visible = true;

			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			return;
			LoadForm_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			oForm01.Update();
			oForm01.Freeze(false);
			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			if ((oForm01 == null) == false) {
				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oForm01 = null;
			}
			MDC_Com.MDC_GF_Message(ref "LoadForm_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		private void CreateItems()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////디비데이터 소스 개체 할당
			//    Set oDS_PS_PP071H = oForm01.DataSources.DBDataSources("@PS_PP071H")
			//    Set oDS_PS_PP071L = oForm01.DataSources.DBDataSources("@PS_PP071L")

			//// 메트릭스 개체 할당
			//    Set oMat01 = oForm01.Items("Mat01").Specific

			oGrid01 = oForm01.Items.Item("Grid01").Specific;
			oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
			oForm01.DataSources.DataTables.Add("ZTEMP");

			return;
			CreateItems_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void ComboBox_Setting()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////콤보에 기본값설정
			SAPbouiCOM.ComboBox oCombo = null;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;

			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("ItemType").Specific.ValidValues.Add("선택", "선택");
			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("ItemType").Specific), ref "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", ref "", ref false, ref false);
			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("Mark").Specific.ValidValues.Add("선택", "선택");
			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("Mark").Specific), ref "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", ref "", ref false, ref false);
			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("Mark").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			return;
			ComboBox_Setting_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			MDC_Com.MDC_GF_Message(ref "ComboBox_Setting_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

//****************************************************************************************************************
//// ItemEventHander
//****************************************************************************************************************
		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			int ErrNum = 0;
			SAPbouiCOM.ProgressBar ProgressBar01 = null;

			////BeforeAction = True
			if ((pval.BeforeAction == true)) {
				switch (pval.EventType) {
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
						////1
						if (pval.ItemUID == "Button01") {
							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
								Search_Grid_Data();
							} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
							} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
							}
						}
						if (pval.ItemUID == "Button02") {
							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
								PS_PP071_SetBaseForm();
								////부모폼에입력
								oForm01.Close();
							} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
							} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
							}
						}
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
						////2
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
						////5
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:
						////6
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
						////7
						if (pval.ItemUID == "Grid01") {
							if (pval.Row == -1) {
								oGrid01.Columns.Item(pval.ColUID).TitleObject.Sortable = true;
							} else {
								if (oGrid01.Rows.SelectedRows.Count > 0) {
									PS_PP071_SetBaseForm();
									////부모폼에입력
									oForm01.Close();
								} else {
									BubbleEvent = false;
								}
							}
						}
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
						////8
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
						////10
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
						////11
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
						////18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
						////19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
						////20
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
						////27
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
						////3
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
						////4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
						////17
						break;
				}

				//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			////BeforeAction = False
			} else if ((pval.BeforeAction == false)) {
				switch (pval.EventType) {
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
						////1
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
						////2
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
						////5
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:
						////6
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
						////7
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
						////8
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
						////10
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
						////11
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
						////18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
						////19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
						////20
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
						////27
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
						////3
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
						////4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
						////17
						SubMain.RemoveForms(oFormUniqueID01);
						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						oForm01 = null;
						//UPGRADE_NOTE: oGrid01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						oGrid01 = null;
						break;
				}
			}
			return;
			Raise_ItemEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgressBar01 = null;
			if (ErrNum == 101) {
				ErrNum = 0;
				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
				BubbleEvent = false;
			} else {
				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
			}
		}

		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;

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
						break;
					case "1281":
						//찾기
						break;
					case "1282":
						//추가
						break;
					case "1285":
						//복원
						break;
					case "1288":
					case "1289":
					case "1290":
					case "1291":
						//레코드이동버튼
						break;
				}

				//-----------------------------------------------------------------------------------------------------------
			////BeforeAction = False
			} else if ((pval.BeforeAction == false)) {
				switch (pval.MenuUID) {
					case "1284":
						//취소
						break;
					case "1286":
						//닫기
						break;
					case "1285":
						//복원
						break;
					case "1293":
						//행삭제
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
			}
			return;
			Raise_MenuEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if ((eventInfo.BeforeAction == true)) {

			} else if ((eventInfo.BeforeAction == false)) {
				////작업
			}
			return;
			Raise_RightClickEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		private void FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string DocNum = null;
			string LineId = null;
			short ErrNum = 0;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet = null;

			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			//--------------------------------------------------------------
			//Header--------------------------------------------------------
			switch (oUID) {

			}

			//--------------------------------------------------------------
			//Line----------------------------------------------------------
			if (oUID == "Mat01") {
				switch (oCol) {
				}
			}

			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			return;
			FlushToItemValue_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void Search_Grid_Data()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			SAPbobsCOM.Recordset oRecordSet = null;
			string sQry = null;

			string Param01 = null;
			string Param02 = null;
			string Param03 = null;
			string Param04 = null;
			string Param05 = null;
			string Param06 = null;

			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			oForm01.Freeze(true);

			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param01 = Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param02 = oForm01.Items.Item("ItemName").Specific.VALUE;
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param03 = oForm01.Items.Item("Size").Specific.VALUE;
			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param04 = oForm01.Items.Item("ItemType").Specific.Selected.VALUE;
			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param05 = oForm01.Items.Item("Mark").Specific.Selected.VALUE;
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param06 = oForm01.Items.Item("OrdNum").Specific.VALUE;


			//sQry = "EXEC PS_PP071_01  '" & oBaseBPLId01 & "','" & oBaseOrdGbn01 & "','" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "', '" & Param05 & "'"
			sQry = "EXEC PS_PP070_03  '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "', '" + Param05 + "', '" + Param06 + "'";
			//sQry = "EXEC PS_PP071_01  '" & 1 & "','" & 101 & "','" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "', '" & Param05 & "'"

			/// Procedure 실행(Grid 사용)
			oForm01.DataSources.DataTables.Item(0).ExecuteQuery((sQry));
			oGrid01.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP");

			GridSetting();


			oForm01.Freeze(false);

			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			return;
			Search_Grid_Data_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			MDC_Com.MDC_GF_Message(ref "Search_Grid_Data_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}


//****************************************************************************************************************
//// Grid 꾸며주기
//****************************************************************************************************************
		private void GridSetting()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			short i = 0;
			string sColsTitle = null;
			string sColsLine = null;

			oForm01.Freeze(true);

			oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

			//    oGrid01.Columns.Item(0).LinkedObjectType = lf_BusinessPartner
			//    oGrid01.Columns.Item(4).LinkedObjectType = lf_Invoice
			//    oGrid01.Columns.Item(2).LinkedObjectType = lf_Items

			for (i = 0; i <= oGrid01.Columns.Count - 1; i++) {
				sColsTitle = oGrid01.Columns.Item(i).TitleObject.Caption;

				oGrid01.Columns.Item(i).Editable = false;

				if (oGrid01.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float) {
					oGrid01.Columns.Item(i).RightJustified = true;
				}

				//        If InStr(1, sColsTitle, "T(mm)") > 0 Or _
				//'           InStr(1, sColsTitle, "W(mm)") > 0 Or _
				//'           InStr(1, sColsTitle, "L(mm)") > 0 Or _
				//'           InStr(1, sColsTitle, "Weight") > 0 Or _
				//'           InStr(1, sColsTitle, "Qty(Kg)") > 0 Then
				//            oGrid01.Columns(i).RightJustified = True
				//            oGrid01.Columns(i).BackColor = &HE0E0E0
				//        End If
				//
				//        If InStr(1, sColsTitle, "수량") > 0 Or _
				//'           InStr(1, sColsTitle, "중량") > 0 Then
				//            oGrid01.Columns(i).RightJustified = True
				//            oGrid01.Columns(i).BackColor = &HFFC0C0
				//        End If
			}

			oForm01.Freeze(false);

			return;
			GridSetting_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "GridSetting_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		private void PS_PP071_SetBaseForm()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string ItemCode01 = null;
			SAPbouiCOM.Matrix oBaseMat01 = null;
			if (oBaseForm01 == null) {
				////DoNothing
			////사용하려는폼의 폼타입
			} else if (oBaseForm01.TypeEx == "PS_PP070") {
				oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
				////부모폼의매트릭스
				////선택된행의수
				for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++) {
					//UPGRADE_WARNING: oBaseMat01.Columns(PP030No).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					//UPGRADE_WARNING: oGrid01.DataTable.Columns().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oBaseMat01.Columns.Item("PP030No").Cells.Item(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns.Item("문서번호").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
					oBaseColRow01 = oBaseColRow01 + 1;
				}
			////사용하려는폼의 폼타입
			} else if (oBaseForm01.TypeEx == "PS_PP080") {
				oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
				////부모폼의매트릭스
				////선택된행의수
				for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++) {
					//UPGRADE_WARNING: oBaseMat01.Columns(PP030No).Cells(oBaseColRow01).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					//UPGRADE_WARNING: oGrid01.DataTable.Columns().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oBaseMat01.Columns.Item("PP030No").Cells.Item(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns.Item("문서번호").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
					oBaseColRow01 = oBaseColRow01 + 1;
				}
				//    ElseIf oBaseForm01.TypeEx = "PS_QM010" Then '//사용하려는폼의 폼타입
				//        oBaseForm01.Items("ItemCode").Specific.Value = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).Value
				//    ElseIf oBaseForm01.TypeEx = "PS_PP077" Then '//사용하려는폼의 폼타입
				//        oBaseForm01.Items("ItemCode").Specific.Value = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).Value
				//    ElseIf oBaseForm01.TypeEx = "PS_PP078" Then '//사용하려는폼의 폼타입
				//        oBaseForm01.Items("ItemCode").Specific.Value = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).Value
			}
			return;
			PS_PP071_SetBaseForm_Error:
			SubMain.Sbo_Application.SetStatusBarMessage("PS_PP071_SetBaseForm_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}
	}
}
