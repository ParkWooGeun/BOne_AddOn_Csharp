//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//	internal class ZPY510
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY510.cls
//////  Module         : 원천징수>근로소득
//////  Desc           : 종전근무지 일괄생성
//////  FormType       : 2010110510
//////  Create Date    : 2010.01.05
//////  Modified Date  :
//////  Creator        : Choi Dong Kwon
//////  Copyright  (c) Morning Data
//////****************************************************************************
//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//		private SAPbouiCOM.Grid oGrid;
//		private SAPbouiCOM.DataTable oDS_ZPY510;
//		private string mJSNYER;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY510.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//------------------------------------------------------------------------
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//------------------------------------------------------------------------
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY510_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//--------------------------------------------------------------------------------------------------------------
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//--------------------------------------------------------------------------------------------------------------

//			SubMain.AddForms(this, oFormUniqueID, "ZPY510");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

//			////////////////////////////////////////////////////////////////////////////////
//			//***************************************************************
//			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//			//oForm.DataBrowser.BrowseBy = "DocNum"
//			//***************************************************************
//			////////////////////////////////////////////////////////////////////////////////

//			CreateItems();

//			oForm.EnableMenu(("1293"), false);
//			/// 행삭제
//			oForm.EnableMenu(("1284"), false);
//			/// 취소

//			oForm.Update();
//			oForm.Visible = true;

//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("종전근무지 일괄생성을 실행시킬 수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			if ((oForm == null) == false) {
//				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oForm = null;
//			}
//		}

////*******************************************************************
//// Item Initial
////*******************************************************************
//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.EditText oEdit = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//// UserDataSources
//			var _with1 = oForm.DataSources.UserDataSources;
//			_with1.Add("JSNYER", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			_with1.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			_with1.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);

//			oEdit = oForm.Items.Item("JSNYER").Specific;
//			oEdit.DataBind.SetBound(true, "", "JSNYER");
//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");
//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");

//			oForm.DataSources.UserDataSources.Item("JSNYER").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY");

//			////사업장
//			oCombo = oForm.Items.Item("FCLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.ValidValues.Add("%", "전체");
//			oForm.Items.Item("FCLTCOD").DisplayDesc = true;
//			oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("TCLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.ValidValues.Add("%", "전체");
//			oForm.Items.Item("TCLTCOD").DisplayDesc = true;
//			oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

//			/// Grid
//			oGrid = oForm.Items.Item("Grid1").Specific;
//			oForm.DataSources.DataTables.Add("ZPY510");

//			oDS_ZPY510 = oForm.DataSources.DataTables.Item("ZPY510");
//			oDS_ZPY510.ExecuteQuery("EXEC ZPY510 '1900', '', '', ''");
//			oGrid.DataTable = oDS_ZPY510;

//			titleSetting();

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

////---------------------------------------------------------------------------------------
//// Procedure : TitleSetting
//// Author    : Choi Dong Kwon
//// Date      : 2008-07-15
//// Purpose   : Grid의 Column Title 지정
////---------------------------------------------------------------------------------------
////
//		private void titleSetting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			short ErrNum = 0;
//			short i = 0;

//			string[] COLNAM = new string[25];

//			/// Initial
//			ErrNum = 0;

//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			///  컬럼명
//			COLNAM[0] = "선택";
//			COLNAM[1] = "사원번호";
//			COLNAM[2] = "성명";
//			COLNAM[3] = "사원순번";
//			COLNAM[4] = "자사코드";
//			COLNAM[5] = "종전자사코드";
//			COLNAM[6] = "사업자번호";
//			COLNAM[7] = "귀속시작일";
//			COLNAM[8] = "귀속종료일";
//			COLNAM[9] = "당해감면시작일";
//			COLNAM[10] = "당해감면종료일";
//			COLNAM[11] = "급여금액";
//			COLNAM[12] = "상여금액";
//			COLNAM[13] = "인정상여";
//			COLNAM[14] = "주식매수선택권행사이익";
//			COLNAM[15] = "우리사주조합인출";
//			COLNAM[16] = "비과세총계";
//			COLNAM[17] = "건강보험";
//			COLNAM[18] = "고용보험";
//			COLNAM[19] = "국민연금";
//			COLNAM[20] = "연금보험료";
//			COLNAM[21] = "소득세";
//			COLNAM[22] = "주민세";
//			COLNAM[23] = "농특세";
//			COLNAM[24] = "퇴직연금";

//			//// 컬럼명 셋팅
//			for (i = 0; i <= 24; i++) {
//				oGrid.Columns.Item(i).TitleObject.Caption = COLNAM[i];

//				if (i >= 11) {
//					oGrid.Columns.Item(i).RightJustified = true;
//				} else {
//					oGrid.Columns.Item(i).RightJustified = false;
//				}

//				if (i > 0) {
//					oGrid.Columns.Item(i).Editable = false;
//				} else {
//					oGrid.Columns.Item(i).Editable = true;
//				}

//			}

//			//// Grid의 컬럼별 ComboBox, CheckBox 세팅
//			Grid_Col_Define();
//			oGrid.AutoResizeColumns();

//			///
//			oForm.Freeze(false);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			if (ErrNum == 1) {

//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("titleSetting 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//		}

//		private void Grid_Display()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			short iRow = 0;

//			string MSTCOD = null;
//			string FCLTCOD = null;
//			string TCLTCOD = null;

//			///  Default Value
//			ErrNum = 0;
//			iRow = 0;

//			mJSNYER = Strings.Trim(oForm.DataSources.UserDataSources.Item("JSNYER").ValueEx);
//			MSTCOD = Strings.Trim(oForm.DataSources.UserDataSources.Item("MSTCOD").ValueEx);

//			/// Check
//			//UPGRADE_WARNING: oForm.Items(TCLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items(FCLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case oForm.Items.Item("FCLTCOD").Specific.Selected == null:
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("TCLTCOD").Specific.Selected == null:
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(mJSNYER):
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//			}
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FCLTCOD = oForm.Items.Item("FCLTCOD").Specific.Selected.Value;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TCLTCOD = oForm.Items.Item("TCLTCOD").Specific.Selected.Value;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//			/// 조회
//			sQry = " Exec ZPY510 '" + Strings.Trim(mJSNYER) + "', '" + Strings.Trim(MSTCOD) + "', '" + Strings.Trim(FCLTCOD) + "', '" + Strings.Trim(TCLTCOD) + "'";
//			Debug.Print(sQry);
//			oDS_ZPY510.ExecuteQuery(sQry);
//			iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

//			if (iRow == 1) {
//				oRecordSet.DoQuery(sQry);
//				iRow = oRecordSet.RecordCount;
//			}

//			MDC_Globals.Sbo_Application.StatusBar.SetText(iRow + " 건이 있습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			titleSetting();

//			return;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사 코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Grid_Display Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{

//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			//// 사번 입력시 성명 조회
//			if (oUID == "MSTCOD") {
//				//UPGRADE_WARNING: oForm.Items(oUID).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.Value)) {
//					oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = "";
//				} else {
//					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", oForm.Items.Item(oUID).Specific.Value);
//				}

//				oForm.Update();
//			}
//			oForm.Freeze(false);
//		}

////*******************************************************************
////// ItemEventHander
////*******************************************************************
//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				//et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					if (pval.BeforeAction) {
//						//// 폴더 열기 버튼
//						if (pval.ItemUID == "Btn1") {
//							Grid_Display();
//							BubbleEvent = false;
//						//// File Upload 버튼
//						} else if (pval.ItemUID == "Btn2") {
//							Grid_Save();
//							BubbleEvent = false;
//						/// ChooseBtn사원리스트
//						} else if (pval.ItemUID == "CBtn1" & oForm.Items.Item("MSTCOD").Enabled == true) {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						/// 전체체크
//						} else if (pval.ItemUID == "Grid1" & pval.ColUID == "U_CHECK" & pval.Row == -1) {
//							CheckAll();
//						}
//					}
//					break;

//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "Grid1" & pval.ColUID == "U_MSTCOD" & pval.CharPressed == 9) {
//							//UPGRADE_WARNING: oDS_ZPY510.GetValue(U_MSTCOD, pval.Row) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oDS_ZPY510.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oDS_ZPY510.GetValue("U_MSTCOD", pval.Row) + "'") == true | string.IsNullOrEmpty(oDS_ZPY510.GetValue("U_MSTCOD", pval.Row))) {
//								oGrid.Columns.Item("U_MSTCOD").Click(pval.Row);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					}
//					break;

//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if (pval.ItemUID == "MSTCOD") {
//							FlushToItemValue(pval.ItemUID);
//						} else if (pval.ItemUID == "Grid1") {
//							FlushToItemValue(pval.ColUID, ref pval.Row);
//						}
//					}
//					break;

//				//et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					//----------------------------------------------------
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//----------------------------------------------------
//					if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oDS_ZPY510 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY510 = null;
//						//UPGRADE_NOTE: oGrid 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oGrid = null;
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//					}
//					break;
//			}

//			return;
//			Raise_FormItemEvent_Error:
//			////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Raise_FormItemEvent_Error:" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

////*******************************************************************
////// MenuEventHander
////*******************************************************************
//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{

//			if (pval.BeforeAction == true) {
//				return;
//			}

//			switch (pval.MenuUID) {
//				case "1287":
//					/// 복제
//					break;
//				case "1281":
//				case "1282":
//					break;
//				case "1288": // TODO: to "1291"
//					break;
//				case "1293":
//					break;
//			}
//			return;
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			int i = 0;
//			string sQry = null;
//			SAPbouiCOM.ComboBox oCombo = null;

//			SAPbobsCOM.Recordset oRecordSet = null;


//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			if ((BusinessObjectInfo.BeforeAction == false)) {
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
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Raise_FormDataEvent_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		private void CheckAll()
//		{
//			string CheckType = null;
//			int oRow = 0;

//			oForm.Freeze(true);
//			CheckType = "Y";
//			for (oRow = 0; oRow <= oGrid.Rows.Count - 1; oRow++) {
//				//UPGRADE_WARNING: oDS_ZPY510.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (Strings.Trim(oDS_ZPY510.GetValue("U_CHECK", oRow)) == "N") {
//					CheckType = "N";
//					break; // TODO: might not be correct. Was : Exit For
//				}
//			}

//			for (oRow = 0; oRow <= oGrid.Rows.Count - 1; oRow++) {
//				oDS_ZPY510.Rows.Offset = oRow;
//				if (CheckType == "N") {
//					oDS_ZPY510.SetValue("U_CHECK", oRow, "Y");
//				} else {
//					oDS_ZPY510.SetValue("U_CHECK", oRow, "N");
//				}
//			}
//			oForm.Freeze(false);

//		}

////---------------------------------------------------------------------------------------
//// Procedure : Grid_Col_Define
//// Author    : Choi Dong Kwon
//// Date      : 2008-07-14
//// Purpose   : Grid의 Column들에 대하여 LinkButton, ComboBox, CheckBox등을 정의
////---------------------------------------------------------------------------------------
////
//		private void Grid_Col_Define()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			short ErrNum = 0;

//			SAPbouiCOM.GridColumn oColumn = null;
//			SAPbouiCOM.EditTextColumn oEditCol = null;
//			SAPbouiCOM.ComboBoxColumn oComboCol = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oColumn = oGrid.Columns.Item("U_CHECK");
//			oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

//			//// 사원순번에 LinkButton 추가
//			oEditCol = oGrid.Columns.Item("EMPID");
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";

//			//// 자사코드
//			//// EditText Column => ComboBox Column으로 변경
//			oColumn = oGrid.Columns.Item("CLTCOD");
//			oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

//			oComboCol = oGrid.Columns.Item("CLTCOD");
//			oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oComboCol.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//				oRecordSet.MoveNext();
//			}

//			oColumn = oGrid.Columns.Item("JCLTCOD");
//			oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

//			oComboCol = oGrid.Columns.Item("JCLTCOD");
//			oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oComboCol.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//				oRecordSet.MoveNext();
//			}

//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oEditCol 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEditCol = null;
//			//UPGRADE_NOTE: oComboCol 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oComboCol = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:

//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oEditCol 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEditCol = null;
//			//UPGRADE_NOTE: oComboCol 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oComboCol = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {

//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Grid_Col_Define 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}

//		}

////---------------------------------------------------------------------------------------
//// Procedure : Grid_Save
//// Author    : Choi Dong Kwon
//// Date      : 2008-07-15
//// Purpose   : Grid의 내용을 일괄 저장하는 프로시저
////---------------------------------------------------------------------------------------
////
//		private void Grid_Save()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short ErrNum = 0;

//			int oRow = 0;
//			int UserId = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			UserId = MDC_Globals.oCompany.UserSignature;

//			ErrNum = 0;
//			for (oRow = 0; oRow <= oGrid.Rows.Count - 1; oRow++) {

//				//// 체크된 행만 저장
//				//UPGRADE_WARNING: oDS_ZPY510.GetValue(U_CHECK, oRow) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (oDS_ZPY510.GetValue("U_CHECK", oRow) == "Y") {

//					MDC_Globals.oCompany.StartTransaction();
//					/// 트랜잭션 시작

//					//UPGRADE_WARNING: oDS_ZPY510.GetValue(JCLTCOD, oRow) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oDS_ZPY510.GetValue(CLTCOD, oRow) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oDS_ZPY510.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = "EXEC ZPY510_1 '" + mJSNYER + "', " + "'" + oDS_ZPY510.GetValue("MSTCOD", oRow) + "', " + "'" + oDS_ZPY510.GetValue("CLTCOD", oRow) + "', " + "'" + oDS_ZPY510.GetValue("JCLTCOD", oRow) + "', " + Convert.ToString(UserId) + " ";
//					oRecordSet.DoQuery(sQry);
//					Debug.Print(sQry);
//					MDC_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
//					/// 트랜잭션 종료
//				}
//			}

//			Grid_Display();
//			MDC_Globals.Sbo_Application.StatusBar.SetText("종전근무지 일괄생성이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			//oForm.Mode = fm_OK_MODE
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:


//			MDC_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
//			/// 트랜잭션 RollBack

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Grid_Save 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

//		}
//	}
//}
