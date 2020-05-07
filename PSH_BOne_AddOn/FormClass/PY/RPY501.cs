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
//	[System.Runtime.InteropServices.ProgId("RPY501_NET.RPY501")]
//	public class RPY501
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : RPY501.cls
//////  Module         : 인사관리>정산관리>정산관련리포트
//////  Desc           : 월별 자료 현황
//////  FormType       : 2010130501
//////  Create Date    : 2006.01.10
//////  Modified Date  : 2006.12.10
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		private void Print_Query()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string WinTitle = null;
//			string ReportName = null;
//			short ErrNum = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			string JSNYER = null;
//			string STRMON = null;
//			string ENDMON = null;
//			string JOBGBN = null;
//			string CLTCOD = null;
//			string Branch = null;
//			string MSTDPT = null;
//			string MSTCOD = null;
//			string PRTGBN = null;

//			/// ODBC 연결 체크
//			if (ConnectODBC() == false) {
//				goto Error_Message;
//			}

//			/// Default
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JSNYER = oForm.Items.Item("JsnYear").Specific.String;
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			STRMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("SMonth").Specific.String, "00");
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ENDMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("EMonth").Specific.String, "00");
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = oForm.Items.Item("MSTCOD").Specific.String;
//			if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
//				MSTCOD = "%";
//			/// Check
//			ErrNum = 0;
//			//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items(Combo03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case MDC_SetMod.ChkYearMonth(ref JSNYER + STRMON) == false:
//				case MDC_SetMod.ChkYearMonth(ref JSNYER + ENDMON) == false:
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("Combo03").Specific.Selected == null:
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("CLTCOD").Specific.Selected == null:
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//			}
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTDPT = oForm.Items.Item("Combo02").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JOBGBN = oForm.Items.Item("Combo03").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			PRTGBN = oForm.Items.Item("PRTGBN").Specific.Selected.VALUE;

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//			if (PRTGBN == "1") {
//				WinTitle = "월별 자료 현황(집계표)";
//				ReportName = "RPY501.rpt";
//			} else {
//				WinTitle = "월별 자료 현황(집계표)";
//				ReportName = "RPY501_1.rpt";
//			}
//			MDC_Globals.gRpt_Formula = new string[3];
//			MDC_Globals.gRpt_Formula_Value = new string[3];

//			/// Formula 수식필드***************************************************/

//			MDC_Globals.gRpt_Formula[1] = "CLTNAM";
//			MDC_Globals.gRpt_Formula_Value[1] = MDC_Globals.oCompany.CompanyName;
//			MDC_Globals.gRpt_Formula[2] = "PRTLMT";
//			MDC_Globals.gRpt_Formula_Value[2] = Strings.Mid(JSNYER, 1, 4) + "년 ";

//			WinTitle = "[RPY501] : " + WinTitle;
//			MDC_Globals.gRpt_SRptSqry = new string[2];
//			MDC_Globals.gRpt_SRptName = new string[2];
//			/// SubReport /

//			MDC_Globals.gRpt_SRptSqry[1] = "";
//			MDC_Globals.gRpt_SRptName[1] = "";

//			/// 조회조건문 /
//			if (PRTGBN == "1") {
//				sQry = "Exec RPY501 " + "'" + Strings.Trim(JSNYER) + "', '" + Strings.Trim(STRMON) + "', " + "'" + Strings.Trim(ENDMON) + "', '" + Strings.Trim(JOBGBN) + "', " + "'" + Strings.Trim(CLTCOD) + "', " + "'" + Strings.Trim(MSTDPT) + "', '" + Strings.Trim(MSTCOD) + "'";
//			} else {
//				sQry = "Exec RPY501_1 " + "'" + Strings.Trim(JSNYER) + "', '" + Strings.Trim(STRMON) + "', " + "'" + Strings.Trim(ENDMON) + "', '" + Strings.Trim(JOBGBN) + "', " + "'" + Strings.Trim(CLTCOD) + "', " + "'" + Strings.Trim(MSTDPT) + "', '" + Strings.Trim(MSTCOD) + "'";
//			}

//			/// Action /
//			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, Convert.ToString(1), "Y", "V", "", 2) == false) {
//				//  SBO_Application.StatusBar.SetText "gCryReport_Action : 실패!", bmt_Short, smt_Error
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			/// Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기준 연월을 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("출력 구분을 선택 하세요..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사 코드를 선택 하세요..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Print_Query : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
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
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								Print_Query();
//								BubbleEvent = false;
//							}
//						} else if (pval.ItemUID == "CBtn1") {
//							if (oForm.Items.Item("MSTCOD").Enabled == true) {
//								oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					}
//					break;
//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					if (pval.BeforeAction == true & pval.ItemUID == "JsnYear" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (Strings.Len(Strings.Trim(oForm.Items.Item("JsnYear").Specific.String)) == 0) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						} else if (Strings.Len(Strings.Trim(oForm.Items.Item(pval.ItemUID).Specific.String)) < 4) {
//							//UPGRADE_WARNING: oForm.Items(pval.ItemUID).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item(pval.ItemUID).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(pval.ItemUID).Specific.VALUE, "2000");
//						}
//					} else if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String))) {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'") == true) {
//								MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//								BubbleEvent = false;
//							} else {
//								//UPGRADE_WARNING: oForm.Items(MSTNAM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oForm.Items.Item("MSTNAM").Specific.VALUE = MDC_SetMod.Get_ReData(ref "U_FullName", ref "Code", ref "[@PH_PY001A]", ref "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", ref "");
//							}
//						}
//					}
//					break;
//				//et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					break;

//				//et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					//----------------------------------------------------
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//----------------------------------------------------
//					if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//					}
//					break;
//			}

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Raise_FormItemEvent_Error:", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm()
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\RPY501.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//------------------------------------------------------------------------
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//------------------------------------------------------------------------
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "RPY501_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			SubMain.AddForms(this, oFormUniqueID, "RPY501");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

//			oForm.Freeze(true);
//			CreateItems();
//			oForm.Freeze(false);

//			oForm.EnableMenu(("1281"), true);
//			/// 찾기
//			oForm.EnableMenu(("1282"), false);
//			/// 추가
//			oForm.EnableMenu(("1284"), false);
//			/// 취소
//			oForm.EnableMenu(("1293"), false);
//			/// 행삭제

//			oForm.Update();
//			oForm.Visible = true;

//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			if ((oForm == null) == false) {
//				oForm.Freeze(false);
//				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oForm = null;
//			}
//		}

////*******************************************************************
////
////*******************************************************************
//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.EditText oEdit = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			var _with1 = oForm.DataSources.UserDataSources;
//			_with1.Add("JsnYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			/// 생성년도
//			_with1.Add("SMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
//			/// 시작월
//			_with1.Add("EMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
//			/// 종료월
//			_with1.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			_with1.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);

//			oEdit = oForm.Items.Item("JsnYear").Specific;
//			oEdit.DataBind.SetBound(true, "", "JsnYear");
//			oEdit = oForm.Items.Item("SMonth").Specific;
//			oEdit.DataBind.SetBound(true, "", "SMonth");
//			oEdit = oForm.Items.Item("EMonth").Specific;
//			oEdit.DataBind.SetBound(true, "", "EMonth");
//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");
//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");

//			//// 자사코드
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			oCombo.ValidValues.Add("%", "모두");
//			while (!(oRecordSet.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}
//			if (oCombo.ValidValues.Count > 0) {
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			}

//			//    '// 지점
//			//    Set oCombo = oForm.Items("Combo01").Specific
//			//    oForm.Items("Combo01").DisplayDesc = True
//			//    sQry = "SELECT Code, Name FROM OUBR WHERE Code <> '-2' OR (Code = '-2' AND Name <> N'주요') ORDER BY Code ASC"
//			//    oRecordSet.DoQuery sQry
//			//    oCombo.ValidValues.Add "%", "모두"
//			//    Do Until oRecordSet.EOF
//			//        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
//			//        oRecordSet.MoveNext
//			//    Loop
//			//    If oCombo.ValidValues.Count > 0 Then
//			//       Call oCombo.Select(0, psk_Index)
//			//    End If

//			//// 부서
//			oCombo = oForm.Items.Item("Combo02").Specific;
//			oForm.Items.Item("Combo02").DisplayDesc = true;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			oCombo.ValidValues.Add("%", "모두");
//			while (!(oRecordSet.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}
//			if (oCombo.ValidValues.Count > 0) {
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			}

//			//// 생성구분
//			oCombo = oForm.Items.Item("Combo03").Specific;
//			oForm.Items.Item("Combo03").DisplayDesc = true;
//			oCombo.ValidValues.Add("1", "연말정산(재직자)");
//			oCombo.ValidValues.Add("2", "중도정산(퇴직자)");
//			oCombo.ValidValues.Add("3", "전체");
//			oCombo.Select("3", SAPbouiCOM.BoSearchKey.psk_ByValue);


//			//// 출력구분
//			oCombo = oForm.Items.Item("PRTGBN").Specific;
//			oForm.Items.Item("PRTGBN").DisplayDesc = true;
//			oCombo.ValidValues.Add("1", "집계표");
//			oCombo.ValidValues.Add("2", "명세서");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			//// Initial Value
//			oForm.DataSources.UserDataSources.Item("JsnYear").ValueEx = Convert.ToString(DateAndTime.Year(DateAndTime.Now));
//			oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = "01";
//			oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = "12";

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
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
//	}
//}
