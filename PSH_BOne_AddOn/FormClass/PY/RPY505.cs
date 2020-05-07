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
//	[System.Runtime.InteropServices.ProgId("RPY505_NET.RPY505")]
//	public class RPY505
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : RPY505.cls
//////  Module         : 인사관리>정산관리>정산관련리포트
//////  Desc           : 소득자료제출집계표
//////  FormType       : 2010130505
//////  Create Date    : 2006.01.23
//////  Modified Date  : 2006.12.23
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

//			string JSNFYMM = null;
//			string JSNTYMM = null;
//			string SINFYMM = null;
//			string SINTYMM = null;
//			string CLTCOD = null;
//			string MSTDPT = null;
//			string MSTCOD = null;
//			string PRTDAT = null;
//			string JOBGBN = null;
//			string PRTTYP = null;

//			/// ODBC 연결 체크
//			if (ConnectODBC() == false) {
//				goto Error_Message;
//			}

//			/// Default
//			var _with1 = oForm.DataSources.UserDataSources;
//			JSNFYMM = _with1.Item("JSNFYMM").ValueEx;
//			JSNTYMM = _with1.Item("JSNTYMM").ValueEx;
//			SINFYMM = _with1.Item("SINFYMM").ValueEx;
//			SINTYMM = _with1.Item("SINTYMM").ValueEx;
//			MSTCOD = _with1.Item("MSTCOD").ValueEx;
//			PRTDAT = _with1.Item("PRTDAT").ValueEx;
//			if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
//				MSTCOD = "%";
//			if (string.IsNullOrEmpty(Strings.Trim(PRTDAT))) {
//				PRTDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd");
//				oForm.DataSources.UserDataSources.Item("PRTDAT").ValueEx = PRTDAT;
//			}

//			if (MDC_SetMod.ChkYearMonth(ref JSNFYMM) == false)
//				JSNFYMM = "";
//			if (MDC_SetMod.ChkYearMonth(ref JSNTYMM) == false)
//				JSNTYMM = "";
//			if (MDC_SetMod.ChkYearMonth(ref SINFYMM) == false)
//				SINFYMM = "";
//			if (MDC_SetMod.ChkYearMonth(ref SINTYMM) == false)
//				SINTYMM = "";

//			/// Check
//			ErrNum = 0;
//			//UPGRADE_WARNING: oForm.Items(Combo01).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items(Combo03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case (string.IsNullOrEmpty(JSNFYMM) | string.IsNullOrEmpty(JSNTYMM)) & (string.IsNullOrEmpty(SINFYMM) | string.IsNullOrEmpty(SINTYMM)):
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("Combo03").Specific.Selected == null:
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("Combo01").Specific.Selected == null:
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//			}
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = oForm.Items.Item("Combo01").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTDPT = oForm.Items.Item("Combo02").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JOBGBN = oForm.Items.Item("Combo03").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			PRTTYP = oForm.Items.Item("Combo04").Specific.Selected.VALUE;
//			/// 소득 구분

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//			WinTitle = "소득자료제출집계표";
//			ReportName = "RPY505.rpt";
//			MDC_Globals.gRpt_Formula = new string[3];
//			MDC_Globals.gRpt_Formula_Value = new string[3];

//			/// Formula 수식필드***************************************************/

//			//    gRpt_Formula(1) = "JSNYER":    gRpt_Formula_Value(1) = JSNYER
//			MDC_Globals.gRpt_Formula[1] = "PRTDAT";
//			MDC_Globals.gRpt_Formula_Value[1] = PRTDAT;

//			if (PRTTYP == "1") {
//				MDC_Globals.gRpt_Formula[2] = "PRTNAM";
//				MDC_Globals.gRpt_Formula_Value[2] = "근로소득 (갑종)";
//			} else if (PRTTYP == "2") {
//				MDC_Globals.gRpt_Formula[2] = "PRTNAM";
//				MDC_Globals.gRpt_Formula_Value[2] = "퇴직소득";
//			} else if (PRTTYP == "3") {
//				MDC_Globals.gRpt_Formula[2] = "PRTNAM";
//				MDC_Globals.gRpt_Formula_Value[2] = "사업소득";
//			} else if (PRTTYP == "4") {
//				MDC_Globals.gRpt_Formula[2] = "PRTNAM";
//				MDC_Globals.gRpt_Formula_Value[2] = "기타소득 (거주자)";
//			} else if (PRTTYP == "5") {
//				MDC_Globals.gRpt_Formula[2] = "PRTNAM";
//				MDC_Globals.gRpt_Formula_Value[2] = "사업·기타소득 (비거주자)";
//			} else if (PRTTYP == "6") {
//				MDC_Globals.gRpt_Formula[2] = "PRTNAM";
//				MDC_Globals.gRpt_Formula_Value[2] = "이자소득";
//			} else if (PRTTYP == "7") {
//				MDC_Globals.gRpt_Formula[2] = "PRTNAM";
//				MDC_Globals.gRpt_Formula_Value[2] = "배당소득";
//			}

//			WinTitle = "[RPY505] : " + WinTitle;
//			MDC_Globals.gRpt_SRptSqry = new string[2];
//			MDC_Globals.gRpt_SRptName = new string[2];
//			MDC_Globals.gRpt_SFormula = new string[2, 2];
//			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];
//			/// SubReport /

//			MDC_Globals.gRpt_SRptSqry[1] = "";
//			MDC_Globals.gRpt_SRptName[1] = "";

//			/// 조회조건문 /
//			sQry = "Exec RPY505 " + "'" + Strings.Trim(JSNFYMM) + "', '" + Strings.Trim(JSNTYMM) + "', " + "'" + Strings.Trim(SINFYMM) + "', '" + Strings.Trim(SINTYMM) + "', " + "'" + Strings.Trim(JOBGBN) + "', '" + Strings.Trim(CLTCOD) + "', '" + Strings.Trim(MSTDPT) + "', " + "'" + Strings.Trim(MSTCOD) + "', '" + Strings.Trim(PRTTYP) + "'";
//			/// Action /
//			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, Convert.ToString(1), "Y", "V", "") == false) {
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 연월, 원천신고 연월이 모두 올바르지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("출력 구분을 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자료 제출자를 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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

//				//et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
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

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\RPY505.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "RPY505_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "RPY505");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

//			////////////////////////////////////////////////////////////////////////////////
//			//***************************************************************
//			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//			// oForm.DataBrowser.BrowseBy = "DocNum"
//			//***************************************************************
//			////////////////////////////////////////////////////////////////////////////////
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

//			var _with2 = oForm.DataSources.UserDataSources;
//			_with2.Add("JSNFYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			/// 귀속년월(From)
//			_with2.Add("JSNTYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			/// 귀속년월(To)
//			_with2.Add("SINFYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			/// 신고년월(From)
//			_with2.Add("SINTYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			/// 신고년월(To)
//			_with2.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			_with2.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			_with2.Add("PRTDAT", SAPbouiCOM.BoDataType.dt_DATE);

//			oEdit = oForm.Items.Item("JSNFYMM").Specific;
//			oEdit.DataBind.SetBound(true, "", "JSNFYMM");
//			oEdit = oForm.Items.Item("JSNTYMM").Specific;
//			oEdit.DataBind.SetBound(true, "", "JSNTYMM");
//			oEdit = oForm.Items.Item("SINFYMM").Specific;
//			oEdit.DataBind.SetBound(true, "", "SINFYMM");
//			oEdit = oForm.Items.Item("SINTYMM").Specific;
//			oEdit.DataBind.SetBound(true, "", "SINTYMM");
//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");
//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");
//			oEdit = oForm.Items.Item("PRTDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "PRTDAT");

//			//// Combo Box Setting
//			//// 사업장
//			oCombo = oForm.Items.Item("Combo01").Specific;
//			oForm.Items.Item("Combo01").DisplayDesc = true;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}
//			if (oCombo.ValidValues.Count > 0) {
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			}

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
//			//// 소득종류
//			oCombo = oForm.Items.Item("Combo04").Specific;
//			oForm.Items.Item("Combo04").DisplayDesc = true;
//			oCombo.ValidValues.Add("1", "근로소득(갑종)");
//			oCombo.ValidValues.Add("2", "퇴직소득");
//			oCombo.ValidValues.Add("3", "사업소득(거주자)");
//			oCombo.ValidValues.Add("4", "기타소득(거주자)");
//			oCombo.ValidValues.Add("5", "사업.기타소득(비거주자)");
//			oCombo.ValidValues.Add("6", "이자소득");
//			oCombo.ValidValues.Add("7", "배당소득");

//			if (oCombo.ValidValues.Count > 0) {
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			}
//			///
//			oForm.DataSources.UserDataSources.Item("JSNFYMM").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateAndTime.Now), "YYYY") + "01";
//			oForm.DataSources.UserDataSources.Item("JSNTYMM").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateAndTime.Now), "YYYY") + "12";
//			oForm.DataSources.UserDataSources.Item("SINFYMM").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateAndTime.Now), "YYYY") + "01";
//			oForm.DataSources.UserDataSources.Item("SINTYMM").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");


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
