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
//	[System.Runtime.InteropServices.ProgId("RPY401_NET.RPY401")]
//	public class RPY401
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

//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString STRDAT = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString ENDDAT = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			string PRTDAT = null;
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString PRTGBN = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(1);
//			string JOBGBN = null;
//			string Branch = null;
//			string MSTDPT = null;
//			string MSTCOD = null;
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString STRJIG = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString ENDJIG = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);

//			/// Default
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			STRDAT.Value = Strings.Mid(oForm.Items.Item("STRDAT").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("STRDAT").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("STRDAT").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ENDDAT.Value = Strings.Mid(oForm.Items.Item("ENDDAT").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("ENDDAT").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("ENDDAT").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("PRTDAT").Specific.String))) {
//				//UPGRADE_WARNING: oForm.Items(PRTDAT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("PRTDAT").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd");
//			}
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			PRTDAT = oForm.Items.Item("PRTDAT").Specific.VALUE;
//			PRTDAT = Strings.Mid(PRTDAT, 1, 4) + "년  " + Strings.Mid(PRTDAT, 5, 2) + "월  " + Strings.Mid(PRTDAT, 7, 2) + "일";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JOBGBN = oForm.Items.Item("Combo03").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items(Combo01).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Branch = (oForm.Items.Item("Combo01").Specific.Selected.VALUE == "-1" ? "%" : oForm.Items.Item("Combo01").Specific.Selected.VALUE);
//			//UPGRADE_WARNING: oForm.Items(Combo02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTDPT = (oForm.Items.Item("Combo02").Specific.Selected.VALUE == "-1" ? "%" : oForm.Items.Item("Combo02").Specific.Selected.VALUE);
//			//UPGRADE_WARNING: oForm.Items.Item(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE) ? "%" : oForm.Items.Item("MSTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			STRJIG.Value = Strings.Mid(oForm.Items.Item("STRJIG").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("STRJIG").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("STRJIG").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ENDJIG.Value = Strings.Mid(oForm.Items.Item("ENDJIG").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("ENDJIG").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("ENDJIG").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));

//			/// Check
//			ErrNum = 0;
//			//UPGRADE_WARNING: oForm.Items(Combo03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case Information.IsDate(STRDAT.Value) == false:
//				case Information.IsDate(ENDDAT.Value) == false:
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case STRDAT.Value > ENDDAT.Value:
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("Combo03").Specific.Selected == null:
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//				case STRDAT.Value >= "2010-01-01":
//					//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//					oRecordSet = null;
//					return;

//					break;
//				case Information.IsDate(STRJIG.Value) == false:
//				case Information.IsDate(ENDJIG.Value) == false:
//					ErrNum = 5;
//					goto Error_Message;
//					break;
//				case STRJIG.Value > ENDJIG.Value:
//					ErrNum = 6;
//					goto Error_Message;
//					break;
//			}
//			PRTGBN.Value = oForm.DataSources.UserDataSources.Item("OptionDS").ValueEx;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//			WinTitle = "퇴직소득원천영수증";

//			/// Formula 수식필드***************************************************/
//			object[] ZRpt_Formula = new object[3];
//			object[] ZRpt_Formula_Value = new object[3];

//			//UPGRADE_WARNING: ZRpt_Formula(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula[1] = "PRTDAT";
//			//UPGRADE_WARNING: ZRpt_Formula_Value(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula_Value[1] = PRTDAT;
//			//UPGRADE_WARNING: ZRpt_Formula(2) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula[2] = "PRTGBN";
//			//UPGRADE_WARNING: ZRpt_Formula_Value(2) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula_Value[2] = PRTGBN.Value;

//			WinTitle = "[RPY401] : " + WinTitle;

//			ReportName = "RPY401.rpt";
//			/// SubReport /
//			object[] ZRpt_SRptSqry = new object[2];
//			object[] ZRpt_SRptName = new object[2];

//			//UPGRADE_WARNING: ZRpt_SRptSqry(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_SRptSqry[1] = "";
//			//UPGRADE_WARNING: ZRpt_SRptName(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_SRptName[1] = "";
//			/// ParameterFields
//			/// 조회조건문 /
//			sQry = "Exec RPY401 " + "'" + Strings.Trim(STRDAT.Value) + "', '" + Strings.Trim(ENDDAT.Value) + "', '" + Strings.Trim(JOBGBN) + "', '" + Strings.Trim(Branch) + "', '" + Strings.Trim(MSTDPT) + "', '" + Strings.Trim(MSTCOD) + "', '" + Strings.Trim(STRJIG.Value) + "', '" + Strings.Trim(ENDJIG.Value) + "'";

//			//    oRecordSet.DoQuery sQry
//			//    If oRecordSet.RecordCount = 0 Then
//			//        ErrNum = 4
//			//        GoTo error_Message
//			//    End If
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("대상 기간을 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("대상종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("출력구분을 선택 하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("해당하는 자료가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급 기간을 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Print_Query : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//		}
//		private void Print_Query2()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string WinTitle = null;
//			string ReportName = null;
//			short ErrNum = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString STRDAT = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString ENDDAT = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			string PRTDAT = null;
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString PRTGBN = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(1);
//			string JOBGBN = null;
//			string Branch = null;
//			string MSTDPT = null;
//			string MSTCOD = null;
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString STRJIG = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString ENDJIG = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);

//			/// Default
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			STRDAT.Value = Strings.Mid(oForm.Items.Item("STRDAT").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("STRDAT").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("STRDAT").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ENDDAT.Value = Strings.Mid(oForm.Items.Item("ENDDAT").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("ENDDAT").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("ENDDAT").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("PRTDAT").Specific.String))) {
//				//UPGRADE_WARNING: oForm.Items(PRTDAT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("PRTDAT").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd");
//			}
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			PRTDAT = oForm.Items.Item("PRTDAT").Specific.VALUE;
//			PRTDAT = Strings.Mid(PRTDAT, 1, 4) + "년  " + Strings.Mid(PRTDAT, 5, 2) + "월  " + Strings.Mid(PRTDAT, 7, 2) + "일";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JOBGBN = oForm.Items.Item("Combo03").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items(Combo01).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Branch = (oForm.Items.Item("Combo01").Specific.Selected.VALUE == "-1" ? "%" : oForm.Items.Item("Combo01").Specific.Selected.VALUE);
//			//UPGRADE_WARNING: oForm.Items(Combo02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTDPT = (oForm.Items.Item("Combo02").Specific.Selected.VALUE == "-1" ? "%" : oForm.Items.Item("Combo02").Specific.Selected.VALUE);
//			//UPGRADE_WARNING: oForm.Items.Item(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE) ? "%" : oForm.Items.Item("MSTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			STRJIG.Value = Strings.Mid(oForm.Items.Item("STRJIG").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("STRJIG").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("STRJIG").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ENDJIG.Value = Strings.Mid(oForm.Items.Item("ENDJIG").Specific.String, Convert.ToInt32("1"), Convert.ToInt32("4")) + "-" + Strings.Mid(oForm.Items.Item("ENDJIG").Specific.String, Convert.ToInt32("6"), Convert.ToInt32("2")) + "-" + Strings.Mid(oForm.Items.Item("ENDJIG").Specific.String, Convert.ToInt32("9"), Convert.ToInt32("2"));

//			/// Check
//			ErrNum = 0;
//			//UPGRADE_WARNING: oForm.Items(Combo03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case Information.IsDate(STRDAT.Value) == false:
//				case Information.IsDate(ENDDAT.Value) == false:
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case STRDAT.Value > ENDDAT.Value:
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("Combo03").Specific.Selected == null:
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//				case Information.IsDate(STRJIG.Value) == false:
//				case Information.IsDate(ENDJIG.Value) == false:
//					ErrNum = 5;
//					goto Error_Message;
//					break;
//				case STRJIG.Value > ENDJIG.Value:
//					ErrNum = 6;
//					goto Error_Message;
//					break;
//			}
//			PRTGBN.Value = oForm.DataSources.UserDataSources.Item("OptionDS").ValueEx;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//			WinTitle = "퇴직소득원천징수영수증";

//			/// Formula 수식필드***************************************************/
//			object[] ZRpt_Formula = new object[3];
//			object[] ZRpt_Formula_Value = new object[3];

//			//UPGRADE_WARNING: ZRpt_Formula(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula[1] = "PRTDAT";
//			//UPGRADE_WARNING: ZRpt_Formula_Value(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula_Value[1] = PRTDAT;
//			//UPGRADE_WARNING: ZRpt_Formula(2) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula[2] = "PRTGBN";
//			//UPGRADE_WARNING: ZRpt_Formula_Value(2) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_Formula_Value[2] = PRTGBN.Value;

//			WinTitle = "[RPY401_2] : " + WinTitle;
//			if (ENDDAT.Value <= "2009-12-31") {
//				ReportName = "RPY401_2.rpt";
//			} else if (ENDDAT.Value <= "2012-07-26") {
//				ReportName = "RPY401_2010.rpt";
//			} else {
//				ReportName = "RPY401_2012.rpt";
//			}
//			/// SubReport /
//			object[] ZRpt_SRptSqry = new object[2];
//			object[] ZRpt_SRptName = new object[2];

//			//UPGRADE_WARNING: ZRpt_SRptSqry(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_SRptSqry[1] = "";
//			//UPGRADE_WARNING: ZRpt_SRptName(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ZRpt_SRptName[1] = "";
//			/// ParameterFields
//			/// 조회조건문 /
//			sQry = "Exec RPY401_2 " + "'" + Strings.Trim(STRDAT.Value) + "', '" + Strings.Trim(ENDDAT.Value) + "', '" + Strings.Trim(JOBGBN) + "', '" + Strings.Trim(Branch) + "', '" + Strings.Trim(MSTDPT) + "', '" + Strings.Trim(MSTCOD) + "', '" + Strings.Trim(STRJIG.Value) + "', '" + Strings.Trim(ENDJIG.Value) + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet = null;
//				return;
//			}


//			/// Action /
//			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, Convert.ToString(2), "Y", "V", "") == false) {
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("대상 기간을 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("대상종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("출력구분을 선택 하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("해당하는 자료가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급 기간을 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//							//                        If oForm.Mode = fm_OK_MODE Or oForm.Mode = fm_UPDATE_MODE Then
//							Print_Query();
//							Print_Query2();
//							BubbleEvent = false;
//							//                        End If
//						} else if (pval.ItemUID == "CBtn1") {
//							if (oForm.Items.Item("MSTCOD").Enabled == true) {
//								oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUID == "STRDAT" | pval.ItemUID == "ENDDAT" | pval.ItemUID == "MSTCOD")) {
//						FlushToItemValue(pval.ItemUID);
//					}
//					break;
//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String))) {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (MDC_SetMod.Value_ChkYn("OHEM", "U_MSTCOD", "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'") == true) {
//								MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//								BubbleEvent = false;
//							} else {
//								//UPGRADE_WARNING: oForm.Items(MSTNAM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oForm.Items.Item("MSTNAM").Specific.VALUE = MDC_SetMod.Get_ReData(ref "LastName+FirstName", ref "U_MSTCOD", ref "OHEM", ref "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", ref "");
//							}
//						}
//					}
//					break;
//				//et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					break;

//				//et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oForm.UniqueID);
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


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\RPY401.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//------------------------------------------------------------------------
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//------------------------------------------------------------------------
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "RPY401_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			SubMain.AddForms(this, oFormUniqueID, "RPY401");
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

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.EditText oEdit = null;
//			string sQry = null;
//			SAPbouiCOM.OptionBtn oOption = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.DataSources.UserDataSources.Add("STRDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			/// 시작월
//			oForm.DataSources.UserDataSources.Add("ENDDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			/// 종료월
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			oForm.DataSources.UserDataSources.Add("PRTDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			oForm.DataSources.UserDataSources.Add("STRJIG", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			/// 지급시작일
//			oForm.DataSources.UserDataSources.Add("ENDJIG", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			/// 지급종료월

//			oEdit = oForm.Items.Item("STRDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "STRDAT");
//			oEdit = oForm.Items.Item("ENDDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "ENDDAT");
//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");
//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");
//			oEdit = oForm.Items.Item("PRTDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "PRTDAT");
//			oEdit = oForm.Items.Item("STRJIG").Specific;
//			oEdit.DataBind.SetBound(true, "", "STRJIG");
//			oEdit = oForm.Items.Item("ENDJIG").Specific;
//			oEdit.DataBind.SetBound(true, "", "ENDJIG");

//			//// Combo Box Setting
//			//// 지점
//			oCombo = oForm.Items.Item("Combo01").Specific;
//			oForm.Items.Item("Combo01").DisplayDesc = true;

//			sQry = "SELECT Code, Name FROM [@PH_PY005A] ";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount > 0) {
//				while (!(oRecordSet.EoF)) {
//					oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//					oRecordSet.MoveNext();
//				}
//			} else {
//				oCombo.ValidValues.Add("", "-");
//			}

//			MDC_SetMod.CLTCOD_Select(oForm, "Combo01");
//			//    sQry = "SELECT Code, Name FROM OUBR WHERE Code <> '-2' OR (Code = '-2' AND Name <> N'주요') ORDER BY Code ASC"
//			//    oRecordSet.DoQuery sQry
//			//    oCombo.ValidValues.Add "%", "모두"
//			//    Do Until oRecordSet.EOF
//			//        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
//			//        oRecordSet.MoveNext
//			//    Loop
//			//    If oCombo.ValidValues.Count > 0 Then
//			//       Call oCombo.Select(0, psk_Index)
//			//    End If
//			//// 부서
//			oCombo = oForm.Items.Item("Combo02").Specific;
//			oForm.Items.Item("Combo02").DisplayDesc = true;
//			//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("Combo01").Specific.VALUE + "'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");

//			//    sQry = " SELECT U_MSTDPT, Name FROM OUDP  WHERE ISNULL(U_MSTDPT, '') <> '' ORDER BY U_MSTDPT ASC "
//			//    oRecordSet.DoQuery sQry
//			//    oCombo.ValidValues.Add "%", "모두"
//			//    Do Until oRecordSet.EOF
//			//        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
//			//        oRecordSet.MoveNext
//			//    Loop
//			//    If oCombo.ValidValues.Count > 0 Then
//			//       Call oCombo.Select(0, psk_Index)
//			//    End If
//			//// 생성구분
//			oCombo = oForm.Items.Item("Combo03").Specific;
//			oForm.Items.Item("Combo03").DisplayDesc = true;
//			oCombo.ValidValues.Add("%", "모두");
//			oCombo.ValidValues.Add("1", "퇴직정산");
//			oCombo.ValidValues.Add("2", "중도정산");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			/// 전체


//			////옵션버튼(생성방법)
//			oForm.DataSources.UserDataSources.Add("OptionDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

//			oForm.Items.Item("Opt1").Visible = true;
//			oForm.Items.Item("Opt2").Visible = true;
//			oForm.Items.Item("Opt3").Visible = true;

//			oOption = oForm.Items.Item("Opt1").Specific;
//			oOption.DataBind.SetBound(true, "", "OptionDS");
//			oOption.ValOn = "1";
//			oOption.ValOff = "N";

//			oOption = oForm.Items.Item("Opt2").Specific;
//			oOption.DataBind.SetBound(true, "", "OptionDS");
//			oOption.GroupWith(("Opt1"));
//			if (oOption.ValOn != "2") {
//				oOption.ValOn = "2";
//			}
//			oOption.ValOff = "N";

//			oOption = oForm.Items.Item("Opt3").Specific;
//			oOption.DataBind.SetBound(true, "", "OptionDS");
//			oOption.GroupWith(("Opt1"));
//			if (oOption.ValOn != "3") {
//				oOption.ValOn = "3";
//			}
//			oOption.ValOff = "N";

//			oOption = oForm.Items.Item("Opt1").Specific;
//			oOption.Selected = true;

//			oForm.DataSources.UserDataSources.Item("STRDAT").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY") + "0101";
//			oForm.DataSources.UserDataSources.Item("ENDDAT").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");


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

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{

//			switch (oUID) {
//				case "STRDAT":
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					MDC_Globals.ZPAY_GBL_JSNYER.Value = Strings.Left(oForm.Items.Item(oUID).Specific.VALUE, 4);
//					oForm.DataSources.UserDataSources.Item("ENDDAT").ValueEx = MDC_Globals.ZPAY_GBL_JSNYER.Value + "1231";
//					oForm.Items.Item("ENDDAT").Update();
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = "";
//					} else {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = MDC_SetMod.Get_ReData(ref "LastName+FirstName", ref "U_MSTCOD", ref "OHEM", ref "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", ref "");
//					}
//					oForm.Items.Item("MSTNAM").Update();
//					break;
//			}
//			oForm.Items.Item(oUID).Update();
//		}
//	}
//}
