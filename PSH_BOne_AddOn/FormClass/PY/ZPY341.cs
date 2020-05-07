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
//	[System.Runtime.InteropServices.ProgId("ZPY341_NET.ZPY341")]
//	public class ZPY341
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY341.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 월별 정산자료 생성
//////  FormType       : 2010110341
//////  Create Date    : 2005.12.12
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;

//		private string oJsnYear;
//		private string oSMonth;
//		private string oEMonth;
//		private string oJsnGbn;

//		private struct TmpZPY341
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//			public char[] JOBDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//			public char[] CLTCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//			public char[] MSTCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//			public char[] MSTNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//			public char[] JIGBIL;
//			//TAMT(1 To 23)  As Double
//			[VBFixedArray(45)]
//				//2009연말정산항목추가(2009.12.31)
//			public double[] TAMT;

//			//UPGRADE_TODO: 해당 구조체의 인스턴스를 초기화하려면 "Initialize"를 호출해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
//			public void Initialize()
//			{
//				//UPGRADE_WARNING: TAMT 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				TAMT = new double[46];
//			}
//		}
////UPGRADE_WARNING: tmpArr 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//		TmpZPY341 tmpArr;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm(string oFromDocEntry01 = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY341.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY341_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY341");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//			oForm.Freeze(true);
//			CreateItems();
//			SetDocument(oFromDocEntry01);
//			oForm.Freeze(false);

//			oForm.EnableMenu(("1281"), false);
//			/// 찾기
//			oForm.EnableMenu(("1282"), true);
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
//						/// ChooseBtn사원리스트
//						if (pval.ItemUID == "CBtn1") {
//							oForm.Items.Item("MstCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						} else if (pval.ItemUID == "Opt1" | pval.ItemUID == "Opt2") {
//							if (oForm.DataSources.UserDataSources.Item("OptionDS").ValueEx == "1") {
//								oForm.Items.Item("MstCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								oForm.Items.Item("Path").Enabled = false;
//								oForm.Items.Item("Btn1").Enabled = false;
//							} else {
//								oForm.Items.Item("Path").Enabled = true;
//								oForm.Items.Item("Btn1").Enabled = true;
//							}
//						//// 경로선택
//						} else if (pval.ItemUID == "Btn1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							Excell_Upload();
//							BubbleEvent = false;
//						/// 월별자료생성실행
//						} else if (pval.ItemUID == "Execution") {
//							if (Execution_Process() == false) {
//								BubbleEvent = false;
//							} else {
//								oForm.ActiveItem = "JsnYear";
//								oForm.Items.Item("JsnGbn").Enabled = false;
//								oForm.Items.Item("STRDAT").Enabled = false;
//								oForm.Items.Item("ENDDAT").Enabled = false;
//								oForm.Items.Item("CLTCOD").Enabled = false;
//								//                        oForm.Items("BPLId").Enabled = False
//								oForm.Items.Item("DptStr").Enabled = false;
//								oForm.Items.Item("DptEnd").Enabled = false;
//								oForm.Items.Item("MstCode").Enabled = false;
//							}
//							if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//							}
//						} else if (pval.ItemUID == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							if (Execution_Save() == false) {
//								BubbleEvent = false;
//								return;
//							} else {
//								oForm.ActiveItem = "JsnYear";
//								oForm.Items.Item("JsnGbn").Enabled = true;
//								oForm.Items.Item("STRDAT").Enabled = true;
//								oForm.Items.Item("ENDDAT").Enabled = true;
//								oForm.Items.Item("CLTCOD").Enabled = true;
//								//                        oForm.Items("BPLId").Enabled = True
//								oForm.Items.Item("DptStr").Enabled = true;
//								oForm.Items.Item("DptEnd").Enabled = true;
//								oForm.Items.Item("MstCode").Enabled = true;
//							}
//						}
//					} else {
//						if (pval.ItemUID == "1" & pval.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//							FormItemEnabled();
//						}
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUID == "JsnYear" | pval.ItemUID == "SMonth" | pval.ItemUID == "EMonth" | pval.ItemUID == "MstCode")) {
//						FlushToItemValue(pval.ItemUID);
//					}
//					break;
//				//et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					if (pval.BeforeAction == true & pval.ItemUID != "1000001" & pval.ItemUID != "2") {
//						///정산년도
//						if (Last_Item == "JsnYear") {
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(Last_Item).Specific.VALUE))) {
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (MDC_SetMod.ChkYearMonth(ref Strings.Trim(Convert.ToString(oForm.Items.Item(Last_Item).Specific.VALUE)) + "01") == false) {
//									oForm.Items.Item(Last_Item).Update();
//									MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//									BubbleEvent = false;
//								}
//							}
//						} else if (Last_Item == "SMonth" | Last_Item == "EMonth") {
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(Last_Item).Specific.VALUE))) {
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (MDC_SetMod.ChkYearMonth(ref oJsnYear + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(Last_Item).Specific.VALUE, "00")) == false) {
//									oForm.Items.Item(Last_Item).Update();
//									MDC_Globals.Sbo_Application.StatusBar.SetText("생성기간을 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//									BubbleEvent = false;
//								}
//							}
//						} else if (Last_Item == "MstCode") {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(Last_Item).Specific.String)) & MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + Strings.Trim(oForm.Items.Item(Last_Item).Specific.String) + "'", ref "") == true) {
//								oForm.Items.Item(Last_Item).Update();
//								MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//								BubbleEvent = false;
//							}
//						}
//					}
//					break;
//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					if (pval.BeforeAction == true & pval.ItemUID == "JsnYear" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (Strings.Len(Strings.Trim(oForm.Items.Item(pval.ItemUID).Specific.String)) < 4) {
//							//UPGRADE_WARNING: oForm.Items(pval.ItemUID).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item(pval.ItemUID).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(pval.ItemUID).Specific.VALUE, "2000");
//						}
//						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (MDC_SetMod.ChkYearMonth(ref Strings.Trim(Convert.ToString(oForm.Items.Item(pval.ItemUID).Specific.VALUE)) + "01") == false) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					} else if (pval.BeforeAction == true & (pval.ItemUID == "SMonth" | pval.ItemUID == "EMonth") & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (MDC_SetMod.ChkYearMonth(ref oJsnYear + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(pval.ItemUID).Specific.VALUE, "00")) == false) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("생성기간을 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					} else if (pval.BeforeAction == true & pval.ItemUID == "MstCode" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MstCode").Specific.String)) & MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + Strings.Trim(oForm.Items.Item("MstCode").Specific.String) + "'", ref "") == true) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					}
//					break;
//				//et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if (pval.ItemUID == "JsnGbn") {
//							FlushToItemValue(pval.ItemUID);
//						}
//						if (pval.ItemUID == "CLTCOD") {
//							////기본사항 - 부서1 (사업장에 따른 부서변경)
//							oCombo = oForm.Items.Item("DptStr").Specific;

//							if (oCombo.ValidValues.Count > 0) {
//								for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//									oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//								}
//								oCombo.ValidValues.Add("", "");
//								oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//							}

//							sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//							//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//							sQry = sQry + " ORDER BY U_Code";
//							MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//							oCombo.ValidValues.Add("%", "전체");
//							oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

//							////기본사항 - 부서2 (사업장에 따른 부서변경)
//							oCombo = oForm.Items.Item("DptEnd").Specific;

//							if (oCombo.ValidValues.Count > 0) {
//								for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//									oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//								}
//								oCombo.ValidValues.Add("", "");
//								oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//							}

//							sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//							//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//							sQry = sQry + " ORDER BY U_Code";
//							MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//							oCombo.ValidValues.Add("%", "전체");
//							oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//						}
//					}
//					break;
//				//et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					if (Last_Item == "Mat1") {
//						if (pval.Row > 0) {
//							Last_Item = pval.ItemUID;
//						}
//					} else {
//						Last_Item = pval.ItemUID;
//					}
//					break;
//				//et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					if (pval.BeforeAction == true) {
//						TmpTable_Delete();
//					}
//					if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//					}
//					break;
//			}

//			return;
//			Raise_FormItemEvent_Error:
//			///////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
//					FormItemEnabled();
//					oForm.Items.Item("JsnYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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

//		private void SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				FormItemEnabled();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY001_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo1 = null;
//			SAPbouiCOM.ComboBox oCombo2 = null;
//			SAPbouiCOM.OptionBtn oOption = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.Column oColumn = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.DataSources.UserDataSources.Add("JsnYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			/// 생성년도
//			oForm.DataSources.UserDataSources.Add("JsnGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 생성구분
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 지점
//			oForm.DataSources.UserDataSources.Add("DptStr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 부서코드
//			oForm.DataSources.UserDataSources.Add("DptEnd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oForm.DataSources.UserDataSources.Add("SMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
//			/// 시작월
//			oForm.DataSources.UserDataSources.Add("EMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
//			/// 종료월
//			oForm.DataSources.UserDataSources.Add("MstCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			oForm.DataSources.UserDataSources.Add("MstName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			oForm.DataSources.UserDataSources.Add("EmpID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oForm.DataSources.UserDataSources.Add("Path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
//			oForm.DataSources.UserDataSources.Add("STRDAT", SAPbouiCOM.BoDataType.dt_DATE);
//			oForm.DataSources.UserDataSources.Add("ENDDAT", SAPbouiCOM.BoDataType.dt_DATE);

//			oEdit = oForm.Items.Item("JsnYear").Specific;
//			oEdit.DataBind.SetBound(true, "", "JsnYear");

//			oEdit = oForm.Items.Item("SMonth").Specific;
//			oEdit.DataBind.SetBound(true, "", "SMonth");
//			oEdit = oForm.Items.Item("EMonth").Specific;
//			oEdit.DataBind.SetBound(true, "", "EMonth");
//			oEdit = oForm.Items.Item("MstCode").Specific;
//			oEdit.DataBind.SetBound(true, "", "MstCode");
//			oEdit = oForm.Items.Item("MstName").Specific;
//			oEdit.DataBind.SetBound(true, "", "MstName");
//			oEdit = oForm.Items.Item("EmpID").Specific;
//			oEdit.DataBind.SetBound(true, "", "EmpID");
//			oEdit = oForm.Items.Item("Path").Specific;
//			oEdit.DataBind.SetBound(true, "", "Path");
//			oEdit = oForm.Items.Item("STRDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "STRDAT");
//			oEdit = oForm.Items.Item("ENDDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "ENDDAT");

//			//// 생성구분
//			oCombo1 = oForm.Items.Item("JsnGbn").Specific;
//			oCombo1.DataBind.SetBound(true, "", "JsnGbn");
//			oCombo1.ValidValues.Add("1", "연말정산(재직자)");
//			oCombo1.ValidValues.Add("2", "중도정산(퇴직자)");
//			oCombo1.ValidValues.Add("3", "전체");

//			oForm.Items.Item("JsnGbn").DisplayDesc = true;

//			//// 사업장
//			oCombo1 = oForm.Items.Item("CLTCOD").Specific;
//			oCombo1.DataBind.SetBound(true, "", "CLTCOD");
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    oRecordSet.DoQuery sQry
//			//    Do Until oRecordSet.EOF
//			//        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
//			//        oRecordSet.MoveNext
//			//    Loop
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			//    '// 지점
//			//    Set oCombo1 = oForm.Items("BPLId").Specific
//			//    sQry = "SELECT Code, Name FROM OUBR WHERE Code <> '-2' OR (Code = '-2' AND Name <> N'주요') ORDER BY Code ASC"
//			//    oRecordSet.DoQuery sQry
//			//    oCombo1.ValidValues.Add "%", "모두"
//			//    Do Until oRecordSet.EOF
//			//        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
//			//        oRecordSet.MoveNext
//			//    Loop
//			//    If oCombo1.ValidValues.Count > 0 Then
//			//        oCombo1.Select 0, psk_Index
//			//    End If

//			//// 부서
//			oCombo1 = oForm.Items.Item("DptStr").Specific;
//			oCombo1.DataBind.SetBound(true, "", "DptStr");
//			oForm.Items.Item("DptStr").DisplayDesc = true;
//			//// 부서
//			oCombo1 = oForm.Items.Item("DptEnd").Specific;
//			oCombo1.DataBind.SetBound(true, "", "DptEnd");
//			oForm.Items.Item("DptEnd").DisplayDesc = true;


//			////옵션버튼(생성방법)
//			oForm.DataSources.UserDataSources.Add("OptionDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			oForm.Items.Item("Opt1").Visible = true;
//			oForm.Items.Item("Opt2").Visible = true;
//			oOption = oForm.Items.Item("Opt1").Specific;
//			oOption.DataBind.SetBound(true, "", "OptionDS");

//			oOption = oForm.Items.Item("Opt2").Specific;
//			oOption.GroupWith(("Opt1"));
//			oOption.DataBind.SetBound(true, "", "OptionDS");

//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Opt1").Specific.Selected = true;

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			oForm.DataSources.UserDataSources.Add("Col1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
//			oColumn = oMat1.Columns.Item("Col1");
//			oColumn.DataBind.SetBound(true, "", "Col1");

//			//UPGRADE_WARNING: oForm.Items(JsnYear).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("JsnYear").Specific.String = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY");
//			oSMonth = "01";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("SMonth").Specific.VALUE = oSMonth;
//			oEMonth = "12";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("EMonth").Specific.VALUE = oEMonth;

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oOption 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oOption = null;
//			//UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo1 = null;
//			//UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo2 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oOption 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oOption = null;
//			//UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo1 = null;
//			//UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo2 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private object Execution_Process()
//		{
//			object functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short ErrNum = 0;
//			int MSTCNT = 0;

//			string STRDPT = null;
//			string ENDDPT = null;
//			string BPLID = null;
//			string CLTCOD = null;
//			string STRDAT = null;
//			string ENDDAT = null;
//			ErrNum = 0;

//			/// 필수Check
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			/// 정산년도
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JsnYear").Specific.String))) {
//				ErrNum = 1;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (MDC_SetMod.ChkYearMonth(ref oJsnYear + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("SMonth").Specific.VALUE, "00")) == false | MDC_SetMod.ChkYearMonth(ref oJsnYear + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("EMonth").Specific.VALUE, "00")) == false) {
//				ErrNum = 2;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items(EMonth).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: oForm.Items(SMonth).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("SMonth").Specific.VALUE > oForm.Items.Item("EMonth").Specific.VALUE) {
//				ErrNum = 3;
//				goto Error_Message;
//			}
//			oMat1.Clear();
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			/// 임시테이블생성
//			sQry = "Exec  ZPY341_1 ";
//			oRecordSet.DoQuery(sQry);
//			/// Matrix Message
//			oForm.DataSources.UserDataSources.Item("Col1").Value = "임시 테이블 생성 완료!";
//			oMat1.AddRow();

//			/// SAP급상여생성 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//			if (oForm.DataSources.UserDataSources.Item("OptionDS").ValueEx == "1") {
//				//         If oForm.Items("BPLId").Specific.Selected Is Nothing Then
//				//             ErrNum = 4
//				//             GoTo Error_Message
//				//UPGRADE_WARNING: oForm.Items(DptEnd).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: oForm.Items(DptStr).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (oForm.Items.Item("DptStr").Specific.Selected == null | oForm.Items.Item("DptEnd").Specific.Selected == null) {
//					ErrNum = 5;
//					goto Error_Message;
//					//UPGRADE_WARNING: oForm.Items(JsnGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				} else if (oForm.Items.Item("JsnGbn").Specific.Selected == null) {
//					ErrNum = 7;
//					goto Error_Message;
//					//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				} else if (oForm.Items.Item("CLTCOD").Specific.Selected == null) {
//					ErrNum = 8;
//					goto Error_Message;
//				}
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oJsnYear = oForm.Items.Item("JsnYear").Specific.String;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oSMonth = oForm.Items.Item("SMonth").Specific.String;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oEMonth = oForm.Items.Item("EMonth").Specific.String;
//				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oJsnGbn = oForm.Items.Item("JsnGbn").Specific.Selected.VALUE;
//				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				STRDPT = oForm.Items.Item("DptStr").Specific.Selected.VALUE;
//				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ENDDPT = oForm.Items.Item("DptEnd").Specific.Selected.VALUE;
//				BPLID = "%";
//				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.VALUE;
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				STRDAT = oForm.Items.Item("STRDAT").Specific.VALUE;
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ENDDAT = oForm.Items.Item("ENDDAT").Specific.VALUE;
//				if (string.IsNullOrEmpty(Strings.Trim(STRDAT)) | string.IsNullOrEmpty(Strings.Trim(ENDDAT))) {
//					ErrNum = 9;
//					goto Error_Message;
//				}
//				if (Strings.Trim(ENDDPT) == "%")
//					ENDDPT = "ZZZZZZZZ";
//				STRDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(STRDAT, "0000-00-00");
//				ENDDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(ENDDAT, "0000-00-00");
//				/// 생성프로시저 실행(년도, 시작월, 종료월, 시작부서, 종료부서, 정산구분, 사원번호
//				MSTCNT = 0;
//				sQry = "Exec ZPY341_3 " + "'" + Strings.Trim(oJsnYear) + "', '" + Strings.Trim(oSMonth) + "','" + Strings.Trim(oEMonth) + "', '" + Strings.Trim(STRDPT) + " ', '" + Strings.Trim(ENDDPT) + "','" + Strings.Trim(oJsnGbn) + "', '" + Strings.Trim(STRDAT) + "',  '" + Strings.Trim(ENDDAT) + "','" + Strings.Trim(CLTCOD) + "', '" + Strings.Trim(BPLID) + "'";
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MstCode").Specific.String))) {
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = sQry + ", '" + oForm.Items.Item("MstCode").Specific.String + "'";
//				} else {
//					sQry = sQry + ", '%'";
//				}
//				oRecordSet.DoQuery(sQry);
//				/// Matrix Message
//				oForm.DataSources.UserDataSources.Item("Col1").Value = "자료 검색중...";
//				oMat1.AddRow();

//				if (oRecordSet.RecordCount > 0) {
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					MSTCNT = oRecordSet.Fields.Item(0).Value;
//				}
//				oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCNT + "명의 월별 자료 집계 완료.";
//				oMat1.AddRow();

//				MDC_Globals.Sbo_Application.StatusBar.SetText("월별 자료 집계 작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//				/// Excel 파일upload /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//			} else if (oForm.DataSources.UserDataSources.Item("OptionDS").ValueEx == "2") {
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Path").Specific.String))) {
//					ErrNum = 6;
//					goto Error_Message;
//				}
//				Execution_Excel();
//			}
//			///
//			oForm.DataSources.UserDataSources.Item("Col1").Value = "생성된 내용은 [추가]하셔야 시스템에 적용됩니다.";
//			oMat1.AddRow();


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			//UPGRADE_WARNING: Execution_Process 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("생성기간을 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("시작월보다 종료월이 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지점을 선택하세요. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("생성할 부서범위를 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Upload할 파일을 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("생성구분을 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 8) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드를 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 9) {
//				//UPGRADE_WARNING: Sbo_Application.S9tStatusBarMessage 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.Sbo_Application.S9tStatusBarMessage("선택 기준일은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Execution_Process 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			//UPGRADE_WARNING: Execution_Process 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = false;
//			return functionReturnValue;
//		}
//		private object Execution_Save()
//		{
//			object functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short ErrNum = 0;
//			string JSNYER = null;
//			string SMonth = null;
//			string EMonth = null;
//			///
//			ErrNum = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "SELECT TOP 1 * FROM [TmpZPY341]";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			}
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JSNYER = oForm.Items.Item("JsnYear").Specific.String;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			SMonth = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("SMonth").Specific.VALUE, "00");
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			EMonth = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("EMonth").Specific.VALUE, "00");
//			/// 임시테이블저장
//			//Exec dbo.ZPY341_4  '2005'
//			sQry = "Exec ZPY341_4 " + "'" + Strings.Trim(JSNYER) + "', '" + Strings.Trim(SMonth) + "', '" + Strings.Trim(EMonth) + "', " + MDC_Globals.oCompany.UserSignature;
//			oRecordSet.DoQuery(sQry);

//			/// 메세지
//			oForm.DataSources.UserDataSources.Item("Col1").Value = "생성된 자료가 시스템에 적용되었습니다!";

//			oMat1.AddRow();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			//UPGRADE_WARNING: Execution_Save 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("추가할 자료가 없습니다. 월자료 생성을 먼저 하십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Execution_Save 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			//UPGRADE_WARNING: Execution_Save 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			ZPAY_g_EmpID MstInfo = default(ZPAY_g_EmpID);

//			switch (oUID) {
//				case "JsnYear":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oJsnYear = oForm.Items.Item(oUID).Specific.String;
//					break;
//				case "SMonth":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oSMonth = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(oUID).Specific.String, "00");
//					oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = oSMonth;
//					break;
//				case "EMonth":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oEMonth = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(oUID).Specific.String, "00");
//					oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = oEMonth;
//					break;
//				case "MstCode":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oForm.DataSources.UserDataSources.Item(oUID).ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("MstName").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = "";
//					} else {
//						oForm.DataSources.UserDataSources.Item(oUID).ValueEx = Strings.UCase(oForm.DataSources.UserDataSources.Item(oUID).ValueEx);
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: MstInfo 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MstInfo = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oForm.DataSources.UserDataSources.Item("MstName").ValueEx = MstInfo.MSTNAM;
//						oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = MstInfo.EmpID;
//					}
//					oForm.Items.Item("MstName").Update();
//					oForm.Items.Item("EmpID").Update();
//					break;
//				case "JsnGbn":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if ((oForm.Items.Item(oUID).Specific.Selected != null)) {
//						//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oJsnGbn = oForm.Items.Item(oUID).Specific.Selected.VALUE;
//					} else {
//						oJsnGbn = "";
//					}
//					oForm.Freeze(true);
//					if (string.IsNullOrEmpty(oJsnYear)) {
//						oJsnYear = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY");
//						//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.Items.Item("JsnYear").Specific.VALUE = oJsnYear;
//					}
//					switch (Strings.Trim(oJsnGbn)) {
//						case "2":
//							/// 중도정산
//							//UPGRADE_WARNING: oForm.Items(STRDAT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("STRDAT").Specific.VALUE = oJsnYear + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "MM" + "01");
//							//UPGRADE_WARNING: oForm.Items(ENDDAT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("ENDDAT").Specific.VALUE = oJsnYear + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "MM") + MDC_SetMod.Month_LastDay(ref oJsnYear + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "MM"));
//							oForm.Items.Item("s10").Visible = true;
//							oForm.Items.Item("s11").Visible = true;
//							oForm.Items.Item("STRDAT").Visible = true;
//							oForm.Items.Item("ENDDAT").Visible = true;
//							break;

//						case "1":
//							/// 연말정산
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("STRDAT").Specific.VALUE = oJsnYear + "1231";
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("ENDDAT").Specific.VALUE = oJsnYear + "1231";
//							oForm.Items.Item("EMonth").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							oForm.Items.Item("s10").Visible = true;
//							oForm.Items.Item("s11").Visible = false;
//							oForm.Items.Item("STRDAT").Visible = true;
//							oForm.Items.Item("ENDDAT").Visible = false;
//							break;
//						case "3":
//							/// 전체
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("STRDAT").Specific.VALUE = oJsnYear + "0101";
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("ENDDAT").Specific.VALUE = oJsnYear + "1231";
//							oForm.Items.Item("EMonth").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							oForm.Items.Item("s10").Visible = false;
//							oForm.Items.Item("s11").Visible = false;
//							oForm.Items.Item("STRDAT").Visible = false;
//							oForm.Items.Item("ENDDAT").Visible = false;
//							break;
//					}
//					oForm.Freeze(false);
//					break;
//			}
//			oForm.Items.Item(oUID).Update();
//			return;
//			Error_Message:
//			MDC_Globals.Sbo_Application.StatusBar.SetText("FlushToItemValue Error : " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private void FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.OptionBtn optBtn = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				oForm.Items.Item("Btn1").Enabled = true;
//				oForm.Items.Item("Execution").Enabled = true;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");


//				////부서1
//				oCombo = oForm.Items.Item("DptStr").Specific;
//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}

//				if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx)) {
//					sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//					//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//					sQry = sQry + " ORDER BY U_Code";
//					MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//					oCombo.ValidValues.Add("%", "전체");
//					oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//				}

//				////부서2
//				oCombo = oForm.Items.Item("DptEnd").Specific;
//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}


//				if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx)) {
//					sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//					//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//					sQry = sQry + " ORDER BY U_Code";
//					MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//					oCombo.ValidValues.Add("%", "전체");
//					oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//				}

//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.Items.Item("Btn1").Enabled = false;
//				oForm.Items.Item("Execution").Enabled = false;
//			}
//		}
//		private void Excell_Upload()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string File_Name = null;

//			//UPGRADE_WARNING: ZP_Form.OpenDialog() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			File_Name = My.MyProject.Forms.ZP_Form.OpenDialog(ref ZP_Form, ref "*.xls;*.xlsx;*.csv", ref "파일선택", ref "C:\\");
//			if (!string.IsNullOrEmpty(File_Name)) {
//				oForm.DataSources.UserDataSources.Item("Path").ValueEx = File_Name;
//			} else {
//				oForm.DataSources.UserDataSources.Item("Path").ValueEx = "";
//			}
//			oForm.Items.Item("Path").Update();
//			return;
//			Error_Message:
//			//////////////////////////////////////////////////////////////////////
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Excell_Upload Error:" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private void Execution_Excel()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short ErrNum = 0;

//			//    Dim x               As Integer
//			short y = 0;
//			object xlApp = null;
//			Microsoft.Office.Interop.Excel.Worksheet xlSheet1 = default(Microsoft.Office.Interop.Excel.Worksheet);
//			string MSTCOD = null;
//			short oRow = 0;

//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = oForm.Items.Item("MstCode").Specific.String;
//			ErrNum = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// Open Work Database
//			xlApp = Interaction.CreateObject("Excel.Application");
//			//UPGRADE_WARNING: xlApp.Workbooks 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			xlApp.Workbooks.Open(oForm.DataSources.UserDataSources.Item("Path").ValueEx);

//			//UPGRADE_WARNING: xlApp.Worksheets 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			xlSheet1 = xlApp.Worksheets(1);
//			//UPGRADE_WARNING: xlApp.DisplayAlerts 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			xlApp.DisplayAlerts = false;
//			//// 컬럼타이틀값이 현재 패치버전과 맞는지 확인
//			//UPGRADE_WARNING: xlSheet1.Cells(2, 42).Text 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: xlSheet1.Cells(2, 50).Text 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if ((string.IsNullOrEmpty(xlSheet1.Cells._Default(2, 50).Text)) | (xlSheet1.Cells._Default(2, 42).Text != "월지급총액")) {
//				ErrNum = 1;
//				goto Error_Message;
//				//UPGRADE_WARNING: xlSheet1.Cells(3, 2).Text 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlSheet1.Cells(2, 2).Text 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (xlSheet1.Cells._Default(2, 2).Text == "자사코드" & string.IsNullOrEmpty(xlSheet1.Cells._Default(3, 2).Text)) {
//				ErrNum = 2;
//				goto Error_Message;
//			}
//			y = 3;
//			oRow = 0;
//			while (1) {
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.JOBDAT = xlSheet1.Cells._Default(y, 1).VALUE;
//				/// 귀속월
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.CLTCOD = xlSheet1.Cells._Default(y, 2).VALUE;
//				/// 자사코드
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.MSTCOD = xlSheet1.Cells._Default(y, 3).VALUE;
//				/// 사원번호
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.MSTNAM = xlSheet1.Cells._Default(y, 4).VALUE;
//				/// 사원성명
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.JIGBIL = xlSheet1.Cells._Default(y, 5).VALUE;
//				/// 지급일자
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[1] = Conversion.Val(xlSheet1.Cells._Default(y, 6).VALUE);
//				/// 과세급여
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[2] = Conversion.Val(xlSheet1.Cells._Default(y, 7).VALUE);
//				/// 과세상여
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[3] = Conversion.Val(xlSheet1.Cells._Default(y, 8).VALUE);
//				/// 인정상여
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[4] = Conversion.Val(xlSheet1.Cells._Default(y, 9).VALUE);
//				/// 주식매수행사이익
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[5] = Conversion.Val(xlSheet1.Cells._Default(y, 10).VALUE);
//				/// 우리사주조합
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[6] = Conversion.Val(xlSheet1.Cells._Default(y, 11).VALUE);
//				/// 과세총계
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[7] = Conversion.Val(xlSheet1.Cells._Default(y, 12).VALUE);
//				/// 비과세-G01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[8] = Conversion.Val(xlSheet1.Cells._Default(y, 13).VALUE);
//				/// 비과세-H01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[9] = Conversion.Val(xlSheet1.Cells._Default(y, 14).VALUE);
//				/// 비과세-H05
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[10] = Conversion.Val(xlSheet1.Cells._Default(y, 15).VALUE);
//				/// 비과세-H06
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[11] = Conversion.Val(xlSheet1.Cells._Default(y, 16).VALUE);
//				/// 비과세-H07
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[12] = Conversion.Val(xlSheet1.Cells._Default(y, 17).VALUE);
//				/// 비과세-H08
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[13] = Conversion.Val(xlSheet1.Cells._Default(y, 18).VALUE);
//				/// 비과세-H09
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[14] = Conversion.Val(xlSheet1.Cells._Default(y, 19).VALUE);
//				/// 비과세-H10
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[15] = Conversion.Val(xlSheet1.Cells._Default(y, 20).VALUE);
//				/// 비과세-H11
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[16] = Conversion.Val(xlSheet1.Cells._Default(y, 21).VALUE);
//				/// 비과세-H12
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[17] = Conversion.Val(xlSheet1.Cells._Default(y, 22).VALUE);
//				/// 비과세-H13
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[18] = Conversion.Val(xlSheet1.Cells._Default(y, 23).VALUE);
//				/// 비과세-I01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[19] = Conversion.Val(xlSheet1.Cells._Default(y, 24).VALUE);
//				/// 비과세-K01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[20] = Conversion.Val(xlSheet1.Cells._Default(y, 25).VALUE);
//				/// 비과세-M01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[21] = Conversion.Val(xlSheet1.Cells._Default(y, 26).VALUE);
//				/// 비과세-M02
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[22] = Conversion.Val(xlSheet1.Cells._Default(y, 27).VALUE);
//				/// 비과세-M03
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[23] = Conversion.Val(xlSheet1.Cells._Default(y, 28).VALUE);
//				/// 비과세-O01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[24] = Conversion.Val(xlSheet1.Cells._Default(y, 29).VALUE);
//				/// 비과세-Q01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[25] = Conversion.Val(xlSheet1.Cells._Default(y, 30).VALUE);
//				/// 비과세-S01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[26] = Conversion.Val(xlSheet1.Cells._Default(y, 31).VALUE);
//				/// 비과세-T01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[27] = Conversion.Val(xlSheet1.Cells._Default(y, 32).VALUE);
//				/// 비과세-X01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[28] = Conversion.Val(xlSheet1.Cells._Default(y, 33).VALUE);
//				/// 비과세-Y01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[29] = Conversion.Val(xlSheet1.Cells._Default(y, 34).VALUE);
//				/// 비과세-Y02
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[30] = Conversion.Val(xlSheet1.Cells._Default(y, 35).VALUE);
//				/// 비과세-Y03
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[31] = Conversion.Val(xlSheet1.Cells._Default(y, 36).VALUE);
//				/// 비과세-Y20
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[32] = Conversion.Val(xlSheet1.Cells._Default(y, 37).VALUE);
//				/// 비과세-Y20
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[33] = Conversion.Val(xlSheet1.Cells._Default(y, 38).VALUE);
//				/// 비과세-Z01
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[34] = Conversion.Val(xlSheet1.Cells._Default(y, 39).VALUE);
//				/// 비과세-차량,식대보조        BIGWA02
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[35] = Conversion.Val(xlSheet1.Cells._Default(y, 40).VALUE);
//				/// 비과세-기타(지급조서포함)   BIGWA04
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[36] = Conversion.Val(xlSheet1.Cells._Default(y, 41).VALUE);
//				/// 비과세-기타(지급조서미포함) BIGWA07
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[37] = Conversion.Val(xlSheet1.Cells._Default(y, 42).VALUE);
//				/// 월지급총액
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[38] = Conversion.Val(xlSheet1.Cells._Default(y, 43).VALUE);
//				/// 국민연금
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[39] = Conversion.Val(xlSheet1.Cells._Default(y, 44).VALUE);
//				/// 건강보험
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[40] = Conversion.Val(xlSheet1.Cells._Default(y, 45).VALUE);
//				/// 고용보험
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[41] = Conversion.Val(xlSheet1.Cells._Default(y, 46).VALUE);
//				/// 소득세
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[42] = Conversion.Val(xlSheet1.Cells._Default(y, 47).VALUE);
//				/// 주민세
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[43] = Conversion.Val(xlSheet1.Cells._Default(y, 48).VALUE);
//				/// 농특세
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[44] = Conversion.Val(xlSheet1.Cells._Default(y, 49).VALUE);
//				/// 기부금
//				//UPGRADE_WARNING: xlSheet1.Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				tmpArr.TAMT[45] = Conversion.Val(xlSheet1.Cells._Default(y, 50).VALUE);
//				/// 장기요양보험

//				if (string.IsNullOrEmpty(Strings.Trim(tmpArr.JOBDAT)) | string.IsNullOrEmpty(Strings.Trim(tmpArr.MSTCOD)) | string.IsNullOrEmpty(Strings.Trim(tmpArr.CLTCOD))) {
//					break; // TODO: might not be correct. Was : Exit Do
//					ErrNum = 3;
//					goto Error_Message;
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)) | (!string.IsNullOrEmpty(Strings.Trim(MSTCOD)) & Strings.Trim(tmpArr.MSTCOD) == Strings.Trim(MSTCOD))) {
//					/// 프로시저실행
//					sQry = "Exec ZPY341_2 '" + tmpArr.JOBDAT + "','" + Strings.Trim(tmpArr.CLTCOD) + "','" + Strings.Trim(tmpArr.MSTCOD) + "','" + tmpArr.JIGBIL + "'," + tmpArr.TAMT[1] + "," + tmpArr.TAMT[2] + "," + tmpArr.TAMT[3] + "," + tmpArr.TAMT[4] + "," + tmpArr.TAMT[5] + "," + tmpArr.TAMT[6] + "," + tmpArr.TAMT[7] + "," + tmpArr.TAMT[8] + "," + tmpArr.TAMT[9] + "," + tmpArr.TAMT[10] + "," + tmpArr.TAMT[11] + "," + tmpArr.TAMT[12] + "," + tmpArr.TAMT[13] + "," + tmpArr.TAMT[14] + "," + tmpArr.TAMT[15] + "," + tmpArr.TAMT[16] + "," + tmpArr.TAMT[17] + "," + tmpArr.TAMT[18] + "," + tmpArr.TAMT[19] + "," + tmpArr.TAMT[20] + "," + tmpArr.TAMT[21] + "," + tmpArr.TAMT[22] + "," + tmpArr.TAMT[23] + "," + tmpArr.TAMT[24] + "," + tmpArr.TAMT[25] + "," + tmpArr.TAMT[26] + "," + tmpArr.TAMT[27] + "," + tmpArr.TAMT[28] + "," + tmpArr.TAMT[29] + "," + tmpArr.TAMT[30] + "," + tmpArr.TAMT[31] + "," + tmpArr.TAMT[32] + "," + tmpArr.TAMT[33] + "," + tmpArr.TAMT[34] + "," + tmpArr.TAMT[35] + "," + tmpArr.TAMT[36] + "," + tmpArr.TAMT[37] + "," + tmpArr.TAMT[38] + "," + tmpArr.TAMT[39] + "," + tmpArr.TAMT[40] + "," + tmpArr.TAMT[41] + "," + tmpArr.TAMT[42] + "," + tmpArr.TAMT[43] + "," + tmpArr.TAMT[44] + "," + tmpArr.TAMT[45];

//					oRecordSet.DoQuery(sQry);
//					if (oRecordSet.RecordCount == 0) {
//						oForm.DataSources.UserDataSources.Item("Col1").Value = "행: (" + y + ") 사원번호:" + tmpArr.MSTCOD + "일치하는 사원이 없습니다.";
//						oMat1.AddRow();
//					}
//					oRow = oRow + 1;
//				}
//				y = y + 1;
//			}
//			oForm.Update();
//			oForm.DataSources.UserDataSources.Item("Col1").Value = "엑셀 upload Line(" + oRow + "/ " + y - 3 + ")작업을 완료하였습니다.";
//			oMat1.AddRow();

//			MDC_Globals.Sbo_Application.StatusBar.SetText("작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			//UPGRADE_WARNING: xlApp.Quit 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			xlApp.Quit();
//			//UPGRADE_NOTE: xlSheet1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlSheet1 = null;
//			//UPGRADE_NOTE: xlApp 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlApp = null;
//			return;
//			Error_Message:
//			/// Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//			 // ERROR: Not supported in C#: OnErrorStatement

//			//UPGRADE_ISSUE: vbNormal을(를) 어떤 상수로 업그레이드할지 결정할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
//			//UPGRADE_ISSUE: Screen 속성 Screen.MousePointer은(는) 사용자 지정 마우스 포인터를 지원하지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
//			//UPGRADE_WARNING: Screen 속성 Screen.MousePointer에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
//			System.Windows.Forms.Cursor.Current = Constants.vbNormal;

//			//UPGRADE_WARNING: xlApp.DisplayAlerts 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			xlApp.DisplayAlerts = false;
//			//UPGRADE_WARNING: xlApp.Quit 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			xlApp.Quit();
//			//UPGRADE_NOTE: xlSheet1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlSheet1 = null;
//			//UPGRADE_NOTE: xlApp 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlApp = null;
//			if (ErrNum == Convert.ToDouble("1")) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("엑셀서식이 패치버전과 일치하지 않습니다. 관리자에게 문의하십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == Convert.ToDouble("2")) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == Convert.ToDouble("3")) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText(y + "번째 행의 자사코드,사원번호, 귀속연월는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("엑셀 작업에 실패 하였습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//		}
//		private void TmpTable_Delete()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			/// 임시테이블생성
//			sQry = "Exec  ZPY341_1 ";
//			oRecordSet.DoQuery(sQry);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("TmpTable_Delete 실행중 오류. 프로시저가 있는지 확인하세요." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
//	}
//}
