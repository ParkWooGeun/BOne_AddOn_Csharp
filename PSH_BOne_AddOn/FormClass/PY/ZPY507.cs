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
//	internal class ZPY507
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY507.cls
//////  Module         : 인사관리>원천징수>근로소득
//////  Desc           : 정산 결과 조회(전체)
//////  FormType       : 2010110507
//////  Create Date    : 2009.12.13
//////  Modified Date  :
//////  Creator        : Choi Dong Kwon
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************
//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//		private SAPbouiCOM.Grid oGrid1;
//		private SAPbouiCOM.DataTable oDS_ZPY507;

//		private void titleSetting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			int iCol = 0;

//			string[] COLNAM = new string[151];
//			SAPbouiCOM.EditTextColumn oEditCol = null;
//			SAPbouiCOM.ComboBoxColumn oComboCol = null;
//			SAPbouiCOM.GridColumn oColumn = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			/// Initial
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			COLNAM[0] = "정산년도";
//			COLNAM[1] = "정산구분";
//			COLNAM[2] = "신고연월";
//			COLNAM[3] = "지급일자";
//			COLNAM[4] = "사업장";
//			COLNAM[5] = "사원번호";
//			COLNAM[6] = "사원순번";
//			COLNAM[7] = "사원명";
//			COLNAM[8] = "월별자료No";
//			COLNAM[9] = "소득항목No";
//			COLNAM[10] = "종전문서No";
//			COLNAM[11] = "세액계산No";
//			COLNAM[12] = "의료비No";
//			COLNAM[13] = "기부금No";
//			COLNAM[14] = "연금.저축No";
//			COLNAM[15] = "현근무지총계";
//			COLNAM[16] = "전근무지총계";
//			COLNAM[17] = "비과세계";
//			COLNAM[18] = "총급여";
//			COLNAM[19] = "근로소득공제";
//			COLNAM[20] = "근로소득금액";
//			COLNAM[21] = "본인공제금액";
//			COLNAM[22] = "배우자유무";
//			COLNAM[23] = "배우자공제액";
//			COLNAM[24] = "부양가족수";
//			COLNAM[25] = "부양가족공제";
//			COLNAM[26] = "경로우대인원";
//			COLNAM[27] = "경로우대공제";
//			COLNAM[28] = "장애인인원";
//			COLNAM[29] = "장애인공제액";
//			COLNAM[30] = "부녀자유무";
//			COLNAM[31] = "부녀자공제액";
//			COLNAM[32] = "자녀양육인원";
//			COLNAM[33] = "자녀양육공제";
//			COLNAM[34] = "출산입양인원";
//			COLNAM[35] = "출산입양공제";
//			COLNAM[36] = "다자녀인원";
//			COLNAM[37] = "다자녀공제";
//			COLNAM[38] = "국민연금";
//			COLNAM[39] = "기타연금(공무원연금)";
//			COLNAM[40] = "기타연금(군인연금)";
//			COLNAM[41] = "기타연금(사립학교교직원연금)";
//			COLNAM[42] = "기타연금(별정우체국연금)";
//			COLNAM[43] = "퇴직연금(근로자퇴직급여보장법)";
//			COLNAM[44] = "퇴직연금(과학기술인공제)";
//			COLNAM[45] = "보험료(건강보험)";
//			COLNAM[46] = "보험료(고용보험)";
//			COLNAM[47] = "보험료(보장성보험)";
//			COLNAM[48] = "보험료(장애인전용)";
//			COLNAM[49] = "의료비공제금액";
//			COLNAM[50] = "교육비공제금액";
//			COLNAM[51] = "주택임차차입금원리금상환-대출기관";
//			COLNAM[52] = "주택임차차입금원리금상환-거주자";
//			COLNAM[53] = "월세액";
//			COLNAM[54] = "장기주택이자상환액-15년미만";
//			COLNAM[55] = "장기주택이자상환액-29년이하";
//			COLNAM[56] = "장기주택이자상환액-30년이상";
//			COLNAM[57] = "기부금공제금액";
//			COLNAM[58] = "혼인,이사,장례비";
//			COLNAM[59] = "특별공제계";
//			COLNAM[60] = "표준공제";
//			COLNAM[61] = "차감소득금액";
//			COLNAM[62] = "개인연금저축공제";
//			COLNAM[63] = "연금저축소득공제";
//			COLNAM[64] = "소기업공제부금소득공제";
//			COLNAM[65] = "주택마련저축(청약저축)";
//			COLNAM[66] = "주택마련저축(주택청약종합저축)";
//			COLNAM[67] = "주택마련저축(장기주택마련저축)";
//			COLNAM[68] = "주택마련저축(근로자주택마련저축)";
//			COLNAM[69] = "투자조합출자공제";
//			COLNAM[70] = "신용카드소득공제";
//			COLNAM[71] = "우리사주조합공제";
//			COLNAM[72] = "장기주식형저축소득공제";
//			COLNAM[73] = "고용유지중소기업공제";
//			COLNAM[74] = "기타소득공제계";
//			COLNAM[75] = "종합소득과세표준";
//			COLNAM[76] = "산출세액";
//			COLNAM[77] = "소득법";
//			COLNAM[78] = "조특법";
//			COLNAM[79] = "조세조약";
//			COLNAM[80] = "감면세액계";
//			COLNAM[81] = "근로소득세액공제";
//			COLNAM[82] = "납세조합공제";
//			COLNAM[83] = "주택차입금";
//			COLNAM[84] = "기부정처자금";
//			COLNAM[85] = "외국납부";
//			COLNAM[86] = "세액공제계";
//			COLNAM[87] = "결정소득세";
//			COLNAM[88] = "결정주민세";
//			COLNAM[89] = "결정농특세";
//			COLNAM[90] = "종(전)근무지_소득세";
//			COLNAM[91] = "종(전)근무지_주민세";
//			COLNAM[92] = "종(전)근무지_농특세";
//			COLNAM[93] = "주(현)근무지_소득세";
//			COLNAM[94] = "주(현)근무지_주민세";
//			COLNAM[95] = "주(현)근무지_농특세";
//			COLNAM[96] = "차감소득세";
//			COLNAM[97] = "차감주민세";
//			COLNAM[98] = "차감농특세";

//			for (iCol = 0; iCol <= 98; iCol++) {
//				oGrid1.Columns.Item(iCol).Editable = false;
//				oGrid1.Columns.Item(iCol).TitleObject.Caption = COLNAM[iCol];
//				if (iCol >= 8) {
//					oGrid1.Columns.Item(iCol).RightJustified = true;
//				}
//				//2007B PL18 이상 일때(2007A 버전은 확인 필요)
//				if (MDC_Globals.oCompany.Version >= Convert.ToDouble("860040")) {
//					oGrid1.Columns.Item(iCol).TitleObject.Sortable = true;
//				}


//			}

//			//// Link Button
//			oEditCol = oGrid1.Columns.Item("EMPID");
//			//// 사원순번
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";

//			oEditCol = oGrid1.Columns.Item("DOCNO1");
//			//// 월별자료No
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";
//			oEditCol = oGrid1.Columns.Item("DOCNO2");
//			//// 소득항목No
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";
//			oEditCol = oGrid1.Columns.Item("DOCNO3");
//			//// 종전문서No
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";
//			oEditCol = oGrid1.Columns.Item("DOCNO4");
//			//// 세액계산No
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";
//			oEditCol = oGrid1.Columns.Item("DOCNO5");
//			//// 의료비No
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";
//			oEditCol = oGrid1.Columns.Item("DOCNO6");
//			//// 기부금No
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";
//			oEditCol = oGrid1.Columns.Item("DOCNO7");
//			//// 연금.저축No
//			oEditCol.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
//			oEditCol.LinkedObjectType = "171";

//			//// ComboBox
//			oColumn = oGrid1.Columns.Item("JSNGBN");
//			//// 정산구분
//			oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

//			oComboCol = oGrid1.Columns.Item("JSNGBN");
//			oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
//			oComboCol.ValidValues.Add("1", "연말정산(재직자)");
//			oComboCol.ValidValues.Add("2", "중도정산(퇴직자)");


//			oColumn = oGrid1.Columns.Item("CLTCOD");
//			//// 사업장
//			oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

//			oComboCol = oGrid1.Columns.Item("CLTCOD");
//			oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				while (!(oRecordSet.EoF)) {
//					oComboCol.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//					oRecordSet.MoveNext();
//				}
//			}

//			oGrid1.AutoResizeColumns();
//			oForm.Freeze(false);

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
//			MDC_Globals.Sbo_Application.StatusBar.SetText("titleSetting 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{

//			switch (oUID) {
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = "";
//					} else {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = MDC_SetMod.Get_ReData(ref "U_FullName", ref "Code", ref "[@PH_PY001A]", ref "'" + oForm.Items.Item(oUID).Specific.String + "'", ref "");
//					}
//					break;
//			}
//			oForm.Update();

//		}
////*******************************************************************
////// ItemEventHander
////*******************************************************************
//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


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
//						//// 찾기 버튼
//						if (pval.ItemUID == "Btn01") {
//							Grid_Display();
//							BubbleEvent = false;
//						/// ChooseBtn 사원리스트
//						} else if (pval.ItemUID == "CBtn01" & oForm.Items.Item("MSTCOD").Enabled == true) {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						}
//					}
//					break;

//				//et_MATRIX_LINK_PRESSED'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					if (pval.BeforeAction) {
//						if (pval.ItemUID == "Grid1" & Strings.Left(pval.ColUID, 5) == "DOCNO") {
//							UserFormLink(ref (pval.ColUID), ref (pval.Row));
//							BubbleEvent = false;
//						}
//					}
//					break;

//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if (pval.ItemUID == "MSTCOD") {
//							FlushToItemValue(pval.ItemUID);
//						}
//					}
//					break;

//				//et_FORM_UNLOAD'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oDS_ZPY507 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY507 = null;
//						//UPGRADE_NOTE: oGrid1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oGrid1 = null;
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

//		private void UserFormLink(ref string LinkName, ref int LineNum)
//		{
//			object oTmpObject = null;
//			string DocNum = null;
//			string JSNYER = null;
//			string CLTCOD = null;
//			string MSTCOD = null;

//			if (!string.IsNullOrEmpty(Strings.Trim(LinkName))) {
//				switch (LinkName) {
//					case "DOCNO1":
//						oTmpObject = new ZPY343();
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						DocNum = oDS_ZPY507.GetValue("DOCNO1", LineNum);
//						if (!string.IsNullOrEmpty(Strings.Trim(DocNum))) {
//							//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oTmpObject.LoadForm(DocNum);
//							MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//						}
//						break;
//					case "DOCNO2":
//						oTmpObject = new ZPY501();
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						JSNYER = Strings.Trim(oDS_ZPY507.GetValue("JSNYMM", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						CLTCOD = Strings.Trim(oDS_ZPY507.GetValue("CLTCOD", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MSTCOD = Strings.Trim(oDS_ZPY507.GetValue("MSTCOD", LineNum));
//						if (!string.IsNullOrEmpty(Strings.Trim(JSNYER)) & !string.IsNullOrEmpty(Strings.Trim(CLTCOD)) & !string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
//							//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD);
//							MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//						}
//						break;
//					case "DOCNO3":
//						oTmpObject = new ZPY502();
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						JSNYER = Strings.Trim(oDS_ZPY507.GetValue("JSNYMM", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						CLTCOD = Strings.Trim(oDS_ZPY507.GetValue("CLTCOD", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MSTCOD = Strings.Trim(oDS_ZPY507.GetValue("MSTCOD", LineNum));
//						if (!string.IsNullOrEmpty(Strings.Trim(JSNYER)) & !string.IsNullOrEmpty(Strings.Trim(CLTCOD)) & !string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
//							//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD);
//							MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//						}
//						break;
//					case "DOCNO4":
//						oTmpObject = new ZPY504();
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						DocNum = oDS_ZPY507.GetValue("DOCNO4", LineNum);
//						if (!string.IsNullOrEmpty(Strings.Trim(DocNum))) {
//							//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oTmpObject.LoadForm(DocNum);
//							MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//						}
//						break;
//					case "DOCNO5":
//						oTmpObject = new ZPY506();
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						JSNYER = Strings.Trim(oDS_ZPY507.GetValue("JSNYMM", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						CLTCOD = Strings.Trim(oDS_ZPY507.GetValue("CLTCOD", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MSTCOD = Strings.Trim(oDS_ZPY507.GetValue("MSTCOD", LineNum));
//						if (!string.IsNullOrEmpty(Strings.Trim(JSNYER)) & !string.IsNullOrEmpty(Strings.Trim(CLTCOD)) & !string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
//							//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD);
//							MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//						}
//						break;
//					case "DOCNO6":
//						oTmpObject = new ZPY505();
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						JSNYER = Strings.Trim(oDS_ZPY507.GetValue("JSNYMM", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						CLTCOD = Strings.Trim(oDS_ZPY507.GetValue("CLTCOD", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MSTCOD = Strings.Trim(oDS_ZPY507.GetValue("MSTCOD", LineNum));
//						if (!string.IsNullOrEmpty(Strings.Trim(JSNYER)) & !string.IsNullOrEmpty(Strings.Trim(CLTCOD)) & !string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
//							//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD);
//							MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//						}
//						break;
//					case "DOCNO7":
//						oTmpObject = new ZPY508();
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						JSNYER = Strings.Trim(oDS_ZPY507.GetValue("JSNYMM", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						CLTCOD = Strings.Trim(oDS_ZPY507.GetValue("CLTCOD", LineNum));
//						//UPGRADE_WARNING: oDS_ZPY507.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MSTCOD = Strings.Trim(oDS_ZPY507.GetValue("MSTCOD", LineNum));
//						if (!string.IsNullOrEmpty(Strings.Trim(JSNYER)) & !string.IsNullOrEmpty(Strings.Trim(CLTCOD)) & !string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
//							//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD);
//							MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//						}
//						break;
//				}

//			}
//			//UPGRADE_NOTE: oTmpObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oTmpObject = null;
//		}

//		private void Grid_Display()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			string FYEAR = null;
//			string TYEAR = null;
//			string JIGFDAT = null;
//			string JIGTDAT = null;
//			string SINFYMM = null;
//			string SINTYMM = null;
//			string MSTCOD = null;
//			string MSTNAM = null;
//			string CLTCOD = null;
//			string JSNGBN = null;
//			double PILMED = 0;
//			double PILGBU = 0;
//			int iRow = 0;
//			/// Check
//			ErrNum = 0;
//			iRow = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			var _with1 = oForm.DataSources.UserDataSources;
//			FYEAR = Strings.Trim(_with1.Item("FYEAR").ValueEx);
//			TYEAR = Strings.Trim(_with1.Item("TYEAR").ValueEx);
//			JIGFDAT = Strings.Trim(_with1.Item("JIGFDAT").ValueEx);
//			JIGTDAT = Strings.Trim(_with1.Item("JIGTDAT").ValueEx);
//			SINFYMM = Strings.Trim(_with1.Item("SINFYMM").ValueEx);
//			SINTYMM = Strings.Trim(_with1.Item("SINTYMM").ValueEx);
//			MSTCOD = Strings.Trim(_with1.Item("MSTCOD").ValueEx);
//			MSTNAM = Strings.Trim(_with1.Item("MSTNAM").ValueEx);
//			PILMED = Conversion.Val(_with1.Item("PILMED").ValueEx);
//			PILGBU = Conversion.Val(_with1.Item("PILGBU").ValueEx);
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JSNGBN = oForm.Items.Item("JSNGBN").Specific.Selected.Value;

//			//// 정산년도 체크(입력 안할 경우 에러)
//			if (string.IsNullOrEmpty(FYEAR) | string.IsNullOrEmpty(TYEAR)) {
//				ErrNum = 1;
//				goto Error_Message;
//			}
//			//// 지급일자 체크(입력 안할 경우 전체조회)
//			if (string.IsNullOrEmpty(JIGFDAT) | string.IsNullOrEmpty(JIGTDAT)) {
//				JIGFDAT = "1900-01-01";
//				JIGTDAT = "2999-12-31";
//			}
//			//// 신고년월 체크(입력 안할 경우 전체조회)
//			if (string.IsNullOrEmpty(SINFYMM) | string.IsNullOrEmpty(SINTYMM)) {
//				SINFYMM = "190001";
//				SINTYMM = "299912";
//			}

//			/// 조회
//			sQry = "Exec ZPY507 N'" + FYEAR + "', N'" + TYEAR + "', " + "N'" + JIGFDAT + "', N'" + JIGTDAT + "', " + "N'" + CLTCOD + "', N'" + MSTCOD + "', " + "N'" + MSTNAM + "', N'" + JSNGBN + "', " + "N'" + SINFYMM + "', N'" + SINTYMM + "', " + Convert.ToString(PILMED) + ", " + Convert.ToString(PILGBU);
//			oDS_ZPY507.ExecuteQuery((sQry));
//			iRow = oDS_ZPY507.Rows.Count;
//			if (iRow == 1) {
//				oRecordSet.DoQuery(sQry);
//				iRow = oRecordSet.RecordCount;
//			}

//			if (iRow > 0) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText(iRow + " 건이 있습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("조회된 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			titleSetting();
//			return;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기준년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Grid_Display Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
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


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY507.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY507_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY507");
//			oForm.SupportedModes = -1;

//			////////////////////////////////////////////////////////////////////////////////
//			//***************************************************************
//			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//			//oForm.DataBrowser.BrowseBy = "Code"
//			//***************************************************************
//			////////////////////////////////////////////////////////////////////////////////
//			oForm.Freeze(true);
//			CreateItems();

//			oForm.EnableMenu(("1281"), false);
//			/// 추가
//			oForm.EnableMenu(("1282"), false);
//			/// 추가

//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//			oForm.Freeze(false);
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

//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.EditText oEdit = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			///UserDataSource 선언
//			var _with2 = oForm.DataSources.UserDataSources;
//			_with2.Add("FYEAR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			_with2.Add("TYEAR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			_with2.Add("JIGFDAT", SAPbouiCOM.BoDataType.dt_DATE);
//			_with2.Add("JIGTDAT", SAPbouiCOM.BoDataType.dt_DATE);
//			_with2.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			_with2.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 200);
//			_with2.Add("SINFYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			_with2.Add("SINTYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			_with2.Add("PILMED", SAPbouiCOM.BoDataType.dt_SUM);
//			_with2.Add("PILGBU", SAPbouiCOM.BoDataType.dt_SUM);

//			oEdit = oForm.Items.Item("FYEAR").Specific;
//			//// 기준년도(From)
//			oEdit.DataBind.SetBound(true, "", "FYEAR");

//			oEdit = oForm.Items.Item("TYEAR").Specific;
//			//// 기준년도(To)
//			oEdit.DataBind.SetBound(true, "", "TYEAR");

//			oEdit = oForm.Items.Item("JIGFDAT").Specific;
//			//// 지급일자(From)
//			oEdit.DataBind.SetBound(true, "", "JIGFDAT");

//			oEdit = oForm.Items.Item("JIGTDAT").Specific;
//			//// 지급일자(To)
//			oEdit.DataBind.SetBound(true, "", "JIGTDAT");

//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			//// 사번
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");

//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			//// 성명
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");

//			oEdit = oForm.Items.Item("SINFYMM").Specific;
//			//// 신고년월(From)
//			oEdit.DataBind.SetBound(true, "", "SINFYMM");

//			oEdit = oForm.Items.Item("SINTYMM").Specific;
//			//// 신고년월(To)
//			oEdit.DataBind.SetBound(true, "", "SINTYMM");

//			oEdit = oForm.Items.Item("PILMED").Specific;
//			//// 의료비공제액
//			oEdit.DataBind.SetBound(true, "", "PILMED");

//			oEdit = oForm.Items.Item("PILGBU").Specific;
//			//// 기부금공제액
//			oEdit.DataBind.SetBound(true, "", "PILGBU");


//			////사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.ValidValues.Add("%", "전체");
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;
//			oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

//			//// 정산구분
//			oCombo = oForm.Items.Item("JSNGBN").Specific;
//			oCombo.ValidValues.Add("%", "모두");
//			oCombo.ValidValues.Add("1", "연말정산(재직자)");
//			oCombo.ValidValues.Add("2", "중도정산(퇴직자)");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


//			oForm.DataSources.UserDataSources.Item("FYEAR").ValueEx = MDC_Globals.ZPAY_GBL_JSNYER.Value;
//			oForm.DataSources.UserDataSources.Item("TYEAR").ValueEx = MDC_Globals.ZPAY_GBL_JSNYER.Value;

//			//// 디비데이터 소스 개체 할당
//			oGrid1 = oForm.Items.Item("Grid1").Specific;
//			oDS_ZPY507 = oForm.DataSources.DataTables.Add("ZPY507");
//			oDS_ZPY507.ExecuteQuery(("Exec ZPY507 '1900', '1900', NULL, NULL, '%', '%', '%', '%', '', '', 0, 0"));
//			oGrid1.DataTable = oDS_ZPY507;

//			titleSetting();

//			oForm.ActiveItem = "FYEAR";

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

//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
//	}
//}
