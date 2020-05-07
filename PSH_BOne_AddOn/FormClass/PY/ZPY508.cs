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
//	internal class ZPY508
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY508.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 연금저축 소득공제 명세 등록
//////  FormType       : 2000060508
//////  Create Date    : 2011.01.03
//////  Modified Date  :
//////  Creator        : Choi Dong Kwon
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************


//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//			//시스템코드 헤더
//		private SAPbouiCOM.DBDataSource oDS_ZPY508H;
//			//시스템코드 라인
//		private SAPbouiCOM.DBDataSource oDS_ZPY508L;
//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string Col_Last_Uid;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int Col_Last_Row;
//		private string oOLDCHK;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다
////*******************************************************************
//		public void LoadForm(ref string JSNYER = "", ref string MSTCOD = "", ref string CLTCOD = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY508.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY508_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY508");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

//			////////////////////////////////////////////////////////////////////////////////
//			//***************************************************************
//			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//			oForm.DataBrowser.BrowseBy = "DocNum";
//			//***************************************************************
//			////////////////////////////////////////////////////////////////////////////////
//			oForm.Freeze(true);
//			CreateItems();

//			oForm.EnableMenu(("1293"), true);
//			/// 행삭제
//			oForm.EnableMenu(("1283"), true);
//			/// 제거
//			oForm.EnableMenu(("1284"), false);
//			/// 취소


//			if (!string.IsNullOrEmpty(JSNYER)) {
//				ShowSource(ref JSNYER, ref MSTCOD, ref CLTCOD);
//			}

//			oForm.Freeze(false);
//			oForm.Update();
//			//oForm.Visible = True

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
//						if (pval.ItemUID == "1") {
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							////추가및 업데이시에
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//									//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (MDC_SetMod.Value_ChkYn(ref "[@ZPY508H]", ref "U_JSNYER", ref "'" + oForm.Items.Item("JSNYER").Specific.String + "'", ref " AND U_MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.String + "'") == false) {
//										MDC_Globals.Sbo_Application.StatusBar.SetText("이미 저장되어져 있는 헤더의 내용과 일치합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//										BubbleEvent = false;
//										return;
//									}
//								}
//								if (Strings.Trim(oDS_ZPY508H.GetValue("U_ENDCHK", 0)) == "Y" & Strings.Trim(oOLDCHK) == "Y") {
//									MDC_Globals.Sbo_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//									BubbleEvent = false;
//									return;
//								} else if (MatrixSpaceLineDel() == false) {
//									BubbleEvent = false;
//								}
//							}
//						/// ChooseBtn사원리스트
//						} else if (pval.ItemUID == "CBtn1" & oForm.Items.Item("MSTCOD").Enabled == true) {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						}
//					} else {
//						if (pval.ItemUID == "1" & pval.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							MDC_Globals.Sbo_Application.ActivateMenuItem("1282");
//						}
//					}
//					break;
//				//et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					if (pval.BeforeAction == true & pval.ItemUID != "1000001" & pval.ItemUID != "2" & oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//						if (Last_Item == "MSTCOD") {
//							//UPGRADE_WARNING: oForm.Items(Last_Item).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + oForm.Items.Item(Last_Item).Specific.String + "'", ref "") == true & !string.IsNullOrEmpty(oForm.Items.Item(Last_Item).Specific.String) & Last_Item != pval.ItemUID) {
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					}
//					if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == true & Last_Item == "Mat1" & Col_Last_Uid == "Col1" & Col_Last_Row > 0 & (Col_Last_Uid != pval.ColUID | Col_Last_Row != pval.Row) & pval.ItemUID != "1000001" & pval.ItemUID != "2") {
//						if (Col_Last_Row > oMat1.VisualRowCount) {
//							return;
//						}
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if ((pval.ItemUID == "MSTCOD" | pval.ItemUID == "JSNYER")) {
//							FlushToItemValue(pval.ItemUID);
//						} else if (pval.ItemUID == "Mat1" & (pval.ColUID == "Col5")) {
//							FlushToItemValue(pval.ColUID, ref pval.Row);
//						}
//					}
//					break;

//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					////추가모드에서 코드이벤트가 코드에서 일어 났을때
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9 & pval.FormMode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + oForm.Items.Item(pval.ItemUID).Specific.String + "'", ref "") == true) {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						} else {
//							if (oMat1.RowCount > 0) {
//								oMat1.Columns.Item("Col5").Cells.Item(oMat1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								BubbleEvent = false;
//							}
//						}
//					}
//					break;
//				//et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					if (Last_Item == "Mat1") {
//						if (pval.Row > 0) {
//							Last_Item = pval.ItemUID;
//							Col_Last_Row = pval.Row;
//							Col_Last_Uid = pval.ColUID;
//						}
//					} else {
//						Last_Item = pval.ItemUID;
//						Col_Last_Row = 0;
//						Col_Last_Uid = "";
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
//						//UPGRADE_NOTE: oDS_ZPY508H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY508H = null;
//						//UPGRADE_NOTE: oDS_ZPY508L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY508L = null;
//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//					}
//					break;
//				//et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					if (pval.BeforeAction == false) {
//						FormItemEnabled();
//						Matrix_AddRow(oMat1.VisualRowCount);
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
//			int i = 0;

//			if (pval.BeforeAction == true) {
//				switch (pval.MenuUID) {
//					case "1283":
//						/// 제거
//						if (Strings.Trim(oDS_ZPY508H.GetValue("U_ENDCHK", 0)) == "Y") {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//							return;
//						} else {
//							if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
//								BubbleEvent = false;
//								return;
//							}
//						}
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@ZPY508H", ref "DocNum");
//						////접속자 권한에 따른 사업장 보기
//						break;


//					default:
//						return;

//						break;
//				}
//			} else {

//				switch (pval.MenuUID) {
//					case "1287":
//						/// 복제
//						break;
//					case "1283":
//						/// 제거
//						FormItemEnabled();
//						break;
//					case "1281":
//					case "1282":
//						FormItemEnabled();
//						if (pval.MenuUID == "1282") {
//							FormClear();
//							Matrix_AddRow(0, ref true);
//							oForm.Items.Item("JSNYER").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						}
//						break;
//					case "1288": // TODO: to "1291"
//						break;
//					case "1293":
//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//							////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							for (i = 0; i <= oMat1.VisualRowCount - 1; i++) {
//								//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat1.Columns.Item("Col0").Cells.Item(i + 1).Specific.Value = i + 1;
//							}

//							oMat1.FlushToDataSource();
//							oDS_ZPY508L.RemoveRecord(oDS_ZPY508L.Size - 1);
//							//// Mat1에 마지막라인(빈라인) 삭제
//							oMat1.Clear();
//							oMat1.LoadFromDataSource();

//						}
//						FlushToItemValue("Col3", ref 1);
//						break;
//				}
//			}
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
////
////*******************************************************************
//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.CheckBox oCheck = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////디비데이터 소스 개체 할당
//			oDS_ZPY508H = oForm.DataSources.DBDataSources("@ZPY508H");
//			oDS_ZPY508L = oForm.DataSources.DBDataSources("@ZPY508L");

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			////사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.ValidValues.Add("%", "전체");
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;
//			oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

//			//// 관계코드
//			oColumn = oMat1.Columns.Item("Col1");
//			oColumn.ValidValues.Add("11", "퇴직연금-근로자퇴직급여보장법");
//			oColumn.ValidValues.Add("12", "퇴직연금-과학기술인공제회");
//			oColumn.ValidValues.Add("21", "연금저축-개인연금저축");
//			oColumn.ValidValues.Add("22", "연금저축-연금저축");
//			oColumn.ValidValues.Add("31", "주택마련-청약저축");
//			oColumn.ValidValues.Add("32", "주택마련-주택청약종합저축");
//			oColumn.ValidValues.Add("33", "주택마련-장기주택마련저축");
//			oColumn.ValidValues.Add("34", "주택마련-근로자주택마련저축");
//			oColumn.ValidValues.Add("41", "장기주식형저축");

//			//// 금융기관
//			oColumn = oMat1.Columns.Item("Col2");
//			sQry = "SELECT BankCode, BankName FROM [ODSC]";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//				oRecordSet.MoveNext();
//			}

//			//// 납입연차
//			oColumn = oMat1.Columns.Item("Col4");
//			oColumn.ValidValues.Add("00", "해당없음");
//			oColumn.ValidValues.Add("01", "1년차");
//			oColumn.ValidValues.Add("02", "2년차");
//			oColumn.ValidValues.Add("03", "3년차");

//			//// 종(전) 여부
//			oColumn = oMat1.Columns.Item("Col7");
//			oColumn.ValOff = "N";
//			oColumn.ValOn = "Y";

//			/// Check 버튼
//			oCheck = oForm.Items.Item("ENDCHK").Specific;
//			oCheck.ValOff = "N";
//			oCheck.ValOn = "Y";

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{
//			int iRow = 0;
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);
//			double TOTAMT = 0;

//			switch (oUID) {
//				case "JSNYER":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oUID).Specific.String))) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MDC_Globals.ZPAY_GBL_JSNYER.Value = oForm.Items.Item(oUID).Specific.String;
//					} else {
//						oDS_ZPY508H.SetValue("U_JSNYER", 0, MDC_Globals.ZPAY_GBL_JSNYER.Value);
//					}
//					oForm.Items.Item(oUID).Update();
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oDS_ZPY508H.SetValue("U_MSTCOD", 0, "");
//						oDS_ZPY508H.SetValue("U_MSTNAM", 0, "");
//						oDS_ZPY508H.SetValue("U_EmpID", 0, "");
//						oDS_ZPY508H.SetValue("U_CLTCOD", 0, "");
//					} else {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oDS_ZPY508H.SetValue("U_MSTCOD", 0, Strings.UCase(oForm.Items.Item(oUID).Specific.String));
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMast = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oDS_ZPY508H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//						oDS_ZPY508H.SetValue("U_EmpID", 0, oMast.EmpID);
//						oDS_ZPY508H.SetValue("U_CLTCOD", 0, oMast.CLTCOD);
//					}
//					oForm.Items.Item("MSTNAM").Update();
//					oForm.Items.Item("EmpID").Update();
//					oForm.Items.Item("CLTCOD").Update();
//					oForm.Items.Item(oUID).Update();
//					break;
//			}
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			switch (oUID) {
//				case "Col5":
//					oMat1.FlushToDataSource();

//					for (iRow = 1; iRow <= oMat1.VisualRowCount; iRow++) {
//						TOTAMT = TOTAMT + Conversion.Val(oDS_ZPY508L.GetValue("U_SAVAMT", iRow - 1));
//					}

//					oDS_ZPY508H.SetValue("U_TOTAMT", 0, Convert.ToString(TOTAMT));
//					oForm.Items.Item("TOTAMT").Update();
//					break;
//			}
//			if (Strings.Left(oUID, 3) == "Col") {
//				oMat1.FlushToDataSource();
//				if (oRow == oMat1.RowCount & Conversion.Val(oDS_ZPY508L.GetValue("U_SAVAMT", oRow - 1)) != 0) {
//					Matrix_AddRow(oRow);
//					oMat1.Columns.Item("Col5").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				}
//			}
//		}
//		private void FormClear()
//		{
//			int DocNum = 0;

//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocNum = MDC_SetMod.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'ZPY508'", ref "");

//			if (DocNum == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.String = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.String = DocNum;
//			}
//			FlushToItemValue("JSNYER");

//		}

//		private void FormItemEnabled()
//		{
//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = true;
//				oForm.Items.Item("MSTCOD").Enabled = true;
//				oForm.Items.Item("MSTNAM").Enabled = true;
//				oForm.Items.Item("DocNum").Enabled = true;
//				oForm.Items.Item("ENDCHK").Enabled = true;
//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = true;
//				oForm.Items.Item("MSTCOD").Enabled = true;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = false;
//				oForm.Items.Item("ENDCHK").Enabled = true;
//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = false;
//				oForm.Items.Item("MSTCOD").Enabled = false;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = false;
//				//// 년마감된것은 비활성화
//				oOLDCHK = oDS_ZPY508H.GetValue("U_ENDCHK", 0);
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(U_ENDCHK, U_JOBYER, [ZPY509L], ' & oDS_ZPY508H.GetValue(U_JSNYER, 0) & ',  AND Code = ' & oDS_ZPY508H.GetValue(U_CLTCOD, 0) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (MDC_SetMod.Get_ReData(ref "U_ENDCHK", ref "U_JOBYER", ref "[@ZPY509L]", ref "'" + oDS_ZPY508H.GetValue("U_JSNYER", 0) + "'", ref " AND Code = '" + oDS_ZPY508H.GetValue("U_CLTCOD", 0) + "'") == "Y") {
//					oForm.Items.Item("ENDCHK").Enabled = false;
//				} else {
//					oForm.Items.Item("ENDCHK").Enabled = true;
//				}

//			}
//		}

//		private bool MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//저장할 데이터의 유효성을 점검한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int iRow = 0;
//			int kRow = 0;
//			short ErrNum = 0;
//			string Chk_Data = null;

//			ErrNum = 0;
//			/// 헤더부분 체크
//			switch (true) {
//				case Strings.Len(Strings.Trim(oDS_ZPY508H.GetValue("U_JSNYER", 0))) != 4:
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY508H.GetValue("U_MSTCOD", 0)):
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY508H.GetValue("U_CLTCOD", 0)):
//					ErrNum = 4;
//					goto Error_Message;
//					break;

//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oMat1.FlushToDataSource();

//			//// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
//			if (oMat1.RowCount == 1) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
//			////oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
//			for (iRow = 0; iRow <= oMat1.VisualRowCount - 2; iRow++) {
//				oDS_ZPY508L.Offset = iRow;
//				if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)))) {
//					ErrNum = 6;
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY508L.GetValue("U_SAVCOD", iRow)))) {
//					ErrNum = 7;
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY508L.GetValue("U_SAVNUM", iRow)))) {
//					ErrNum = 8;
//					oMat1.Columns.Item("Col3").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (Strings.Trim(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) == "41" & Strings.Trim(oDS_ZPY508L.GetValue("U_STYEAR", iRow)) == "00") {
//					ErrNum = 9;
//					oMat1.Columns.Item("Col3").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else {
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//중복체크작업
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					Chk_Data = Strings.Trim(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) + Strings.Trim(oDS_ZPY508L.GetValue("U_SAVCOD", iRow)) + Strings.Trim(oDS_ZPY508L.GetValue("U_SAVNUM", iRow)) + Strings.Trim(oDS_ZPY508L.GetValue("U_STYEAR", iRow));
//					for (kRow = iRow + 1; kRow <= oMat1.VisualRowCount - 2; kRow++) {
//						oDS_ZPY508L.Offset = kRow;
//						if (Strings.Trim(Chk_Data) == Strings.Trim(oDS_ZPY508L.GetValue("U_SAVGBN", kRow)) + Strings.Trim(oDS_ZPY508L.GetValue("U_SAVCOD", kRow)) + Strings.Trim(oDS_ZPY508L.GetValue("U_SAVNUM", kRow)) + Strings.Trim(oDS_ZPY508L.GetValue("U_STYEAR", kRow))) {
//							ErrNum = 5;
//							oMat1.Columns.Item("Col3").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							goto Error_Message;
//						}
//					}
//				}

//				if (Strings.Trim(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) != "41" & Strings.Trim(oDS_ZPY508L.GetValue("U_STYEAR", iRow)) != "00") {
//					oDS_ZPY508L.Offset = iRow;
//					oDS_ZPY508L.SetValue("U_STYEAR", iRow, "00");
//					oMat1.SetLineData((iRow + 1));
//				}
//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//			////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oDS_ZPY508L.RemoveRecord(oDS_ZPY508L.Size - 1);
//			//// Mat1에 마지막라인(빈라인) 삭제

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			switch (ErrNum) {
//				case 1:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 2:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 3:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("의료비 지급금액이 0입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 4:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 5:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("중복입력되었습니다. 저축구분/금융기관/계좌번호/납입연차별로 집계하여 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 6:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("저축구분은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 7:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("금융기관은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 8:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("계좌번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 9:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("장기주식형 저축인 경우 납입연차를 1년차~3년차로 선택하여야 합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				default:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("MatrixSpaceLineDel Error:" + Err().Number + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void ShowSource(ref string JSNYER, ref string MSTCOD, ref string CLTCOD)
//		{
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string DocEntry = null;
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			sQry = "SELECT DocEntry FROM [@ZPY508H]";
//			sQry = sQry + "   WHERE U_JSNYER = N'" + JSNYER + "'";
//			sQry = sQry + "   AND   U_MSTCOD = N'" + MSTCOD + "'";
//			sQry = sQry + "   AND   U_CLTCOD = N'" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				while (!(oRecordSet.EoF)) {
//					//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					DocEntry = oRecordSet.Fields.Item(0).Value;
//					oRecordSet.MoveNext();
//				}
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("JSNYER").Specific.Value = JSNYER;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("MSTCOD").Specific.String = MSTCOD;

//				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("CLTCOD").Specific.Select(CLTCOD, SAPbouiCOM.BoSearchKey.psk_ByValue);
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.Value = DocEntry;

//				oForm.Items.Item("DocNum").Update();
//				oMat1.LoadFromDataSource();
//				oForm.Update();
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//				MDC_Globals.Sbo_Application.ActivateMenuItem("1282");

//				oDS_ZPY508H.SetValue("U_JSNYER", 0, JSNYER);
//				oDS_ZPY508H.SetValue("U_MSTCOD", 0, MSTCOD);
//				oDS_ZPY508H.SetValue("U_CLTCOD", 0, CLTCOD);
//				//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oMast = MDC_SetMod.Get_EmpID_InFo(ref MSTCOD);
//				oDS_ZPY508H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//				oDS_ZPY508H.SetValue("U_EmpID", 0, oMast.EmpID);

//				oForm.Update();

//				MDC_Globals.Sbo_Application.SendKeys("{TAB}");
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//		}

////*******************************************************************
////// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
////*******************************************************************
//		private void Matrix_AddRow(int oRow, ref bool Insert_YN = false)
//		{
//			if (Insert_YN == false) {
//				oDS_ZPY508L.InsertRecord((oRow));
//			}
//			oDS_ZPY508L.Offset = oRow;
//			oDS_ZPY508L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//			oDS_ZPY508L.SetValue("U_SAVGBN", oRow, "");
//			oDS_ZPY508L.SetValue("U_SAVCOD", oRow, "");
//			oDS_ZPY508L.SetValue("U_JONGYN", oRow, "N");
//			oDS_ZPY508L.SetValue("U_SAVNAM", oRow, "");
//			oDS_ZPY508L.SetValue("U_SAVNUM", oRow, "");
//			oDS_ZPY508L.SetValue("U_STYEAR", oRow, "00");
//			oDS_ZPY508L.SetValue("U_SAVAMT", oRow, Convert.ToString(0));
//			oDS_ZPY508L.SetValue("U_SARAMT", oRow, Convert.ToString(0));
//			oMat1.LoadFromDataSource();
//		}
//	}
//}
