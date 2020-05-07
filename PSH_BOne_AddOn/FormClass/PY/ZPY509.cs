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
//	internal class ZPY509
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY509.cls
//////  Module         : 원천징수>근로소득
//////  Desc           : 정산자료 마감작업
//////  FormType       : 2010110509
//////  Create Date    : 2009.02.13
//////  Modified Date  :
//////  Creator        : Choi Dong Kwon
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//		private SAPbouiCOM.DBDataSource oDS_ZPY509H;
//		private SAPbouiCOM.DBDataSource oDS_ZPY509L;

//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string Col_Last_Uid;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int Col_Last_Row;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY509.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY509_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY509");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

//			////////////////////////////////////////////////////////////////////////////////
//			//***************************************************************
//			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//			oForm.DataBrowser.BrowseBy = "Code";
//			//***************************************************************
//			////////////////////////////////////////////////////////////////////////////////
//			oForm.Freeze(true);
//			CreateItems();
//			FormItemEnabled();

//			oForm.EnableMenu(("1293"), false);
//			/// 행삭제
//			oForm.EnableMenu(("1283"), true);
//			/// 제거
//			oForm.EnableMenu(("1287"), false);
//			/// 복제
//			oForm.EnableMenu(("1284"), false);
//			/// 취소

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

////*******************************************************************
////
////*******************************************************************
//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// Matrix
//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			////디비데이터 소스 개체 할당
//			oDS_ZPY509H = oForm.DataSources.DBDataSources("@ZPY509H");
//			oDS_ZPY509L = oForm.DataSources.DBDataSources("@ZPY509L");

//			////사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			//    oCombo.ValidValues.Add "%", "전체"
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;
//			//    oCombo.Select "%", psk_ByValue

//			/// 월별자료
//			oColumn = oMat1.Columns.Item("Col5");
//			oColumn.ValOff = "N";
//			oColumn.ValOn = "Y";

//			/// 정산자료
//			oColumn = oMat1.Columns.Item("Col2");
//			oColumn.ValOff = "N";
//			oColumn.ValOn = "Y";

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

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
//							////추가및 업데이트시에
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//									//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (MDC_SetMod.Value_ChkYn("[@ZPY509H]", "Code", "'" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "'") == false) {
//										MDC_Globals.Sbo_Application.StatusBar.SetText("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//										BubbleEvent = false;
//										return;
//									}
//								}
//								if (HeaderSpaceLineDel() == false) {
//									BubbleEvent = false;
//									return;
//								} else if (MatrixSpaceLineDel() == false) {
//									BubbleEvent = false;
//									return;
//								} else {
//									Batch_EndCheck();
//								}
//							}
//						} else if (pval.ItemUID == "Btn1") {
//							Create_Year();
//						}
//					} else {
//						if (pval.ItemUID == "1" & pval.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							MDC_Globals.Sbo_Application.ActivateMenuItem("1282");
//						} else if (pval.ItemUID == "Mat1" & (pval.ColUID == "Col2" | pval.ColUID == "Col5")) {
//							FlushToItemValue(pval.ColUID, ref pval.Row);
//						}
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.ItemUID == "Mat1" & pval.ColUID == "Col1") {
//						FlushToItemValue(pval.ColUID, ref pval.Row);
//					}
//					break;

//				//et_CLICK'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == true & Last_Item == "Mat1" & Col_Last_Uid == "Col1" & Col_Last_Row > 0 & (Col_Last_Uid != pval.ColUID | Col_Last_Row != pval.Row) & pval.ItemUID != "1000001" & pval.ItemUID != "2") {
//						if (Col_Last_Row > oMat1.VisualRowCount) {
//							return;
//						}
//					} else if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == true & pval.ItemUID == "Mat1" & pval.Row > 0) {
//						Last_Item = pval.ItemUID;
//						Col_Last_Row = pval.Row;
//						Col_Last_Uid = pval.ColUID;
//					}
//					break;
//				//et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
//						//UPGRADE_NOTE: oDS_ZPY509H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY509H = null;
//						//UPGRADE_NOTE: oDS_ZPY509L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY509L = null;
//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//					}
//					break;
//				//et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					if (pval.BeforeAction == false) {
//						FormItemEnabled();
//						Matrix_AddRow(oMat1.VisualRowCount, ref false);
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
//				switch (pval.MenuUID) {
//					case "1283":
//						/// 제거
//						if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
//							BubbleEvent = false;
//							return;
//						}
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@ZPY509H", ref "Code");
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
//					case "1281":
//					case "1282":
//						FormItemEnabled();
//						if (pval.MenuUID == "1282") {
//							Matrix_AddRow(0, ref true);

//						}
//						break;
//					//        Case "1283" '/ 제거
//					//             FormItemEnabled
//					case "1288": // TODO: to "1291"
//						FormItemEnabled();
//						break;
//					//        Case "1293" '/ 행삭제

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
////// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
////*******************************************************************
//		private void Matrix_AddRow(int oRow, ref bool Insert_YN = false)
//		{
//			if (Insert_YN == false) {
//				oDS_ZPY509L.InsertRecord((oRow));
//			}
//			oDS_ZPY509L.Offset = oRow;
//			oDS_ZPY509L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//			oDS_ZPY509L.SetValue("U_JOBYER", oRow, "");
//			oDS_ZPY509L.SetValue("U_MONCHK", oRow, "N");
//			oDS_ZPY509L.SetValue("U_ENDCHK", oRow, "N");
//			oDS_ZPY509L.SetValue("U_LGNADM", oRow, "");
//			oDS_ZPY509L.SetValue("U_MODDAT", oRow, "");
//			oMat1.LoadFromDataSource();
//		}
//		private void FormItemEnabled()
//		{
//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//				oForm.Items.Item("CLTCOD").Enabled = true;
//			} else {
//				oForm.Items.Item("CLTCOD").Enabled = false;
//				oForm.Items.Item("Btn1").Enabled = false;
//			}
//			if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//				oForm.Items.Item("Btn1").Enabled = true;
//			} else {
//				oForm.Items.Item("Btn1").Enabled = false;
//			}
//			//// 접속자에 따른 권한별 사업장 콤보박스세팅
//			MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//		}

//		private bool HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;

//			ErrNum = 0;

//			/// Check
//			//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case oForm.Items.Item("CLTCOD").Specific.Selected == null:
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//			}

//			oDS_ZPY509H.SetValue("Code", 0, oDS_ZPY509H.GetValue("U_CLTCOD", 0));

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 선택하여 주십시오", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}
//		private bool MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//저장할 데이터의 유효성을 점검한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			int k = 0;
//			short ErrNum = 0;
//			string Chk_Data = null;

//			ErrNum = 0;
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
//			for (i = 0; i <= oMat1.VisualRowCount - 2; i++) {
//				oDS_ZPY509L.Offset = i;
//				if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY509L.GetValue("U_JOBYER", i)))) {
//					ErrNum = 2;
//					oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;

//				} else {
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//중복체크작업
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					Chk_Data = Strings.Trim(oDS_ZPY509L.GetValue("U_JOBYER", i));
//					for (k = i + 1; k <= oMat1.VisualRowCount - 2; k++) {
//						oDS_ZPY509L.Offset = k;
//						if (Strings.Trim(Chk_Data) == Strings.Trim(oDS_ZPY509L.GetValue("U_JOBYER", k))) {
//							ErrNum = 3;
//							oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							goto Error_Message;
//						}
//					}
//				}
//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//			////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oDS_ZPY509L.RemoveRecord(oDS_ZPY509L.Size - 1);
//			//// Mat1에 마지막라인(빈라인) 삭제

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도가 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("MatrixSpaceLineDel Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

////---------------------------------------------------------------------------------------
//// Procedure : Create_Year
//// DateTime  : 2009-02-16
//// Author    :
//// Purpose   : 연도 생성
////---------------------------------------------------------------------------------------
////
//		private void Create_Year()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string CLTCOD = null;
//			int iRow = 0;
//			int kRow = 0;
//			int MaxRow = 0;
//			string check = null;

//			short NowYer = 0;

//			CLTCOD = Strings.Trim(oDS_ZPY509H.GetValue("U_CLTCOD", 0));
//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			NowYer = MDC_SetMod.Get_ReData("CONVERT(CHAR(4), GETDATE(), 120)", "1", "OADM", "1");
//			MaxRow = oMat1.VisualRowCount - 1;

//			oMat1.FlushToDataSource();
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//// Matrix 맨밑에 있는 빈줄 삭제
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY509L.GetValue("U_JOBYER", MaxRow)))) {
//				oDS_ZPY509L.RemoveRecord(MaxRow);
//				MaxRow = MaxRow - 1;
//			}
//			Create_Step1:

//			/// 기존 정산데이터에 대한 마감연도를 생성
//			sQry = "EXEC ZPY509_1 '" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount == 0) {
//				goto Create_Step2;
//			}

//			while (!(oRecordSet.EoF)) {
//				check = "N";
//				//// 생성하려는 연도가 이미 Matrix에 존재하는지 비교
//				if (oDS_ZPY509L.Size > 0) {
//					for (kRow = 0; kRow <= oDS_ZPY509L.Size - 1; kRow++) {
//						if (Conversion.Val(oDS_ZPY509L.GetValue("U_JOBYER", kRow)) == Conversion.Val(oRecordSet.Fields.Item("U_JSNYER").Value)) {
//							check = "Y";
//						}
//					}
//				}
//				//// 생성하려는 연도가 이미 Matrix에 존재하는 경우 생성 안함
//				if (check == "N") {
//					MaxRow = MaxRow + 1;
//					oDS_ZPY509L.InsertRecord(MaxRow);
//					oDS_ZPY509L.Offset = MaxRow;
//					oDS_ZPY509L.SetValue("U_LINENUM", MaxRow, Convert.ToString(MaxRow + 1));
//					oDS_ZPY509L.SetValue("U_JOBYER", MaxRow, Convert.ToString(Conversion.Val(oRecordSet.Fields.Item("U_JSNYER").Value)));
//					oDS_ZPY509L.SetValue("U_MONCHK", MaxRow, "N");
//					oDS_ZPY509L.SetValue("U_ENDCHK", MaxRow, "N");
//					oDS_ZPY509L.SetValue("U_LGNADM", MaxRow, "");
//					oDS_ZPY509L.SetValue("U_MODDAT", MaxRow, "");
//				}
//				oRecordSet.MoveNext();
//			}
//			Create_Step2:

//			/// 올해부터 앞으로 10년간에 대한 마감연도를 생성

//			for (iRow = NowYer; iRow <= NowYer + 10; iRow++) {
//				check = "N";
//				//// 생성하려는 연도가 이미 Matrix에 존재하는지 비교
//				if (oDS_ZPY509L.Size > 0) {
//					for (kRow = 0; kRow <= oDS_ZPY509L.Size - 1; kRow++) {
//						if (Conversion.Val(oDS_ZPY509L.GetValue("U_JOBYER", kRow)) == iRow) {
//							check = "Y";
//						}
//					}
//				}
//				//// 생성하려는 연도가 이미 Matrix에 존재하는 경우 생성 안함
//				if (check == "N") {
//					MaxRow = MaxRow + 1;
//					oDS_ZPY509L.InsertRecord(MaxRow);
//					oDS_ZPY509L.Offset = MaxRow;
//					oDS_ZPY509L.SetValue("U_LINENUM", MaxRow, Convert.ToString(MaxRow + 1));
//					oDS_ZPY509L.SetValue("U_JOBYER", MaxRow, Convert.ToString(iRow));
//					oDS_ZPY509L.SetValue("U_MONCHK", MaxRow, "N");
//					oDS_ZPY509L.SetValue("U_ENDCHK", MaxRow, "N");
//					oDS_ZPY509L.SetValue("U_LGNADM", MaxRow, "");
//					oDS_ZPY509L.SetValue("U_MODDAT", MaxRow, "");
//				}
//			}

//			//// Matrix밑에 빈줄 추가
//			Matrix_AddRow(MaxRow + 1, ref false);

//			return;
//			Error_Message:
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Create_Year Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{
//			string MONCHK = null;
//			string ENDCHK = null;

//			switch (oUID) {
//				case "Col1":
//					oMat1.FlushToDataSource();

//					oDS_ZPY509L.Offset = oRow - 1;

//					if (oRow == oMat1.RowCount & !string.IsNullOrEmpty(Strings.Trim(oDS_ZPY509L.GetValue("U_JOBYER", oRow - 1)))) {
//						Matrix_AddRow(oRow);
//						oMat1.Columns.Item("Col1").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					}
//					oMat1.LoadFromDataSource();
//					break;
//				case "Col2":
//				case "Col5":

//					oForm.Freeze(true);
//					oMat1.FlushToDataSource();

//					MONCHK = Strings.Trim(oDS_ZPY509L.GetValue("U_MONCHK", oRow - 1));
//					ENDCHK = Strings.Trim(oDS_ZPY509L.GetValue("U_ENDCHK", oRow - 1));

//					oDS_ZPY509L.Offset = oRow - 1;
//					//// 정산자료가 Y이면 월별자료는 자동으로 Y로 변경
//					//// 월별자료 또는 정산자료에 Y를 체크하면 자동으로 사용자와 수정일자를 현재기준으로 변경
//					if (ENDCHK == "Y") {
//						oDS_ZPY509L.SetValue("U_MONCHK", oRow - 1, "Y");
//					}
//					oDS_ZPY509L.SetValue("U_LGNADM", oRow - 1, MDC_Globals.oCompany.UserName);
//					oDS_ZPY509L.SetValue("U_MODDAT", oRow - 1, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD"));
//					oMat1.LoadFromDataSource();
//					oForm.Freeze(false);
//					break;

//			}

//		}

////---------------------------------------------------------------------------------------
//// Procedure : Batch_EndCheck
//// DateTime  : 2009-02-16
//// Author    :
//// Purpose   : 마감작업 처리
////---------------------------------------------------------------------------------------
////
//		private void Batch_EndCheck()
//		{
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			int iRow = 0;

//			string CLTCOD = null;
//			string JOBYER = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			CLTCOD = Strings.Trim(oDS_ZPY509H.GetValue("U_CLTCOD", 0));

//			MDC_Globals.oCompany.StartTransaction();
//			for (iRow = 0; iRow <= oMat1.VisualRowCount - 1; iRow++) {
//				if (Strings.Trim(oDS_ZPY509L.GetValue("U_MODDAT", iRow)) == Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD")) {
//					JOBYER = Strings.Trim(oDS_ZPY509L.GetValue("U_JOBYER", iRow));
//					if (Strings.Trim(oDS_ZPY509L.GetValue("U_ENDCHK", iRow)) == "Y") {
//						sQry = "EXEC ZPY509 '" + CLTCOD + "', '" + JOBYER + "', '1'";
//						oRecordSet.DoQuery(sQry);
//					} else if (Strings.Trim(oDS_ZPY509L.GetValue("U_MONCHK", iRow)) == "Y") {
//						sQry = "EXEC ZPY509 '" + CLTCOD + "', '" + JOBYER + "', '2'";
//						oRecordSet.DoQuery(sQry);
//					} else {
//						sQry = "EXEC ZPY509 '" + CLTCOD + "', '" + JOBYER + "', '3'";
//						oRecordSet.DoQuery(sQry);
//					}
//				}
//			}
//			MDC_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Error_Message:
//			if (MDC_Globals.oCompany.InTransaction) {
//				MDC_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Batch_EndCheck Error : " + Strings.Space(5) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

//		}
//	}
//}
