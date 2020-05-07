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
//	[System.Runtime.InteropServices.ProgId("ZPY501_NET.ZPY501")]
//	public class ZPY501
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY501.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 소득공제항목 등록
//////  FormType       : 2010110501
//////  Create Date    : 2006.01.19
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//			//시스템코드 헤더
//		private SAPbouiCOM.DBDataSource oDS_ZPY501H;
//			//시스템코드 라인
//		private SAPbouiCOM.DBDataSource oDS_ZPY501L;
//		private SAPbouiCOM.Matrix oMat1;
//		private SAPbouiCOM.CheckBox oCheck;
//		private bool MsterChk;
//		private string oOLDCHK;
//			/// 종전전액보험료
//		private double BOHAL1;
//			/// 주현전액보험료
//		private double BOHAL2;


//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string Col_Last_Uid;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int Col_Last_Row;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm(ref string JSNYER = "", ref string MSTCOD = "", ref string CLTCOD = "")
//		{
//			//Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY501.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY501_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY501");
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
//			oForm.Items.Item("Folder1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

//			oForm.EnableMenu(("1293"), true);
//			/// 행삭제
//			oForm.EnableMenu(("1284"), false);
//			/// 취소

//			if (!string.IsNullOrEmpty(Strings.Trim(JSNYER))) {
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
////// RaiseDataEvent
////*******************************************************************
//		public void RaiseDataEvent(SAPbouiCOM.IBusinessObjectInfo pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (oForm == null) {
//				return;
//			}

//			switch (pval.EventType) {
//				//et_FORM_DATA_ADD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//					if (pval.BeforeAction == false) {
//						if (pval.ActionSuccess == true) {
//							if (MsterChk == true) {
//								MasterUpdate();
//							}
//						}
//					}
//					break;
//				//et_FORM_DATA_UPDATE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//					if (pval.BeforeAction == false) {
//						if (pval.ActionSuccess == true) {
//							if (MsterChk == true) {
//								MasterUpdate();
//							}
//						}
//					}
//					break;
//				//et_FORM_DATA_DELETE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//					break;
//				//et_FORM_DATA_LOAD ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//					break;

//			}

//			return;
//			RaiseDataEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			MDC_Globals.Sbo_Application.StatusBar.SetText("RaiseDataEvent_Error:", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//								//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//								////추가및 업데이시에
//								//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//								if (HeaderSpaceLineDel() == true) {
//									if (MatrixSpaceLineDel() == false) {
//										BubbleEvent = false;
//										return;
//									}
//								} else {
//									BubbleEvent = false;
//									return;
//								}
//							}
//						/// ChooseBtn사원리스트
//						} else if (pval.ItemUID == "CBtn1" & oForm.Items.Item("MSTCOD").Enabled == true) {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						} else if (pval.ItemUID == "Folder1") {
//							oForm.PaneLevel = 1;
//						} else if (pval.ItemUID == "Folder2") {
//							oForm.PaneLevel = 2;
//						} else if (pval.ItemUID == "Folder3") {
//							oForm.PaneLevel = 3;
//						/// 가족사항등록 창 띄우기
//						} else if (pval.ItemUID == "Btn1") {
//							CallSource(ref 0);
//						}
//					} else {
//						if (pval.ItemUID == "1" & pval.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							MDC_Globals.Sbo_Application.ActivateMenuItem("1282");
//							oForm.Items.Item("Folder1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						} else if (pval.ItemUID == "1" & pval.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//							FormItemEnabled();
//						/// 기존부양가족 정보가져오기
//						} else if (pval.ItemUID == "Btn2") {
//							//UPGRADE_WARNING: oForm.Items(JSNYER).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(oForm.Items.Item("JSNYER").Specific.VALUE)) {
//								//                    If oDS_ZPY501H.GetValue("U_JSNYER", 0) = "" Then
//								oForm.Items.Item("JSNYER").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도는 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							} else if (string.IsNullOrEmpty(oDS_ZPY501H.GetValue("U_CLTCOD", 0))) {
//								MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							} else {
//								Family_Data_Display();
//							}
//						/// 부양가족명세 공제항목에 적용
//						} else if (pval.ItemUID == "Btn3") {
//							Family_Total(ref "All");
//						/// 연금.저축 명세 가져오기
//						} else if (pval.ItemUID == "Btn4") {
//							SavingAmount_Data_Display();
//						}
//					}
//					break;
//				//et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					if (pval.BeforeAction == true & pval.ItemUID != "2" & oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//						if (Last_Item == "MSTCOD") {
//							//UPGRADE_WARNING: oForm.Items(Last_Item).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + oForm.Items.Item(Last_Item).Specific.String + "'", ref "") == true & !string.IsNullOrEmpty(oForm.Items.Item(Last_Item).Specific.String) & Last_Item != pval.ItemUID) {
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					}
//					if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == true & Last_Item == "Mat1" & Col_Last_Uid == "Col1" & Col_Last_Row > 0 & (Col_Last_Uid != pval.ColUID | Col_Last_Row != pval.Row) & pval.ItemUID != "2") {
//						if (Col_Last_Row > oMat1.VisualRowCount) {
//							return;
//						}
//					} else if (pval.BeforeAction == false & pval.ItemUID == "Mat1" & pval.Row > 0 & pval.ColUID == "Col0") {
//						oForm.DataSources.UserDataSources.Item("FAMNAM").ValueEx = oDS_ZPY501L.GetValue("U_FamNam", pval.Row - 1);
//						oForm.DataSources.UserDataSources.Item("FAMPER").ValueEx = oDS_ZPY501L.GetValue("U_FamPer", pval.Row - 1);
//						oForm.Items.Item("FAMNAM").Update();
//						oForm.Items.Item("FAMPER").Update();
//					} else if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == false & pval.ItemUID == "Mat1" & (pval.ColUID == "Col5" | pval.ColUID == "Col6" | pval.ColUID == "Col7" | pval.ColUID == "Col8" | pval.ColUID == "Col9" | pval.ColUID == "Col10") & pval.Row > 0) {
//						FlushToItemValue(pval.ColUID, ref pval.Row);
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUID == "MSTCOD" | pval.ItemUID == "JSNYER" | pval.ItemUID == "BAEWOO" | pval.ItemUID == "BUYN20" | pval.ItemUID == "BUYN60" | pval.ItemUID == "GYNGLO" | pval.ItemUID == "GYNGL2" | pval.ItemUID == "JANGAE" | pval.ItemUID == "MZBURI" | pval.ItemUID == "BUYN06" | pval.ItemUID == "BONSCH" | pval.ItemUID == "CSHSAV" | pval.ItemUID == "LAWGBU" | pval.ItemUID == "POCGBU" | pval.ItemUID == "SP1GBU" | pval.ItemUID == "SP2GBU" | pval.ItemUID == "USJGBU" | pval.ItemUID == "JIJGBU" | pval.ItemUID == "BONCSH1" | pval.ItemUID == "BONEDC1")) {
//						FlushToItemValue(pval.ItemUID);
//					} else if (pval.BeforeAction == false & pval.ItemChanged == true & pval.ItemUID == "Mat1" & (pval.ColUID == "Col1" | pval.ColUID == "Col2" | pval.ColUID == "Col5" | pval.ColUID == "Col6" | pval.ColUID == "Col7" | pval.ColUID == "Col8" | pval.ColUID == "Col9" | pval.ColUID == "Col23" | pval.ColUID == "Col10")) {
//						FlushToItemValue(pval.ColUID, ref pval.Row);
//					}
//					break;
//				//et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if (pval.ItemUID == "MZBURI" | pval.ItemUID == "BAEWOO" | pval.ItemUID == "HUSMAN") {
//							MsterChk = true;
//						} else if (pval.ItemUID == "Mat1" & (pval.ColUID == "Col3" | pval.ColUID == "Col4") & pval.Row > 0) {
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
//							//                    If oMat1.RowCount > 0 Then
//							//                        oMat1.Columns("Col1").Cells(oMat1.VisualRowCount).Click ct_Regular
//							//                        BubbleEvent = False
//							//                    End If
//						}
//					} else if (pval.BeforeAction == true & pval.ItemUID == "JSNYER" & pval.CharPressed == 9 & pval.FormMode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(pval.ItemUID).Specific.String))) {
//							oForm.Items.Item("JSNYER").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도는 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
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
//						//UPGRADE_NOTE: oDS_ZPY501H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY501H = null;
//						//UPGRADE_NOTE: oDS_ZPY501L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY501L = null;
//						//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oCheck = null;
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
//				//et_FORM_RESIZE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					if (pval.BeforeAction == false) {
//						oForm.Freeze(true);
//						//oForm.Items("Reg1").Width = oForm.Width - 26
//						//oForm.Items("Reg1").Height = oForm.Height - 358
//						oMat1.AutoResizeColumns();
//						oForm.Freeze(false);
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
//						if (Strings.Trim(oDS_ZPY501H.GetValue("U_Check", 0)) == "Y") {
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
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@ZPY501H", ref "DocNum");
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
//					// oForm.Items("Btn1").Visible = True
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
//						FormItemEnabled();
//						break;
//					case "1293":
//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//							////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							for (i = 0; i <= oMat1.VisualRowCount - 1; i++) {
//								//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat1.Columns.Item("Col0").Cells.Item(i + 1).Specific.VALUE = i + 1;
//							}
//							oMat1.FlushToDataSource();
//							oDS_ZPY501L.RemoveRecord(oDS_ZPY501L.Size - 1);
//							//// Mat1에 마지막라인(빈라인) 삭제
//							oMat1.Clear();
//							oMat1.LoadFromDataSource();
//							Family_Total(ref "All");
//						}
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

//		private void CallSource(ref int oRow)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			object oTmpObject = null;

//			//    If oForm.Mode <> fm_FIND_MODE Then
//			//        Set oTmpObject = New ZPY121     'ZPY121(2010110121) : 가족사항등록
//			//        Call oTmpObject.LoadForm(oForm.uniqueID, oForm.Items("MSTCOD").Specific.String, oForm.Items("JSNYER").Specific.String)
//			//        Sbo_Application.Forms.ActiveForm.Select
//			//    End If
//			if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//				oTmpObject = new PH_PY001();
//				//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oTmpObject.LoadForm();
//				MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//			}
//			//UPGRADE_NOTE: oTmpObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oTmpObject = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oTmpObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oTmpObject = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CallSource Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
////*******************************************************************
////
////*******************************************************************
//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			SAPbouiCOM.Folder oFolder = null;
//			SAPbouiCOM.Column oColumn = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////디비데이터 소스 개체 할당
//			oDS_ZPY501H = oForm.DataSources.DBDataSources("@ZPY501H");
//			oDS_ZPY501L = oForm.DataSources.DBDataSources("@ZPY501L");
//			/// UserDataSource
//			oForm.DataSources.UserDataSources.Add("FAMNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			oForm.DataSources.UserDataSources.Add("FAMPER", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 40);
//			oEdit = oForm.Items.Item("FAMNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "FAMNAM");
//			oEdit = oForm.Items.Item("FAMPER").Specific;
//			oEdit.DataBind.SetBound(true, "", "FAMPER");

//			oForm.PaneLevel = 1;

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			/// Folder /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//			//    oForm.DataSources.UserDataSources.Add "FolderDS", dt_SHORT_TEXT, 1
//			//    For i = 1 To 2
//			//        Set oFolder = oForm.Items("Folder" & i).Specific
//			//        oFolder.DataBind.SetBound True, "", "FolderDS"
//			//        If i = 1 Then
//			//            oFolder.Select
//			//        Else
//			//            oFolder.GroupWith ("Folder" & i - 1)
//			//        End If
//			//    Next i
//			oForm.Items.Item("Folder1").AffectsFormMode = false;
//			oForm.Items.Item("Folder2").AffectsFormMode = false;
//			oForm.Items.Item("Folder3").AffectsFormMode = false;
//			oForm.Items.Item("Folder1").Enabled = true;
//			oForm.Items.Item("Folder2").Enabled = true;
//			oForm.Items.Item("Folder3").Enabled = true;

//			/// Check 버튼
//			oCheck = oForm.Items.Item("Check1").Specific;
//			oCheck.ValOn = "Y";
//			oCheck.ValOff = "N";
//			//    Set oCheck = oForm.Items("BONJAN").Specific
//			//    oCheck.ValOn = "1": oCheck.ValOff = "0"

//			oCombo = oForm.Items.Item("BAEWOO").Specific;
//			oCombo.ValidValues.Add("1", "YES");
//			oCombo.ValidValues.Add("0", "NO");
//			oCombo = oForm.Items.Item("MZBURI").Specific;
//			oCombo.ValidValues.Add("1", "YES");
//			oCombo.ValidValues.Add("0", "NO");
//			oCombo = oForm.Items.Item("HUSMAN").Specific;
//			oCombo.ValidValues.Add("1", "세대주");
//			oCombo.ValidValues.Add("2", "세대원");

//			//// 사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT Code, Name FROM [@PH_PY005A]";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;


//			//// 교육대상구분
//			oColumn = oMat1.Columns.Item("Col16");
//			oColumn.DisplayDesc = true;
//			oColumn.ValidValues.Add("0", "본인");
//			oColumn.ValidValues.Add("1", "영유아");
//			oColumn.ValidValues.Add("2", "초중고");
//			oColumn.ValidValues.Add("3", "대학생");
//			oColumn.ValidValues.Add("4", "대학원생");
//			oColumn.ValidValues.Add("5", "특수교육");
//			//// 관계코드
//			oColumn = oMat1.Columns.Item("Col3");
//			oColumn.ValidValues.Add("0", "본인");
//			oColumn.ValidValues.Add("1", "소득자의 직계존속");
//			oColumn.ValidValues.Add("2", "배우자의 직계존속");
//			oColumn.ValidValues.Add("3", "배우자");
//			oColumn.ValidValues.Add("4", "직계비속 자녀.입양자");
//			//    oColumn.ValidValues.Add "5", "형제자매"
//			//    oColumn.ValidValues.Add "6", "기타"
//			/// (2010.01.19) 2009년귀속 관계코드변경
//			oColumn.ValidValues.Add("5", "직계비속 자녀.입양자외");
//			oColumn.ValidValues.Add("6", "형제자매");
//			//oColumn.ValidValues.Add "7", "기타"
//			/// (2010.01.04) 2010년귀속 관계코드변경
//			oColumn.ValidValues.Add("7", "수급자");
//			oColumn.ValidValues.Add("8", "위탁아동");

//			//// 내외국인
//			oColumn = oMat1.Columns.Item("Col4");
//			oColumn.ValidValues.Add("1", "내국인");
//			oColumn.ValidValues.Add("9", "외국인");
//			//// 보험료구분
//			oColumn = oMat1.Columns.Item("Col11");
//			oColumn.ValidValues.Add("1", "일반인보험");
//			oColumn.ValidValues.Add("2", "장애인보험");
//			//// 부양가족 인적공제내용
//			oColumn = oMat1.Columns.Item("Col5");
//			/// 기본
//			oColumn.ValOn = "1";
//			oColumn.ValOff = "0";
//			oColumn = oMat1.Columns.Item("Col6");
//			/// 부녀자
//			oColumn.ValOn = "1";
//			oColumn.ValOff = "0";
//			oColumn = oMat1.Columns.Item("Col7");
//			/// 장애인
//			oColumn.ValOn = "1";
//			oColumn.ValOff = "0";
//			oColumn = oMat1.Columns.Item("Col8");
//			/// 경로우대
//			oColumn.ValOn = "1";
//			oColumn.ValOff = "0";
//			oColumn = oMat1.Columns.Item("Col9");
//			/// 양육
//			oColumn.ValOn = "1";
//			oColumn.ValOff = "0";
//			oColumn = oMat1.Columns.Item("Col10");
//			/// 다자녀
//			oColumn.ValOn = "1";
//			oColumn.ValOff = "0";
//			oColumn.Editable = false;
//			oColumn = oMat1.Columns.Item("Col23");
//			/// 출산입양
//			oColumn.ValOn = "1";
//			oColumn.ValOff = "0";

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oFolder 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oFolder = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oFolder 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oFolder = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private string Exist_YN(ref string JOBYER, ref string MSTCOD, ref string CLTCOD)
//		{
//			string functionReturnValue = null;
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//저장할 데이터의 기존데이터가 있는지 확인한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "SELECT Top 1 T1.DocNum FROM [@ZPY501H] T1 ";
//			sQry = sQry + " WHERE T1.U_JSNYER = N'" + Strings.Trim(JOBYER) + "'";
//			sQry = sQry + "   AND T1.U_MSTCOD = N'" + Strings.Trim(MSTCOD) + "'";
//			sQry = sQry + "   AND T1.U_CLTCOD = N'" + Strings.Trim(CLTCOD) + "'";
//			oRecordSet.DoQuery(sQry);

//			while (!(oRecordSet.EoF)) {
//				//UPGRADE_WARNING: oRecordSet().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = oRecordSet.Fields.Item(0).Value;
//				oRecordSet.MoveNext();
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(Exist_YN()))) {
//				functionReturnValue = "";
//				return functionReturnValue;
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);
//			string INTGBN = null;

//			switch (oUID) {
//				case "JSNYER":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oUID).Specific.String))) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MDC_Globals.ZPAY_GBL_JSNYER.Value = oForm.Items.Item(oUID).Specific.String;
//					} else {
//						oDS_ZPY501H.SetValue("U_JSNYER", 0, MDC_Globals.ZPAY_GBL_JSNYER.Value);
//					}
//					oForm.Items.Item(oUID).Update();
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oDS_ZPY501H.SetValue("U_MSTCOD", 0, "");
//						oDS_ZPY501H.SetValue("U_MSTNAM", 0, "");
//						oDS_ZPY501H.SetValue("U_EmpID", 0, "");
//						oDS_ZPY501H.SetValue("U_CLTCOD", 0, "");
//					} else {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oDS_ZPY501H.SetValue("U_MSTCOD", 0, Strings.UCase(oForm.Items.Item(oUID).Specific.String));
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMast = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oDS_ZPY501H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//						oDS_ZPY501H.SetValue("U_EmpID", 0, oMast.EmpID);
//						oDS_ZPY501H.SetValue("U_CLTCOD", 0, oMast.CLTCOD);
//					}
//					/// 추가모드일경우 인적공제 뿌려줌
//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//						/// 사원마스터의 부양가족수
//						oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//						sQry = "SELECT U_BAEWOO, U_BUYNSU, U_BUYN20, U_BUYN60, U_GYNGLO, U_GYNGL2, U_JANGAE, U_MZBURI, U_BUYN06, U_DAGYSU, U_INTGBN ";
//						sQry = sQry + " FROM [@PH_PY001A]  WHERE Code = N'" + oDS_ZPY501H.GetValue("U_MSTCOD", 0) + "'";
//						oRecordSet.DoQuery(sQry);
//						if (oRecordSet.RecordCount > 0) {
//							oDS_ZPY501H.SetValue("U_BAEWOO", 0, oRecordSet.Fields.Item("U_BAEWOO").Value);
//							oDS_ZPY501H.SetValue("U_BUYN20", 0, oRecordSet.Fields.Item("U_BUYN20").Value);
//							oDS_ZPY501H.SetValue("U_BUYN60", 0, oRecordSet.Fields.Item("U_BUYN60").Value);
//							oDS_ZPY501H.SetValue("U_GYNGLO", 0, oRecordSet.Fields.Item("U_GYNGLO").Value);
//							oDS_ZPY501H.SetValue("U_GYNGL2", 0, oRecordSet.Fields.Item("U_GYNGL2").Value);
//							oDS_ZPY501H.SetValue("U_JANGAE", 0, oRecordSet.Fields.Item("U_JANGAE").Value);
//							oDS_ZPY501H.SetValue("U_MZBURI", 0, oRecordSet.Fields.Item("U_MZBURI").Value);
//							oDS_ZPY501H.SetValue("U_BUYN06", 0, oRecordSet.Fields.Item("U_BUYN06").Value);
//							oDS_ZPY501H.SetValue("U_DAGYSU", 0, oRecordSet.Fields.Item("U_DAGYSU").Value);
//							oDS_ZPY501H.SetValue("U_CHLSAN", 0, Convert.ToString(0));
//							oForm.Items.Item("BAEWOO").Update();
//							oForm.Items.Item("BUYN20").Update();
//							oForm.Items.Item("BUYN60").Update();
//							oForm.Items.Item("GYNGLO").Update();
//							oForm.Items.Item("GYNGL2").Update();
//							oForm.Items.Item("JANGAE").Update();
//							oForm.Items.Item("MZBURI").Update();
//							oForm.Items.Item("BUYN06").Update();
//							oForm.Items.Item("DAGYSU").Update();
//							oForm.Items.Item("CHLSAN").Update();
//							//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							INTGBN = oRecordSet.Fields.Item("U_INTGBN").Value;
//							MsterChk = false;
//						}
//					}
//					oForm.Items.Item("MSTNAM").Update();
//					oForm.Items.Item("EmpID").Update();
//					oForm.Items.Item(oUID).Update();
//					break;
//				case "BAEWOO":
//				case "BUYN20":
//				case "GYNGLO":
//				case "GYNGL2":
//				case "JANGAE":
//				case "MZBURI":
//				case "BUYN06":
//				case "HUSMAN":
//					MsterChk = true;
//					break;
//			}
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			if (Strings.Left(oUID, 3) == "Col") {
//				oMat1.FlushToDataSource();
//				oDS_ZPY501L.Offset = oRow - 1;
//				switch (oUID) {
//					case "Col1":
//					case "Col2":
//					case "Col3":
//					case "Col4":
//					case "Col5":
//					case "Col6":
//					case "Col7":
//					case "Col8":
//					case "Col9":
//					case "Col10":
//					case "Col23":
//						MsterChk = true;
//						break;
//				}
//				oDS_ZPY501L.Offset = oRow - 1;
//				///
//				if (oRow == oMat1.VisualRowCount & !string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", oRow - 1)))) {
//					Matrix_AddRow(oRow);
//					oMat1.Columns.Item("Col1").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				}
//			}
//		}

//		private void FormClear()
//		{
//			int DocNum = 0;

//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocNum = MDC_SetMod.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'ZPY501'", ref "");

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
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				if (Strings.Len(MDC_Globals.ZPAY_GBL_JSNYER.Value) > 0) {
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("JSNYER").Specific.VALUE = MDC_Globals.ZPAY_GBL_JSNYER.Value;
//				}
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = true;
//				oForm.Items.Item("MSTCOD").Enabled = true;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = false;
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				if (Strings.Len(MDC_Globals.ZPAY_GBL_JSNYER.Value) > 0) {
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("JSNYER").Specific.VALUE = MDC_Globals.ZPAY_GBL_JSNYER.Value;
//				}
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = false;
//				oForm.Items.Item("MSTCOD").Enabled = false;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = false;
//				oForm.Items.Item("CLTCOD").Enabled = false;
//				oOLDCHK = oDS_ZPY501H.GetValue("U_Check", 0);

//			}
//			oForm.Items.Item("Folder1").Enabled = true;
//			oForm.Items.Item("Folder2").Enabled = true;
//			oForm.Items.Item("Folder3").Enabled = true;

//		}

//		private bool HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string DocNum = null;

//			ErrNum = 0;

//			Family_Total(ref "ErrorCheck");

//			/// Check
//			switch (true) {
//				case string.IsNullOrEmpty(oDS_ZPY501H.GetValue("U_JSNYER", 0)):
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY501H.GetValue("U_MSTCOD", 0)):
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case (Conversion.Val(oDS_ZPY501H.GetValue("U_BOHAMT", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JGABOA", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BOHAL1", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BOHAL2", 0))) != (Conversion.Val(oDS_ZPY501H.GetValue("U_BONBOH1", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BONBOH2", 0))):
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//				case (Conversion.Val(oDS_ZPY501H.GetValue("U_JGAMED", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_GENMED", 0))) != (Conversion.Val(oDS_ZPY501H.GetValue("U_BONMED1", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BONMED2", 0))):
//					ErrNum = 4;
//					goto Error_Message;
//					break;
//				case (Conversion.Val(oDS_ZPY501H.GetValue("U_BONSCH", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JGASCH", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JICSCH", 0))) > (Conversion.Val(oDS_ZPY501H.GetValue("U_BONEDC1", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BONEDC2", 0))):
//					ErrNum = 5;
//					goto Error_Message;
//					break;
//				case (Conversion.Val(oDS_ZPY501H.GetValue("U_CADSAV", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_GIRSAV", 0))) != (Conversion.Val(oDS_ZPY501H.GetValue("U_BONCAD1", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BONCAD2", 0))):
//					ErrNum = 6;
//					goto Error_Message;
//					break;
//				case Conversion.Val(oDS_ZPY501H.GetValue("U_CA1SAV", 0)) != (Conversion.Val(oDS_ZPY501H.GetValue("U_BONCA11", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BONCA12", 0))):
//					ErrNum = 13;
//					goto Error_Message;
//					break;
//				case Conversion.Val(oDS_ZPY501H.GetValue("U_CSHSAV", 0)) != Conversion.Val(oDS_ZPY501H.GetValue("U_BONCSH1", 0)):
//					ErrNum = 7;
//					goto Error_Message;
//					break;
//				case (Conversion.Val(oDS_ZPY501H.GetValue("U_LAWGBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_POCGBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_SP1GBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_SP2GBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_USJGBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JIJGBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JI2GBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JNGGBU", 0))) != (Conversion.Val(oDS_ZPY501H.GetValue("U_BONGBU1", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BONGBU2", 0))):
//					ErrNum = 8;
//					goto Error_Message;
//					break;
//				case Strings.Trim(oDS_ZPY501H.GetValue("U_Check", 0)) == "Y" & Strings.Trim(oOLDCHK) == "Y":
//					ErrNum = 9;
//					goto Error_Message;
//					break;
//				case Conversion.Val(oDS_ZPY501H.GetValue("U_JHEJA1", 0)) != 0 & Conversion.Val(oDS_ZPY501H.GetValue("U_JHEJA2", 0)) != 0:
//					ErrNum = 10;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY501H.GetValue("U_CLTCOD", 0)):
//					ErrNum = 11;
//					goto Error_Message;
//					break;
//			}

//			DocNum = Exist_YN(ref oDS_ZPY501H.GetValue("U_JSNYER", 0), ref oDS_ZPY501H.GetValue("U_MSTCOD", 0), ref oDS_ZPY501H.GetValue("U_CLTCOD", 0));
//			if (!string.IsNullOrEmpty(Strings.Trim(DocNum)) & Strings.Trim(oDS_ZPY501H.GetValue("DocNum", 0)) != Strings.Trim(DocNum)) {
//				//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//				//같은데이터가 존재하는데 자기 자신이 현재 자기자신이 아니라면(같은월에는 취소한거 아니면 하나만 존재해야함)
//				//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//				ErrNum = 12;
//				goto Error_Message;
//			}

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 연도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("보험료금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("의료비금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("교육비금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("신용카드금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("현금영수증금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 8) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부금금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 9) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 10) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("장기주택저당차입금이자상환액 600만원한도금액 또는 1000만원 한도 금액 중 하나만 입력가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 11) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 12) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("문서번호" + DocNum + " 와(과) 데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 13) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("직불카드 금액이 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void MasterUpdate()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;

//			/// Question
//			if (MDC_Globals.Sbo_Application.MessageBox("변경된 인적공제내역을 사원마스터-급여기본사항에 반영하시겠습니까?", 2, "&Yes!", "&No") == 2) {
//				return;
//			}

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "UPDATE [@PH_PY001A] SET U_BAEWOO = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_BAEWOO", 0)) + "'";
//			sQry = sQry + "         , U_BUYN20 = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN20", 0)) + "'";
//			sQry = sQry + "         , U_BUYN60 = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN60", 0)) + "'";
//			sQry = sQry + "         , U_BUYNSU = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN20", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN60", 0)) + "'";
//			sQry = sQry + "         , U_GYNGLO = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_GYNGLO", 0)) + "'";
//			sQry = sQry + "         , U_GYNGL2 = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_GYNGL2", 0)) + "'";
//			sQry = sQry + "         , U_JANGAE = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_JANGAE", 0)) + "'";
//			sQry = sQry + "         , U_MZBURI = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_MZBURI", 0)) + "'";
//			sQry = sQry + "         , U_BUYN06 = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN06", 0)) + "'";
//			sQry = sQry + "         , U_DAGYSU = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_DAGYSU", 0)) + "'";
//			sQry = sQry + "         , U_HUSMAN = '" + Conversion.Val(oDS_ZPY501H.GetValue("U_HUSMAN", 0)) + "'";
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + "   WHERE Code = N'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'";
//			oRecordSet.DoQuery(sQry);
//			/// 본인주민번호도 수정
//			if (!string.IsNullOrEmpty(oDS_ZPY501L.GetValue("U_FamPer", 0)) & Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", 0)) == "0") {
//				/// 인사마스터
//				sQry = "UPDATE [@PH_PY001A] SET GOVID = '" + oDS_ZPY501L.GetValue("U_FamPer", 0) + "'";
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "   WHERE Code = N'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'";
//				oRecordSet.DoQuery(sQry);
//				/// 급여기본등록
//				sQry = "UPDATE [@PH_PY001A] SET  U_INTGBN = '" + oDS_ZPY501L.GetValue("U_ChkInt", 0) + "'";
//				/// 내외국인
//				sQry = sQry + "         ,      U_BJNGAE = '" + oDS_ZPY501L.GetValue("U_ChkJan", 0) + "'";
//				/// 본인장애유무
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "   WHERE Code = N'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'";
//				oRecordSet.DoQuery(sQry);
//			}

//			/// 가족사항등록화면에 부양가족정보 적용
//			if (MDC_Globals.Sbo_Application.MessageBox("변경된 인적공제내역을 가족사항등록에 반영하시겠습니까?", 2, "&Yes!", "&No") == 1) {
//				sQry = "EXEC ZPY501_1 '" + oDS_ZPY501H.GetValue("U_JSNYER", 0) + "', '" + oDS_ZPY501H.GetValue("U_CLTCOD", 0) + "', '" + oDS_ZPY501H.GetValue("U_MSTCOD", 0) + "', '" + MDC_Globals.oCompany.UserSignature + "'";
//				oRecordSet.DoQuery(sQry);
//			}

//			/// 메세지
//			// SBO_Application.StatusBar.SetText "마스터에 적용하였습니다.", bmt_Short, smt_Success
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("MasterUpdate Error:" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

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

//			short U_BAEWOO = 0;
//			short U_GYNGLO = 0;
//			short U_DAGYSU = 0;
//			short U_MZBURI = 0;
//			short U_JANGAE = 0;
//			short U_BUYN06 = 0;
//			short U_INJTOT = 0;
//			string GovidChk = null;
//			string CLTCOD = null;

//			U_BAEWOO = 0;
//			U_GYNGLO = 0;
//			U_JANGAE = 0;
//			U_MZBURI = 0;
//			U_BUYN06 = 0;
//			U_DAGYSU = 0;
//			U_INJTOT = 0;

//			ErrNum = 0;
//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = MDC_SetMod.Get_ReData(ref "ISNULL(U_CLTCOD,'')", ref "Code", ref "[@PH_PY001A]", ref "'" + oDS_ZPY501H.GetValue("U_MSTCOD", 0) + "'", ref "");
//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			GovidChk = MDC_SetMod.Get_ReData(ref "ISNULL(U_GovIDChk,'N')", ref "Code", ref "[@PH_PY005A]", ref "'" + CLTCOD + "'", ref "");

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oMat1.FlushToDataSource();

//			//// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
//			if (oMat1.RowCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//				return functionReturnValue;
//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
//			////oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
//			for (i = 0; i <= oMat1.VisualRowCount - 1; i++) {
//				oDS_ZPY501L.Offset = i;
//				//// 가족성명이 있을 경우만
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", i)))) {
//					if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkBas", i)) == "1") {
//						U_INJTOT = U_INJTOT + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkBas", i));
//						U_BAEWOO = U_BAEWOO + (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", i)) == "3" ? 1 : 0);
//						U_GYNGLO = U_GYNGLO + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkJeL", i));
//						U_MZBURI = U_MZBURI + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkbuY", i));
//						U_DAGYSU = U_DAGYSU + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkDaJ", i));
//					}
//					U_JANGAE = U_JANGAE + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkJan", i));
//					U_BUYN06 = U_BUYN06 + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkChl", i));

//					if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamPer", i)))) {
//						ErrNum = 5;
//						oMat1.Columns.Item("Col2").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						goto Error_Message;
//					} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", i)))) {
//						ErrNum = 6;
//						oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						goto Error_Message;
//					} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_ChkInt", i)))) {
//						ErrNum = 7;
//						oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						goto Error_Message;
//						/// 교육비공제는 기본공제대상자가 아니여도 가능함.(기본공제대상자에서 나이제한으로 기본공제대상자가 안될경우도 포함됨. 20세이상자녀대학교)
//						//        ElseIf Trim$(oDS_ZPY501L.GetValue("U_ChkBas", i)) <> "1" And (Val(oDS_ZPY501L.GetValue("U_EDCAMT1", i)) + Val(oDS_ZPY501L.GetValue("U_EDCAMT1", i))) > 0 Then
//						//            ErrNum = 15
//						//            oMat1.Columns("Col1").Cells(i + 1).Click ct_Regular
//						//            GoTo Error_Message

//					} else if ((Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT1", i)) > 0 | Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT2", i)) > 0) & string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_SCHGBN", i)))) {
//						ErrNum = 17;
//						oMat1.Columns.Item("Col17").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						goto Error_Message;
//					} else {
//						//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//						//중복체크작업
//						//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//						Chk_Data = Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", i));
//						for (k = i + 1; k <= oMat1.VisualRowCount - 1; k++) {
//							oDS_ZPY501L.Offset = k;
//							if (Strings.Trim(Chk_Data) == Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", k))) {
//								ErrNum = 3;
//								oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								goto Error_Message;
//							}
//						}
//					}
//					//// 6.주민번호 오류 체크
//					if (Strings.Trim(GovidChk) == "Y" & Strings.Len(oDS_ZPY501L.GetValue("U_FamPer", i)) > 0) {
//						if (MDC_Com.GovIDCheck(ref oDS_ZPY501L.GetValue("U_FamPer", i)) == false) {
//							ErrNum = 16;
//							oMat1.Columns.Item("Col2").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							goto Error_Message;
//						}
//					}

//					//// 가족성명이 없을경우
//				} else {
//					/// 주민등록번호는 있는데 가족성명이 안들어오는 경우
//					if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", i))) & !string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamPer", i)))) {
//						ErrNum = 4;
//						oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						goto Error_Message;
//					}
//				}

//			}
//			///
//			switch (true) {
//				case U_INJTOT != (Conversion.Val(oDS_ZPY501H.GetValue("U_BAEWOO", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN20", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN60", 0)) + 1):
//					ErrNum = 8;
//					goto Error_Message;
//					break;
//				case U_BAEWOO != Conversion.Val(oDS_ZPY501H.GetValue("U_BAEWOO", 0)):
//					ErrNum = 9;
//					goto Error_Message;
//					break;
//				case U_GYNGLO != (Conversion.Val(oDS_ZPY501H.GetValue("U_GYNGLO", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_GYNGL2", 0))):
//					ErrNum = 10;
//					goto Error_Message;
//					break;
//				case U_JANGAE != Conversion.Val(oDS_ZPY501H.GetValue("U_JANGAE", 0)):
//					ErrNum = 11;
//					goto Error_Message;
//					break;
//				case U_MZBURI != Conversion.Val(oDS_ZPY501H.GetValue("U_MZBURI", 0)):
//					ErrNum = 12;
//					goto Error_Message;
//					break;
//				case U_BUYN06 != Conversion.Val(oDS_ZPY501H.GetValue("U_BUYN06", 0)):
//					ErrNum = 13;
//					goto Error_Message;
//					break;
//				case U_DAGYSU != Conversion.Val(oDS_ZPY501H.GetValue("U_DAGYSU", 0)):
//					ErrNum = 14;
//					goto Error_Message;
//					break;
//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//			////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", oDS_ZPY501L.Size - 1)))) {
//				oDS_ZPY501L.RemoveRecord(oDS_ZPY501L.Size - 1);
//				//// Mat1에 마지막라인(빈라인) 삭제
//			}
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("부양가족명세에 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 연도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("중복된 자료가 있습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("부양가족 성명은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("부양가족 주민번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("부양가족 관계코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("부양가족 내외국인구분은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 8) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("부양가족수가 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 9) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("배우자가 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 10) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("경로우대인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 11) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("장애인인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 12) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("부녀자인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 13) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자녀양육인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 14) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("다자녀공제인원수가 부양가족명세와 다릅니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 15) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("교육비 공제는 기본공제대상자만 가능합니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 16) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("주민등록번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 17) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("교육비가 입력된 경우 교육대상 구분은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Error_Message:" + Err().Number + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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

//			sQry = "SELECT DocEntry FROM [@ZPY501H]";
//			sQry = sQry + "   WHERE U_JSNYER = N'" + JSNYER + "'";
//			sQry = sQry + "   AND   U_MSTCOD = N'" + MSTCOD + "'";
//			sQry = sQry + "   AND   U_CLTCOD = N'" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				while (!(oRecordSet.EoF)) {
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					DocEntry = oRecordSet.Fields.Item(0).Value;
//					oRecordSet.MoveNext();
//				}
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("JSNYER").Specific.VALUE = JSNYER;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("MSTCOD").Specific.String = MSTCOD;
//				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("CLTCOD").Specific.Select(CLTCOD, SAPbouiCOM.BoSearchKey.psk_ByValue);
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.VALUE = DocEntry;

//				oForm.Items.Item("DocNum").Update();
//				oMat1.LoadFromDataSource();
//				oForm.Update();
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//				MDC_Globals.Sbo_Application.ActivateMenuItem("1282");

//				oDS_ZPY501H.SetValue("U_JSNYER", 0, JSNYER);
//				oDS_ZPY501H.SetValue("U_MSTCOD", 0, MSTCOD);
//				oDS_ZPY501H.SetValue("U_CLTCOD", 0, CLTCOD);
//				//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oMast = MDC_SetMod.Get_EmpID_InFo(ref MSTCOD);
//				oDS_ZPY501H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//				oDS_ZPY501H.SetValue("U_EmpID", 0, oMast.EmpID);

//				oForm.Update();

//				MDC_Globals.Sbo_Application.SendKeys("{TAB}");
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//		}

//		private void SavingAmount_Data_Display()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short ErrNum = 0;
//			string JSNYER = null;
//			string MSTCOD = null;
//			string CLTCOD = null;

//			if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE & oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE & oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//				return;
//			}

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			JSNYER = Strings.Trim(oDS_ZPY501H.GetValue("U_JSNYER", 0));
//			MSTCOD = Strings.Trim(oDS_ZPY501H.GetValue("U_MSTCOD", 0));
//			CLTCOD = Strings.Trim(oDS_ZPY501H.GetValue("U_CLTCOD", 0));

//			if (string.IsNullOrEmpty(JSNYER)) {
//				ErrNum = 2;
//				goto Error_Message;
//			} else if (string.IsNullOrEmpty(MSTCOD)) {
//				ErrNum = 3;
//				goto Error_Message;
//			} else if (string.IsNullOrEmpty(CLTCOD)) {
//				ErrNum = 4;
//				goto Error_Message;
//			}

//			sQry = "EXEC ZPY501_2 '" + JSNYER + "', '" + MSTCOD + "', '" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount == 0) {

//				ErrNum = 1;

//				//2012.02.16 기존에 연금저축명세가 존재 해서 값이 들어온 이후에 연금저축명세를
//				//삭제했을때 값을 초기화 되도록 수정(V12.01.08 버젼 이후로 적용되도록 소스만 수정 해 놓음)

//				oDS_ZPY501H.SetValue("U_RETSAV", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_RETSA1", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_GYNSAV", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_YUNSAV", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_HUSAMT", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_HU1AMT", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_HU2AMT", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_HU3AMT", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_JFDAM1", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_JFDAM2", 0, Convert.ToString(0));
//				oDS_ZPY501H.SetValue("U_JFDAM3", 0, Convert.ToString(0));

//				goto Error_Message;

//			}

//			oDS_ZPY501H.SetValue("U_RETSAV", 0, oRecordSet.Fields.Item("RETSAV").Value);
//			oDS_ZPY501H.SetValue("U_RETSA1", 0, oRecordSet.Fields.Item("RETSA1").Value);
//			oDS_ZPY501H.SetValue("U_GYNSAV", 0, oRecordSet.Fields.Item("GYNSAV").Value);
//			oDS_ZPY501H.SetValue("U_YUNSAV", 0, oRecordSet.Fields.Item("YUNSAV").Value);
//			oDS_ZPY501H.SetValue("U_HUSAMT", 0, oRecordSet.Fields.Item("HUSAMT").Value);
//			oDS_ZPY501H.SetValue("U_HU1AMT", 0, oRecordSet.Fields.Item("HU1AMT").Value);
//			oDS_ZPY501H.SetValue("U_HU2AMT", 0, oRecordSet.Fields.Item("HU2AMT").Value);
//			oDS_ZPY501H.SetValue("U_HU3AMT", 0, oRecordSet.Fields.Item("HU3AMT").Value);
//			oDS_ZPY501H.SetValue("U_JFDAM1", 0, oRecordSet.Fields.Item("JFDAM1").Value);
//			oDS_ZPY501H.SetValue("U_JFDAM2", 0, oRecordSet.Fields.Item("JFDAM2").Value);
//			oDS_ZPY501H.SetValue("U_JFDAM3", 0, oRecordSet.Fields.Item("JFDAM3").Value);

//			oForm.Update();

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Error_Message:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			switch (ErrNum) {
//				case 1:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("가져올 연금,저축 명세내역이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 2:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 3:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("사원 번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				case 4:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("자사 코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//				default:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("SavingAmount_Data_Display Error : " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					break;
//			}

//		}

//		private void Family_Data_Display()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short i = 0;
//			short oRow = 0;
//			short cnt = 0;
//			short JSNYER = 0;
//			double GBUAMT = 0;
//			string MSTCOD = null;
//			string CLTCOD = null;

//			/// Question
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			i = 0;
//			oRow = 0;
//			///
//			cnt = oDS_ZPY501L.Size;
//			if (cnt > 0) {
//				for (i = 0; i <= cnt - 1; i++) {
//					oDS_ZPY501L.RemoveRecord(oDS_ZPY501L.Size - 1);
//				}
//				Matrix_AddRow(0, ref true);
//			} else {
//				oMat1.LoadFromDataSource();
//			}
//			///
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JSNYER = Conversion.Val(oForm.Items.Item("JSNYER").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String);
//			CLTCOD = Strings.Trim(oDS_ZPY501H.GetValue("U_CLTCOD", 0));
//			if (JSNYER <= 0) {
//				return;
//			}

//			i = 0;
//			sQry = "EXEC ZPY501 " + "'" + JSNYER + "', N'" + MSTCOD + "', '" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				if (i + 1 > oDS_ZPY501L.Size) {
//					oDS_ZPY501L.InsertRecord((i));
//				}

//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				BOHAL1 = oRecordSet.Fields.Item("U_JONBOH").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				BOHAL2 = oRecordSet.Fields.Item("U_BOHAMT3").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				GBUAMT = oRecordSet.Fields.Item("U_GBUAMT3").Value;
//				/// 본인일 경우 전액공제보험료표시
//				if (Strings.Trim(oRecordSet.Fields.Item("U_ChkCod").Value) == "0") {
//					oDS_ZPY501H.SetValue("U_BOHAL1", 0, Convert.ToString(BOHAL1));
//					oDS_ZPY501H.SetValue("U_BOHAL2", 0, Convert.ToString(BOHAL2));
//					oDS_ZPY501H.SetValue("U_JI2GBU", 0, Convert.ToString(GBUAMT));
//				}
//				oDS_ZPY501L.Offset = i;
//				oDS_ZPY501L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_ZPY501L.SetValue("U_FamNam", i, oRecordSet.Fields.Item("U_FamNam").Value);
//				oDS_ZPY501L.SetValue("U_FamPer", i, oRecordSet.Fields.Item("U_FamPer").Value);
//				oDS_ZPY501L.SetValue("U_ChkCod", i, oRecordSet.Fields.Item("U_ChkCod").Value);
//				oDS_ZPY501L.SetValue("U_ChkInt", i, oRecordSet.Fields.Item("U_ChkInt").Value);
//				oDS_ZPY501L.SetValue("U_ChkBas", i, oRecordSet.Fields.Item("U_ChkBas").Value);
//				oDS_ZPY501L.SetValue("U_ChkbuY", i, oRecordSet.Fields.Item("U_ChkbuY").Value);
//				oDS_ZPY501L.SetValue("U_ChkJan", i, oRecordSet.Fields.Item("U_ChkJan").Value);
//				oDS_ZPY501L.SetValue("U_ChkJeL", i, oRecordSet.Fields.Item("U_ChkJeL").Value);
//				oDS_ZPY501L.SetValue("U_ChkChl", i, oRecordSet.Fields.Item("U_ChkChl").Value);
//				oDS_ZPY501L.SetValue("U_ChkDaJ", i, oRecordSet.Fields.Item("U_ChkDaJ").Value);
//				oDS_ZPY501L.SetValue("U_ChkCHS", i, "0");
//				//출산/입양: 해당년도의 입양자,출산자이니 전년도자료 클리어
//				oDS_ZPY501L.SetValue("U_BOHGBN", i, oRecordSet.Fields.Item("U_BOHGBN").Value);
//				oDS_ZPY501L.SetValue("U_BOHAMT1", i, oRecordSet.Fields.Item("U_BOHAMT1").Value);
//				oDS_ZPY501L.SetValue("U_BOHAMT2", i, oRecordSet.Fields.Item("U_BOHAMT2").Value);
//				oDS_ZPY501L.SetValue("U_BOHAMT3", i, Convert.ToString(BOHAL1 + BOHAL2));
//				oDS_ZPY501L.SetValue("U_MEDAMT1", i, oRecordSet.Fields.Item("U_MEDAMT1").Value);
//				oDS_ZPY501L.SetValue("U_MEDAMT2", i, oRecordSet.Fields.Item("U_MEDAMT2").Value);
//				oDS_ZPY501L.SetValue("U_SCHGBN", i, oRecordSet.Fields.Item("U_SCHGBN").Value);
//				oDS_ZPY501L.SetValue("U_EDCAMT1", i, oRecordSet.Fields.Item("U_EDCAMT1").Value);
//				oDS_ZPY501L.SetValue("U_EDCAMT2", i, oRecordSet.Fields.Item("U_EDCAMT2").Value);
//				oDS_ZPY501L.SetValue("U_CADAMT1", i, oRecordSet.Fields.Item("U_CADAMT1").Value);
//				oDS_ZPY501L.SetValue("U_CADAMT2", i, oRecordSet.Fields.Item("U_CADAMT2").Value);
//				oDS_ZPY501L.SetValue("U_CSHCAD1", i, oRecordSet.Fields.Item("U_CSHCAD1").Value);
//				oDS_ZPY501L.SetValue("U_CSHCAD2", i, oRecordSet.Fields.Item("U_CSHCAD2").Value);
//				oDS_ZPY501L.SetValue("U_CSHAMT1", i, oRecordSet.Fields.Item("U_CSHAMT1").Value);
//				oDS_ZPY501L.SetValue("U_GBUAMT1", i, oRecordSet.Fields.Item("U_GBUAMT1").Value);
//				oDS_ZPY501L.SetValue("U_GBUAMT2", i, oRecordSet.Fields.Item("U_GBUAMT2").Value);
//				oDS_ZPY501L.SetValue("U_GBUAMT3", i, oRecordSet.Fields.Item("U_GBUAMT3").Value);
//				oRecordSet.MoveNext();
//				i = i + 1;
//			}

//			/// 메세지
//			oMat1.LoadFromDataSource();
//			i = oMat1.VisualRowCount - 1;
//			if (!string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", i)))) {
//				Matrix_AddRow(i + 1);
//			}

//			if (Strings.Trim(oDS_ZPY501H.GetValue("U_HUSMAN", 0)) != "1" & Strings.Trim(oDS_ZPY501H.GetValue("U_HUSMAN", 0)) != "2") {
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oDS_ZPY501H.SetValue("U_HUSMAN", 0, MDC_SetMod.Get_ReData("U_HUSMAN", "Code", "[@PH_PY001A]", "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'"));
//			}

//			sQry = "EXEC ZPY501_3 '" + JSNYER + "', '" + MSTCOD + "', '" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount > 0) {
//				oDS_ZPY501H.SetValue("U_LAWGBU", 0, oRecordSet.Fields.Item("LAWGBU").Value);
//				oDS_ZPY501H.SetValue("U_POCGBU", 0, oRecordSet.Fields.Item("POCGBU").Value);
//				oDS_ZPY501H.SetValue("U_SP1GBU", 0, oRecordSet.Fields.Item("SP1GBU").Value);
//				oDS_ZPY501H.SetValue("U_SP2GBU", 0, oRecordSet.Fields.Item("SP2GBU").Value);
//				oDS_ZPY501H.SetValue("U_USJGBU", 0, oRecordSet.Fields.Item("USJGBU").Value);
//				oDS_ZPY501H.SetValue("U_JIJGBU", 0, oRecordSet.Fields.Item("JIJGBU").Value);
//				oDS_ZPY501H.SetValue("U_JI2GBU", 0, oRecordSet.Fields.Item("JI2GBU").Value);
//				oDS_ZPY501H.SetValue("U_JNGGBU", 0, oRecordSet.Fields.Item("JNGGBU").Value);
//				oDS_ZPY501H.SetValue("U_BONGBU1", 0, oRecordSet.Fields.Item("BONGBU1").Value);
//				oDS_ZPY501H.SetValue("U_BONGBU2", 0, oRecordSet.Fields.Item("BONGBU2").Value);
//			}

//			oForm.Update();
//			////
//			// SBO_Application.StatusBar.SetText "마스터에 적용하였습니다.", bmt_Short, smt_Success
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Family_Data_Display Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

//		}

//		private void Family_Total(ref string TotalOption = "")
//		{
//			short iCol = 0;
//			double[] BOHAMT = new double[2];
//			double[] MEDAMT = new double[2];
//			double[] EDCAMT = new double[2];
//			double[] CADAMT = new double[2];
//			double[] CSHCAD = new double[2];
//			double[] CSHAMT = new double[2];
//			double[] GBUAMT = new double[2];
//			PerSonNo_Info Per_Info = default(PerSonNo_Info);

//			double U_JGAMED = 0;
//			double U_BOHAMT = 0;
//			double U_JGABOA = 0;
//			double U_GENMED = 0;
//			double U_JGASCH = 0;
//			double U_BONSCH = 0;
//			double U_JICSCH = 0;
//			double SCHKUM = 0;
//			double TMPAMT = 0;
//			short U_GYNGLO = 0;
//			short U_BUYN20 = 0;
//			short U_BAEWOO = 0;
//			short U_BUYN60 = 0;
//			short U_GYNGL2 = 0;
//			short U_DAGYSU = 0;
//			short U_MZBURI = 0;
//			short U_JANGAE = 0;
//			short U_BUYN06 = 0;
//			short U_INJTOT = 0;
//			short U_CHLSAN = 0;
//			double U_JI2GBU = 0;
//			///
//			for (iCol = 0; iCol <= 1; iCol++) {
//				BOHAMT[iCol] = 0;
//				MEDAMT[iCol] = 0;
//				EDCAMT[iCol] = 0;
//				CADAMT[iCol] = 0;
//				CSHCAD[iCol] = 0;
//				CSHAMT[iCol] = 0;
//				GBUAMT[iCol] = 0;
//			}

//			U_BOHAMT = 0;
//			U_JGABOA = 0;
//			U_JGAMED = 0;
//			U_GENMED = 0;
//			U_BONSCH = 0;
//			U_JGASCH = 0;
//			U_JICSCH = 0;
//			TMPAMT = 0;
//			///
//			U_BAEWOO = 0;
//			U_BUYN20 = 0;
//			U_BUYN60 = 0;
//			U_GYNGLO = 0;
//			U_GYNGL2 = 0;
//			U_JANGAE = 0;
//			U_MZBURI = 0;
//			U_BUYN06 = 0;
//			U_DAGYSU = 0;
//			U_INJTOT = 0;
//			U_DAGYSU = 0;
//			U_CHLSAN = 0;
//			///
//			oMat1.FlushToDataSource();
//			for (iCol = 0; iCol <= oDS_ZPY501L.Size - 1; iCol++) {
//				oDS_ZPY501L.Offset = iCol;
//				//// 인적공제 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", iCol)))) {
//					//// 나이확인
//					//UPGRADE_WARNING: Per_Info 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					Per_Info = MDC_Com.Age_Chk(ref oDS_ZPY501L.GetValue("U_FamPer", iCol), ref Strings.Trim(oDS_ZPY501H.GetValue("U_JSNYER", 0)) + "1231");
//					/// 기본공제자가 아니더라도 공제 가능.
//					U_JANGAE = U_JANGAE + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkJan", iCol));
//					/// 장애인공제수
//					/// 본인/배우자/직계존속은 안됨(자녀양육,출산입양)
//					//
//					if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) > "3") {
//						/// 출산.입양공제는 기본공제대상자는 다른 근로자가 공제, 본인은 자녀양육비추가공제로 받을수 있슴.
//						U_CHLSAN = U_CHLSAN + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkCHS", iCol));
//						/// 출산입양자공제수
//						/// 6세이하비속의 추가공제는 기본공제대상자는 다른 근로자가 공제, 본인은 자녀양육비추가공제로 받을수 있슴.
//						if (Per_Info.ManAge <= 6)
//							U_BUYN06 = U_BUYN06 + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkChl", iCol));
//					}
//					/// 기본공제자일경우만 공제일때만 가능.
//					if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkBas", iCol)) == "1") {
//						U_INJTOT = U_INJTOT + 1;

//						if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) == "0" & Per_Info.Sex == Convert.ToDouble("1")) {
//							U_MZBURI = U_MZBURI + Conversion.Val(oDS_ZPY501L.GetValue("U_ChkbuY", iCol));
//							///부녀자공제수(본인이고 여성일경우만가능)
//						}

//						/// 본인이아닐경우
//						if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) != "0") {
//							if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) == "3")
//								U_BAEWOO = U_BAEWOO + 1;
//							if (Per_Info.ManAge <= 20 & Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) != "3")
//								U_BUYN20 = U_BUYN20 + 1;
//							if (Per_Info.ManAge >= 60 & Per_Info.Sex == 0 & Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) != "3") {
//								U_BUYN60 = U_BUYN60 + 1;
//							}
//							//// 2009년부터 여자일경우도 55세이상-> 60세이상으로 상향조정
//							if (Per_Info.ManAge >= 60 & Per_Info.Sex == 1 & Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) != "3") {
//								U_BUYN60 = U_BUYN60 + 1;
//							}
//							/// 나이한도에서는 제외되나 장애자여서 기본공제인원에 들어가는 사람(배우자는 배우자인원수에 포함하니 제외)
//							if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) != "3") {
//								if (Conversion.Val(oDS_ZPY501L.GetValue("U_ChkJan", iCol)) == 1) {
//									if (Per_Info.ManAge > 20 & ((Per_Info.ManAge < 60 & Per_Info.Sex == 1) | (Per_Info.ManAge < 60 & Per_Info.Sex != 1)))
//										U_BUYN20 = U_BUYN20 + 1;
//								}
//							}
//							//// 2009년 65세이상 69세이하 경로우대 폐지됨.
//							//If Per_Info.ManAge >= 65 And Per_Info.ManAge <= 69 Then U_GYNGLO = U_GYNGLO + 1
//							U_GYNGLO = 0;
//							if (Per_Info.ManAge >= 70)
//								U_GYNGL2 = U_GYNGL2 + 1;
//							if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) == "4" & Per_Info.ManAge <= 20) {
//								U_DAGYSU = U_DAGYSU + 1;
//							}
//						}
//					}
//				}
//				//// 보험료 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				BOHAMT[0] = BOHAMT[0] + Conversion.Val(oDS_ZPY501L.GetValue("U_BOHAMT1", iCol));
//				BOHAMT[1] = BOHAMT[1] + Conversion.Val(oDS_ZPY501L.GetValue("U_BOHAMT2", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_BOHAMT3", iCol));
//				TMPAMT = 0;
//				TMPAMT = Conversion.Val(oDS_ZPY501L.GetValue("U_BOHAMT1", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_BOHAMT2", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_BOHAMT3", iCol));
//				/// 장애인보험
//				if (Strings.Trim(oDS_ZPY501L.GetValue("U_BOHGBN", iCol)) == "2") {
//					U_JGABOA = U_JGABOA + TMPAMT;
//					/// 본인일경우
//					if (Strings.Trim(oDS_ZPY501L.GetValue("U_CHKCOD", iCol)) == "0") {
//						U_JGABOA = U_JGABOA - Conversion.Val(oDS_ZPY501H.GetValue("U_BOHAL1", 0)) - Conversion.Val(oDS_ZPY501H.GetValue("U_BOHAL2", 0));
//					}
//				/// 일반인보험
//				} else {
//					U_BOHAMT = U_BOHAMT + TMPAMT;
//					/// 본인일경우
//					if (Strings.Trim(oDS_ZPY501L.GetValue("U_CHKCOD", iCol)) == "0") {
//						U_BOHAMT = U_BOHAMT - Conversion.Val(oDS_ZPY501H.GetValue("U_BOHAL1", 0)) - Conversion.Val(oDS_ZPY501H.GetValue("U_BOHAL2", 0));
//					}
//				}
//				/// 의료비 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				MEDAMT[0] = MEDAMT[0] + Conversion.Val(oDS_ZPY501L.GetValue("U_MEDAMT1", iCol));
//				MEDAMT[1] = MEDAMT[1] + Conversion.Val(oDS_ZPY501L.GetValue("U_MEDAMT2", iCol));
//				TMPAMT = Conversion.Val(oDS_ZPY501L.GetValue("U_MEDAMT1", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_MEDAMT2", iCol));
//				//        If Trim$(oDS_ZPY501L.GetValue("U_ChkCod", icol)) = "0" Or _
//				//'           Trim$(oDS_ZPY501L.GetValue("U_ChkJeL", icol)) = "1" Or _
//				//'           Trim$(oDS_ZPY501L.GetValue("U_ChkJan", icol)) = "1" Then   '/ 본인, 경로, 장애인
//				/// 본인, 65세이상, 장애인
//				if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) == "0" | Per_Info.ManAge >= 65 | Strings.Trim(oDS_ZPY501L.GetValue("U_ChkJan", iCol)) == "1") {
//					U_JGAMED = U_JGAMED + TMPAMT;
//				} else {
//					U_GENMED = U_GENMED + TMPAMT;
//				}
//				/// 교육비 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				EDCAMT[0] = EDCAMT[0] + Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol));
//				EDCAMT[1] = EDCAMT[1] + Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol));
//				TMPAMT = Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol));
//				/// 본인전액공제(대학원포함)
//				if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) == "0") {
//					U_BONSCH = Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol));
//				} else {
//					/// 장애인특수교육
//					if (Strings.Trim(oDS_ZPY501L.GetValue("U_SCHGBN", iCol)) == "5") {
//						U_JGASCH = U_JGASCH + Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT1", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_EDCAMT2", iCol));
//					} else {
//						switch (Strings.Trim(oDS_ZPY501L.GetValue("U_SCHGBN", iCol))) {
//							case "1":
//							case "2":
//								SCHKUM = 3000000;
//								/// 영유아, 초중고
//								break;
//							case "3":
//								SCHKUM = 9000000;
//								/// 대학교
//								break;
//							case "4":
//								SCHKUM = 0;
//								/// 대학원생
//								break;
//						}
//						if (TMPAMT <= SCHKUM) {
//							U_JICSCH = U_JICSCH + TMPAMT;
//						} else {
//							U_JICSCH = U_JICSCH + SCHKUM;
//						}
//					}
//				}
//				/// 신용카드 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				CADAMT[0] = CADAMT[0] + Conversion.Val(oDS_ZPY501L.GetValue("U_CADAMT1", iCol));
//				CADAMT[1] = CADAMT[1] + Conversion.Val(oDS_ZPY501L.GetValue("U_CADAMT2", iCol));
//				/// 직불카드
//				CSHCAD[0] = CSHCAD[0] + Conversion.Val(oDS_ZPY501L.GetValue("U_CSHCAD1", iCol));
//				CSHCAD[1] = CSHCAD[1] + Conversion.Val(oDS_ZPY501L.GetValue("U_CSHCAD2", iCol));
//				/// 현금영수증 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				CSHAMT[0] = CSHAMT[0] + Conversion.Val(oDS_ZPY501L.GetValue("U_CSHAMT1", iCol));
//				/// 기부금 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				GBUAMT[0] = GBUAMT[0] + Conversion.Val(oDS_ZPY501L.GetValue("U_GBUAMT1", iCol));
//				GBUAMT[1] = GBUAMT[1] + Conversion.Val(oDS_ZPY501L.GetValue("U_GBUAMT2", iCol)) + Conversion.Val(oDS_ZPY501L.GetValue("U_GBUAMT3", iCol));
//				U_JI2GBU = U_JI2GBU + Conversion.Val(oDS_ZPY501L.GetValue("U_GBUAMT3", iCol));
//			}

//			if (TotalOption == "All") {
//				/// 다자녀일경우 소득공제명세에도 체크자동으로
//				for (iCol = 0; iCol <= oDS_ZPY501L.Size - 1; iCol++) {
//					oDS_ZPY501L.Offset = iCol;
//					if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkBas", iCol)) == "1" & !string.IsNullOrEmpty(Strings.Trim(oDS_ZPY501L.GetValue("U_FamNam", iCol))) & U_DAGYSU >= 2) {
//						//UPGRADE_WARNING: Per_Info 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Per_Info = MDC_Com.Age_Chk(ref oDS_ZPY501L.GetValue("U_FamPer", iCol), ref Strings.Trim(oDS_ZPY501H.GetValue("U_JSNYER", 0)) + "1231");
//						if (Strings.Trim(oDS_ZPY501L.GetValue("U_ChkCod", iCol)) == "4" & Per_Info.ManAge <= 20) {
//							oDS_ZPY501L.SetValue("U_ChkDaJ", iCol, Convert.ToString(1));
//						} else {
//							oDS_ZPY501L.SetValue("U_ChkDaJ", iCol, Convert.ToString(0));
//						}
//					} else {
//						oDS_ZPY501L.SetValue("U_ChkDaJ", iCol, Convert.ToString(0));
//					}
//				}
//			}
//			oForm.Freeze(true);
//			oMat1.LoadFromDataSource();
//			if (TotalOption == "All") {
//				/// 소득공제항목 //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				oDS_ZPY501H.SetValue("U_JGABOA", 0, Convert.ToString(U_JGABOA));
//				/// 장애인보험료
//				oDS_ZPY501H.SetValue("U_BOHAMT", 0, Convert.ToString(U_BOHAMT));
//				/// 일반인보험료
//				oDS_ZPY501H.SetValue("U_JGAMED", 0, Convert.ToString(U_JGAMED));
//				/// 본인,경로,장애인의료비
//				oDS_ZPY501H.SetValue("U_GENMED", 0, Convert.ToString(U_GENMED));
//				/// 그외 의료비
//				oDS_ZPY501H.SetValue("U_BONSCH", 0, Convert.ToString(U_BONSCH));
//				/// 본인교육비
//				oDS_ZPY501H.SetValue("U_JGASCH", 0, Convert.ToString(U_JGASCH));
//				/// 장애특수교육비
//				oDS_ZPY501H.SetValue("U_JICSCH", 0, Convert.ToString(U_JICSCH));
//				/// 가족교육비
//				if ((Conversion.Val(oDS_ZPY501H.GetValue("U_CADSAV", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_GIRSAV", 0))) != (CADAMT[0] + CADAMT[1])) {
//					oDS_ZPY501H.SetValue("U_CADSAV", 0, Convert.ToString(CADAMT[0] + CADAMT[1]));
//					/// 신용카드
//					//oDS_ZPY501H.SetValue "U_GIRSAV", 0, CADAMT(1)   '/ 신용카드(학원,지로)
//				}
//				oDS_ZPY501H.SetValue("U_CA1SAV", 0, Convert.ToString(CSHCAD[0] + CSHCAD[1]));
//				oDS_ZPY501H.SetValue("U_CSHSAV", 0, Convert.ToString(CSHAMT[0]));
//				/// 현금영수증

//				/// 기부금 재계산
//				TMPAMT = Conversion.Val(oDS_ZPY501H.GetValue("U_LAWGBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_POCGBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_SP1GBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_SP2GBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_USJGBU", 0)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JNGGBU", 0));

//				if ((Conversion.Val(Convert.ToString(TMPAMT)) + Conversion.Val(oDS_ZPY501H.GetValue("U_JIJGBU", 0))) != (GBUAMT[0] + GBUAMT[1] - U_JI2GBU)) {
//					/// 기부금총액> 총기부금-지정기부금
//					/// 그외기부금
//					if ((Conversion.Val(Convert.ToString(GBUAMT[0])) + Conversion.Val(Convert.ToString(GBUAMT[1]))) > (TMPAMT + U_JI2GBU + Conversion.Val(oDS_ZPY501H.GetValue("U_JIJGBU", 0)))) {
//						oDS_ZPY501H.SetValue("U_JIJGBU", 0, Convert.ToString(GBUAMT[0] + GBUAMT[1] - U_JI2GBU - TMPAMT));
//						/// 기부금
//					} else if (Conversion.Val(Convert.ToString(GBUAMT[0])) == 0 & Conversion.Val(Convert.ToString(GBUAMT[1])) == 0) {
//						oDS_ZPY501H.SetValue("U_LAWGBU", 0, Convert.ToString(0));
//						oDS_ZPY501H.SetValue("U_POCGBU", 0, Convert.ToString(0));
//						oDS_ZPY501H.SetValue("U_SP1GBU", 0, Convert.ToString(0));
//						oDS_ZPY501H.SetValue("U_SP2GBU", 0, Convert.ToString(0));
//						oDS_ZPY501H.SetValue("U_USJGBU", 0, Convert.ToString(0));
//						oDS_ZPY501H.SetValue("U_JNGGBU", 0, Convert.ToString(0));
//						oDS_ZPY501H.SetValue("U_JIJGBU", 0, Convert.ToString(0));
//						oDS_ZPY501H.SetValue("U_JI2GBU", 0, Convert.ToString(0));
//					}
//				}

//				oDS_ZPY501H.SetValue("U_JI2GBU", 0, Convert.ToString(U_JI2GBU));
//				/// 기부금(노조)

//				if (U_DAGYSU < 2)
//					U_DAGYSU = 0;
//				/// 인적공제
//				oDS_ZPY501H.SetValue("U_BAEWOO", 0, Convert.ToString(U_BAEWOO));
//				oDS_ZPY501H.SetValue("U_BUYN20", 0, Convert.ToString(U_BUYN20));
//				oDS_ZPY501H.SetValue("U_BUYN60", 0, Convert.ToString(U_BUYN60));
//				oDS_ZPY501H.SetValue("U_GYNGLO", 0, Convert.ToString(U_GYNGLO));
//				oDS_ZPY501H.SetValue("U_GYNGL2", 0, Convert.ToString(U_GYNGL2));
//				oDS_ZPY501H.SetValue("U_GYNGLO", 0, Convert.ToString(U_GYNGLO));
//				oDS_ZPY501H.SetValue("U_JANGAE", 0, Convert.ToString(U_JANGAE));
//				oDS_ZPY501H.SetValue("U_MZBURI", 0, Convert.ToString(U_MZBURI));
//				oDS_ZPY501H.SetValue("U_BUYN06", 0, Convert.ToString(U_BUYN06));
//				oDS_ZPY501H.SetValue("U_DAGYSU", 0, Convert.ToString(U_DAGYSU));
//				oDS_ZPY501H.SetValue("U_CHLSAN", 0, Convert.ToString(U_CHLSAN));
//				oDS_ZPY501H.SetValue("U_INJTOT", 0, Convert.ToString(U_INJTOT));


//			}

//			/// 부양가족 명세 집계
//			oDS_ZPY501H.SetValue("U_BONBOH1", 0, Convert.ToString(BOHAMT[0]));
//			oDS_ZPY501H.SetValue("U_BONBOH2", 0, Convert.ToString(BOHAMT[1]));
//			oDS_ZPY501H.SetValue("U_BONMED1", 0, Convert.ToString(MEDAMT[0]));
//			oDS_ZPY501H.SetValue("U_BONMED2", 0, Convert.ToString(MEDAMT[1]));
//			oDS_ZPY501H.SetValue("U_BONEDC1", 0, Convert.ToString(EDCAMT[0]));
//			oDS_ZPY501H.SetValue("U_BONEDC2", 0, Convert.ToString(EDCAMT[1]));
//			oDS_ZPY501H.SetValue("U_BONCAD1", 0, Convert.ToString(CADAMT[0]));
//			oDS_ZPY501H.SetValue("U_BONCAD2", 0, Convert.ToString(CADAMT[1]));
//			oDS_ZPY501H.SetValue("U_BONCA11", 0, Convert.ToString(CSHCAD[0]));
//			oDS_ZPY501H.SetValue("U_BONCA12", 0, Convert.ToString(CSHCAD[1]));
//			oDS_ZPY501H.SetValue("U_BONCSH1", 0, Convert.ToString(CSHAMT[0]));
//			oDS_ZPY501H.SetValue("U_BONGBU1", 0, Convert.ToString(GBUAMT[0]));
//			oDS_ZPY501H.SetValue("U_BONGBU2", 0, Convert.ToString(GBUAMT[1]));
//			oForm.Freeze(false);
//			oForm.Update();

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//		}

////*******************************************************************
////// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
////*******************************************************************
//		private void Matrix_AddRow(int oRow, ref bool Insert_YN = false)
//		{
//			if (Insert_YN == false) {
//				oDS_ZPY501L.InsertRecord((oRow));
//			}
//			oDS_ZPY501L.Offset = oRow;
//			oDS_ZPY501L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//			oDS_ZPY501L.SetValue("U_FamNam", oRow, "");
//			oDS_ZPY501L.SetValue("U_FamPer", oRow, "");
//			if (oRow == 0) {
//				oDS_ZPY501L.SetValue("U_ChkCod", oRow, "0");
//			} else {
//				oDS_ZPY501L.SetValue("U_ChkCod", oRow, "");
//			}
//			oDS_ZPY501L.SetValue("U_ChkInt", oRow, "1");
//			oDS_ZPY501L.SetValue("U_ChkBas", oRow, "1");
//			oDS_ZPY501L.SetValue("U_ChkbuY", oRow, "0");
//			oDS_ZPY501L.SetValue("U_ChkJan", oRow, "0");
//			oDS_ZPY501L.SetValue("U_ChkJeL", oRow, "0");
//			oDS_ZPY501L.SetValue("U_ChkChl", oRow, "0");
//			oDS_ZPY501L.SetValue("U_ChkDaJ", oRow, "0");
//			oDS_ZPY501L.SetValue("U_ChkCHS", oRow, "0");
//			oDS_ZPY501L.SetValue("U_BOHGBN", oRow, "1");
//			oDS_ZPY501L.SetValue("U_BOHAMT1", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_BOHAMT2", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_MEDAMT1", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_MEDAMT2", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_SCHGBN", oRow, "1");
//			oDS_ZPY501L.SetValue("U_EDCAMT1", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_EDCAMT2", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_CADAMT1", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_CADAMT2", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_CSHCAD1", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_CSHCAD2", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_CSHAMT1", oRow, Convert.ToString(0));
//			oDS_ZPY501L.SetValue("U_GBUAMT2", oRow, Convert.ToString(0));

//			oMat1.LoadFromDataSource();
//		}
//	}
//}
