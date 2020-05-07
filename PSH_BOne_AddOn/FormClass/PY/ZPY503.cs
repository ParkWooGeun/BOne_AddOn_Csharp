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
//	[System.Runtime.InteropServices.ProgId("ZPY503_NET.ZPY503")]
//	public class ZPY503
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY503.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 정산 세액 계산
//////  FormType       : 2000060503
//////  Create Date    : 2006.01.20
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		private string oJsnYear;
//		private string oSMonth;
//		private string oEMonth;

//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm()
//		{
//			//Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY503.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY503_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY503");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//			////////////////////////////////////////////////////////////////////////////////
//			//***************************************************************
//			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//			//oForm.DataBrowser.BrowseBy = "DocNum"
//			//***************************************************************
//			////////////////////////////////////////////////////////////////////////////////
//			oForm.Freeze(true);
//			CreateItems();
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
//							if (oForm.Items.Item("MSTCOD").Enabled == true) {
//								oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						} else if (pval.ItemUID == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							if (Execution() == false) {
//								BubbleEvent = false;
//							} else {
//								BubbleEvent = false;
//								oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							}
//						}
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUID == "JSNYER" | pval.ItemUID == "SMonth" | pval.ItemUID == "EMonth" | pval.ItemUID == "MSTCOD" | pval.ItemUID == "JSNMON" | pval.ItemUID == "JIGDAT")) {
//						FlushToItemValue(pval.ItemUID);
//					}
//					break;
//				//et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if (pval.ItemUID == "JSNGBN" | pval.ItemUID == "JSNMON") {
//							FlushToItemValue(pval.ItemUID);
//						}
//						if (pval.ItemUID == "CLTCOD") {
//							////기본사항 - 부서1 (사업장에 따른 부서변경)
//							oCombo = oForm.Items.Item("DPTSTR").Specific;

//							if (oCombo.ValidValues.Count > 0) {
//								for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//									oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//								}
//								oCombo.ValidValues.Add("%", "전체");
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
//							oCombo = oForm.Items.Item("DPTEND").Specific;

//							if (oCombo.ValidValues.Count > 0) {
//								for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//									oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//								}
//								oCombo.ValidValues.Add("%", "전체");
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
//				//et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					if (pval.BeforeAction == true & pval.ItemUID != "1000001" & pval.ItemUID != "2") {
//						///정산년도
//						if (Last_Item == "JSNYER") {
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
//							//                    If Trim$(oForm.Items(Last_Item).Specific.Value) <> "" Then
//							//                        If MDC_SetMod.ChkYearMonth(oJsnYear & Format$(oForm.Items(Last_Item).Specific.Value, "00")) = False Then
//							//                            oForm.Items(Last_Item).Update
//							//                            Sbo_Application.StatusBar.SetText "생성기간을 확인하여 주십시오.", bmt_Short, smt_Error
//							//                            BubbleEvent = False
//							//                        End If
//							//                    End If
//						} else if (Last_Item == "MSTCOD") {
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
//					if (pval.BeforeAction == true & pval.ItemUID == "JSNYER" & pval.CharPressed == 9) {
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
//					} else if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String)) & MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", ref "") == true) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
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
//					oForm.Items.Item("JSNYER").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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

//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo1 = null;
//			SAPbouiCOM.ComboBox oCombo2 = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.Column oColumn = null;
//			string sQry = null;
//			short i = 0;
//			string STDMON = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			oForm.DataSources.UserDataSources.Add("JSNYER", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			/// 귀속년도
//			oForm.DataSources.UserDataSources.Add("JSNMON", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
//			/// 귀속월
//			oForm.DataSources.UserDataSources.Add("JSNGBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 구분
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 사업장
//			oForm.DataSources.UserDataSources.Add("DPTSTR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 부서코드
//			oForm.DataSources.UserDataSources.Add("DPTEND", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oForm.DataSources.UserDataSources.Add("SMonth", SAPbouiCOM.BoDataType.dt_DATE);
//			/// 시작일
//			oForm.DataSources.UserDataSources.Add("EMonth", SAPbouiCOM.BoDataType.dt_DATE);
//			/// 종료일
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			oForm.DataSources.UserDataSources.Add("SINYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			/// 신고연월
//			oForm.DataSources.UserDataSources.Add("JIGDAT", SAPbouiCOM.BoDataType.dt_DATE);
//			/// 지급일자

//			oEdit = oForm.Items.Item("JSNYER").Specific;
//			oEdit.DataBind.SetBound(true, "", "JSNYER");
//			oCombo1 = oForm.Items.Item("JSNMON").Specific;
//			oCombo1.DataBind.SetBound(true, "", "JSNMON");
//			oEdit = oForm.Items.Item("SMonth").Specific;
//			oEdit.DataBind.SetBound(true, "", "SMonth");
//			oEdit = oForm.Items.Item("EMonth").Specific;
//			oEdit.DataBind.SetBound(true, "", "EMonth");
//			oEdit = oForm.Items.Item("SINYMM").Specific;
//			oEdit.DataBind.SetBound(true, "", "SINYMM");
//			oEdit = oForm.Items.Item("JIGDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "JIGDAT");
//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");
//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");
//			oCombo1 = oForm.Items.Item("JSNGBN").Specific;
//			oCombo1.DataBind.SetBound(true, "", "JSNGBN");

//			oForm.DataSources.UserDataSources.Add("Col0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
//			oForm.DataSources.UserDataSources.Add("Col1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

//			oColumn = oMat1.Columns.Item("Col0");
//			oColumn.DataBind.SetBound(true, "", "Col0");

//			oColumn = oMat1.Columns.Item("Col1");
//			oColumn.DataBind.SetBound(true, "", "Col1");

//			//// 정산구분
//			oCombo1 = oForm.Items.Item("JSNGBN").Specific;
//			oCombo1.ValidValues.Add("1", "연말정산(재직자)");
//			oCombo1.ValidValues.Add("2", "중도정산(퇴직자)");
//			//    sQry = " SELECT U_Minor, U_CdName FROM [@ZPY001L] WHERE Code='P192' ORDER BY U_Minor "
//			//    oRecordSet.DoQuery sQry
//			//    Do Until oRecordSet.EOF
//			//        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
//			//        oRecordSet.MoveNext
//			//    Loop
//			//// 귀속연월
//			oCombo1 = oForm.Items.Item("JSNMON").Specific;
//			for (i = 1; i <= 12; i++) {
//				STDMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(i, "00");
//				oCombo1.ValidValues.Add(STDMON, STDMON);
//			}
//			//// 사업장
//			oCombo1 = oForm.Items.Item("CLTCOD").Specific;
//			oCombo1.DataBind.SetBound(true, "", "CLTCOD");
//			sQry = "SELECT Code,Name FROM [@PH_PY005A] ";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oCombo1.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}



//			//// 부서
//			oCombo1 = oForm.Items.Item("DPTSTR").Specific;
//			oCombo1.DataBind.SetBound(true, "", "DPTSTR");
//			oForm.Items.Item("DPTSTR").DisplayDesc = true;
//			//// 부서
//			oCombo1 = oForm.Items.Item("DPTEND").Specific;
//			oCombo1.DataBind.SetBound(true, "", "DPTEND");
//			oForm.Items.Item("DPTEND").DisplayDesc = true;

//			oForm.DataSources.UserDataSources.Item("JSNMON").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "MM");
//			oForm.ActiveItem = "JSNYER";

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
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
//			//UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo1 = null;
//			//UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo2 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private bool Execution()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short ErrNum = 0;
//			int TOTCNT = 0;
//			int MSTCNT = 0;
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString DPTSTR = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(8);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString DPTEND = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(8);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString MSTCOD = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(8);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString STRDAT = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString ENDDAT = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(10);
//			string CLTCOD = null;
//			string BPLID = null;
//			string JSNGBN = null;

//			ErrNum = 0;
//			/// 필수Check /
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			/// 정산년도
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JSNYER").Specific.String))) {
//				ErrNum = 1;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items(JSNGBN).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("JSNGBN").Specific.Selected == null) {
//				ErrNum = 7;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("SMonth").Specific.VALUE)) | string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("EMonth").Specific.VALUE))) {
//				ErrNum = 2;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items(EMonth).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: oForm.Items(SMonth).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("SMonth").Specific.VALUE > oForm.Items.Item("EMonth").Specific.VALUE) {
//				ErrNum = 3;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("CLTCOD").Specific.Selected == null) {
//				ErrNum = 8;
//				goto Error_Message;
//				//    ElseIf oForm.Items("BPLId").Specific.Selected Is Nothing Then
//				//        ErrNum = 4
//				//        GoTo Error_Message
//				//UPGRADE_WARNING: oForm.Items(DPTEND).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: oForm.Items(DPTSTR).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("DPTSTR").Specific.Selected == null | oForm.Items.Item("DPTEND").Specific.Selected == null) {
//				ErrNum = 5;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (Strings.Len(oForm.Items.Item("SINYMM").Specific.VALUE) != 6) {
//				ErrNum = 9;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items(JSNMON).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("JSNMON").Specific.Selected == null) {
//				ErrNum = 10;
//				goto Error_Message;
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (Strings.Len(oForm.Items.Item("JIGDAT").Specific.VALUE) == 0) {
//				ErrNum = 11;
//				goto Error_Message;
//			}
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DPTSTR.Value = oForm.Items.Item("DPTSTR").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DPTEND.Value = oForm.Items.Item("DPTEND").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD.Value = oForm.Items.Item("MSTCOD").Specific.String;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			STRDAT.Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("SMonth").Specific.VALUE, "0000-00-00");
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ENDDAT.Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("EMonth").Specific.VALUE, "0000-00-00");
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.VALUE;
//			//    BPLID = oForm.Items("BPLId").Specific.Selected.VALUE
//			if (Strings.Trim(DPTSTR.Value) == "-1")
//				DPTSTR.Value = "00000001";
//			if (Strings.Trim(DPTEND.Value) == "-1")
//				DPTEND.Value = "ZZZZZZZZ";
//			if (string.IsNullOrEmpty(Strings.Trim(MSTCOD.Value)))
//				MSTCOD.Value = "%";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JSNGBN = oForm.Items.Item("JSNGBN").Specific.Selected.VALUE;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// 해당년도 정산마감인지 확인여부
//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData(U_ENDCHK, U_JOBYER, [ZPY509L], ' & Trim$(oJsnYear) & ',  AND Code = ' & Trim$(CLTCOD) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_SetMod.Get_ReData(ref "U_ENDCHK", ref "U_JOBYER", ref "[@ZPY509L]", ref "'" + Strings.Trim(oJsnYear) + "'", ref " AND Code = '" + Strings.Trim(CLTCOD) + "'") == "Y") {
//				ErrNum = 13;
//				goto Error_Message;
//			}

//			/// 정산세액계산 대상자 조회
//			sQry = " EXEC ZPY503_1 '" + Strings.Trim(oJsnYear) + "', '" + Strings.Trim(JSNGBN) + "', '" + STRDAT.Value + "', '" + ENDDAT.Value + "','" + Strings.Trim(CLTCOD) + "', '" + Strings.Trim(DPTSTR.Value) + "', '" + Strings.Trim(DPTEND.Value) + "','" + Strings.Trim(MSTCOD.Value) + "' ";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 6;
//				goto Error_Message;
//			}
//			////
//			oMat1.Clear();
//			TOTCNT = 0;
//			MSTCNT = 0;
//			while (!(oRecordSet.EoF)) {
//				TOTCNT = TOTCNT + 1;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MSTCOD.Value = oRecordSet.Fields.Item(0).Value;

//				/// 해당년도 정산마감인지 확인여부
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(U_ENDCHK, U_JSNYER, [ZPY504H], ' & Trim$(oJsnYear) & ',  AND U_MSTCOD = ' & Trim$(MSTCOD) & ' AND U_CLTCOD = ' & Trim$(CLTCOD) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (MDC_SetMod.Get_ReData(ref "U_ENDCHK", ref "U_JSNYER", ref "[@ZPY504H]", ref "'" + Strings.Trim(oJsnYear) + "'", ref " AND U_MSTCOD = '" + Strings.Trim(MSTCOD.Value) + "' AND U_CLTCOD = '" + Strings.Trim(CLTCOD) + "'") == "Y") {
//					oForm.DataSources.UserDataSources.Item("Col0").Value = Convert.ToString(TOTCNT);
//					oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD.Value + ": " + oRecordSet.Fields.Item("MSTNAM").Value + " 세액계산 제외! 잠금자료************";
//					oMat1.AddRow();
//				} else {
//					/// 월별자료관리에 데이터유무확인
//					if (MDC_SetMod.Value_ChkYn(ref "[@ZPY343H]", ref "U_JsnYear", ref "'" + Strings.Trim(oJsnYear) + "'", ref " AND  U_MstCode = '" + Strings.Trim(MSTCOD.Value) + "' AND U_CLTCOD = '" + Strings.Trim(CLTCOD) + "'") == true) {
//						oForm.DataSources.UserDataSources.Item("Col0").Value = Convert.ToString(TOTCNT);
//						oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD.Value + ": " + oRecordSet.Fields.Item("MSTNAM").Value + " 세액계산 실패! 월별자료관리 누락";
//						oMat1.AddRow();
//						/// 소득공제항목등록에 데이터유무확인
//					} else if (MDC_SetMod.Value_ChkYn(ref "[@ZPY501H]", ref "U_JSNYER", ref "'" + Strings.Trim(oJsnYear) + "'", ref " AND  U_MSTCOD = '" + Strings.Trim(MSTCOD.Value) + "' AND U_CLTCOD = '" + Strings.Trim(CLTCOD) + "'") == true) {
//						oForm.DataSources.UserDataSources.Item("Col0").Value = Convert.ToString(TOTCNT);
//						oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD.Value + ": " + oRecordSet.Fields.Item("MSTNAM").Value + " 세액계산 실패! 소득공제항목등록 누락";
//						oMat1.AddRow();
//						/// 급여기본등록 데이터유무확인 (2010.03.03 최동권 추가)
//					} else if (MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + Strings.Trim(MSTCOD.Value) + "'") == true) {
//						oForm.DataSources.UserDataSources.Item("Col0").Value = Convert.ToString(TOTCNT);
//						oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD.Value + ": " + oRecordSet.Fields.Item("MSTNAM").Value + " 세액계산 실패! 급여기본등록 누락";
//						oMat1.AddRow();
//					} else {
//						/// 정산세액결과로직
//						if (Execution_Save(ref oJsnYear, ref (oRecordSet.Fields.Item("CLTCOD").Value), ref MSTCOD.Value) == true) {
//							MSTCNT = MSTCNT + 1;
//							oForm.DataSources.UserDataSources.Item("Col0").Value = Convert.ToString(TOTCNT);
//							oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD.Value + ": " + oRecordSet.Fields.Item("MSTNAM").Value + " 세액계산 완료.";
//							oMat1.AddRow();
//						} else {
//							oForm.DataSources.UserDataSources.Item("Col0").Value = Convert.ToString(TOTCNT);
//							oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD.Value + ": " + oRecordSet.Fields.Item("MSTNAM").Value + " 세액계산 실패! **************";
//							oMat1.AddRow();
//						}
//					}

//				}
//				oRecordSet.MoveNext();
//			}
//			///
//			MDC_Globals.Sbo_Application.StatusBar.SetText("( " + MSTCNT + "/" + TOTCNT + " )의 작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("생성기간을 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("시작월보다 종료월이 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지점을 선택하세요. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("생성할 부서범위를 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("세액계산할 대상자료가 없습니다. 월자료 생성을 먼저 하십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("정산 구분은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 8) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사 코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 9) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("신고 연월은 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 10) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 11) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급 일자는 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 13) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("Execution 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}
//		private bool Execution_Save(ref string JSNYER, ref string CLTCOD, ref string MSTCOD)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset sRecordset = null;
//			string sQry = null;

//			string JSNGBN = null;
//			string JSNMON = null;
//			string SINYMM = null;
//			string JIGDAT = null;

//			//// Default
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JSNGBN = oForm.Items.Item("JSNGBN").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JSNMON = oForm.Items.Item("JSNMON").Specific.Selected.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			SINYMM = oForm.Items.Item("SINYMM").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JIGDAT = oForm.Items.Item("JIGDAT").Specific.VALUE;
//			JIGDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(JIGDAT, "0000-00-00");
//			JSNMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(JSNMON, "00");

//			//// 사원별 정산 세액계산시작
//			sRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//Exec dbo.MDC_ZPY503_05  '2005', '9603004', '1','12','2008-01-01'
//			sQry = "Exec ZPY503_" + Strings.Mid(JSNYER, 3, 2) + Strings.Space(1) + "'" + JSNYER + "', '" + Strings.Trim(CLTCOD) + "','" + Strings.Trim(MSTCOD) + "', '" + Strings.Trim(JSNGBN) + "' , '" + Strings.Trim(JSNMON) + "', '" + Strings.Trim(SINYMM) + "', '" + Strings.Trim(JIGDAT) + "', " + MDC_Globals.oCompany.UserSignature;

//			sRecordset.DoQuery(sQry);
//			if (sRecordset.RecordCount <= 0) {
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Execution_Save 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{
//			ZPAY_g_EmpID MstInfo = default(ZPAY_g_EmpID);
//			string JIGDAT = null;
//			switch (oUID) {
//				case "JSNYER":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oUID).Specific.String))) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MDC_Globals.ZPAY_GBL_JSNYER.Value = oForm.Items.Item(oUID).Specific.String;
//					} else {
//						oForm.DataSources.UserDataSources.Item("JSNYER").ValueEx = MDC_Globals.ZPAY_GBL_JSNYER.Value;
//					}
//					oJsnYear = oForm.DataSources.UserDataSources.Item("JSNYER").ValueEx;
//					break;
//				case "SMonth":
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oSMonth = oForm.Items.Item(oUID).Specific.VALUE;
//					oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = oSMonth;
//					break;
//				case "EMonth":
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oEMonth = oForm.Items.Item(oUID).Specific.VALUE;
//					oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = oEMonth;
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.Items.Item(oUID).Specific.String = "";
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = "";
//					} else {
//						//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.Items.Item(oUID).Specific.String = Strings.UCase(oForm.Items.Item(oUID).Specific.String);
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: MstInfo 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MstInfo = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = MstInfo.MSTNAM;
//					}
//					oForm.Items.Item("MSTNAM").Update();
//					break;
//				case "JSNGBN":
//					if (string.IsNullOrEmpty(Strings.Trim(oJsnYear)))
//						oJsnYear = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyy");
//					oForm.Freeze(true);
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oForm.Items.Item(oUID).Specific.Selected.VALUE == "1") {
//						oForm.DataSources.UserDataSources.Item("JSNMON").ValueEx = "12";
//						oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = oJsnYear + "0101";
//						oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = oJsnYear + "1231";
//					} else {
//						oForm.DataSources.UserDataSources.Item("JSNMON").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = "";
//					}
//					oForm.Items.Item("SMonth").Update();
//					oForm.Items.Item("EMonth").Update();
//					oForm.Freeze(false);
//					break;
//				case "JSNMON":
//					if (string.IsNullOrEmpty(Strings.Trim(oForm.DataSources.UserDataSources.Item(oUID).ValueEx))) {
//						oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("SINYMM").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("JIGDAT").ValueEx = "";
//					} else {
//						//UPGRADE_WARNING: oForm.Items(JSNGBN).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if ((oForm.Items.Item("JSNGBN").Specific.Selected != null)) {
//							oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = oJsnYear + Strings.Trim(oForm.DataSources.UserDataSources.Item("JSNMON").ValueEx) + "01";
//							//UPGRADE_WARNING: MDC_SetMod.Lday(oJsnYear & Trim$(oForm.DataSources.UserDataSources(JSNMON).ValueEx)) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = oJsnYear + Strings.Trim(oForm.DataSources.UserDataSources.Item("JSNMON").ValueEx) + MDC_SetMod.Lday(ref oJsnYear + Strings.Trim(oForm.DataSources.UserDataSources.Item("JSNMON").ValueEx));
//							oForm.DataSources.UserDataSources.Item("SINYMM").ValueEx = "";
//							oForm.DataSources.UserDataSources.Item("JIGDAT").ValueEx = "";
//						}
//					}
//					oForm.Items.Item("SMonth").Update();
//					oForm.Items.Item("EMonth").Update();
//					break;
//				case "JIGDAT":
//					JIGDAT = oForm.DataSources.UserDataSources.Item("JIGDAT").ValueEx;
//					if (string.IsNullOrEmpty(Strings.Trim(JIGDAT))) {
//						oForm.DataSources.UserDataSources.Item("JIGDAT").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("SINYMM").ValueEx = "";
//					} else {
//						if (Strings.Right(JIGDAT, 2) <= "10") {
//							oForm.DataSources.UserDataSources.Item("SINYMM").ValueEx = Strings.Left(JIGDAT, 6);
//						} else {
//							oForm.DataSources.UserDataSources.Item("SINYMM").ValueEx = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Left(JIGDAT, 6) + "01", "0000-00-00"))), "YYYYMM");
//						}
//					}
//					break;

//			}
//			oForm.Items.Item(oUID).Update();
//		}
//	}
//}
