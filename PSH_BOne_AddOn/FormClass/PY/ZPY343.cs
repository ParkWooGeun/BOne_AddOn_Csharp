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
//	[System.Runtime.InteropServices.ProgId("ZPY343_NET.ZPY343")]
//	public class ZPY343
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY343.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 월별자료관리
//////  FormType       : 2000060343
//////  Create Date    : 2005.12.10
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//		private SAPbouiCOM.CheckBox oCheck;
//			//시스템코드 헤더
//		private SAPbouiCOM.DBDataSource oDS_ZPY343H;
//			//시스템코드 라인
//		private SAPbouiCOM.DBDataSource oDS_ZPY343L;

//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;
////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm(ref string DocNum = "")
//		{
//			//Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY343.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//------------------------------------------------------------------------
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//------------------------------------------------------------------------
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY343_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//--------------------------------------------------------------------------------------------------------------
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//--------------------------------------------------------------------------------------------------------------
//			SubMain.AddForms(this, oFormUniqueID, "ZPY343");
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
//			FormItemEnabled();

//			oForm.EnableMenu(("1281"), true);
//			/// 찾기
//			oForm.EnableMenu(("1282"), false);
//			/// 추가
//			oForm.EnableMenu(("1284"), false);
//			/// 취소
//			oForm.EnableMenu(("1293"), false);
//			/// 행삭제

//			if (!string.IsNullOrEmpty(DocNum)) {
//				ShowSource(ref DocNum);
//			}

//			oForm.Update();
//			oForm.Visible = true;
//			oForm.Freeze(false);

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
//							//------------------------------------------------------------
//							////추가및 업데이트시에
//							//------------------------------------------------------------
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (MatrixSpaceLineDel() == false) {
//									BubbleEvent = false;
//									return;
//								}
//							}
//						} else if (pval.ItemUID == "CBtn1") {
//							if (oForm.Items.Item("MstCode").Enabled == true) {
//								oForm.Items.Item("MstCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					} else {
//						if (pval.ItemUID == "Check1") {
//							if (oCheck.Checked == true) {
//								oForm.ActiveItem = "Focus";
//								oForm.Items.Item("Mat1").Enabled = false;
//							} else {
//								oForm.Items.Item("Mat1").Enabled = true;
//							}
//						}
//					}
//					break;
//				//et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if (pval.ItemUID == "CLTCOD") {
//							////기본사항 - 부서1 (사업장에 따른 부서변경)
//							oCombo = oForm.Items.Item("DptCode").Specific;

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
//						}
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						if (pval.ItemUID == "Mat1" & (Strings.Left(pval.ColUID, 3) == "Col")) {
//							FlushToItemValue(pval.ColUID, ref pval.Row);
//						} else if (pval.ItemUID == "MstCode" | pval.ItemUID == "JnsYear") {
//							FlushToItemValue(pval.ItemUID);
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
//						}
//					} else if (pval.BeforeAction == true & pval.ItemUID == "MstCode" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (Strings.Len(Strings.Trim(oForm.Items.Item("MstCode").Specific.String)) == 0) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//					//----------------------------------------------------
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//----------------------------------------------------
//					if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_ZPY343H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY343H = null;
//						//UPGRADE_NOTE: oDS_ZPY343L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY343L = null;
//						//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oCheck = null;
//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//					}
//					break;
//				//et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					if (pval.BeforeAction == false) {
//						//UPGRADE_WARNING: MDC_SetMod.Get_UserName() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.DataSources.UserDataSources.Item("User").ValueEx = MDC_SetMod.Get_UserName(ref oDS_ZPY343H.GetValue("UserSign", 0));
//						//                Call Matrix_TitleSetting
//						oMat1.AutoResizeColumns();
//						FormItemEnabled();
//						if (oDS_ZPY343H.GetValue("U_Check", 0) == "Y") {
//							oForm.Items.Item("Mat1").Enabled = false;
//						} else {
//							oForm.Items.Item("Mat1").Enabled = true;
//						}
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
//						if (Strings.Trim(oDS_ZPY343H.GetValue("U_Check", 0)) == "Y") {
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
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@ZPY343H", ref "DocNum");
//						////접속자 권한에 따른 사업장 보기
//						break;
//					default:
//						return;

//						break;
//				}
//			} else {
//				switch (pval.MenuUID) {
//					case "1287":
//						//// 복제
//						break;
//					case "1283":
//						//// 제거
//						FormItemEnabled();
//						break;
//					case "1281":
//					case "1282":
//						//// 찾기, 추가
//						FormItemEnabled();
//						Matrix_TitleSetting();
//						Clear_UserDb();
//						//           oForm.Update
//						oForm.Items.Item("JsnYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1288": // TODO: to "1291"
//						break;
//					case "1293":
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
//			SAPbouiCOM.EditText oEdit = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////디비데이터 소스 개체 할당
//			oDS_ZPY343H = oForm.DataSources.DBDataSources("@ZPY343H");
//			oDS_ZPY343L = oForm.DataSources.DBDataSources("@ZPY343L");

//			oForm.DataSources.UserDataSources.Add("User", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			oEdit = oForm.Items.Item("UserSign").Specific;
//			oEdit.DataBind.SetBound(true, "", "User");

//			////사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			////부서
//			oForm.Items.Item("DptCode").DisplayDesc = true;

//			//// 직책
//			oCombo = oForm.Items.Item("StpCode").Specific;
//			sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("StpCode").DisplayDesc = true;

//			/// 체크버튼(잠금체크 )
//			//Call oForm.DataSources.UserDataSources.Add("CheckDS", dt_SHORT_TEXT, 1)

//			oCheck = oForm.Items.Item("Check1").Specific;
//			//oCheck.DataBind.SetBound True, "", "CheckDS"
//			oCheck.ValOff = "N";
//			oCheck.ValOn = "Y";

//			//
//			oMat1 = oForm.Items.Item("Mat1").Specific;

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
//			int iRow = 0;
//			/// 계산시
//			double GWASEE = 0;
//			double MONPAY = 0;
//			//UPGRADE_WARNING: SUMArr 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//			double[] SUMArr = new double[52];
//			ZPAY_g_EmpID MstInfo = default(ZPAY_g_EmpID);

//			if (Strings.Left(oUID, 3) != "Col") {
//				switch (oUID) {
//					case "MstCode":
//						//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//							oDS_ZPY343H.SetValue("U_MstCode", 0, "");
//							oDS_ZPY343H.SetValue("U_EmpId", 0, "");
//							oDS_ZPY343H.SetValue("U_MstName", 0, "");
//						} else {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_ZPY343H.SetValue("U_MstCode", 0, Strings.UCase(oForm.Items.Item(oUID).Specific.String));
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: MstInfo 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							MstInfo = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//							oDS_ZPY343H.SetValue("U_EmpId", 0, MstInfo.EmpID);
//							oDS_ZPY343H.SetValue("U_MstName", 0, MstInfo.MSTNAM);
//						}
//						oForm.Items.Item(oUID).Update();
//						break;

//					case "JsnYear":
//						//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							MDC_Globals.ZPAY_GBL_JSNYER.Value = oForm.Items.Item(oUID).Specific.String;
//						}
//						Matrix_TitleSetting();
//						break;
//				}
//				return;
//			}
//			GWASEE = 0;
//			MONPAY = 0;
//			for (iRow = 1; iRow <= 51; iRow++) {
//				SUMArr[iRow] = 0;
//			}

//			//------------------------------------------------------
//			oMat1.FlushToDataSource();
//			switch (oUID) {
//				case "Col1":
//				case "Col2":
//				case "Col3":
//				case "Col4":
//				case "Col6":
//				case "Col7":
//				case "Col8":
//				case "Col9":
//				case "Col19":
//				case "Col20":
//				case "Col21":
//				case "Col22":
//				case "Col23":
//				case "Col24":
//				case "Col25":
//				case "Col26":
//				case "Col27":
//				case "Col28":
//				case "Col29":
//				case "Col30":
//				case "Col31":
//				case "Col32":
//				case "Col33":
//				case "Col34":
//				case "Col35":
//				case "Col36":
//				case "Col37":
//				case "Col38":
//				case "Col39":
//				case "Col40":
//				case "Col41":
//				case "Col42":
//				case "Col43":
//				case "Col44":
//				case "Col45":
//				case "Col46":
//				case "Col47":
//				case "Col48":
//				case "Col49":
//				case "Col50":
//				case "Col51":
//				case "Col52":

//					oMat1.FlushToDataSource();

//					oDS_ZPY343L.Offset = oRow - 1;
//					/// 과세총계
//					GWASEE = Conversion.Val(oDS_ZPY343L.GetValue("U_GwaPay", oRow - 1));
//					GWASEE = GWASEE + Conversion.Val(oDS_ZPY343L.GetValue("U_GwaBns", oRow - 1));
//					GWASEE = GWASEE + Conversion.Val(oDS_ZPY343L.GetValue("U_InJBns", oRow - 1));
//					GWASEE = GWASEE + Conversion.Val(oDS_ZPY343L.GetValue("U_JUSBNS", oRow - 1));
//					///2007주식행사이익추가
//					GWASEE = GWASEE + Conversion.Val(oDS_ZPY343L.GetValue("U_URIBNS", oRow - 1));
//					///2009우리사주조합인출추가
//					oDS_ZPY343L.SetValue("U_GwaSee", oRow - 1, Convert.ToString(GWASEE));
//					/// 월지급총계
//					MONPAY = GWASEE;
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa02", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa03", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa05", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa06", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwu03", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa04", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa07", oRow - 1));

//					/// 2009년추가.
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGG01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH05", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH06", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH07", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH08", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH09", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH10", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH11", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH12", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH13", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGI01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGK01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGM01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGM02", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGM03", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGO01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGQ01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGS01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGT01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGX01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY01", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY02", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY03", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY20", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY21", oRow - 1));
//					MONPAY = MONPAY + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGZ01", oRow - 1));

//					oDS_ZPY343L.SetValue("U_JigTotal", oRow - 1, Convert.ToString(MONPAY));
//					oMat1.SetLineData(oRow);
//					break;
//			}
//			/// 항목별 합계
//			for (iRow = 1; iRow <= oMat1.VisualRowCount - 1; iRow++) {
//				oDS_ZPY343L.Offset = iRow - 1;
//				SUMArr[1] = SUMArr[1] + Conversion.Val(oDS_ZPY343L.GetValue("U_GwaPay", iRow - 1));
//				/// 과세급여
//				SUMArr[2] = SUMArr[2] + Conversion.Val(oDS_ZPY343L.GetValue("U_GwaBns", iRow - 1));
//				/// 과세상여
//				SUMArr[3] = SUMArr[3] + Conversion.Val(oDS_ZPY343L.GetValue("U_InJBns", iRow - 1));
//				/// 인정상여
//				SUMArr[4] = SUMArr[4] + Conversion.Val(oDS_ZPY343L.GetValue("U_GwaSee", iRow - 1));
//				/// 과세총계
//				SUMArr[5] = SUMArr[5] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa01", iRow - 1));
//				/// 생산비과
//				SUMArr[6] = SUMArr[6] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa02", iRow - 1));
//				/// 기타비과
//				SUMArr[7] = SUMArr[7] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa03", iRow - 1));
//				/// 국외비과(\)
//				SUMArr[8] = SUMArr[8] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa04", iRow - 1));
//				/// 국외비과($)
//				SUMArr[9] = SUMArr[9] + Conversion.Val(oDS_ZPY343L.GetValue("U_JiGTotal", iRow - 1));
//				/// 월지급총액
//				SUMArr[10] = SUMArr[10] + Conversion.Val(oDS_ZPY343L.GetValue("U_KukAmt", iRow - 1));
//				/// 국민연금
//				SUMArr[11] = SUMArr[11] + Conversion.Val(oDS_ZPY343L.GetValue("U_MedAmt", iRow - 1));
//				/// 건강보험
//				SUMArr[12] = SUMArr[12] + Conversion.Val(oDS_ZPY343L.GetValue("U_GBHAmt", iRow - 1));
//				/// 고용보험
//				SUMArr[13] = SUMArr[13] + Conversion.Val(oDS_ZPY343L.GetValue("U_GabGun", iRow - 1));
//				/// 소득세
//				SUMArr[14] = SUMArr[14] + Conversion.Val(oDS_ZPY343L.GetValue("U_Jumin", iRow - 1));
//				/// 주민세
//				SUMArr[15] = SUMArr[15] + Conversion.Val(oDS_ZPY343L.GetValue("U_NonTK", iRow - 1));
//				/// 농특세
//				//SUMArr(16) = SUMArr(16) + Val(oDS_ZPY343L.GetValue("U_GamSe", iRow - 1)) '/ 감면세액
//				SUMArr[17] = SUMArr[17] + Conversion.Val(oDS_ZPY343L.GetValue("U_GbuAmt", iRow - 1));
//				/// 기부금
//				SUMArr[18] = SUMArr[18] + Conversion.Val(oDS_ZPY343L.GetValue("U_JUSBNS", iRow - 1));
//				/// 주식행사이익
//				SUMArr[19] = SUMArr[19] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa05", iRow - 1));
//				/// 기타비과
//				SUMArr[20] = SUMArr[20] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa06", iRow - 1));
//				/// 연구비과세
//				/// 2009년추가
//				SUMArr[21] = SUMArr[21] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGG01", iRow - 1));
//				SUMArr[22] = SUMArr[22] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH01", iRow - 1));
//				SUMArr[23] = SUMArr[23] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH05", iRow - 1));
//				SUMArr[24] = SUMArr[24] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH06", iRow - 1));
//				SUMArr[25] = SUMArr[25] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH07", iRow - 1));
//				SUMArr[26] = SUMArr[26] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH08", iRow - 1));
//				SUMArr[27] = SUMArr[27] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH09", iRow - 1));
//				SUMArr[28] = SUMArr[28] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH10", iRow - 1));
//				SUMArr[29] = SUMArr[29] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH11", iRow - 1));
//				SUMArr[30] = SUMArr[30] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH12", iRow - 1));
//				SUMArr[31] = SUMArr[31] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGH13", iRow - 1));
//				SUMArr[32] = SUMArr[32] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGI01", iRow - 1));
//				SUMArr[33] = SUMArr[33] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGK01", iRow - 1));
//				SUMArr[34] = SUMArr[34] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGM01", iRow - 1));
//				SUMArr[35] = SUMArr[35] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGM02", iRow - 1));
//				SUMArr[36] = SUMArr[36] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGM03", iRow - 1));
//				SUMArr[37] = SUMArr[37] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGO01", iRow - 1));
//				SUMArr[38] = SUMArr[38] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGQ01", iRow - 1));
//				SUMArr[39] = SUMArr[39] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGS01", iRow - 1));
//				SUMArr[40] = SUMArr[40] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGT01", iRow - 1));
//				SUMArr[41] = SUMArr[41] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGX01", iRow - 1));
//				SUMArr[42] = SUMArr[42] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY01", iRow - 1));
//				SUMArr[43] = SUMArr[43] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY02", iRow - 1));
//				SUMArr[44] = SUMArr[44] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY03", iRow - 1));
//				SUMArr[45] = SUMArr[45] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY20", iRow - 1));
//				SUMArr[46] = SUMArr[46] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGZ01", iRow - 1));

//				SUMArr[47] = SUMArr[47] + Conversion.Val(oDS_ZPY343L.GetValue("U_URIBNS", iRow - 1));
//				SUMArr[48] = SUMArr[48] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwu03", iRow - 1));
//				SUMArr[49] = SUMArr[49] + Conversion.Val(oDS_ZPY343L.GetValue("U_BiGwa07", iRow - 1));
//				SUMArr[50] = SUMArr[50] + Conversion.Val(oDS_ZPY343L.GetValue("U_NGYAMT", iRow - 1));

//				//// 2010년 추가
//				SUMArr[51] = SUMArr[51] + Conversion.Val(oDS_ZPY343L.GetValue("U_BIGY21", iRow - 1));

//			}
//			iRow = 12;
//			oDS_ZPY343L.Offset = iRow;
//			oDS_ZPY343L.SetValue("U_GwaPay", iRow, Convert.ToString(SUMArr[1]));
//			oDS_ZPY343L.SetValue("U_GwaBns", iRow, Convert.ToString(SUMArr[2]));
//			oDS_ZPY343L.SetValue("U_InJBns", iRow, Convert.ToString(SUMArr[3]));
//			oDS_ZPY343L.SetValue("U_JUSBNS", iRow, Convert.ToString(SUMArr[18]));
//			oDS_ZPY343L.SetValue("U_GwaSee", iRow, Convert.ToString(SUMArr[4]));
//			oDS_ZPY343L.SetValue("U_BiGwa01", iRow, Convert.ToString(SUMArr[5]));
//			oDS_ZPY343L.SetValue("U_BiGwa02", iRow, Convert.ToString(SUMArr[6]));
//			oDS_ZPY343L.SetValue("U_BiGwa03", iRow, Convert.ToString(SUMArr[7]));
//			oDS_ZPY343L.SetValue("U_BiGwa04", iRow, Convert.ToString(SUMArr[8]));
//			oDS_ZPY343L.SetValue("U_BiGwa05", iRow, Convert.ToString(SUMArr[19]));
//			oDS_ZPY343L.SetValue("U_BiGwa06", iRow, Convert.ToString(SUMArr[20]));

//			oDS_ZPY343L.SetValue("U_JiGTotal", iRow, Convert.ToString(SUMArr[9]));
//			oDS_ZPY343L.SetValue("U_KukAmt", iRow, Convert.ToString(SUMArr[10]));
//			oDS_ZPY343L.SetValue("U_MedAmt", iRow, Convert.ToString(SUMArr[11]));
//			oDS_ZPY343L.SetValue("U_GBHAmt", iRow, Convert.ToString(SUMArr[12]));
//			oDS_ZPY343L.SetValue("U_GabGun", iRow, Convert.ToString(SUMArr[13]));
//			oDS_ZPY343L.SetValue("U_Jumin", iRow, Convert.ToString(SUMArr[14]));
//			oDS_ZPY343L.SetValue("U_NonTK", iRow, Convert.ToString(SUMArr[15]));
//			//Call oDS_ZPY343L.SetValue("U_GamSe", iRow, SUMArr(16))
//			oDS_ZPY343L.SetValue("U_GbuAmt", iRow, Convert.ToString(SUMArr[17]));
//			/// 2009년추가
//			oDS_ZPY343L.SetValue("U_BIGG01", iRow, Convert.ToString(SUMArr[21]));
//			oDS_ZPY343L.SetValue("U_BIGH01", iRow, Convert.ToString(SUMArr[22]));
//			oDS_ZPY343L.SetValue("U_BIGH05", iRow, Convert.ToString(SUMArr[23]));
//			oDS_ZPY343L.SetValue("U_BIGH06", iRow, Convert.ToString(SUMArr[24]));
//			oDS_ZPY343L.SetValue("U_BIGH07", iRow, Convert.ToString(SUMArr[25]));
//			oDS_ZPY343L.SetValue("U_BIGH08", iRow, Convert.ToString(SUMArr[26]));
//			oDS_ZPY343L.SetValue("U_BIGH09", iRow, Convert.ToString(SUMArr[27]));
//			oDS_ZPY343L.SetValue("U_BIGH10", iRow, Convert.ToString(SUMArr[28]));
//			oDS_ZPY343L.SetValue("U_BIGH11", iRow, Convert.ToString(SUMArr[29]));
//			oDS_ZPY343L.SetValue("U_BIGH12", iRow, Convert.ToString(SUMArr[30]));
//			oDS_ZPY343L.SetValue("U_BIGH13", iRow, Convert.ToString(SUMArr[31]));
//			oDS_ZPY343L.SetValue("U_BIGI01", iRow, Convert.ToString(SUMArr[32]));
//			oDS_ZPY343L.SetValue("U_BIGK01", iRow, Convert.ToString(SUMArr[33]));
//			oDS_ZPY343L.SetValue("U_BIGM01", iRow, Convert.ToString(SUMArr[34]));
//			oDS_ZPY343L.SetValue("U_BIGM02", iRow, Convert.ToString(SUMArr[35]));
//			oDS_ZPY343L.SetValue("U_BIGM03", iRow, Convert.ToString(SUMArr[36]));
//			oDS_ZPY343L.SetValue("U_BIGO01", iRow, Convert.ToString(SUMArr[37]));
//			oDS_ZPY343L.SetValue("U_BIGQ01", iRow, Convert.ToString(SUMArr[38]));
//			oDS_ZPY343L.SetValue("U_BIGS01", iRow, Convert.ToString(SUMArr[39]));
//			oDS_ZPY343L.SetValue("U_BIGT01", iRow, Convert.ToString(SUMArr[40]));
//			oDS_ZPY343L.SetValue("U_BIGX01", iRow, Convert.ToString(SUMArr[41]));
//			oDS_ZPY343L.SetValue("U_BIGY01", iRow, Convert.ToString(SUMArr[42]));
//			oDS_ZPY343L.SetValue("U_BIGY02", iRow, Convert.ToString(SUMArr[43]));
//			oDS_ZPY343L.SetValue("U_BIGY03", iRow, Convert.ToString(SUMArr[44]));
//			oDS_ZPY343L.SetValue("U_BIGY20", iRow, Convert.ToString(SUMArr[45]));
//			oDS_ZPY343L.SetValue("U_BIGZ01", iRow, Convert.ToString(SUMArr[46]));

//			oDS_ZPY343L.SetValue("U_URIBNS", iRow, Convert.ToString(SUMArr[47]));
//			oDS_ZPY343L.SetValue("U_BiGwu03", iRow, Convert.ToString(SUMArr[48]));
//			oDS_ZPY343L.SetValue("U_BiGwa07", iRow, Convert.ToString(SUMArr[49]));
//			oDS_ZPY343L.SetValue("U_NGYAMT", iRow, Convert.ToString(SUMArr[50]));

//			oDS_ZPY343L.SetValue("U_BIGY21", iRow, Convert.ToString(SUMArr[51]));

//			oMat1.SetLineData(iRow + 1);
//			oForm.Update();
//		}

//		private void FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.OptionBtn optBtn = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//				oForm.Items.Item("DocNum").Enabled = true;
//				oForm.Items.Item("JsnYear").Enabled = true;
//				oForm.Items.Item("MstCode").Enabled = true;
//				oForm.Items.Item("MstName").Enabled = true;
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				if (Strings.Len(MDC_Globals.ZPAY_GBL_JSNYER.Value) > 0) {
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("JsnYear").Specific.VALUE = MDC_Globals.ZPAY_GBL_JSNYER.Value;
//				}

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				////기본사항 - 부서 (사업장에 따른 부서변경)
//				oCombo = oForm.Items.Item("DptCode").Specific;

//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}

//				if (!string.IsNullOrEmpty(oDS_ZPY343H.GetValue("U_CLTCOD", 0))) {
//					sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//					//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//					sQry = sQry + " ORDER BY U_Code";
//					MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//				}

//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				oForm.Items.Item("DocNum").Enabled = true;
//				oForm.Items.Item("JsnYear").Enabled = true;
//				oForm.Items.Item("MstCode").Enabled = true;
//				oForm.Items.Item("MstName").Enabled = false;
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				if (Strings.Len(MDC_Globals.ZPAY_GBL_JSNYER.Value) > 0) {
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("JsnYear").Specific.VALUE = MDC_Globals.ZPAY_GBL_JSNYER.Value;
//				}

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				////부서
//				oCombo = oForm.Items.Item("TeamCode").Specific;
//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "-");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}

//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				oForm.Items.Item("DocNum").Enabled = false;
//				oForm.Items.Item("JsnYear").Enabled = false;
//				oForm.Items.Item("MstCode").Enabled = false;
//				oForm.Items.Item("MstName").Enabled = false;
//				oForm.Items.Item("CLTCOD").Enabled = false;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);
//			}
//		}

//		private bool MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//----------------------------------------------------------------------------
//			//저장할 데이터의 유효성을 점검한다
//			//----------------------------------------------------------------------------
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			int k = 0;
//			short ErrNum = 0;
//			string Chk_Data = null;

//			ErrNum = 0;
//			/// 헤더부분 체크
//			switch (true) {
//				case string.IsNullOrEmpty(oDS_ZPY343H.GetValue("U_JsnYear", 0)):
//					ErrNum = 4;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY343H.GetValue("U_MstCode", 0)):
//					ErrNum = 5;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY343H.GetValue("U_CLTCOD", 0)):
//					ErrNum = 7;
//					goto Error_Message;
//					break;
//			}

//			//----------------------------------------------------------------------------
//			//화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
//			//----------------------------------------------------------------------------
//			oMat1.FlushToDataSource();

//			//// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
//			if (oMat1.RowCount == 1) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			//----------------------------------------------------------------------------
//			////마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
//			////oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
//			//----------------------------------------------------------------------------
//			//// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
//			for (i = 0; i <= oMat1.VisualRowCount - 2; i++) {
//				oDS_ZPY343L.Offset = i;
//				if (Conversion.Val(oDS_ZPY343L.GetValue("U_JIGTotal", i)) != 0 & string.IsNullOrEmpty(Strings.Trim(oDS_ZPY343L.GetValue("U_JIGDate", i)))) {
//					ErrNum = 2;
//					oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (!string.IsNullOrEmpty(Strings.Trim(oDS_ZPY343L.GetValue("U_JIGDate", i))) & Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_ZPY343L.GetValue("U_JIGDate", i), "0000-00-00")) == false) {
//					ErrNum = 6;
//					oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else {
//					//------------------------------------------------
//					//중복체크작업
//					//------------------------------------------------
//					Chk_Data = Strings.Trim(oDS_ZPY343L.GetValue("U_JIGDate", i));
//					for (k = i + 1; k <= oMat1.VisualRowCount - 2; k++) {
//						oDS_ZPY343L.Offset = k;
//						if (Strings.Trim(Chk_Data) == Strings.Trim(oDS_ZPY343L.GetValue("U_JIGDate", k)) & Conversion.Val(oDS_ZPY343L.GetValue("U_JIGTotal", i)) != 0) {
//							ErrNum = 3;
//							oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							goto Error_Message;
//						}
//					}
//				}
//			}

//			//--------------------------------------------------------------------------------------------
//			////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//			////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//			//--------------------------------------------------------------------------------------------
//			//  oDS_ZPY343L.RemoveRecord oDS_ZPY343L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제

//			//--------------------------------------------------------------------------------------------
//			//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
//			//--------------------------------------------------------------------------------------------
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급일자가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급일자가 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급일자를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}
//		private void ShowSource(ref string DocNum)
//		{
//			oForm.Items.Item("DocNum").Enabled = true;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("DocNum").Specific.VALUE = DocNum;

//			oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//		}
//		private void Clear_UserDb()
//		{
//			oForm.DataSources.UserDataSources.Item("User").Value = "";
//		}

////---------------------------------------------------------------------------------------
//// Procedure : Matrix_TitleSetting
//// DateTime  : 2009-12-29 11:08
//// Author    : Choi Dong Kwon
//// Purpose   : 비과세 코드 설정의 데이터를 읽어와 Matrix Title의 비과세 컬럼에 대하여 표시여부와 타이틀을 적용한다
////             단, 찾기모드에서는 전체컬럼을 표시한다
////---------------------------------------------------------------------------------------
////
//		private void Matrix_TitleSetting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string JSNYER = null;
//			string sQry = null;
//			int iCol = 0;
//			string COLNAM = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			JSNYER = Strings.Trim(oDS_ZPY343H.GetValue("U_JsnYear", 0));

//			if ((Information.IsNumeric(JSNYER) == false | string.IsNullOrEmpty(JSNYER)) == false) {
//				//// 2008년 이전
//				if (Conversion.Val(JSNYER) <= 2008) {
//					sQry = "SELECT  U_BTXCOD, U_BTXNAM, ISNULL(U_MONCHK,'Y') AS U_MONCHK " + "FROM    [@ZPY117L] T0 " + "WHERE   T0.CODE = (SELECT MAX(CODE) FROM [@ZPY117L] T1 WHERE CODE <= '" + JSNYER + "') " + "AND     T0.CODE <= '2008' " + "ORDER   BY U_BTXCOD ";
//				//// 2009년 이후
//				} else {
//					sQry = "SELECT  U_BTXCOD, U_BTXNAM, ISNULL(U_MONCHK,'Y') AS U_MONCHK  " + "FROM    [@ZPY117L] T0 " + "WHERE   T0.CODE = (SELECT MAX(CODE) FROM [@ZPY117L] T1 WHERE CODE <= '" + JSNYER + "') " + "AND     T0.CODE >= '2009' " + "ORDER   BY U_BTXCOD ";
//				}
//				oRecordSet.DoQuery(sQry);
//			}

//			oForm.Freeze(true);
//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE | Information.IsNumeric(JSNYER) == false | string.IsNullOrEmpty(JSNYER) | oRecordSet.RecordCount == 0) {
//				oMat1.Columns.Item("Col6").TitleObject.Caption = "비과세-생산";
//				oMat1.Columns.Item("Col7").TitleObject.Caption = "비과세-식대.차량";
//				oMat1.Columns.Item("Col8").TitleObject.Caption = "비과세-국외";
//				oMat1.Columns.Item("Col23").TitleObject.Caption = "비과세-외국인";
//				oMat1.Columns.Item("Col21").TitleObject.Caption = "비과세-보육";
//				oMat1.Columns.Item("Col20").TitleObject.Caption = "비과세-연구";
//				oMat1.Columns.Item("Col9").TitleObject.Caption = "비과세-기타(제출)";
//				oMat1.Columns.Item("Col24").TitleObject.Caption = "비과세-기타(미제출)";

//				oMat1.Columns.Item("Col26").TitleObject.Caption = "비과세(G01)";
//				oMat1.Columns.Item("Col27").TitleObject.Caption = "비과세(H01)";
//				oMat1.Columns.Item("Col28").TitleObject.Caption = "비과세(H05)";
//				oMat1.Columns.Item("Col29").TitleObject.Caption = "비과세(H06)";
//				oMat1.Columns.Item("Col30").TitleObject.Caption = "비과세(H07)";
//				oMat1.Columns.Item("Col31").TitleObject.Caption = "비과세(H08)";
//				oMat1.Columns.Item("Col32").TitleObject.Caption = "비과세(H09)";
//				oMat1.Columns.Item("Col33").TitleObject.Caption = "비과세(H10)";
//				oMat1.Columns.Item("Col34").TitleObject.Caption = "비과세(H11)";
//				oMat1.Columns.Item("Col35").TitleObject.Caption = "비과세(H12)";
//				oMat1.Columns.Item("Col36").TitleObject.Caption = "비과세(H13)";
//				oMat1.Columns.Item("Col37").TitleObject.Caption = "비과세(I01)";
//				oMat1.Columns.Item("Col38").TitleObject.Caption = "비과세(K01)";
//				oMat1.Columns.Item("Col39").TitleObject.Caption = "비과세(M01)";
//				oMat1.Columns.Item("Col40").TitleObject.Caption = "비과세(M02)";
//				oMat1.Columns.Item("Col41").TitleObject.Caption = "비과세(M03)";
//				oMat1.Columns.Item("Col42").TitleObject.Caption = "비과세(O01)";
//				oMat1.Columns.Item("Col43").TitleObject.Caption = "비과세(Q01)";
//				oMat1.Columns.Item("Col44").TitleObject.Caption = "비과세(S01)";
//				oMat1.Columns.Item("Col45").TitleObject.Caption = "비과세(T01)";
//				oMat1.Columns.Item("Col46").TitleObject.Caption = "비과세(X01)";
//				oMat1.Columns.Item("Col47").TitleObject.Caption = "비과세(R10,Y01)";
//				oMat1.Columns.Item("Col48").TitleObject.Caption = "비과세(Y02)";
//				oMat1.Columns.Item("Col49").TitleObject.Caption = "비과세(Y03)";
//				oMat1.Columns.Item("Col50").TitleObject.Caption = "비과세(Y22,Y20)";
//				oMat1.Columns.Item("Col51").TitleObject.Caption = "비과세(Z01)";
//				oMat1.Columns.Item("Col52").TitleObject.Caption = "비과세(Y21)";

//				//// 비과세 컬럼 전체 표시
//				for (iCol = 6; iCol <= 52; iCol++) {
//					if ((iCol >= 6 & iCol <= 9) | iCol == 20 | iCol == 21 | iCol == 23 | iCol == 24 | (iCol >= 26 & iCol <= 52)) {
//						oMat1.Columns.Item("Col" + Convert.ToString(iCol)).Visible = false;
//					}
//				}

//			} else if (Conversion.Val(JSNYER) <= 2008) {
//				while (!(oRecordSet.EoF)) {
//					//// 비과세 코드에 따라 컬럼UID 확인
//					switch (oRecordSet.Fields.Item("U_BTXCOD").Value) {
//						case "01":
//							COLNAM = "Col6";
//							/// 비과세-생산
//							break;
//						case "02":
//							COLNAM = "Col7";
//							/// 비과세-식대.차량
//							break;
//						case "03":
//							COLNAM = "Col8";
//							/// 비과세-국외
//							break;
//						case "04":
//							COLNAM = "Col23";
//							/// 비과세-외국인
//							break;
//						case "05":
//							COLNAM = "Col9";
//							/// 비과세-기타(제출)
//							break;
//						case "06":
//							COLNAM = "Col20";
//							/// 비과세-연구
//							break;
//						case "07":
//							COLNAM = "Col21";
//							/// 비과세-보육
//							break;
//						case "08":
//							COLNAM = "Col24";
//							/// 비과세-기타(미제출)
//							break;
//					}

//					//// 컬럼명, 화면표시 여부 적용
//					var _with2 = oMat1.Columns.Item(COLNAM);
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					_with2.TitleObject.Caption = oRecordSet.Fields.Item("U_BTXNAM").Value;
//					if (oRecordSet.Fields.Item("U_MONCHK").Value == "N") {
//						_with2.Visible = false;
//					} else {
//						_with2.Visible = true;
//					}

//					oRecordSet.MoveNext();
//				}

//				//// 2009년 이후 생성 컬럼은 화면 표시 안함
//				for (iCol = 26; iCol <= 52; iCol++) {
//					oMat1.Columns.Item("Col" + Convert.ToString(iCol)).Visible = false;
//				}
//			} else {
//				while (!(oRecordSet.EoF)) {
//					//// 비과세 코드에 따라 컬럼UID 확인
//					switch (oRecordSet.Fields.Item("U_BTXCOD").Value) {
//						case "G01":
//							COLNAM = "Col26";
//							//// 비과세(G01)
//							break;
//						case "H01":
//							COLNAM = "Col27";
//							//// 비과세(H01)
//							break;
//						case "H05":
//							COLNAM = "Col28";
//							//// 비과세(H05)
//							break;
//						case "H06":
//							COLNAM = "Col29";
//							//// 비과세(H06)
//							break;
//						case "H07":
//							COLNAM = "Col30";
//							//// 비과세(H07)
//							break;
//						case "H08":
//							COLNAM = "Col31";
//							//// 비과세(H08)
//							break;
//						case "H09":
//							COLNAM = "Col32";
//							//// 비과세(H09)
//							break;
//						case "H10":
//							COLNAM = "Col33";
//							//// 비과세(H10)
//							break;
//						case "H11":
//							COLNAM = "Col34";
//							//// 비과세(H11)
//							break;
//						case "H12":
//							COLNAM = "Col35";
//							//// 비과세(H12)
//							break;
//						case "H13":
//							COLNAM = "Col36";
//							//// 비과세(H13)
//							break;
//						case "I01":
//							COLNAM = "Col37";
//							//// 비과세(I01)
//							break;
//						case "K01":
//							COLNAM = "Col38";
//							//// 비과세(K01)
//							break;
//						case "M01":
//							COLNAM = "Col39";
//							//// 비과세(M01)
//							break;
//						case "M02":
//							COLNAM = "Col40";
//							//// 비과세(M02)
//							break;
//						case "M03":
//							COLNAM = "Col41";
//							//// 비과세(M03)
//							break;
//						case "O01":
//							COLNAM = "Col42";
//							//// 비과세(O01)
//							break;
//						case "Q01":
//							COLNAM = "Col43";
//							//// 비과세(Q01)
//							break;
//						case "S01":
//							COLNAM = "Col44";
//							//// 비과세(S01)
//							break;
//						case "T01":
//							COLNAM = "Col45";
//							//// 비과세(T01)
//							break;
//						case "X01":
//							COLNAM = "Col46";
//							//// 비과세(X01)
//							break;
//						case "Y01":
//						case "R10":
//							COLNAM = "Col47";
//							//// 비과세(Y01)
//							break;
//						case "Y02":
//							COLNAM = "Col48";
//							//// 비과세(Y02)
//							break;
//						case "Y03":
//							COLNAM = "Col49";
//							//// 비과세(Y03)
//							break;
//						case "Y20":
//						case "Y22":
//							COLNAM = "Col50";
//							//// 비과세(Y20)
//							break;
//						case "Y21":
//							COLNAM = "Col52";
//							//// 비과세(Y21)
//							break;
//						case "Z01":
//							COLNAM = "Col51";
//							//// 비과세(Z01)
//							break;
//					}

//					//// 컬럼명, 화면표시 여부 적용
//					var _with1 = oMat1.Columns.Item(COLNAM);
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					_with1.TitleObject.Caption = oRecordSet.Fields.Item("U_BTXNAM").Value;
//					if (oRecordSet.Fields.Item("U_MONCHK").Value == "Y") {
//						_with1.Visible = true;
//					} else {
//						_with1.Visible = false;
//					}

//					oRecordSet.MoveNext();
//				}
//				oMat1.Columns.Item("Col6").Visible = false;
//				//비과세-생산
//				oMat1.Columns.Item("Col7").Visible = true;
//				//비과세-식대차량
//				oMat1.Columns.Item("Col8").Visible = false;
//				//비과세-국외
//				oMat1.Columns.Item("Col23").Visible = false;
//				//비과세-외국인
//				oMat1.Columns.Item("Col21").Visible = false;
//				//비과세-보육
//				oMat1.Columns.Item("Col20").Visible = false;
//				//비과세-연구
//				oMat1.Columns.Item("Col9").Visible = true;
//				//비과세-기타(제출)
//				oMat1.Columns.Item("Col24").Visible = true;
//				//비과세-기타(미제출)

//			}

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Matrix_TitleSetting 실행 중 오류가 발생하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
//	}
//}
