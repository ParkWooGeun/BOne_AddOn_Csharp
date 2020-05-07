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
//	[System.Runtime.InteropServices.ProgId("ZPY504_NET.ZPY504")]
//	public class ZPY504
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY504.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 정산결과조회
//////  FormType       : 2010110504
//////  Create Date    : 2005.01.20
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//			//시스템코드 헤더
//		private SAPbouiCOM.DBDataSource oDS_ZPY504H;

//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;
//		private string oOLDCHK;

//		private void FormItemEnabled()
//		{
//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				oForm.Items.Item("DocNum").Enabled = true;
//				oForm.Items.Item("JSNYER").Enabled = true;
//				oForm.Items.Item("MSTCOD").Enabled = true;
//				oForm.Items.Item("MSTNAM").Enabled = true;
//				oForm.Items.Item("JSNGBN").Enabled = true;
//				oForm.Items.Item("ENDCHK").Enabled = true;
//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				oForm.Items.Item("DocNum").Enabled = false;
//				oForm.Items.Item("JSNYER").Enabled = false;
//				oForm.Items.Item("MSTCOD").Enabled = false;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("JSNGBN").Enabled = false;
//				//// 급여월마감된것은 비활성화
//				oOLDCHK = oDS_ZPY504H.GetValue("U_ENDCHK", 0);
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(U_ENDCHK, U_JOBYER, [ZPY509L], ' & oDS_ZPY504H.GetValue(U_JSNYER, 0) & ',  AND Code = ' & oDS_ZPY504H.GetValue(U_CLTCOD, 0) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (MDC_SetMod.Get_ReData(ref "U_ENDCHK", ref "U_JOBYER", ref "[@ZPY509L]", ref "'" + oDS_ZPY504H.GetValue("U_JSNYER", 0) + "'", ref " AND Code = '" + oDS_ZPY504H.GetValue("U_CLTCOD", 0) + "'") == "Y") {
//					oForm.Items.Item("ENDCHK").Enabled = false;
//				} else {
//					oForm.Items.Item("ENDCHK").Enabled = true;
//				}

//			}
//		}

//		private void UserFormLink(ref string LinkName)
//		{
//			object oTmpObject = null;

//			//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (oForm.Items.Item("CLTCOD").Specific.Selected == null)
//				return;

//			if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//				if (LinkName == "LBtn1") {
//					//            DOCNUM = MDC_SetMod.Get_ReData("DocNum", "U_JSNYER", "[@ZPY501H]", _
//					//'                    "'" & oForm.Items("JSNYER").Specific.String & "'", _
//					//'                    " AND U_MSTCOD = '" & oForm.Items("MSTCOD").Specific.String & "'")
//					//            Call MDC_SetMod.UserLink(DOCNUM, "ZPY501", "D")
//					oTmpObject = new ZPY501();
//					//ZPY501(2000060501) :
//					//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oTmpObject.LoadForm(oForm.Items.Item("JSNYER").Specific.String, oForm.Items.Item("MSTCOD").Specific.String, oForm.Items.Item("CLTCOD").Specific.Selected.Value);
//					MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//					MDC_SetMod.Get_FormColor();
//				} else if (LinkName == "LBtn2") {
//					oTmpObject = new ZPY505();
//					//ZPY505(2000060505) :
//					//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oTmpObject.LoadForm(oForm.Items.Item("JSNYER").Specific.String, oForm.Items.Item("MSTCOD").Specific.String, oForm.Items.Item("CLTCOD").Specific.Selected.Value);
//					MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//					MDC_SetMod.Get_FormColor();
//				} else if (LinkName == "LBtn3") {
//					oTmpObject = new ZPY506();
//					//ZPY506(2000060506) :
//					//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oTmpObject.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oTmpObject.LoadForm(oForm.Items.Item("JSNYER").Specific.String, oForm.Items.Item("MSTCOD").Specific.String, oForm.Items.Item("CLTCOD").Specific.Selected.Value);
//					MDC_Globals.Sbo_Application.Forms.ActiveForm.Select();
//					MDC_SetMod.Get_FormColor();
//				}
//			}
//			//UPGRADE_NOTE: oTmpObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oTmpObject = null;
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
//						if (pval.ItemUID == "CBtn1") {
//							if (oForm.Items.Item("MSTCOD").Enabled == true) {
//								oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						//// Link Button
//						} else if (pval.ItemUID == "LBtn1" | pval.ItemUID == "LBtn2" | pval.ItemUID == "LBtn3") {
//							UserFormLink(ref (pval.ItemUID));
//						} else if (pval.ItemUID == "1" & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)) {
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							////추가 및 업데이트시에
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							if (HeaderSpaceLineDel() == false) {
//								BubbleEvent = false;
//								return;
//							}
//						}
//					} else {
//						if (pval.ItemUID == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//							FormItemEnabled();
//						} else if (pval.ItemUID == "Folder1") {
//							oForm.PaneLevel = 1;
//						} else if (pval.ItemUID == "Folder2") {
//							oForm.PaneLevel = 2;
//						}
//					}
//					break;

//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUID == "MSTCOD" | pval.ItemUID == "JSNYER" | pval.ItemUID == "JSNMON")) {
//						FlushToItemValue(pval.ItemUID);
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
//					break;

//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					if (pval.BeforeAction == true & pval.ItemUID == "JSNYER" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (Strings.Len(Strings.Trim(oForm.Items.Item("JSNYER").Specific.String)) == 0) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도를 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					} else if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String))) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						} else {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + oForm.Items.Item(pval.ItemUID).Specific.String + "'", ref "") == true) {
//								oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
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
//						//UPGRADE_NOTE: oDS_ZPY504H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY504H = null;
//					}
//					break;
//				//et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					break;
//			}

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Raise_FormItemEvent_Error:", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private bool HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string DocNum = null;

//			ErrNum = 0;
//			/// Check
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items(JSNGBN).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case Strings.Trim(oDS_ZPY504H.GetValue("U_ENDCHK", 0)) == "Y" & Strings.Trim(oOLDCHK) == "Y":
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY504H.GetValue("U_MSTCOD", 0)):
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("JSNGBN").Specific.Selected == null:
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JSNMON").Specific.Value)):
//					ErrNum = 4;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("SINYMM").Specific.Value)):
//					ErrNum = 5;
//					goto Error_Message;
//					break;
//				case oForm.Items.Item("CLTCOD").Specific.Selected == null:
//					ErrNum = 6;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JSNYER").Specific.Value)):
//					ErrNum = 7;
//					goto Error_Message;
//					break;
//			}

//			DocNum = Exist_YN(ref oDS_ZPY504H.GetValue("U_JSNYER", 0), ref oDS_ZPY504H.GetValue("U_MSTCOD", 0), ref oDS_ZPY504H.GetValue("U_CLTCOD", 0));
//			if (!string.IsNullOrEmpty(Strings.Trim(DocNum)) & Strings.Trim(oDS_ZPY504H.GetValue("DocNum", 0)) != Strings.Trim(DocNum)) {
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("정산 구분은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지급 연월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사 코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("정산 년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 8) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("문서번호" + DocNum + " 와(과) 데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
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

//			sQry = "SELECT Top 1 T1.DocNum FROM [@ZPY504H] T1 ";
//			sQry = sQry + " WHERE T1.U_JSNYER = N'" + Strings.Trim(JOBYER) + "'";
//			sQry = sQry + "   AND T1.U_MSTCOD = N'" + Strings.Trim(MSTCOD) + "'";
//			sQry = sQry + "   AND T1.U_CLTCOD = N'" + Strings.Trim(CLTCOD) + "'";
//			oRecordSet.DoQuery(sQry);

//			while (!(oRecordSet.EoF)) {
//				//UPGRADE_WARNING: oRecordSet().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);

//			switch (oUID) {
//				case "JSNYER":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oUID).Specific.String))) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MDC_Globals.ZPAY_GBL_JSNYER.Value = oForm.Items.Item(oUID).Specific.String;
//					} else {
//						oDS_ZPY504H.SetValue("U_JSNYER", 0, MDC_Globals.ZPAY_GBL_JSNYER.Value);
//					}
//					oForm.Items.Item(oUID).Update();
//					break;
//				case "JSNMON":
//					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oUID).Specific.Value))) {
//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oDS_ZPY504H.SetValue("U_JSNMON", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(oUID).Specific.Value, "00"));
//					}
//					oForm.Items.Item(oUID).Update();
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oDS_ZPY504H.SetValue("U_MSTCOD", 0, "");
//						oDS_ZPY504H.SetValue("U_MSTNAM", 0, "");
//						oDS_ZPY504H.SetValue("U_EmpID", 0, "");
//					} else {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oDS_ZPY504H.SetValue("U_MSTCOD", 0, Strings.UCase(oForm.Items.Item(oUID).Specific.String));
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMast = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oDS_ZPY504H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//						oDS_ZPY504H.SetValue("U_EmpID", 0, oMast.EmpID);
//					}

//					oForm.Items.Item("MSTNAM").Update();
//					oForm.Items.Item("EmpID").Update();
//					oForm.Items.Item(oUID).Update();
//					break;
//			}

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
//						if (Strings.Trim(oDS_ZPY504H.GetValue("U_ENDCHK", 0)) == "Y") {
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
//						/// 찾기, 추가
//						FormItemEnabled();
//						oDS_ZPY504H.SetValue("U_JSNYER", 0, MDC_Globals.ZPAY_GBL_JSNYER.Value);
//						oForm.Items.Item("JSNYER").Update();
//						oForm.Items.Item("JSNYER").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1288": // TODO: to "1291"
//						FormItemEnabled();
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
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm(ref string DocNum = "")
//		{
//			//Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY504.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY504_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY504");
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

//		private void ShowSource(ref string DocNum)
//		{
//			oForm.Items.Item("DocNum").Enabled = true;
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("DocNum").Specific.Value = DocNum;

//			oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//		}

////*******************************************************************
////
////*******************************************************************
//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.CheckBox oCheck = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Folder oFolder = null;
//			short iCol = 0;

//			oForm.PaneLevel = 2;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oFolder = oForm.Items.Item("Folder1").Specific;

//			//// DB DataSource 개체 할당
//			oDS_ZPY504H = oForm.DataSources.DBDataSources("@ZPY504H");

//			//// Folder
//			oForm.DataSources.UserDataSources.Add("FolderD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			for (iCol = 1; iCol <= 2; iCol++) {
//				oFolder = oForm.Items.Item("Folder" + iCol).Specific;
//				oFolder.DataBind.SetBound(true, "", "FolderD");
//				if (iCol == 2) {
//					oFolder.Select();
//					oFolder.GroupWith(("Folder" + iCol - 1));

//				}
//				oForm.Items.Item("Folder" + iCol).Enabled = true;
//				oForm.Items.Item("Folder" + iCol).AffectsFormMode = false;
//			}

//			//// 정산구분
//			oCombo = oForm.Items.Item("JSNGBN").Specific;
//			oCombo.ValidValues.Add("1", "연말정산(재직자)");
//			oCombo.ValidValues.Add("2", "중도정산(퇴직자)");


//			//// 사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}

//			/// Check 버튼
//			oCheck = oForm.Items.Item("ENDCHK").Specific;
//			oCheck.ValOff = "N";
//			oCheck.ValOn = "Y";

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
//	}
//}
