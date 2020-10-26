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
//	internal class PH_PY012
//	{
//////********************************************************************************
//////  File           : PH_PY012.cls
//////  Module         : 인사관리 > 인사
//////  Desc           : 출장등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY012A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY012B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//			//이전 출장구분 값 저장용(전역변수)
//		string DestDivValue;

//		public void LoadForm(string oFormDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY012.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY012_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY012");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			oForm.DataBrowser.BrowseBy = "DocEntry";

//			oForm.Freeze(true);
//			PH_PY012_CreateItems();
//			PH_PY012_EnableMenus();
//			PH_PY012_SetDocument(oFormDocEntry01);
//			//    Call PH_PY012_FormResize

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY012_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oDS_PH_PY012A = oForm.DataSources.DBDataSources("@PH_PY012A");
//			oDS_PH_PY012B = oForm.DataSources.DBDataSources("@PH_PY012B");


//			oMat1 = oForm.Items.Item("Mat01").Specific;

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
//			oMat1.AutoResizeColumns();


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			//사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			//출장구분
//			oCombo = oForm.Items.Item("DestDiv").Specific;
//			//    oCombo.ValidValues.Add "", ""
//			oCombo.ValidValues.Add("01", "공용");
//			oCombo.ValidValues.Add("02", "국내출장");
//			oCombo.ValidValues.Add("03", "해외출장");
//			//    oCombo.Select 0, psk_Index
//			oForm.Items.Item("DestDiv").DisplayDesc = true;

//			//출장지역
//			oCombo = oForm.Items.Item("DestCode").Specific;
//			sQry = "            SELECT      '' AS [Code],";
//			sQry = sQry + "                 '' AS [Name],";
//			sQry = sQry + "                 -1 AS [Seq]";
//			sQry = sQry + "  UNION ALL";
//			sQry = sQry + "  SELECT      T1.U_Code AS [Code],";
//			sQry = sQry + "                 T1.U_CodeNm AS [Name],";
//			sQry = sQry + "                 T1.U_Seq AS [Seq]";
//			sQry = sQry + "  FROM       [@PS_HR200H] AS T0";
//			sQry = sQry + "                 INNER JOIN";
//			sQry = sQry + "                 [@PS_HR200L] AS T1";
//			sQry = sQry + "                     ON T0.Code = T1.Code";
//			sQry = sQry + "  WHERE      T0.Code = 'P217'";
//			sQry = sQry + "                 AND T1.U_UseYN = 'Y'";
//			sQry = sQry + "  ORDER BY  Seq";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("DestCode").DisplayDesc = true;

//			//매트릭스-차량구분
//			oColumn = oMat1.Columns.Item("Vehicle");
//			oColumn.ValidValues.Add("", "");
//			sQry = "           SELECT      U_Code AS [Code],";
//			sQry = sQry + "                U_CodeNm As Name";
//			sQry = sQry + " FROM       [@PS_HR200L]";
//			sQry = sQry + " WHERE      Code = 'P218'";
//			sQry = sQry + "                AND U_UseYN = 'Y'";
//			sQry = sQry + " ORDER BY  U_Seq";
//			MDC_SetMod.GP_MatrixSetMatComboList(ref oColumn, ref sQry, ref Convert.ToString(false), ref Convert.ToString(false));
//			oColumn.DisplayDesc = true;

//			//매트릭스-유종
//			oColumn = oMat1.Columns.Item("FuelType");
//			oColumn.ValidValues.Add("", "");
//			oColumn.ValidValues.Add("1", "휘발유");
//			oColumn.ValidValues.Add("2", "가스");
//			oColumn.ValidValues.Add("3", "경유");
//			oColumn.DisplayDesc = true;

//			//매트릭스-통화
//			oColumn = oMat1.Columns.Item("Currency");
//			oColumn.ValidValues.Add("", "");
//			sQry = "           SELECT      T0.CurrCode AS [Code],";
//			sQry = sQry + "                T0.CurrName AS [Name]";
//			sQry = sQry + " FROM       [OCRN] AS T0";
//			sQry = sQry + " ORDER BY  CurrCode";
//			MDC_SetMod.GP_MatrixSetMatComboList(ref oColumn, ref sQry, ref Convert.ToString(false), ref Convert.ToString(false));

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY012_CreateItems_Error:

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY012_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", false);
//			//// 삭제
//			oForm.EnableMenu("1287", false);
//			//// 복제
//			oForm.EnableMenu("1286", true);
//			//// 닫기
//			oForm.EnableMenu("1284", true);
//			//// 취소
//			oForm.EnableMenu("1293", true);
//			//// 행삭제

//			return;
//			PH_PY012_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY012_SetDocument(string oFormDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFormDocEntry01))) {
//				PH_PY012_FormItemEnabled();
//				PH_PY012_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY012_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY012_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY012_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;

//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//폼 DocEntry 세팅
//				PH_PY012_FormClear();

//				//        '사업장 세팅
//				//        Call oDS_PH_PY012A.setValue("U_CLTCOD", 0, MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'"))
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				//출장구분 세팅
//				oCombo = oForm.Items.Item("DestDiv").Specific;
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY012_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string BPLID = null;
//			//사업장 코드
//			string StdYear = null;
//			//기준년도
//			string StdMonth = null;
//			//기준월

//			string FuelType = null;
//			//유종
//			double FuelPrc = 0;
//			//단가
//			double Distance = 0;
//			//거리
//			double FoodNum = 0;
//			//식수
//			int TransExp = 0;
//			//교통비
//			double InsurExp = 0;
//			//보험료
//			double AirpExp = 0;
//			//공항세
//			double DayExp = 0;
//			//일비
//			double LodgExp = 0;
//			//숙박비
//			double FoodExp = 0;
//			//식비
//			short TotFoodExp = 0;
//			//총식비
//			double ParkExp = 0;
//			//주차비
//			double TollExp = 0;
//			//도로비
//			double TotalExp = 0;
//			//합계

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY012_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}

//								////해야할일 작업
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (PH_PY012_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}
//								////해야할일 작업

//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//							}
//						}
//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (pval.ActionSuccess == true) {
//									PH_PY012_FormItemEnabled();
//									PH_PY012_AddMatrixRow();
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (pval.ActionSuccess == true) {
//									PH_PY012_FormItemEnabled();
//									PH_PY012_AddMatrixRow();
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//								if (pval.ActionSuccess == true) {
//									PH_PY012_FormItemEnabled();
//								}
//							}
//						}
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2

//					if (pval.BeforeAction == true) {

//						if (pval.ItemUID == "Mat01") {

//							//                    If pval.ColUID = "Name" And pval.CharPressed = "9" Then
//							//
//							//                        If oMat1.Columns.Item("Name").Cells(pval.Row).Specific.Value = "" Then
//							//                            Call Sbo_Application.ActivateMenuItem("7425")
//							//                            BubbleEvent = False
//							//                        End If
//							//
//							//                    End If

//						} else if (pval.ItemUID == "MSTCOD" & pval.CharPressed == Convert.ToDouble("9")) {

//							//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value)) {
//								MDC_Globals.Sbo_Application.ActivateMenuItem("7425");
//								BubbleEvent = false;
//							}

//						}

//					} else if (pval.Before_Action == false) {

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
//							if (pval.Row > 0) {
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = pval.ColUID;
//								oLastColRow = pval.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pval.ItemUID;
//							oLastColUID = "";
//							oLastColRow = 0;
//							break;
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5

//					oForm.Freeze(true);

//					if (pval.BeforeAction == true) {

//						if (pval.ItemUID == "DestDiv") {

//							if (oMat1.RowCount > 1) {

//								//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								DestDivValue = oForm.Items.Item("DestDiv").Specific.Value;
//								//"계속 진행하겠습니까?" 라는 질문에 No를 선택했을 경우 DestDiv 콤보박스의 Value를 돌려놓기 위해 이전 값을 저장

//							}

//						}

//					} else if (pval.BeforeAction == false) {

//						if (pval.ItemChanged == true) {

//							if (pval.ItemUID == "Mat01") {

//								if (pval.ColUID == "Vehicle") {

//									PH_PY012_AddMatrixRow();

//								//유종을 선택하면 유종에 따른 단가를 조회
//								} else if (pval.ColUID == "FuelType") {

//									oMat1.FlushToDataSource();

//									BPLID = Strings.Trim(oDS_PH_PY012A.GetValue("U_CLTCOD", 0));
//									//사업장
//									StdYear = Strings.Left(Strings.Trim(oDS_PH_PY012A.GetValue("U_FrDate", 0)), 4);
//									StdMonth = Strings.Mid(Strings.Trim(oDS_PH_PY012A.GetValue("U_FrDate", 0)), 5, 2);
//									FuelType = Strings.Trim(oDS_PH_PY012B.GetValue("U_FuelType", pval.Row - 1));

//									FuelPrc = PH_PY012_GetFuelPrc(BPLID, StdYear, StdMonth, FuelType);

//									oDS_PH_PY012B.SetValue("U_FuelPrc", pval.Row - 1, Convert.ToString(FuelPrc));

//									oMat1.LoadFromDataSource();

//								}

//							} else if (pval.ItemUID == "DestDiv") {

//								if (oMat1.RowCount > 1) {

//									if (MDC_Globals.Sbo_Application.MessageBox("저장하지 않은 데이터는 모두 삭제됩니다. 계속 진행하겠습니까?", 1, "Yes", "No") == 1) {

//										oMat1.Clear();
//										oMat1.FlushToDataSource();
//										oMat1.LoadFromDataSource();

//										PH_PY012_AddMatrixRow();

//									} else {

//										//이전 상태의 콤보값으로 되돌림
//										oDS_PH_PY012A.SetValue("U_DestDiv", 0, DestDivValue);
//										//콤보박스를 바로 Select 하면 ComboSelect 이벤트가 발생하기 때문에 UDO를 이용하여 Binding
//										oForm.Freeze(false);

//										return;

//									}

//								}

//								//출장구분에 따라 매트릭스 컬럼 Visible 설정
//								PH_PY012_ColumnStatusChange();

//							}

//							oMat1.AutoResizeColumns();

//						}
//					}

//					oForm.Freeze(false);
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Mat01":
//								if (pval.Row > 0) {
//									oMat1.SelectRow(pval.Row, true, false);
//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Mat01":
//								if (pval.Row > 0) {
//									oLastItemUID = pval.ItemUID;
//									oLastColUID = pval.ColUID;
//									oLastColRow = pval.Row;
//								}
//								break;
//							default:
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pval.BeforeAction == false) {

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//					////9
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {

//						if (pval.ItemChanged == true) {

//							switch (pval.ItemUID) {

//								case "MSTCOD":

//									//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oDS_PH_PY012A.SetValue("U_MSTNAM", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pval.ItemUID).Specific.Value + "'"));
//									break;

//								case "Mat01":

//									//UPGRADE_WARNING: oForm(DestDiv).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//차량구분 선택 시
//									if (pval.ColUID == "Vehicle" & oForm.Items.Item("DestDiv").Specific.Value == "01") {

//										//                                oMat1.FlushToDataSource

//										//                                oMat1.LoadFromDataSource

//										PH_PY012_AddMatrixRow();

//										//UPGRADE_WARNING: oForm(DestDiv).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//기간시작
//									} else if (pval.ColUID == "FrDate" & (oForm.Items.Item("DestDiv").Specific.Value == "02" | oForm.Items.Item("DestDiv").Specific.Value == "03")) {

//										//                                oMat1.FlushToDataSource
//										//
//										//                                oMat1.LoadFromDataSource

//										PH_PY012_AddMatrixRow();

//									//단가 입력 시
//									} else if (pval.ColUID == "FuelPrc") {

//										oMat1.FlushToDataSource();

//										FuelPrc = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FuelPrc", pval.Row - 1));
//										Distance = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_Distance", pval.Row - 1));

//										TransExp = FuelPrc * Distance * 0.1;

//										oDS_PH_PY012B.SetValue("U_TransExp", pval.Row - 1, Convert.ToString(TransExp));

//										//합계 계산(교통비+일비+식비+주차비+도로비)
//										DayExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1));
//										//일비
//										FoodExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1));
//										//식비
//										ParkExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1));
//										//주차비
//										TollExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1));
//										//도로비

//										TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp;

//										oDS_PH_PY012B.SetValue("U_TotalExp", pval.Row - 1, Convert.ToString(TotalExp));
//										//합계

//										oMat1.LoadFromDataSource();

//									//거리 입력 시
//									} else if (pval.ColUID == "Distance") {

//										oMat1.FlushToDataSource();

//										FuelPrc = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FuelPrc", pval.Row - 1));
//										Distance = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_Distance", pval.Row - 1));

//										TransExp = FuelPrc * Distance * 0.1;

//										oDS_PH_PY012B.SetValue("U_TransExp", pval.Row - 1, Convert.ToString(TransExp));

//										//합계 계산(교통비+일비+식비+주차비+도로비)
//										DayExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1));
//										//일비
//										FoodExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1));
//										//식비
//										ParkExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1));
//										//주차비
//										TollExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1));
//										//도로비

//										TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp;

//										oDS_PH_PY012B.SetValue("U_TotalExp", pval.Row - 1, Convert.ToString(TotalExp));
//										//합계

//										oMat1.LoadFromDataSource();

//									//식수 입력 시
//									} else if (pval.ColUID == "FoodNum") {

//										oMat1.FlushToDataSource();

//										FoodNum = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FoodNum", pval.Row - 1));
//										FoodExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1));

//										TotFoodExp = FoodNum * FoodExp;

//										oDS_PH_PY012B.SetValue("U_FoodExp", pval.Row - 1, Convert.ToString(FoodExp));

//										//합계 계산(교통비+일비+식비+주차비+도로비)
//										TransExp = Convert.ToInt32(oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1));
//										//교통비
//										DayExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1));
//										//일비
//										ParkExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1));
//										//주차비
//										TollExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1));
//										//도로비

//										TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp;

//										oDS_PH_PY012B.SetValue("U_TotalExp", pval.Row - 1, Convert.ToString(TotalExp));
//										//합계

//										oMat1.LoadFromDataSource();

//									//식비 입력 시
//									} else if (pval.ColUID == "FoodExp") {

//										oMat1.FlushToDataSource();

//										FoodNum = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FoodNum", pval.Row - 1));
//										FoodExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1));

//										TotFoodExp = FoodNum * FoodExp;

//										oDS_PH_PY012B.SetValue("U_FoodExp", pval.Row - 1, Convert.ToString(FoodExp));

//										//합계 계산(교통비+일비+식비+주차비+도로비)
//										TransExp = Convert.ToInt32(oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1));
//										//교통비
//										DayExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1));
//										//일비
//										ParkExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1));
//										//주차비
//										TollExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1));
//										//도로비

//										TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp;

//										oDS_PH_PY012B.SetValue("U_TotalExp", pval.Row - 1, Convert.ToString(TotalExp));
//										//합계

//										oMat1.LoadFromDataSource();

//									//주차비 입력 시
//									} else if (pval.ColUID == "ParkExp") {

//										//                                oMat1.FlushToDataSource
//										//
//										//                                '합계(교통비+일비+식비+주차비+도로비+숙박비+보험료+공항세)
//										//                                TransExp = oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1) '교통비
//										//                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
//										//                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1) '식비
//										//                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
//										//                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비
//										//
//										//                                TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp
//										//
//										//                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계
//										//
//										//                                oMat1.LoadFromDataSource

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									//도로비 입력 시
//									} else if (pval.ColUID == "TollExp") {

//										//                                oMat1.FlushToDataSource
//										//
//										//                                '합계(교통비+일비+식비+주차비+도로비+숙박비+보험료+공항세)
//										//                                TransExp = oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1) '교통비
//										//                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
//										//                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1) '식비
//										//                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
//										//                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비
//										//
//										//                                TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp
//										//
//										//                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계
//										//
//										//                                oMat1.LoadFromDataSource

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									//일비 입력 시
//									} else if (pval.ColUID == "DayExp") {

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									//숙박비 입력 시
//									} else if (pval.ColUID == "LodgExp") {

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									//교통비 입력 시
//									} else if (pval.ColUID == "TransExp") {

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									//보험료 입력 시
//									} else if (pval.ColUID == "InsurExp") {

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									//식비 입력 시
//									} else if (pval.ColUID == "FoodExp") {

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									//공항세 입력 시
//									} else if (pval.ColUID == "AirpExp") {

//										PH_PY012_CalculateTotalExp(pval.Row - 1);
//										//합계 계산

//									}

//									oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//									oMat1.AutoResizeColumns();
//									break;

//							}

//						}

//					}
//					oForm.Freeze(false);
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						oMat1.LoadFromDataSource();

//						PH_PY012_FormItemEnabled();
//						PH_PY012_AddMatrixRow();
//						oMat1.AutoResizeColumns();

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//					////12
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//					////16
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY012A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY012A = null;
//						//UPGRADE_NOTE: oDS_PH_PY012B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY012B = null;

//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//					////20
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////21
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {

//						oMat1.AutoResizeColumns();

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//					////22
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//					////23
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					if (pval.BeforeAction == true) {

//					} else if (pval.Before_Action == false) {
//						//                If pval.ItemUID = "Code" Then
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY012A", "Code")
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//					////37
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//					////38
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_Drag:
//					////39
//					break;

//			}

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			oForm.Freeze((false));
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short loopCount = 0;
//			double FeeTot = 0;
//			double TuiTot = 0;
//			double Total = 0;

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
//							BubbleEvent = false;
//							return;
//						}
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					case "1293":
//						break;
//					case "1281":
//						break;
//					case "1282":
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						break;

//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY012_FormItemEnabled();
//						PH_PY012_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY012_FormItemEnabled();
//						PH_PY012_AddMatrixRow();
//						oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						PH_PY012_ColumnStatusChange();
//						break;
//					case "1282":
//						////문서추가
//						PH_PY012_FormItemEnabled();
//						PH_PY012_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY012_FormItemEnabled();
//						PH_PY012_AddMatrixRow();
//						PH_PY012_ColumnStatusChange();
//						break;
//					case "1293":
//						//// 행삭제

//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							oMat1.FlushToDataSource();

//							while ((i <= oDS_PH_PY012B.Size - 1)) {
//								if (string.IsNullOrEmpty(oDS_PH_PY012B.GetValue("U_LineNum", i))) {
//									oDS_PH_PY012B.RemoveRecord((i));
//									i = 0;
//								} else {
//									i = i + 1;
//								}
//							}

//							for (i = 0; i <= oDS_PH_PY012B.Size; i++) {
//								oDS_PH_PY012B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//							}

//							oMat1.LoadFromDataSource();
//						}
//						PH_PY012_AddMatrixRow();
//						break;

//					//            '복제
//					//            Case "1287"
//					//
//					//                Call oForm.Freeze(True)
//					//                Call oDS_PH_PY012A.setValue("DocEntry", 0, "")
//					//
//					//                For i = 0 To oMat1.VisualRowCount - 1
//					//                    Call oMat1.FlushToDataSource
//					//                    Call oDS_PH_PY012B.setValue("DocEntry", i, "")
//					//                    Call oDS_PH_PY012B.setValue("U_PayYN", i, "N")
//					//                    Call oMat1.LoadFromDataSource
//					//                Next i
//					//                Call oForm.Freeze(False)

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((BusinessObjectInfo.BeforeAction == true)) {
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
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
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
//			return;
//			Raise_FormDataEvent_Error:


//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			switch (pval.ItemUID) {
//				case "Mat01":
//					if (pval.Row > 0) {
//						oLastItemUID = pval.ItemUID;
//						oLastColUID = pval.ColUID;
//						oLastColRow = pval.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pval.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY012_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			string FirstColumnName = null;

//			//출장구분에 따른 매트릭스 컬럼 설정이 변경되므로 첫 컬럼의 컬럼명을 저장
//			//UPGRADE_WARNING: oForm.Items(DestDiv).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (oForm.Items.Item("DestDiv").Specific.Value == "01") {

//				FirstColumnName = "U_Vehicle";

//				//UPGRADE_WARNING: oForm.Items(DestDiv).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("DestDiv").Specific.Value == "02" | oForm.Items.Item("DestDiv").Specific.Value == "03") {

//				FirstColumnName = "U_FrDate";

//			}


//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY012B.GetValue(FirstColumnName, oRow - 1)))) {
//					if (oDS_PH_PY012B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY012B.InsertRecord((oRow));
//					}
//					oDS_PH_PY012B.Offset = oRow;
//					oDS_PH_PY012B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					//라인번호
//					oDS_PH_PY012B.SetValue("U_FrDate", oRow, "");
//					//기간시작일자
//					oDS_PH_PY012B.SetValue("U_ToDate", oRow, "");
//					//기간종료일자
//					oDS_PH_PY012B.SetValue("U_Destinat", oRow, "");
//					//행선지
//					oDS_PH_PY012B.SetValue("U_PayPstg", oRow, Convert.ToString(0));
//					//지급율
//					oDS_PH_PY012B.SetValue("U_Currency", oRow, "");
//					//통화
//					oDS_PH_PY012B.SetValue("U_Rate", oRow, Convert.ToString(0));
//					//환율
//					oDS_PH_PY012B.SetValue("U_Vehicle", oRow, "");
//					//차량구분
//					oDS_PH_PY012B.SetValue("U_FuelType", oRow, "");
//					//유종
//					oDS_PH_PY012B.SetValue("U_FuelPrc", oRow, Convert.ToString(0));
//					//단가
//					oDS_PH_PY012B.SetValue("U_Distance", oRow, Convert.ToString(0));
//					//거리
//					oDS_PH_PY012B.SetValue("U_FoodNum", oRow, Convert.ToString(0));
//					//식수
//					oDS_PH_PY012B.SetValue("U_TransExp", oRow, Convert.ToString(0));
//					//교통비
//					oDS_PH_PY012B.SetValue("U_InsurExp", oRow, Convert.ToString(0));
//					//보험료
//					oDS_PH_PY012B.SetValue("U_AirpExp", oRow, Convert.ToString(0));
//					//공항세
//					oDS_PH_PY012B.SetValue("U_DayExp", oRow, Convert.ToString(0));
//					//일비
//					oDS_PH_PY012B.SetValue("U_FDayExp", oRow, Convert.ToString(0));
//					//일비(외화)
//					oDS_PH_PY012B.SetValue("U_LodgExp", oRow, Convert.ToString(0));
//					//숙박비
//					oDS_PH_PY012B.SetValue("U_FoodExp", oRow, Convert.ToString(0));
//					//숙박비(외화)
//					oDS_PH_PY012B.SetValue("U_ParkExp", oRow, Convert.ToString(0));
//					//주차비
//					oDS_PH_PY012B.SetValue("U_TollExp", oRow, Convert.ToString(0));
//					//도로비
//					oDS_PH_PY012B.SetValue("U_TotalExp", oRow, Convert.ToString(0));
//					//합계
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY012B.Offset = oRow - 1;
//					oDS_PH_PY012B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					//라인번호
//					oDS_PH_PY012B.SetValue("U_FrDate", oRow - 1, "");
//					//기간시작일자
//					oDS_PH_PY012B.SetValue("U_ToDate", oRow - 1, "");
//					//기간종료일자
//					oDS_PH_PY012B.SetValue("U_Destinat", oRow - 1, "");
//					//행선지
//					oDS_PH_PY012B.SetValue("U_PayPstg", oRow - 1, Convert.ToString(0));
//					//지급율
//					oDS_PH_PY012B.SetValue("U_Currency", oRow - 1, "");
//					//통화
//					oDS_PH_PY012B.SetValue("U_Rate", oRow - 1, Convert.ToString(0));
//					//환율
//					oDS_PH_PY012B.SetValue("U_Vehicle", oRow - 1, "");
//					//차량구분
//					oDS_PH_PY012B.SetValue("U_FuelType", oRow - 1, "");
//					//유종
//					oDS_PH_PY012B.SetValue("U_FuelPrc", oRow - 1, Convert.ToString(0));
//					//단가
//					oDS_PH_PY012B.SetValue("U_Distance", oRow - 1, Convert.ToString(0));
//					//거리
//					oDS_PH_PY012B.SetValue("U_FoodNum", oRow - 1, Convert.ToString(0));
//					//식수
//					oDS_PH_PY012B.SetValue("U_TransExp", oRow - 1, Convert.ToString(0));
//					//교통비
//					oDS_PH_PY012B.SetValue("U_InsurExp", oRow - 1, Convert.ToString(0));
//					//보험료
//					oDS_PH_PY012B.SetValue("U_AirpExp", oRow - 1, Convert.ToString(0));
//					//공항세
//					oDS_PH_PY012B.SetValue("U_DayExp", oRow - 1, Convert.ToString(0));
//					//일비
//					oDS_PH_PY012B.SetValue("U_FDayExp", oRow - 1, Convert.ToString(0));
//					//일비(외화)
//					oDS_PH_PY012B.SetValue("U_LodgExp", oRow - 1, Convert.ToString(0));
//					//숙박비
//					oDS_PH_PY012B.SetValue("U_FoodExp", oRow - 1, Convert.ToString(0));
//					//숙박비(외화)
//					oDS_PH_PY012B.SetValue("U_ParkExp", oRow - 1, Convert.ToString(0));
//					//주차비
//					oDS_PH_PY012B.SetValue("U_TollExp", oRow - 1, Convert.ToString(0));
//					//도로비
//					oDS_PH_PY012B.SetValue("U_TotalExp", oRow - 1, Convert.ToString(0));
//					//합계
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY012B.Offset = oRow;
//				oDS_PH_PY012B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				//라인번호
//				oDS_PH_PY012B.SetValue("U_FrDate", oRow, "");
//				//기간시작일자
//				oDS_PH_PY012B.SetValue("U_ToDate", oRow, "");
//				//기간종료일자
//				oDS_PH_PY012B.SetValue("U_Destinat", oRow, "");
//				//행선지
//				oDS_PH_PY012B.SetValue("U_PayPstg", oRow, Convert.ToString(0));
//				//지급율
//				oDS_PH_PY012B.SetValue("U_Currency", oRow, "");
//				//통화
//				oDS_PH_PY012B.SetValue("U_Rate", oRow, Convert.ToString(0));
//				//환율
//				oDS_PH_PY012B.SetValue("U_Vehicle", oRow, "");
//				//차량구분
//				oDS_PH_PY012B.SetValue("U_FuelType", oRow, "");
//				//유종
//				oDS_PH_PY012B.SetValue("U_FuelPrc", oRow, Convert.ToString(0));
//				//단가
//				oDS_PH_PY012B.SetValue("U_Distance", oRow, Convert.ToString(0));
//				//거리
//				oDS_PH_PY012B.SetValue("U_FoodNum", oRow, Convert.ToString(0));
//				//식수
//				oDS_PH_PY012B.SetValue("U_TransExp", oRow, Convert.ToString(0));
//				//교통비
//				oDS_PH_PY012B.SetValue("U_InsurExp", oRow, Convert.ToString(0));
//				//보험료
//				oDS_PH_PY012B.SetValue("U_AirpExp", oRow, Convert.ToString(0));
//				//공항세
//				oDS_PH_PY012B.SetValue("U_DayExp", oRow, Convert.ToString(0));
//				//일비
//				oDS_PH_PY012B.SetValue("U_FDayExp", oRow, Convert.ToString(0));
//				//일비(외화)
//				oDS_PH_PY012B.SetValue("U_LodgExp", oRow, Convert.ToString(0));
//				//숙박비
//				oDS_PH_PY012B.SetValue("U_FoodExp", oRow, Convert.ToString(0));
//				//숙박비(외화)
//				oDS_PH_PY012B.SetValue("U_ParkExp", oRow, Convert.ToString(0));
//				//주차비
//				oDS_PH_PY012B.SetValue("U_TollExp", oRow, Convert.ToString(0));
//				//도로비
//				oDS_PH_PY012B.SetValue("U_TotalExp", oRow, Convert.ToString(0));
//				//합계
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY012_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY012_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY012'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.Value = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
//			}
//			return;
//			PH_PY012_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

////Private Sub PH_PY012_FormResize()
////'******************************************************************************
////'Function ID : PH_PY012_FormResize()
////'해당모듈 : PH_PY012
////'기능 : Form의 크기 변경 시 아이템들의 위치 및 크기 동적 변경
////'인수 : 없음
////'반환값 : 없음
////'특이사항 : 없음
////'******************************************************************************
////On Error GoTo PH_PY012_FormResize_Error
////
////    oForm.Items("Mat01").Left = 10
////    oForm.Items("Mat01").Top = 110
////    oForm.Items("Mat01").Height = oForm.Height - 330
////    oForm.Items("Mat01").Width = oForm.Width - 20
////
////
////    Exit Sub
////
////PH_PY012_FormResize_Error:
////    Sbo_Application.SetStatusBarMessage "PH_PY012_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
////End Sub

//		private void PH_PY012_ColumnStatusChange()
//		{
//			//******************************************************************************
//			//Function ID : PH_PY012_FormResize()
//			//해당모듈 : PH_PY012
//			//기능 : 출장구분에 따른 매트릭스 컬럼 내용 변경
//			//인수 : 없음
//			//반환값 : 없음
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			//UPGRADE_WARNING: oForm.Items(DestDiv).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//공용
//			if (oForm.Items.Item("DestDiv").Specific.Value == "01") {

//				oMat1.Columns.Item("FrDate").Visible = false;
//				//기간시작일자
//				oMat1.Columns.Item("ToDate").Visible = false;
//				//기간종료일자
//				oMat1.Columns.Item("Destinat").Visible = false;
//				//행선지
//				oMat1.Columns.Item("PayPstg").Visible = false;
//				//지급율
//				oMat1.Columns.Item("Currency").Visible = false;
//				//통화
//				oMat1.Columns.Item("Rate").Visible = false;
//				//환율
//				oMat1.Columns.Item("Vehicle").Visible = true;
//				//차량구분
//				oMat1.Columns.Item("FuelType").Visible = true;
//				//유종
//				oMat1.Columns.Item("FuelPrc").Visible = true;
//				//단가
//				oMat1.Columns.Item("Distance").Visible = true;
//				//거리
//				oMat1.Columns.Item("FoodNum").Visible = true;
//				//식수
//				oMat1.Columns.Item("TransExp").Visible = true;
//				//교통비
//				oMat1.Columns.Item("InsurExp").Visible = false;
//				//보험료
//				oMat1.Columns.Item("AirpExp").Visible = false;
//				//공항세
//				oMat1.Columns.Item("DayExp").Visible = true;
//				//일비
//				oMat1.Columns.Item("FDayExp").Visible = false;
//				//일비(외화)
//				oMat1.Columns.Item("LodgExp").Visible = false;
//				//숙박비
//				oMat1.Columns.Item("FLodgExp").Visible = false;
//				//숙박비(외화)
//				oMat1.Columns.Item("FoodExp").Visible = true;
//				//식비
//				oMat1.Columns.Item("ParkExp").Visible = true;
//				//주차비
//				oMat1.Columns.Item("TollExp").Visible = true;
//				//도로비
//				oMat1.Columns.Item("TotalExp").Visible = true;
//				//합계

//				//        oMat1.Columns("TransExp").Editable = False '교통비 Editable 설정
//				//        oMat1.Columns("FoodExp").Editable = False '식비 Editable 설정

//				oMat1.AutoResizeColumns();

//				//UPGRADE_WARNING: oForm.Items(DestDiv).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//국내출장
//			} else if (oForm.Items.Item("DestDiv").Specific.Value == "02") {

//				oMat1.Columns.Item("FrDate").Visible = true;
//				//기간시작일자
//				oMat1.Columns.Item("ToDate").Visible = true;
//				//기간종료일자
//				oMat1.Columns.Item("Destinat").Visible = true;
//				//행선지
//				oMat1.Columns.Item("PayPstg").Visible = true;
//				//지급율
//				oMat1.Columns.Item("Currency").Visible = false;
//				//통화
//				oMat1.Columns.Item("Rate").Visible = false;
//				//환율
//				oMat1.Columns.Item("Vehicle").Visible = false;
//				//차량구분
//				oMat1.Columns.Item("FuelType").Visible = false;
//				//유종
//				oMat1.Columns.Item("FuelPrc").Visible = false;
//				//단가
//				oMat1.Columns.Item("Distance").Visible = false;
//				//거리
//				oMat1.Columns.Item("FoodNum").Visible = false;
//				//식수
//				oMat1.Columns.Item("TransExp").Visible = true;
//				//교통비
//				oMat1.Columns.Item("InsurExp").Visible = true;
//				//보험료
//				oMat1.Columns.Item("AirpExp").Visible = true;
//				//공항세
//				oMat1.Columns.Item("DayExp").Visible = true;
//				//일비
//				oMat1.Columns.Item("FDayExp").Visible = false;
//				//일비(외화)
//				oMat1.Columns.Item("LodgExp").Visible = true;
//				//숙박비
//				oMat1.Columns.Item("FLodgExp").Visible = false;
//				//숙박비(외화)
//				oMat1.Columns.Item("FoodExp").Visible = true;
//				//식비
//				oMat1.Columns.Item("ParkExp").Visible = false;
//				//주차비
//				oMat1.Columns.Item("TollExp").Visible = false;
//				//도로비
//				oMat1.Columns.Item("TotalExp").Visible = true;
//				//합계

//				//        oMat1.Columns("TransExp").Editable = True '교통비 Editable 설정
//				//        oMat1.Columns("FoodExp").Editable = True '식비 Editable 설정

//				oMat1.AutoResizeColumns();

//				//UPGRADE_WARNING: oForm.Items(DestDiv).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//해외출장
//			} else if (oForm.Items.Item("DestDiv").Specific.Value == "03") {

//				oMat1.Columns.Item("FrDate").Visible = true;
//				//기간시작일자
//				oMat1.Columns.Item("ToDate").Visible = true;
//				//기간종료일자
//				oMat1.Columns.Item("Destinat").Visible = true;
//				//행선지
//				oMat1.Columns.Item("PayPstg").Visible = true;
//				//지급율
//				oMat1.Columns.Item("Currency").Visible = true;
//				//통화
//				oMat1.Columns.Item("Rate").Visible = true;
//				//환율
//				oMat1.Columns.Item("Vehicle").Visible = false;
//				//차량구분
//				oMat1.Columns.Item("FuelType").Visible = false;
//				//유종
//				oMat1.Columns.Item("FuelPrc").Visible = false;
//				//단가
//				oMat1.Columns.Item("Distance").Visible = false;
//				//거리
//				oMat1.Columns.Item("FoodNum").Visible = false;
//				//식수
//				oMat1.Columns.Item("TransExp").Visible = true;
//				//교통비
//				oMat1.Columns.Item("InsurExp").Visible = true;
//				//보험료
//				oMat1.Columns.Item("AirpExp").Visible = true;
//				//공항세
//				oMat1.Columns.Item("DayExp").Visible = true;
//				//일비
//				oMat1.Columns.Item("FDayExp").Visible = true;
//				//일비(외화)
//				oMat1.Columns.Item("LodgExp").Visible = true;
//				//숙박비
//				oMat1.Columns.Item("FLodgExp").Visible = true;
//				//숙박비(외화)
//				oMat1.Columns.Item("FoodExp").Visible = true;
//				//식비
//				oMat1.Columns.Item("ParkExp").Visible = false;
//				//주차비
//				oMat1.Columns.Item("TollExp").Visible = false;
//				//도로비
//				oMat1.Columns.Item("TotalExp").Visible = true;
//				//합계

//				//        oMat1.Columns("TransExp").Editable = True '교통비 Editable 설정
//				//        oMat1.Columns("FoodExp").Editable = True '식비 Editable 설정

//				oMat1.AutoResizeColumns();

//			}

//			return;
//			PH_PY012_ColumnStatusChange_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_ColumnStatusChange_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY012_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//사업장
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY012A.GetValue("U_CLTCOD", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//사원
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY012A.GetValue("U_MSTCOD", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사원정보는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//기간(From)
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY012A.GetValue("U_FrDate", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("기간 시작은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//기간(To)
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY012A.GetValue("U_ToDate", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("기간 종료는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("ToDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//라인
//			if (oMat1.VisualRowCount > 1) {
//				for (i = 1; i <= oMat1.VisualRowCount - 1; i++) {

//					//            '학교
//					//            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
//					//                Sbo_Application.SetStatusBarMessage "학교는 필수입니다.", bmt_Short, True
//					//                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
//					//                PH_PY012_DataValidCheck = False
//					//                Exit Function
//					//            End If
//					//
//					//            '학교명
//					//            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
//					//                Sbo_Application.SetStatusBarMessage "학교명은 필수입니다.", bmt_Short, True
//					//                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
//					//                PH_PY012_DataValidCheck = False
//					//                Exit Function
//					//            End If
//					//
//					//            '학년
//					//            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
//					//                Sbo_Application.SetStatusBarMessage "학년은 필수입니다.", bmt_Short, True
//					//                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
//					//                PH_PY012_DataValidCheck = False
//					//                Exit Function
//					//            End If
//					//
//					//            '회차
//					//            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
//					//                Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
//					//                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
//					//                PH_PY012_DataValidCheck = False
//					//                Exit Function
//					//            End If

//				}
//			} else {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			oMat1.FlushToDataSource();
//			//// Matrix 마지막 행 삭제(DB 저장시)
//			if (oDS_PH_PY012B.Size > 1)
//				oDS_PH_PY012B.RemoveRecord((oDS_PH_PY012B.Size - 1));

//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY012_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY012_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = oForm.Items.Item("Param01").Specific.Value;
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Param01").Specific.Value;
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("Param01").Specific.Value;
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param04 = oForm.Items.Item("Param01").Specific.Value;

//			sQry = "SELECT 10";
//			oRecordSet.DoQuery(sQry);

//			oMat1.Clear();
//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();

//			if ((oRecordSet.RecordCount == 0)) {
//				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//				goto PH_PY012_MTX01_Exit;
//			}

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				if (i != 0) {
//					oDS_PH_PY012B.InsertRecord((i));
//				}
//				oDS_PH_PY012B.Offset = i;
//				oDS_PH_PY012B.SetValue("U_COL01", i, oRecordSet.Fields.Item(0).Value);
//				oDS_PH_PY012B.SetValue("U_COL02", i, oRecordSet.Fields.Item(1).Value);
//				oRecordSet.MoveNext();
//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
//			}
//			oMat1.LoadFromDataSource();
//			oMat1.AutoResizeColumns();
//			oForm.Update();

//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY012_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			return;
//			PH_PY012_MTX01_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY012_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY012A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY012A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY012_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY012_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY012_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private double PH_PY012_GetFuelPrc(string pBPLId, string pStdYear, string pStdMonth, string pFuelType)
//		{
//			double functionReturnValue = 0;
//			//******************************************************************************
//			//Function ID : PH_PY012_GetFuelPrc()
//			//해당모듈 : PH_PY012
//			//기능 : 유류단가 가져오기
//			//인수 : pBPLId:사업장, pStdYear:기준년도, pStdMonth:기준월, pFuelType:유종
//			//반환값 : 유류단가
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short loopCount = 0;
//			string sQry = null;
//			object CheckAmt = null;

//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "           SELECT      T0.U_Year AS [StdYear],";
//			sQry = sQry + "                T1.U_Month AS [StdMonth],";
//			sQry = sQry + "                T1.U_Gasoline AS [Gasoline],";
//			sQry = sQry + "                T1.U_Diesel AS [Diesel],";
//			sQry = sQry + "                T1.U_LPG AS [LPG]";
//			sQry = sQry + " FROM       [@PH_PY007A] AS T0";
//			sQry = sQry + "                INNER JOIN";
//			sQry = sQry + "                [@PH_PY007B] AS T1";
//			sQry = sQry + "                    ON T0.Code = T1.Code";
//			sQry = sQry + " WHERE      T0.U_CLTCOD = '" + pBPLId + "'";
//			sQry = sQry + "                AND T0.U_Year = '" + pStdYear + "'";
//			sQry = sQry + "                AND T1.U_Month = '" + pStdMonth + "'";

//			oRecordSet.DoQuery(sQry);

//			//휘발유
//			if (pFuelType == "1") {
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = oRecordSet.Fields.Item("Gasoline").Value;
//			//가스
//			} else if (pFuelType == "2") {
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = oRecordSet.Fields.Item("LPG").Value;
//			//경유
//			} else if (pFuelType == "3") {
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = oRecordSet.Fields.Item("Diesel").Value;
//			}
//			return functionReturnValue;
//			PH_PY012_GetFuelPrc_Error:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_GetFuelPrc_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY012_CalculateTotalExp(short pRow)
//		{
//			//******************************************************************************
//			//Function ID : PH_PY012_CalculateTotalExp()
//			//해당모듈 : PH_PY012
//			//기능 : 합계 계산
//			//인수 : pRow : pval.Row 값
//			//반환값 : 비용 합계
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string FuelType = null;
//			//유종
//			double FuelPrc = 0;
//			//단가
//			double Distance = 0;
//			//거리
//			double FoodNum = 0;
//			//식수
//			double TransExp = 0;
//			//교통비
//			double InsurExp = 0;
//			//보험료
//			double AirpExp = 0;
//			//공항세
//			double DayExp = 0;
//			//일비
//			double LodgExp = 0;
//			//숙박비
//			double FoodExp = 0;
//			//식비
//			double ParkExp = 0;
//			//주차비
//			double TollExp = 0;
//			//도로비
//			double TotalExp = 0;
//			//합계

//			oMat1.FlushToDataSource();

//			//합계(교통비+일비+식비+주차비+도로비+숙박비+보험료+공항세)
//			TransExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_TransExp", pRow));
//			//교통비
//			DayExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_DayExp", pRow));
//			//일비
//			FoodExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_FoodExp", pRow));
//			//식비
//			ParkExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_ParkExp", pRow));
//			//주차비
//			TollExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_TollExp", pRow));
//			//도로비
//			LodgExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_LodgExp", pRow));
//			//숙박비
//			InsurExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_InsurExp", pRow));
//			//보험료
//			AirpExp = Convert.ToDouble(oDS_PH_PY012B.GetValue("U_AirpExp", pRow));
//			//공항세

//			TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp + LodgExp + InsurExp + AirpExp;

//			oDS_PH_PY012B.SetValue("U_TotalExp", pRow, Convert.ToString(TotalExp));
//			//합계

//			oMat1.LoadFromDataSource();

//			return;
//			PH_PY012_CalculateTotalExp_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY012_CalculateTotalExp_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//	}
//}
