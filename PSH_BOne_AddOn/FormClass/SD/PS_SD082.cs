//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_PS_Addon
//{
//	internal class PS_SD082
//	{
////****************************************************************************************************************
//////  File           : PS_SD082.cls
//////  Module         : SD
//////  Description    : 여신한도 초과승인
//////  FormType       : PS_SD082
//////  Create Date    : 2010.10.22
//////  Modified Date  :
//////  Creator        : Ryu Yung Jo
//////  Company        : Poongsan Holdings
////****************************************************************************************************************

//		public string oFormUniqueID01;
//		public SAPbouiCOM.Form oForm01;
//		public SAPbouiCOM.Matrix oMat01;
//		public SAPbouiCOM.Matrix oMat02;
//			//등록라인
//		private SAPbouiCOM.DBDataSource oDS_PS_SD082L;
//			//등록라인
//		private SAPbouiCOM.DBDataSource oDS_PS_SD082M;

//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string oLast_Item_UID;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string oLast_Col_UID;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int oLast_Col_Row;

//		private int oLast_Mode;
//		private int oSeq;

////****************************************************************************************************************
//// .srf 파일로부터 폼을 로드한다.
////****************************************************************************************************************
//		public void LoadForm(ref SAPbouiCOM.Form oForm02 = null, string oItemUID02 = "", string oColUID02 = "", int oColRow02 = 0, string oTradeType02 = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			string oInnerXml01 = null;
//			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

//			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_SD082.srf");
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

//			//매트릭스의 타이틀높이와 셀높이를 고정
//			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}

//			oFormUniqueID01 = "PS_SD082_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID01);
//			////폼추가
//			SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));

//			//폼 할당
//			oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

//			oForm01.SupportedModes = -1;

//			oForm01.Freeze(true);

//			CreateItems();
//			ComboBox_Setting();
//			Initialization();
//			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//			LoadCaption();

//			oForm01.EnableMenu(("1283"), false);
//			//// 삭제
//			oForm01.EnableMenu(("1286"), false);
//			//// 닫기
//			oForm01.EnableMenu(("1287"), false);
//			//// 복제
//			oForm01.EnableMenu(("1284"), false);
//			//// 취소
//			oForm01.EnableMenu(("1293"), false);
//			//// 행삭제

//			oForm01.Update();
//			oForm01.Freeze(false);
//			oForm01.Visible = true;

//			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc01 = null;
//			return;
//			LoadForm_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			oForm01.Update();
//			oForm01.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc01 = null;
//			if ((oForm01 == null) == false) {
//				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oForm01 = null;
//			}
//			MDC_Com.MDC_GF_Message(ref "LoadForm_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

////****************************************************************************************************************
////// ItemEventHander
////****************************************************************************************************************
//		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			int ErrNum = 0;
//			object TempForm01 = null;
//			SAPbouiCOM.ProgressBar ProgressBar01 = null;

//			string ItemType = null;
//			string RequestDate = null;
//			string Size = null;
//			string ItemCode = null;
//			string ItemName = null;
//			string Unit = null;
//			string DueDate = null;
//			string RequestNo = null;
//			int Qty = 0;
//			decimal Weight = default(decimal);
//			string RFC_Sender = null;
//			double Calculate_Weight = 0;
//			int Seq = 0;

//			////BeforeAction = True
//			if ((pval.BeforeAction == true)) {
//				switch (pval.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//						////1
//						break;
//					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//						////2
//						break;
//					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//						////5
//						break;
//					case SAPbouiCOM.BoEventTypes.et_CLICK:
//						////6
//						break;
//					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//						////7
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//						////8
//						break;
//					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//						////10
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//						////11
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//						////18
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//						////19
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//						////20
//						break;
//					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//						////27
//						break;
//					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//						////3
//						break;
//					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//						////4
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//						////17
//						break;
//				}
//			////BeforeAction = False
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.EventType) {
//					//et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//						////1
//						if (pval.ItemUID == "Btn01") {
//							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								Update_SD080(ref pval);
//								oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//								oMat01.Clear();
//								oDS_PS_SD082L.Clear();
//								oMat02.Clear();
//								oDS_PS_SD082M.Clear();
//								LoadCaption();
//							} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//								oForm01.Close();
//							}
//						} else if (pval.ItemUID == "Btn02") {
//							LoadData();
//							oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							LoadCaption();
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//						////2
//						break;
//					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//						////5
//						break;
//					//et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_CLICK:
//						////6
//						if (pval.ItemUID == "Radio01") {
//							oForm01.Freeze(true);
//							oForm01.Settings.MatrixUID = "Mat01";
//							oForm01.Settings.EnableRowFormat = true;
//							oForm01.Settings.Enabled = true;
//							oForm01.Freeze(false);
//						} else if (pval.ItemUID == "Radio02") {
//							oForm01.Freeze(true);
//							oForm01.Settings.MatrixUID = "Mat02";
//							oForm01.Settings.EnableRowFormat = true;
//							oForm01.Settings.Enabled = true;
//							oForm01.Freeze(false);
//						} else if (pval.ItemUID == "Mat01") {
//							if (pval.ColUID == "LineNum") {
//								LoadData_Mat02((Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg02", pval.Row - 1))));
//							} else if (pval.ColUID == "Check") {
//								oForm01.Freeze(true);
//								oMat01.FlushToDataSource();
//								for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
//									if (Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg01", i)) == "Y") {
//										oForm01.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//										LoadCaption();
//										oForm01.Freeze(false);
//										BubbleEvent = false;
//										return;
//									}
//								}
//								oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//								LoadCaption();
//								oForm01.Freeze(false);
//								BubbleEvent = false;
//								return;
//							}
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//						////7
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//						////8
//						break;
//					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//						////10
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//						////11
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//						////18
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//						////19
//						break;
//					//et_FORM_RESIZE /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//						////20
//						oForm01.Freeze(true);

//						oForm01.Items.Item("Mat01").Top = 50;
//						oForm01.Items.Item("Mat01").Left = 6;
//						oForm01.Items.Item("Mat01").Width = oForm01.Width * 0.4 - 6;
//						oForm01.Items.Item("Mat01").Height = oForm01.Height - 110;

//						oForm01.Items.Item("Mat02").Top = oForm01.Items.Item("Mat01").Top;
//						oForm01.Items.Item("Mat02").Left = oForm01.Width * 0.4 + 6 + 10;
//						oForm01.Items.Item("Mat02").Width = oForm01.Width * 0.6 - 6 - 22;
//						oForm01.Items.Item("Mat02").Height = oForm01.Height - 110;

//						oForm01.Items.Item("Radio01").Left = 6;
//						oForm01.Items.Item("Radio02").Left = oForm01.Width * 0.4 + 6 + 10;

//						oMat01.AutoResizeColumns();
//						oMat02.AutoResizeColumns();

//						//                oMat01.Columns("Check").Width = 40
//						//                oMat01.Columns("DocNum").Width = 60
//						//                oMat01.Columns("BPLId").Width = 50
//						//                oMat01.Columns("CntcCode").Width = 60
//						//                oMat01.Columns("DocDate").Width = 80
//						//
//						//                oMat02.Columns("CardCode").Width = 80
//						//                oMat02.Columns("CardName").Width = 80
//						//                oMat02.Columns("RequestP").Width = 80
//						//                oMat02.Columns("CreditP").Width = 80
//						//                oMat02.Columns("MiSuP").Width = 80
//						//                oMat02.Columns("Balance").Width = 80
//						//                oMat02.Columns("OutPreP").Width = 80
//						//                oMat02.Columns("Comment").Width = 80

//						oForm01.Freeze(false);
//						break;
//					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//						////27
//						break;
//					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//						////3
//						break;
//					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//						////4
//						break;
//					//et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//						////17
//						SubMain.RemoveForms(oFormUniqueID01);
//						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm01 = null;
//						//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat01 = null;
//						//UPGRADE_NOTE: oMat02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat02 = null;
//						//UPGRADE_NOTE: oDS_PS_SD082L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PS_SD082L = null;
//						//UPGRADE_NOTE: oDS_PS_SD082M 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PS_SD082M = null;
//						break;
//				}
//			}
//			return;
//			Raise_ItemEvent_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			if (ErrNum == 101) {
//				ErrNum = 0;
//				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
//				BubbleEvent = false;
//			} else {
//				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//		}

//		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;

//			////BeforeAction = True
//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					case "1282":
//						//추가
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;
//				}
//			////BeforeAction = False
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					case "1282":
//						//추가
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;
//				}
//			}
//			return;
//			Raise_MenuEvent_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			////BeforeAction = True
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
//			////BeforeAction = False
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
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((eventInfo.BeforeAction == true)) {
//				////작업
//			} else if ((eventInfo.BeforeAction == false)) {
//				////작업
//			}
//			return;
//			Raise_RightClickEvent_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			////디비데이터 소스 개체 할당
//			oDS_PS_SD082L = oForm01.DataSources.DBDataSources("@PS_USERDS01");
//			oDS_PS_SD082M = oForm01.DataSources.DBDataSources("@PS_USERDS02");

//			//// 메트릭스 개체 할당
//			oMat01 = oForm01.Items.Item("Mat01").Specific;
//			oMat02 = oForm01.Items.Item("Mat02").Specific;

//			oForm01.DataSources.UserDataSources.Add("Radio01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("Radio01").Specific.DataBind.SetBound(true, "", "Radio01");

//			oForm01.DataSources.UserDataSources.Add("Radio02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("Radio02").Specific.DataBind.SetBound(true, "", "Radio02");

//			//UPGRADE_WARNING: oForm01.Items().Specific.GroupWith 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("Radio01").Specific.GroupWith("Radio02");

//			//    oDS_PS_SD082L.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
//			//    oDS_PS_SD082L.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")

//			return;
//			CreateItems_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void ComboBox_Setting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			////콤보에 기본값설정
//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;

//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm01.DataSources.UserDataSources.Add("OkYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("OkYN").Specific.DataBind.SetBound(true, "", "OkYN");

//			//// 승인상태
//			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("OkYN").Specific.ValidValues.Add("Y", "승인");
//			//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("OkYN").Specific.ValidValues.Add("N", "미승인");
//			//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("OkYN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_Index);

//			//// 사업장
//			oCombo = oForm01.Items.Item("BPLId").Specific;
//			sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId";
//			oRecordSet01.DoQuery(sQry);
//			while (!(oRecordSet01.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
//				oMat01.Columns.Item("BPLId").ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
//				oRecordSet01.MoveNext();
//			}

//			//// 사용자
//			sQry = "Select empID, lastName + firstName From OHEM Order by empID";
//			oRecordSet01.DoQuery(sQry);
//			while (!(oRecordSet01.EoF)) {
//				oMat01.Columns.Item("CntcCode").ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
//				oRecordSet01.MoveNext();
//			}

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			return;
//			ComboBox_Setting_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "ComboBox_Setting_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void Initialization()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;

//			////아이디별 사업장 세팅
//			oCombo = oForm01.Items.Item("BPLId").Specific;
//			oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);

//			////아이디별 사번 세팅
//			//    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD

//			////아이디별 부서 세팅
//			//    Set oCombo = oForm01.Items("DeptCode").Specific
//			//    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			return;
//			Initialization_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			MDC_Com.MDC_GF_Message(ref "Initialization_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		private void LoadCaption()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				//UPGRADE_WARNING: oForm01.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm01.Items.Item("Btn01").Specific.Caption = "확인";
//			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				//UPGRADE_WARNING: oForm01.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm01.Items.Item("Btn01").Specific.Caption = "확인";
//			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//				//UPGRADE_WARNING: oForm01.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm01.Items.Item("Btn01").Specific.Caption = "승인";
//			}

//			return;
//			LoadCaption_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "Delete_EmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void LoadData()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string OkYN = null;
//			string BPLId = null;
//			string DocNum = null;

//			//UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			BPLId = Strings.Trim(oForm01.Items.Item("BPLId").Specific.Value);
//			OkYN = Strings.Trim(oForm01.DataSources.UserDataSources.Item("OkYN").Value);

//			if (string.IsNullOrEmpty(OkYN))
//				OkYN = "%";

//			sQry = "EXEC [PS_SD082_01] '" + BPLId + "','" + OkYN + "','" + DocNum + "','01'";
//			oRecordSet01.DoQuery(sQry);

//			oMat01.Clear();
//			oDS_PS_SD082L.Clear();

//			oMat02.Clear();
//			oDS_PS_SD082M.Clear();

//			if (oRecordSet01.RecordCount == 0) {
//				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
//				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet01 = null;
//				return;
//			}

//			oForm01.Freeze(true);
//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
//				if (i + 1 > oDS_PS_SD082L.Size) {
//					oDS_PS_SD082L.InsertRecord((i));
//				}

//				oMat01.AddRow();
//				oDS_PS_SD082L.Offset = i;
//				oDS_PS_SD082L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PS_SD082L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("DocNum").Value));
//				oDS_PS_SD082L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("U_BPLId").Value));
//				oDS_PS_SD082L.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("U_CntcCode").Value));
//				oDS_PS_SD082L.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("U_DocDate").Value));

//				oRecordSet01.MoveNext();
//				ProgBar01.Value = ProgBar01.Value + 1;
//				ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
//			}
//			oMat01.LoadFromDataSource();
//			oMat01.AutoResizeColumns();
//			ProgBar01.Stop();
//			oForm01.Freeze(false);

//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			return;
//			LoadData_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "LoadData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void LoadData_Mat02(string sDocNum)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string OkYN = null;
//			string BPLId = null;
//			string DocNum = null;

//			//UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			BPLId = Strings.Trim(oForm01.Items.Item("BPLId").Specific.Value);
//			OkYN = Strings.Trim(oForm01.DataSources.UserDataSources.Item("OkYN").Value);

//			sQry = "EXEC [PS_SD082_01] '" + BPLId + "','" + OkYN + "','" + sDocNum + "','02'";
//			oRecordSet01.DoQuery(sQry);

//			oMat02.Clear();
//			oDS_PS_SD082M.Clear();

//			if (oRecordSet01.RecordCount == 0) {
//				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
//				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet01 = null;
//				return;
//			}

//			oForm01.Freeze(true);
//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
//				if (i + 1 > oDS_PS_SD082M.Size) {
//					oDS_PS_SD082M.InsertRecord((i));
//				}

//				oMat02.AddRow();
//				oDS_PS_SD082M.Offset = i;
//				oDS_PS_SD082M.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PS_SD082M.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("U_CardCode").Value));
//				oDS_PS_SD082M.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("U_CardName").Value));
//				oDS_PS_SD082M.SetValue("U_ColSum01", i, Strings.Trim(oRecordSet01.Fields.Item("U_RequestP").Value));
//				oDS_PS_SD082M.SetValue("U_ColSum02", i, Strings.Trim(oRecordSet01.Fields.Item("U_CreditP").Value));
//				oDS_PS_SD082M.SetValue("U_ColSum03", i, Strings.Trim(oRecordSet01.Fields.Item("U_MiSuP").Value));
//				oDS_PS_SD082M.SetValue("U_ColSum04", i, Strings.Trim(oRecordSet01.Fields.Item("U_Balance").Value));
//				oDS_PS_SD082M.SetValue("U_ColSum05", i, Strings.Trim(oRecordSet01.Fields.Item("U_OutPreP").Value));
//				oDS_PS_SD082M.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("U_Comment").Value));

//				oRecordSet01.MoveNext();
//				ProgBar01.Value = ProgBar01.Value + 1;
//				ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
//			}
//			oMat02.LoadFromDataSource();
//			oMat02.AutoResizeColumns();
//			ProgBar01.Stop();
//			oForm01.Freeze(false);

//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			return;
//			LoadData_Mat02_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "LoadData_Mat02_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public bool Update_SD080(ref SAPbouiCOM.ItemEvent pval)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string DocNum = null;
//			string OkDate = null;

//			oMat01.FlushToDataSource();

//			for (i = 0; i <= oMat01.RowCount - 1; i++) {
//				if (Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg01", i)) == "Y") {
//					DocNum = Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg02", i));
//					OkDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");

//					sQry = "UPDATE [@PS_SD080H] ";
//					sQry = sQry + "SET ";
//					sQry = sQry + "U_OkYN = 'Y', ";
//					sQry = sQry + "U_OkDate = '" + OkDate + "'";
//					sQry = sQry + "Where DocNum = '" + DocNum + "'";

//					RecordSet01.DoQuery(sQry);
//				}
//			}

//			MDC_Com.MDC_GF_Message(ref "여신한도 초과승인 완료!", ref "S");

//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//			Update_JakNum_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "Update_JakNum_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}
//	}
//}
