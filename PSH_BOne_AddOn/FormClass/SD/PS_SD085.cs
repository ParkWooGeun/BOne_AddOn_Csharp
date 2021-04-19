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
//	internal class PS_SD085
//	{
////****************************************************************************************************************
//////  File           : PS_SD085.cls
//////  Module         : SD
//////  Description    : 입금조회
//////  FormType       : PS_SD085
//////  Create Date    : 2011.03.24
//////  Modified Date  :
//////  Creator        : N.G.Y
//////  Company        : Poongsan Holdings
////****************************************************************************************************************

//		public string oFormUniqueID01;
//		public SAPbouiCOM.Form oForm01;
//		public SAPbouiCOM.Matrix oMat01;
//			//등록라인
//		private SAPbouiCOM.DBDataSource oDS_PS_USERDS01;

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

//			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_SD085.srf");
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

//			//매트릭스의 타이틀높이와 셀높이를 고정
//			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}

//			oFormUniqueID01 = "PS_SD085_" + GetTotalFormsCount();
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

//			oForm01.EnableMenu(("1281"), false);
//			//// 찾기
//			oForm01.EnableMenu(("1282"), false);
//			//// 추가
//			//    oForm01.EnableMenu ("1283"), False        '// 삭제
//			//    oForm01.EnableMenu ("1286"), False         '// 닫기
//			//    oForm01.EnableMenu ("1287"), False        '// 복제
//			//    oForm01.EnableMenu ("1284"), False         '// 취소
//			//    oForm01.EnableMenu ("1293"), False         '// 행삭제

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
//					//et_KEY_DOWN ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//						////2
//						if (pval.CharPressed == 9) {
//							if (pval.ItemUID == "CardCode") {
//								//UPGRADE_WARNING: oForm01.Items(CardCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm01.Items.Item("CardCode").Specific.VALUE)) {
//									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//									BubbleEvent = false;
//								}
//							}
//						}
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
//					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//						////1
//						if (pval.ItemUID == "Btn01") {
//							LoadData();
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//						////2
//						break;
//					//et_COMBO_SELECT ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//						////5
//						if (pval.ItemChanged == true) {
//							if (pval.ItemUID == "BPLId") {
//								FlushToItemValue(pval.ItemUID);
//							}
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_CLICK:
//						////6
//						break;
//					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//						////7
//						break;
//					//et_MATRIX_LINK_PRESSED /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//						////8
//						break;
//					//                If pval.ItemUID = "Mat01" And pval.ColUID = "TrandId" Then
//					//                   'Set TempForm01 = New "392"
//					//                ElseIf pval.ItemUID = "Mat01" And pval.ColUID = "Ref1" Then
//					//                        Set TempForm01 = New PS_PP040
//					//                End If
//					//
//					//                Call TempForm01.LoadForm(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)
//					//                Set TempForm01 = Nothing

//					//et_VALIDATE ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//						////10
//						if (pval.ItemChanged == true) {
//							if (pval.ItemUID == "CardCode") {
//								FlushToItemValue(pval.ItemUID);
//							}
//						}
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
//					//et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//						////17
//						SubMain.RemoveForms(oFormUniqueID01);
//						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm01 = null;
//						//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat01 = null;
//						//UPGRADE_NOTE: oDS_PS_USERDS01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PS_USERDS01 = null;
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
//			oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01");

//			//// 메트릭스 개체 할당
//			oMat01 = oForm01.Items.Item("Mat01").Specific;

//			oForm01.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
//			oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = Convert.ToString(DateAndTime.Today);

//			oForm01.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
//			oForm01.DataSources.UserDataSources.Item("DocDateTo").Value = Convert.ToString(DateAndTime.Today);

//			return;
//			CreateItems_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void ComboBox_Setting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			MDC_SetMod.Set_ComboList(ref (oForm01.Items.Item("BPLId").Specific), ref "SELECT BPLId, BPLName FROM OBPL order by BPLId", ref "1", ref false, ref false);

//			return;
//			ComboBox_Setting_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "ComboBox_Setting_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void Initialization()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;

//			////아이디별 사업장 세팅
//			oCombo = oForm01.Items.Item("BPLId").Specific;
//			oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);


//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			return;
//			Initialization_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			MDC_Com.MDC_GF_Message(ref "Initialization_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			string CardCode = null;

//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (oUID) {
//				case "CardCode":
//					oForm01.Freeze(true);
//					if (oUID == "CarCode") {
//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = "Select CardName From OCRD Where CardCode = '" + Strings.Trim(oForm01.Items.Item("CardCode").Specific.VALUE) + "'";
//						oRecordSet01.DoQuery(sQry);

//						//UPGRADE_WARNING: oForm01.Items(CardName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm01.Items.Item("CardName").Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
//					}
//					oForm01.Freeze(false);
//					break;
//			}

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			return;
//			FlushToItemValue_Error:
//			oForm01.Freeze(false);
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}


//////조회데이타 가져오기
//		public void LoadData()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			short ErrNum = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			string DocDateTo = null;
//			string BPLId = null;
//			string DocDateFr = null;
//			string CardCode = null;

//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oMat01.Clear();
//			oDS_PS_USERDS01.Clear();

//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			BPLId = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocDateFr = Strings.Trim(oForm01.Items.Item("DocDateFr").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocDateTo = Strings.Trim(oForm01.Items.Item("DocDateTo").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CardCode = Strings.Trim(oForm01.Items.Item("CardCode").Specific.VALUE);

//			if (string.IsNullOrEmpty(BPLId))
//				BPLId = "%";
//			if (string.IsNullOrEmpty(DocDateFr))
//				DocDateFr = "18990101";
//			if (string.IsNullOrEmpty(DocDateTo))
//				DocDateTo = "20991231";
//			if (string.IsNullOrEmpty(CardCode))
//				CardCode = "%";

//			sQry = "EXEC [PS_SD085_01] '" + BPLId + "', '" + DocDateFr + "', '" + DocDateTo + "', '" + CardCode + "'";
//			oRecordSet01.DoQuery(sQry);

//			if (oRecordSet01.RecordCount == 0) {
//				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
//				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet01 = null;
//				oForm01.Freeze(false);
//				return;
//			}

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
//				if (i + 1 > oDS_PS_USERDS01.Size) {
//					oDS_PS_USERDS01.InsertRecord((i));
//				}

//				oMat01.AddRow();
//				oDS_PS_USERDS01.Offset = i;
//				oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PS_USERDS01.SetValue("U_ColDt01", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("DocDate").Value), "YYYYMMDD"));
//				oDS_PS_USERDS01.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("TransId").Value));
//				oDS_PS_USERDS01.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("Ref1").Value));
//				oDS_PS_USERDS01.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("CardCode").Value));
//				oDS_PS_USERDS01.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("CardName").Value));
//				oDS_PS_USERDS01.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("LineMemo").Value));
//				oDS_PS_USERDS01.SetValue("U_ColReg06", i, Strings.Trim(oRecordSet01.Fields.Item("Account").Value));
//				oDS_PS_USERDS01.SetValue("U_ColSum01", i, Strings.Trim(oRecordSet01.Fields.Item("Amt").Value));
//				oDS_PS_USERDS01.SetValue("U_ColSum02", i, Strings.Trim(oRecordSet01.Fields.Item("RefAmt").Value));
//				oDS_PS_USERDS01.SetValue("U_ColReg07", i, Strings.Trim(oRecordSet01.Fields.Item("RefNum").Value));
//				//----------------------------------------------------------------------------------------------------------
//				oRecordSet01.MoveNext();
//				ProgBar01.Value = ProgBar01.Value + 1;
//				ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
//			}
//			oMat01.LoadFromDataSource();
//			//            oMat01.AutoResizeColumns
//			ProgBar01.Stop();
//			oForm01.Freeze(false);


//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			return;
//			LoadData_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			oForm01.Freeze(false);
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "LoadData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}
//	}
//}
