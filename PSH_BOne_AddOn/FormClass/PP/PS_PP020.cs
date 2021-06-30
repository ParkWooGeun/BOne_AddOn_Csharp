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
//	internal class PS_PP020
//	{
////****************************************************************************************************************
//////  File           : PS_PP020.cls
//////  Module         : PP
//////  Description    : 작번등록
//////  FormType       : PS_PP020
//////  Create Date    : 2010.10.05
//////  Modified Date  : 2018.07.26 필드 요청자 필드 추
//////  Creator        : Ryu Yung Jo
//////  Company        : Poongsan Holdings
////****************************************************************************************************************

//		public string oFormUniqueID01;
//		public SAPbouiCOM.Form oForm01;
//		public SAPbouiCOM.Matrix oMat01;
//		public SAPbouiCOM.Matrix oMat02;
//			//등록헤더
//		private SAPbouiCOM.DBDataSource oDS_PS_PP020H;
//		private SAPbouiCOM.DBDataSource oDS_PS_TEMPTABLE;

//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string oLast_ItemUID;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string oLast_ColUID;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int oLast_ColRow;

//		private int oLast_Mode;
//		private string oLast_RightClick_DocEntry;

////****************************************************************************************************************
//// .srf 파일로부터 폼을 로드한다.
////****************************************************************************************************************
//		public void LoadForm()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			string oInnerXml01 = null;
//			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

//			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_PP020.srf");
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

//			//매트릭스의 타이틀높이와 셀높이를 고정
//			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}

//			oFormUniqueID01 = "PS_PP020_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID01);
//			//폼추가
//			SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));

//			//폼 할당
//			oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

//			oForm01.SupportedModes = -1;
//			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//			oForm01.Freeze(true);
//			CreateItems();
//			ComboBox_Setting();
//			Initialization();
//			Add_MatrixRow(0, true);
//			LoadCaption();
//			FormItemEnabled();

//			oForm01.EnableMenu("1283", false);
//			//삭제
//			oForm01.EnableMenu("1286", false);
//			//닫기
//			oForm01.EnableMenu("1287", false);
//			//복제
//			oForm01.EnableMenu("1285", false);
//			//복원
//			oForm01.EnableMenu("1284", true);
//			//취소
//			oForm01.EnableMenu("1293", true);
//			//행삭제
//			oForm01.EnableMenu("1299", true);
//			//행닫기

//			oForm01.Update();
//			oForm01.Freeze(false);
//			oForm01.Visible = true;

//			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc01 = null;
//			return;
//			LoadForm_Error:

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
////ItemEventHander
////****************************************************************************************************************
//		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			int j = 0;
//			int ErrNum = 0;
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
//			int Seq = 0;
//			object ChildForm01 = null;
//			ChildForm01 = new PS_SM010();
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;

//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string SubNo1 = null;
//			string JakName = null;
//			string SubNo2 = null;
//			string FirstInOutGbn = null;
//			if ((pval.BeforeAction == true)) {
//				switch (pval.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//						//1
//						if (pval.ItemUID == "Btn01") {
//							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (HeaderSpaceLineDel() == false) {
//									BubbleEvent = false;
//									return;
//								}
//								if (MatrixSpaceLineDel() == false) {
//									BubbleEvent = false;
//									return;
//								}

//								//                        If Add_JakNum(pval) = False Then
//								//                            BubbleEvent = False
//								//                            Exit Sub
//								//                        End If

//								if (PS_PP020_AddData() == false) {

//									BubbleEvent = false;
//									return;

//								}

//								oMat01.Clear();
//								oMat01.FlushToDataSource();
//								oMat01.LoadFromDataSource();
//								Add_MatrixRow(0, true);

//								//                        Call Delete_EmptyRow
//								oLast_Mode = oForm01.Mode;
//							} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (MatrixSpaceLineDel() == false) {
//									BubbleEvent = false;
//									return;
//								}

//								if (PS_PP020_AddData() == false) {

//									BubbleEvent = false;
//									return;

//								}

//								//                        If Update_JakNum(pval) = False Then
//								//                            BubbleEvent = False
//								//                            Exit Sub
//								//                        End If

//								oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//								LoadCaption();
//							} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//								oLast_Mode = oForm01.Mode;
//							}
//						} else if (pval.ItemUID == "Btn02") {
//							if (HeaderSpaceLineDel() == false) {
//								BubbleEvent = false;
//								return;
//							}
//							oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							LoadCaption();
//							LoadData();
//						} else if (pval.ItemUID == "Btn03") {
//							if (Add_SubJakNum(ref pval) == false) {
//								BubbleEvent = false;
//								return;
//							}
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//						//2
//						if (pval.CharPressed == 9) {
//							if (pval.ItemUID == "CardCode") {
//								//UPGRADE_WARNING: oForm01.Items(CardCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm01.Items.Item("CardCode").Specific.VALUE)) {
//									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//									BubbleEvent = false;
//								}
//							} else if (pval.ItemUID == "ItemCode") {
//								//UPGRADE_WARNING: oForm01.Items(ItemCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm01.Items.Item("ItemCode").Specific.VALUE)) {
//									//UPGRADE_WARNING: ChildForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row);
//									BubbleEvent = false;
//								}
//							} else if (pval.ItemUID == "Mat01") {
//								if (pval.ColUID == "JakName") {
//									//UPGRADE_WARNING: oMat01.Columns(JakName).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (string.IsNullOrEmpty(oMat01.Columns.Item("JakName").Cells.Item(pval.Row).Specific.VALUE)) {
//										SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//										BubbleEvent = false;
//									}
//								} else if (pval.ColUID == "ReqCod") {
//									//UPGRADE_WARNING: oMat01.Columns(ReqCod).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (string.IsNullOrEmpty(oMat01.Columns.Item("ReqCod").Cells.Item(pval.Row).Specific.VALUE)) {
//										SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//										BubbleEvent = false;
//									}
//								} else if (pval.ColUID == "CardCode") {
//									//UPGRADE_WARNING: oMat01.Columns(CardCode).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (string.IsNullOrEmpty(oMat01.Columns.Item("CardCode").Cells.Item(pval.Row).Specific.VALUE)) {
//										SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//										BubbleEvent = false;
//									}

//								} else if (pval.ColUID == "ShipCode") {
//									//UPGRADE_WARNING: oMat01.Columns(ShipCode).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (string.IsNullOrEmpty(oMat01.Columns.Item("ShipCode").Cells.Item(pval.Row).Specific.VALUE)) {
//										SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//										BubbleEvent = false;
//									}
//								}
//							}
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//						//5
//						if (pval.ItemUID == "ReWork") {
//							oDS_PS_TEMPTABLE.Clear();
//							oMat02.Clear();
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_CLICK:
//						//6
//						if (pval.ItemUID == "RadioMat01") {
//							oForm01.Freeze(true);
//							oForm01.Settings.MatrixUID = "Mat01";
//							oForm01.Settings.EnableRowFormat = true;
//							oForm01.Settings.Enabled = true;
//							oForm01.Freeze(false);
//						} else if (pval.ItemUID == "RadioMat02") {
//							oForm01.Freeze(true);
//							oForm01.Settings.MatrixUID = "Mat02";
//							oForm01.Settings.EnableRowFormat = true;
//							oForm01.Settings.Enabled = true;
//							oForm01.Freeze(false);
//						}

//						if (pval.ItemUID == "Mat01") {
//							if (pval.Row > 0) {
//								oMat01.SelectRow(pval.Row, true, false);
//							}
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//						//7
//						if (pval.ItemUID == "Mat01" & pval.Row != Convert.ToDouble("0")) {
//							j = 0;
//							if (oMat02.VisualRowCount == 0) {
//								oDS_PS_TEMPTABLE.Clear();
//							}

//							JakName = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakName", pval.Row - 1));
//							if (Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)) == "00") {
//								//UPGRADE_WARNING: oForm01.Items(ReWork).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//정상
//								if (oForm01.Items.Item("ReWork").Specific.VALUE == "10") {
//									if (Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) == "000") {
//										sQry = "        SELECT  MAX(ISNULL(U_SubNo1, '00')) ";
//										sQry = sQry + " FROM    [@PS_PP020H] ";
//										sQry = sQry + " WHERE   U_JakName = '" + JakName + "' ";
//										sQry = sQry + "         AND U_SubNo2 = '" + Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) + "'";
//										oRecordSet01.DoQuery(sQry);

//										SubNo1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Convert.ToDouble(Strings.Trim(oRecordSet01.Fields.Item(0).Value)) + 1, "00");
//										SubNo2 = "000";
//									} else {
//										MDC_Com.MDC_GF_Message(ref "해당 작번은 서브작번을 만들 수 없습니다. 확인하세요." + Err().Number + " - " + Err().Description, ref "W");
//										j = 1;
//									}
//								} else {
//									if (Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) == "000") {
//										sQry = "        SELECT  MAX(ISNULL(U_SubNo1, '00')) ";
//										sQry = sQry + " FROM    [@PS_PP020H] ";
//										sQry = sQry + " WHERE   U_JakName = '" + JakName + "' ";
//										sQry = sQry + "         AND U_SubNo1 >= '80' ";
//										sQry = sQry + "         AND U_SubNo2 = '" + Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) + "'";
//										oRecordSet01.DoQuery(sQry);

//										if (string.IsNullOrEmpty(Strings.Trim(oRecordSet01.Fields.Item(0).Value))) {
//											SubNo1 = "80";
//										} else {
//											SubNo1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Convert.ToDouble(Strings.Trim(oRecordSet01.Fields.Item(0).Value)) + 1, "00");
//										}
//										SubNo2 = "000";
//									} else {
//										MDC_Com.MDC_GF_Message(ref "해당 작번은 서브작번을 만들 수 없습니다. 확인하세요." + Err().Number + " - " + Err().Description, ref "W");
//										j = 1;
//									}

//								}
//							} else if (Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)) != "00") {
//								if (Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", pval.Row - 1)) == "000") {
//									sQry = "        SELECT  MAX(ISNULL(U_SubNo2, '000')) ";
//									sQry = sQry + " FROM    [@PS_PP020H] Where U_JakName = '" + JakName + "' ";
//									sQry = sQry + "         AND U_SubNo1 = '" + Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)) + "' ";

//									oRecordSet01.DoQuery(sQry);

//									SubNo1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo1", pval.Row - 1)), "00");
//									SubNo2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Convert.ToDouble(Strings.Trim(oRecordSet01.Fields.Item(0).Value)) + 1, "000");
//								} else {
//									MDC_Com.MDC_GF_Message(ref "해당 작번은 서브작번을 만들 수 없습니다. 확인하세요." + Err().Number + " - " + Err().Description, ref "W");
//									j = 1;
//								}
//							}

//							for (i = 0; i <= oMat02.VisualRowCount - 1; i++) {
//								oMat02.LoadFromDataSource();
//								if (Strings.Trim(oDS_PS_TEMPTABLE.GetValue("U_sField01", i)) == JakName & Strings.Trim(oDS_PS_TEMPTABLE.GetValue("U_sField02", i)) == SubNo1 & Strings.Trim(oDS_PS_TEMPTABLE.GetValue("U_sField03", i)) == SubNo2) {
//									MDC_Com.MDC_GF_Message(ref "같은 행을 두번 선택할 수 없습니다. 확인하세요." + Err().Number + " - " + Err().Description, ref "W");
//									j = 1;
//								}
//							}

//							if (j == 0) {
//								oForm01.Freeze(true);
//								Add_MatrixRow(oMat02.VisualRowCount, ref false, ref "Mat02");
//								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("JakName").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = JakName;
//								//UPGRADE_WARNING: oMat02.Columns(SubNo1).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("SubNo1").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(SubNo1, "00");
//								//UPGRADE_WARNING: oMat02.Columns(SubNo2).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("SubNo2").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(SubNo2, "000");
//								//UPGRADE_WARNING: oMat02.Columns(ItemCode).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("ItemCode").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_ItemCode", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(ItemName).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("ItemName").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_ItemName", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(Material).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("Material").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_Material", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(Unit).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("Unit").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_Unit", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(Size).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("Size").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_Size", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("ItmBSort").Cells.Item(oMat02.VisualRowCount).Specific.Select(Strings.Trim(oDS_PS_PP020H.GetValue("U_ItmBSort", pval.Row - 1)), SAPbouiCOM.BoSearchKey.psk_ByValue);
//								//UPGRADE_WARNING: oMat02.Columns(SjDocNum).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("SjDocNum").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDocNum", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(SjLinNum).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("SjLinNum").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjLinNum", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(CardCode).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("CardCode").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardCode", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(CardName).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("CardName").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardName", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(ShipCode).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("ShipCode").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipCode", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(ShipName).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("ShipName").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipName", pval.Row - 1));
//								if (string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", pval.Row - 1)))) {

//								} else {
//									//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oMat02.Columns.Item("InOutGbn").Cells.Item(oMat02.VisualRowCount).Specific.Select(Strings.Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", pval.Row - 1)), SAPbouiCOM.BoSearchKey.psk_ByValue);
//								}
//								//UPGRADE_WARNING: oMat02.Columns(JakDate).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("JakDate").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
//								//UPGRADE_WARNING: oMat02.Columns(ProDate).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("ProDate").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
//								//UPGRADE_WARNING: oMat02.Columns(ReDate).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("ReDate").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
//								//UPGRADE_WARNING: oMat02.Columns(SjDcDate).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("SjDcDate").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDcDate", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(SjDuDate).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("SjDuDate").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDuDate", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns(SlePrice).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("SlePrice").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_SlePrice", pval.Row - 1));
//								//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("WorkGbn").Cells.Item(oMat02.VisualRowCount).Specific.Select(Strings.Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", pval.Row - 1)));

//								//UPGRADE_WARNING: oMat02.Columns(OrderAmt).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("OrderAmt").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", pval.Row - 1));
//								//수주금액
//								//UPGRADE_WARNING: oMat02.Columns(NegoAmt).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("NegoAmt").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", pval.Row - 1));
//								//Nego금액
//								//UPGRADE_WARNING: oMat02.Columns(TrgtAmt).Cells(oMat02.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat02.Columns.Item("TrgtAmt").Cells.Item(oMat02.VisualRowCount).Specific.VALUE = Strings.Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", pval.Row - 1));
//								//목표금액

//								oMat02.FlushToDataSource();
//								oMat02.LoadFromDataSource();
//								oMat02.AutoResizeColumns();

//								oForm01.Freeze(false);
//								j = 0;
//							}
//							BubbleEvent = false;

//							//외주구분의 첫행을 선택한 후 컬럼 타이틀을 더블클릭하면 첫행의 값으로 첫행 이외의 값을 자동으로 선택하는 기능 추가(2012.08.20 송명규 수정)
//						} else if (pval.ItemUID == "Mat01" & pval.Row == Convert.ToDouble("0") & pval.ColUID == "InOutGbn") {

//							oForm01.Freeze(true);

//							oMat01.FlushToDataSource();
//							FirstInOutGbn = Strings.Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", 0));
//							for (i = 1; i <= oMat01.VisualRowCount - 2; i++) {

//								oDS_PS_PP020H.SetValue("U_InOutgbn", i, FirstInOutGbn);

//							}
//							oMat01.LoadFromDataSource();
//							oForm01.Freeze(false);

//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//						//8
//						break;
//					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//						//10
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//						//11
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//						//18
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//						//19
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//						//20
//						break;
//					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//						//27
//						break;
//					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//						//3
//						break;
//					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//						//4
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//						//17
//						break;
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//						//1
//						break;
//					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//						//2
//						break;
//					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//						//3
//						break;
//					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//						//4
//						break;
//					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//						//5
//						if (pval.ItemUID == "Mat01") {
//							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							} else {
//								oForm01.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//								LoadCaption();
//							}
//						}

//						if (pval.ItemUID == "WorkGbn" & oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							oForm01.Freeze(true);
//							oMat01.Clear();
//							oMat01.FlushToDataSource();
//							oMat01.LoadFromDataSource();

//							Add_MatrixRow(0, true);
//							//UPGRADE_WARNING: oForm01.Items(WorkGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (oForm01.Items.Item("WorkGbn").Specific.Selected.VALUE == "10") {
//								oMat01.Columns.Item("RegNum").Editable = true;
//								oMat01.Columns.Item("ItemCode").Editable = false;
//							} else {
//								oMat01.Columns.Item("RegNum").Editable = false;
//								oMat01.Columns.Item("ItemCode").Editable = true;
//							}
//							oForm01.Freeze(false);
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_CLICK:
//						//6
//						break;
//					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//						//7
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//						//8
//						break;
//					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//						//10
//						if (pval.ItemChanged == true) {
//							if (pval.ItemUID == "CntcCode") {
//								FlushToItemValue(pval.ItemUID);
//							} else if (pval.ItemUID == "CardCode") {
//								FlushToItemValue(pval.ItemUID);
//							} else if (pval.ItemUID == "ItemCode") {
//								FlushToItemValue(pval.ItemUID);
//							} else if (pval.ItemUID == "Mat01") {
//								if (pval.ColUID == "JakName") {
//									FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//								} else if (pval.ColUID == "ItemCode") {
//									FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//								} else if (pval.ColUID == "ShipCode") {
//									FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//								} else if (pval.ColUID == "ReqCod") {
//									FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//								}

//								if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								} else {
//									oForm01.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//									LoadCaption();
//								}
//							}
//						}
//						break;
//					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//						//11
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//						//17
//						SubMain.RemoveForms(oFormUniqueID01);
//						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm01 = null;
//						//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat01 = null;
//						//UPGRADE_NOTE: oDS_PS_PP020H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PS_PP020H = null;
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//						//18
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//						//19
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//						//20

//						oForm01.Items.Item("Mat01").Top = 82;
//						oForm01.Items.Item("Mat01").Left = 6;
//						oForm01.Items.Item("Mat01").Width = oForm01.Width - 18;
//						oForm01.Items.Item("Mat01").Height = (oForm01.Height - oForm01.Items.Item("Mat01").Top - (oForm01.Height - oForm01.Items.Item("Btn01").Top)) / 3 * 2 - 25;

//						oForm01.Items.Item("RadioMat02").Top = oForm01.Items.Item("Mat01").Height + oForm01.Items.Item("Mat01").Top - 4;
//						oForm01.Items.Item("RadioMat02").Left = 6;
//						oForm01.Items.Item("RadioMat02").Height = 20;

//						oForm01.Items.Item("33").Top = oForm01.Items.Item("RadioMat02").Top;
//						oForm01.Items.Item("ReWork").Top = oForm01.Items.Item("RadioMat02").Top;

//						oForm01.Items.Item("Mat02").Top = oForm01.Items.Item("Mat01").Height + oForm01.Items.Item("Mat01").Top + 15;
//						oForm01.Items.Item("Mat02").Left = oForm01.Items.Item("Mat01").Left;
//						oForm01.Items.Item("Mat02").Width = oForm01.Items.Item("Mat01").Width;
//						oForm01.Items.Item("Mat02").Height = oForm01.Items.Item("Mat01").Height / 2 + 5;
//						break;

//					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//						//27
//						break;

//				}
//			}

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			return;
//			Raise_ItemEvent_Error:

//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
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
//			int ReturnValue = 0;
//			string sQry = null;
//			string ErrNum = null;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
//						if (oLast_ItemUID == "Mat01") {
//							ReturnValue = SubMain.Sbo_Application.MessageBox("해당 라인의 작번을 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소");
//							switch (ReturnValue) {
//								case 1:
//									//작업지시등록여부 체크
//									sQry = "        SELECT  COUNT(*) AS [Cnt]";
//									sQry = sQry + " FROM    [@PS_PP030H]";
//									sQry = sQry + " WHERE   U_BaseNum = '" + oLast_RightClick_DocEntry + "'";
//									sQry = sQry + "         AND [Canceled] = 'N'";
//									sQry = sQry + "         AND U_OrdGbn IN ('105','106')";

//									oRecordSet01.DoQuery(sQry);

//									//작업지시등록이 존재하면
//									if (oRecordSet01.Fields.Item("Cnt").Value > 0) {
//										ErrNum = "1";
//										BubbleEvent = false;
//										goto Raise_MenuEvent_Error;
//									}

//									sQry = "        DELETE ";
//									sQry = sQry + " FROM    [@PS_PP020H] ";
//									sQry = sQry + " WHERE   DocEntry = '" + oLast_RightClick_DocEntry + "'";
//									oRecordSet01.DoQuery(sQry);

//									oLast_RightClick_DocEntry = Convert.ToString(0);
//									MDC_Com.MDC_GF_Message(ref "작번이 삭제되었습니다.", ref "S");
//									break;
//								case 2:
//									MDC_Com.MDC_GF_Message(ref "실행이 취소되었습니다.", ref "S");
//									BubbleEvent = false;
//									return;

//									break;
//							}
//						}
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
//					case "1299":
//						//행닫기

//						if (oLast_ItemUID == "Mat01") {
//							ReturnValue = SubMain.Sbo_Application.MessageBox("해당 라인의 작번등록을 [닫기]처리합니다. 복원할 수 없습니다. 진행하시겠습니까?", 1, "&확인", "&취소");
//							switch (ReturnValue) {
//								case 1:
//									sQry = "        UPDATE  [@PS_PP020H] ";
//									sQry = sQry + " SET     Status = 'C',";
//									sQry = sQry + "         UpdateDate = GETDATE(),";
//									sQry = sQry + "         UserSign = '" + SubMain.Sbo_Company.UserSignature + "'";
//									sQry = sQry + " Where   DocEntry = '" + oLast_RightClick_DocEntry + "'";
//									oRecordSet01.DoQuery(sQry);

//									oLast_RightClick_DocEntry = Convert.ToString(0);
//									MDC_Com.MDC_GF_Message(ref "작번등록이 [닫기]처리되었습니다.", ref "S");
//									break;
//								case 2:
//									MDC_Com.MDC_GF_Message(ref "실행이 취소되었습니다.", ref "S");
//									BubbleEvent = false;
//									return;

//									break;
//							}
//						}
//						break;
//				}
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
//						if (oMat01.RowCount != oMat01.VisualRowCount) {
//							oForm01.Freeze(true);
//							for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat01.Columns.Item("DocNum").Cells.Item(i + 1).Specific.VALUE = i + 1;
//							}

//							oMat01.FlushToDataSource();
//							oDS_PS_PP020H.RemoveRecord(oDS_PS_PP020H.Size - 1);
//							oMat01.Clear();
//							oMat01.LoadFromDataSource();
//							oForm01.Freeze(false);
//						}

//						if (oMat02.RowCount != oMat02.VisualRowCount - 1) {
//							oForm01.Freeze(true);
//							for (i = 0; i <= oMat02.VisualRowCount; i++) {
//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat01.Columns.Item("DocNum").Cells.Item(i + 1).Specific.VALUE = i + 1;
//							}

//							oMat02.FlushToDataSource();
//							oDS_PS_TEMPTABLE.RemoveRecord(oDS_PS_TEMPTABLE.Size - 1);
//							oMat02.Clear();
//							oMat02.LoadFromDataSource();
//							oForm01.Freeze(false);
//						}
//						break;
//					case "1281":
//						//찾기
//						FormItemEnabled();
//						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//						LoadCaption();
//						break;
//					case "1282":
//						//추가
//						oForm01.Freeze(true);
//						FormItemEnabled();

//						//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (Strings.Trim(oForm01.Items.Item("WorkGbn").Specific.VALUE) == "10") {
//							oMat01.Columns.Item("RegNum").Editable = true;
//							oMat01.Columns.Item("ItemCode").Editable = false;
//						} else {
//							oMat01.Columns.Item("RegNum").Editable = false;
//							oMat01.Columns.Item("ItemCode").Editable = true;
//						}

//						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						oMat01.Clear();
//						oMat01.FlushToDataSource();
//						oMat01.LoadFromDataSource();
//						Add_MatrixRow(0, true);
//						LoadCaption();
//						oForm01.Freeze(false);
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						FormItemEnabled();
//						if (oMat01.VisualRowCount > 0) {
//							//UPGRADE_WARNING: oMat01.Columns(CGNo).Cells(oMat01.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(oMat01.VisualRowCount).Specific.VALUE)) {
//								if (oDS_PS_PP020H.GetValue("Status", 0) == "O") {
//									Add_MatrixRow(oMat01.RowCount, false);
//								}
//							}
//						}
//						break;

//				}
//			}

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			return;
//			Raise_MenuEvent_Error:

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			if (ErrNum == "1") {
//				MDC_Com.MDC_GF_Message(ref "작업지시등록이 존재하는 작번입니다. 삭제할 수 없습니다.", ref "E");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((BusinessObjectInfo.BeforeAction == true)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						//33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						//34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						//35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						//36
//						break;
//				}
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						//33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						//34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						//35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						//36
//						break;
//				}
//			}
//			return;
//			Raise_FormDataEvent_Error:

//			MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((eventInfo.BeforeAction == true)) {
//				if (eventInfo.Row > 0) {
//					oLast_ItemUID = eventInfo.ItemUID;
//					if (oLast_ItemUID == "Mat01") {
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oLast_RightClick_DocEntry = Strings.Trim(oMat01.Columns.Item("DocEntry").Cells.Item(eventInfo.Row).Specific.VALUE);
//					}
//				}
//			} else if ((eventInfo.BeforeAction == false)) {
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			//디비데이터 소스 개체 할당
//			oDS_PS_PP020H = oForm01.DataSources.DBDataSources("@PS_PP020H");
//			oDS_PS_TEMPTABLE = oForm01.DataSources.DBDataSources("@PS_TEMPTABLE");

//			//메트릭스 개체 할당
//			oMat01 = oForm01.Items.Item("Mat01").Specific;
//			oMat02 = oForm01.Items.Item("Mat02").Specific;

//			oForm01.DataSources.UserDataSources.Add("PuDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("PuDateFr").Specific.DataBind.SetBound(true, "", "PuDateFr");

//			oForm01.DataSources.UserDataSources.Add("PuDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("PuDateTo").Specific.DataBind.SetBound(true, "", "PuDateTo");

//			oForm01.DataSources.UserDataSources.Add("JakDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("JakDateFr").Specific.DataBind.SetBound(true, "", "JakDateFr");

//			oForm01.DataSources.UserDataSources.Add("JakDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("JakDateTo").Specific.DataBind.SetBound(true, "", "JakDateTo");

//			oForm01.DataSources.UserDataSources.Add("RadioMat01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("RadioMat01").Specific.DataBind.SetBound(true, "", "RadioMat01");

//			oForm01.DataSources.UserDataSources.Add("RadioMat02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("RadioMat02").Specific.DataBind.SetBound(true, "", "RadioMat02");

//			//UPGRADE_WARNING: oForm01.Items().Specific.GroupWith 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm01.Items.Item("RadioMat01").Specific.GroupWith("RadioMat02");

//			oMat01.Columns.Item("OrderAmt").Visible = false;
//			//메인작번등록-견적금액 숨김
//			oMat01.Columns.Item("NegoAmt").Visible = false;
//			//메인작번등록-네고금액 숨김
//			oMat01.Columns.Item("TrgtAmt").Visible = false;
//			//메인작번등록-목표금액 숨김

//			oMat02.Columns.Item("OrderAmt").Visible = false;
//			//서브작번등록-견적금액 숨김
//			oMat02.Columns.Item("NegoAmt").Visible = false;
//			//서브작번등록-네고금액 숨김
//			oMat02.Columns.Item("TrgtAmt").Visible = false;
//			//서브작번등록-목표금액 숨김

//			return;
//			CreateItems_Error:

//			MDC_Com.MDC_GF_Message(ref "CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void ComboBox_Setting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//사업장
//			oCombo = oForm01.Items.Item("BPLId").Specific;
//			sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
//			oRecordSet01.DoQuery(sQry);
//			while (!(oRecordSet01.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
//				oRecordSet01.MoveNext();
//			}

//			//품목대분류
//			oCombo = oForm01.Items.Item("ItmBSort").Specific;
//			sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('105', '106') Order by Code";
//			oRecordSet01.DoQuery(sQry);
//			while (!(oRecordSet01.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
//				oMat01.Columns.Item("ItmBSort").ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
//				oRecordSet01.MoveNext();
//			}
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			//작업구분
//			oCombo = oForm01.Items.Item("WorkGbn").Specific;
//			oCombo.ValidValues.Add("10", "영업");
//			oCombo.ValidValues.Add("20", "정비");
//			oCombo.ValidValues.Add("30", "멀티");
//			oCombo.ValidValues.Add("40", "신동");
//			oCombo.ValidValues.Add("50", "R/D");
//			oCombo.ValidValues.Add("60", "견본");

//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			oMat01.Columns.Item("WorkGbn").ValidValues.Add("10", "영업");
//			oMat01.Columns.Item("WorkGbn").ValidValues.Add("20", "정비");
//			oMat01.Columns.Item("WorkGbn").ValidValues.Add("30", "멀티");
//			oMat01.Columns.Item("WorkGbn").ValidValues.Add("40", "신동");
//			oMat01.Columns.Item("WorkGbn").ValidValues.Add("50", "R/D");
//			oMat01.Columns.Item("WorkGbn").ValidValues.Add("60", "견본");

//			oMat02.Columns.Item("WorkGbn").ValidValues.Add("10", "영업");
//			oMat02.Columns.Item("WorkGbn").ValidValues.Add("20", "정비");
//			oMat02.Columns.Item("WorkGbn").ValidValues.Add("30", "멀티");
//			oMat02.Columns.Item("WorkGbn").ValidValues.Add("40", "신동");
//			oMat02.Columns.Item("WorkGbn").ValidValues.Add("50", "R/D");
//			oMat02.Columns.Item("WorkGbn").ValidValues.Add("60", "견본");

//			//작번구분
//			oCombo = oForm01.Items.Item("JakGbn").Specific;
//			oCombo.ValidValues.Add("99", "전체");
//			oCombo.ValidValues.Add("00", "메인작번");
//			oCombo.ValidValues.Add("01", "서브작번");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			//재작업구분
//			oCombo = oForm01.Items.Item("ReWork").Specific;
//			oCombo.ValidValues.Add("10", "정상");
//			oCombo.ValidValues.Add("20", "재작업");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			//작번구분
//			oMat01.Columns.Item("InOutGbn").ValidValues.Add("%", "선택");
//			oMat01.Columns.Item("InOutGbn").ValidValues.Add("IN", "자체");
//			oMat01.Columns.Item("InOutGbn").ValidValues.Add("OUT", "외주");

//			oMat02.Columns.Item("InOutGbn").ValidValues.Add("%", "선택");
//			oMat02.Columns.Item("InOutGbn").ValidValues.Add("IN", "자체");
//			oMat02.Columns.Item("InOutGbn").ValidValues.Add("OUT", "외주");

//			//재작업 사유(Mat01)
//			sQry = "        SELECT  B.U_Minor, ";
//			sQry = sQry + "         B.U_CdName";
//			sQry = sQry + " FROM    [@PS_SY001H] AS A";
//			sQry = sQry + "         INNER JOIN";
//			sQry = sQry + "         [@PS_SY001L] AS B";
//			sQry = sQry + "             ON A.Code = B.Code";
//			sQry = sQry + "             AND A.Code = 'P202'";

//			MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ReWkRn"), sQry);

//			//재작업 사유(Mat02)
//			MDC_Com.MDC_GP_MatrixSetMatComboList(oMat02.Columns.Item("ReWkRn"), sQry);

//			//연간품여부
//			oMat01.Columns.Item("YearPdYN").ValidValues.Add("N", "N");
//			oMat01.Columns.Item("YearPdYN").ValidValues.Add("Y", "Y");

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			return;
//			ComboBox_Setting_Error:

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

//			//아이디별 사업장 세팅
//			oCombo = oForm01.Items.Item("BPLId").Specific;
//			oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			return;
//			Initialization_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			MDC_Com.MDC_GF_Message(ref "Initialization_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void CF_ChooseFromList()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			CF_ChooseFromList_Error:

//			MDC_Com.MDC_GF_Message(ref "CF_ChooseFromList_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//				oForm01.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				oForm01.Items.Item("BPLId").Enabled = true;
//				oForm01.Items.Item("ItmBSort").Enabled = true;
//				oForm01.Items.Item("WorkGbn").Enabled = true;
//				oForm01.Items.Item("CardCode").Enabled = true;
//				oForm01.Items.Item("ItemCode").Enabled = true;
//				oForm01.Items.Item("PuDateFr").Enabled = true;
//				oForm01.Items.Item("PuDateTo").Enabled = true;
//				oForm01.Items.Item("JakDateFr").Enabled = false;
//				oForm01.Items.Item("JakDateTo").Enabled = false;
//				oForm01.Items.Item("JakGbn").Enabled = false;
//				oForm01.Items.Item("Btn02").Enabled = false;

//				oMat01.Columns.Item("JakName").Editable = true;
//				oMat01.Columns.Item("SubNo1").Editable = false;
//				oMat01.Columns.Item("SubNo2").Editable = false;
//				oMat01.Columns.Item("JakMyung").Editable = true;
//				oMat01.Columns.Item("JakSize").Editable = true;
//				oMat01.Columns.Item("JakUnit").Editable = true;
//				oMat01.Columns.Item("RegNum").Editable = false;
//				oMat01.Columns.Item("ItemCode").Editable = false;
//				oMat01.Columns.Item("CardCode").Editable = true;
//				oMat01.Columns.Item("ShipCode").Editable = true;
//				oMat01.Columns.Item("InOutGbn").Editable = true;
//				oMat01.Columns.Item("ProDate").Editable = true;
//				oMat01.Columns.Item("ReDate").Editable = true;
//				oMat01.Columns.Item("WrWeight").Editable = true;
//				oMat01.Columns.Item("Comments").Editable = true;
//				oMat01.Columns.Item("Status").Editable = true;

//				oForm01.Items.Item("RadioMat02").Visible = false;
//				oForm01.Items.Item("Mat02").Visible = false;
//			} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				oForm01.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				oForm01.Items.Item("BPLId").Enabled = true;
//				oForm01.Items.Item("ItmBSort").Enabled = true;
//				oForm01.Items.Item("WorkGbn").Enabled = true;
//				oForm01.Items.Item("CardCode").Enabled = true;
//				oForm01.Items.Item("ItemCode").Enabled = true;
//				oForm01.Items.Item("PuDateFr").Enabled = false;
//				oForm01.Items.Item("PuDateTo").Enabled = false;
//				oForm01.Items.Item("JakDateFr").Enabled = true;
//				oForm01.Items.Item("JakDateTo").Enabled = true;
//				oForm01.Items.Item("JakGbn").Enabled = true;
//				oForm01.Items.Item("Btn02").Enabled = true;

//				oMat01.Columns.Item("JakName").Editable = false;
//				oMat01.Columns.Item("SubNo1").Editable = false;
//				oMat01.Columns.Item("SubNo2").Editable = false;
//				oMat01.Columns.Item("JakMyung").Editable = true;
//				oMat01.Columns.Item("JakSize").Editable = true;
//				oMat01.Columns.Item("JakUnit").Editable = true;
//				oMat01.Columns.Item("RegNum").Editable = false;
//				oMat01.Columns.Item("ItemCode").Editable = false;
//				oMat01.Columns.Item("CardCode").Editable = true;
//				oMat01.Columns.Item("ShipCode").Editable = true;
//				oMat01.Columns.Item("InOutGbn").Editable = true;
//				oMat01.Columns.Item("ProDate").Editable = true;
//				oMat01.Columns.Item("ReDate").Editable = true;
//				oMat01.Columns.Item("WrWeight").Editable = true;
//				oMat01.Columns.Item("Comments").Editable = true;
//				oMat01.Columns.Item("Status").Editable = true;

//				oForm01.Items.Item("RadioMat02").Visible = true;
//				oForm01.Items.Item("Mat02").Visible = true;
//			} else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {

//			}
//			return;
//			FormItemEnabled_Error:

//			MDC_Com.MDC_GF_Message(ref "FormItemEnabled_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}
//		public void Add_MatrixRow(int oRow, ref bool RowIserted = false, ref string ItemUID = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (ItemUID == "Mat02") {
//				//행추가여부
//				if (RowIserted == false) {
//					oDS_PS_TEMPTABLE.InsertRecord((oRow));
//				}
//				oMat02.AddRow();
//				oDS_PS_TEMPTABLE.Offset = oRow;
//				oDS_PS_TEMPTABLE.SetValue("U_iField01", oRow, Convert.ToString(oRow + 1));
//				oMat02.LoadFromDataSource();
//			} else {
//				//행추가여부
//				if (RowIserted == false) {
//					oDS_PS_PP020H.InsertRecord((oRow));
//				}
//				oMat01.AddRow();
//				oDS_PS_PP020H.Offset = oRow;
//				oDS_PS_PP020H.SetValue("DocNum", oRow, Convert.ToString(oRow + 1));
//				oMat01.LoadFromDataSource();
//			}
//			return;
//			Add_MatrixRow_Error:

//			MDC_Com.MDC_GF_Message(ref "Add_MatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			short ErrNum = 0;
//			string sQry = null;
//			string ItemCode = null;
//			int Qty = 0;
//			int RegNum = 0;
//			double Calculate_Weight = 0;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string JakName = null;
//			switch (oUID) {
//				case "ShipCode":
//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = "Select CardName From OCRD Where CardCode = '" + Strings.Trim(oForm01.Items.Item("ShipCode").Specific.VALUE) + "'";
//					oRecordSet01.DoQuery(sQry);
//					//UPGRADE_WARNING: oForm01.Items(ShipName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm01.Items.Item("ShipName").Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
//					break;
//				case "ItemCode":
//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = "Select ItemName From OITM Where ItemCode = '" + Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE) + "'";
//					oRecordSet01.DoQuery(sQry);
//					//UPGRADE_WARNING: oForm01.Items(ItemName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm01.Items.Item("ItemName").Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
//					break;
//				case "CardCode":
//					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = "Select CardName From OCRD Where CardCode = '" + Strings.Trim(oForm01.Items.Item("CardCode").Specific.VALUE) + "'";
//					oRecordSet01.DoQuery(sQry);
//					//UPGRADE_WARNING: oForm01.Items(CardName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm01.Items.Item("CardName").Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
//					break;
//				case "Mat01":
//					if (oCol == "JakName") {
//						oForm01.Freeze(true);
//						if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if ((oRow == oMat01.RowCount | oMat01.VisualRowCount == 0) & !string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item("JakName").Cells.Item(oRow).Specific.VALUE))) {
//								oMat01.FlushToDataSource();
//								Add_MatrixRow(oMat01.RowCount, false);
//							}
//						}

//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						JakName = Strings.Trim(oMat01.Columns.Item("JakName").Cells.Item(oRow).Specific.VALUE);
//						sQry = "        Select  a.DocEntry,";
//						sQry = sQry + "         a.DocNum,";
//						sQry = sQry + "         a.Period,";
//						sQry = sQry + "         a.Instance,";
//						sQry = sQry + "         a.Series,";
//						sQry = sQry + "         a.Handwrtten,";
//						sQry = sQry + "         a.Canceled,";
//						sQry = sQry + "         a.Object,";
//						sQry = sQry + "         a.LogInst,";
//						sQry = sQry + "         a.UserSign,";
//						sQry = sQry + "         a.Transfered,";
//						sQry = sQry + "         a.Status,";
//						sQry = sQry + "         a.CreateDate,";
//						sQry = sQry + "         a.CreateTime,";
//						sQry = sQry + "         a.UpdateDate,";
//						sQry = sQry + "         a.UpdateTime,";
//						sQry = sQry + "         a.DataSource,";
//						sQry = sQry + "         a.U_BPLId,";
//						sQry = sQry + "         a.U_RegNum,";
//						sQry = sQry + "         a.U_ItemCode,";
//						sQry = sQry + "         a.U_ItemName,";
//						sQry = sQry + "         a.U_Material,";
//						sQry = sQry + "         a.U_Unit,";
//						sQry = sQry + "         a.U_Size,";
//						sQry = sQry + "         a.U_ItmBSort,";
//						sQry = sQry + "         a.U_CpName,";
//						sQry = sQry + "         a.U_SjDocNum,";
//						sQry = sQry + "         a.U_SjLinNum,";
//						sQry = sQry + "         a.U_SjQty,";
//						sQry = sQry + "         a.U_SjWeight,";
//						sQry = sQry + "         a.U_SjDcDate,";
//						sQry = sQry + "         a.U_SjDuDate,";
//						sQry = sQry + "         a.U_SlePrice,";
//						sQry = sQry + "         a.U_WorkGbn,";
//						sQry = sQry + "         a.U_CardCode,";
//						sQry = sQry + "         a.U_CardName,";
//						sQry = sQry + "         a.U_ShipCode,";
//						sQry = sQry + "         a.U_ShipName,";
//						sQry = sQry + "         a.U_PuDate,";
//						sQry = sQry + "         a.U_ProDate,";
//						sQry = sQry + "         a.U_Comments,";
//						sQry = sQry + "         a.U_JakName,";
//						sQry = sQry + "         a.U_SubNo1,";
//						sQry = sQry + "         a.U_SubNo2,";
//						sQry = sQry + "         a.U_Status,";
//						sQry = sQry + "         U_ReqCod = '" + MDC_PS_Common.User_MSTCOD() + "',";
//						sQry = sQry + "         a.U_UseDept,";
//						sQry = sQry + "         ISNULL(a.U_ReWeight, 0) AS U_ReWeight";
//						sQry = sQry + " FROM    [@PS_PP010H] a ";
//						sQry = sQry + "         LEFT JOIN ";
//						sQry = sQry + "         [@PS_SD010H] b ";
//						sQry = sQry + "             ON a.U_RegNum = b.U_RegNum";
//						sQry = sQry + " WHERE   a.U_JakName = '" + JakName + "'";

//						oRecordSet01.DoQuery(sQry);

//						if (!string.IsNullOrEmpty(Strings.Trim(oRecordSet01.Fields.Item(0).Value))) {
//							sQry = "        Select  a.DocEntry,";
//							sQry = sQry + "         a.DocNum,";
//							sQry = sQry + "         a.Period,";
//							sQry = sQry + "         a.Instance,";
//							sQry = sQry + "         a.Series,";
//							sQry = sQry + "         a.Handwrtten,";
//							sQry = sQry + "         a.Canceled,";
//							sQry = sQry + "         a.Object,";
//							sQry = sQry + "         a.LogInst,";
//							sQry = sQry + "         a.UserSign,";
//							sQry = sQry + "         a.Transfered,";
//							sQry = sQry + "         a.Status,";
//							sQry = sQry + "         a.CreateDate,";
//							sQry = sQry + "         a.CreateTime,";
//							sQry = sQry + "         a.UpdateDate,";
//							sQry = sQry + "         a.UpdateTime,";
//							sQry = sQry + "         a.DataSource,";
//							sQry = sQry + "         a.U_BPLId,";
//							sQry = sQry + "         a.U_RegNum,";
//							sQry = sQry + "         a.U_ItemCode,";
//							sQry = sQry + "         a.U_ItemName,";
//							sQry = sQry + "         a.U_Material,";
//							sQry = sQry + "         a.U_Unit,";
//							sQry = sQry + "         a.U_Size,";
//							sQry = sQry + "         a.U_ItmBSort,";
//							sQry = sQry + "         a.U_CpName,";
//							sQry = sQry + "         a.U_SjDocNum,";
//							sQry = sQry + "         a.U_SjLinNum,";
//							sQry = sQry + "         a.U_SjQty,";
//							sQry = sQry + "         a.U_SjWeight,";
//							sQry = sQry + "         a.U_SjDcDate,";
//							sQry = sQry + "         a.U_SjDuDate,";
//							sQry = sQry + "         a.U_SlePrice,";
//							sQry = sQry + "         a.U_WorkGbn,";
//							sQry = sQry + "         a.U_CardCode,";
//							sQry = sQry + "         a.U_CardName,";
//							sQry = sQry + "         a.U_ShipCode,";
//							sQry = sQry + "         a.U_ShipName,";
//							sQry = sQry + "         a.U_PuDate,";
//							sQry = sQry + "         a.U_ProDate,";
//							sQry = sQry + "         a.U_Comments,";
//							sQry = sQry + "         a.U_JakName,";
//							sQry = sQry + "         a.U_SubNo1,";
//							sQry = sQry + "         a.U_SubNo2,";
//							sQry = sQry + "         a.U_Status,";
//							sQry = sQry + "         U_ReqCod = '" + MDC_PS_Common.User_MSTCOD() + "',";
//							sQry = sQry + "         a.U_UseDept,";
//							sQry = sQry + "         ISNULL(a.U_ReWeight, 0) AS U_ReWeight,";
//							sQry = sQry + "         B.SalUnitMsr ";
//							sQry = sQry + " FROM    [@PS_PP010H] a ";
//							sQry = sQry + "         INNER JOIN ";
//							sQry = sQry + "         [OITM] b ";
//							sQry = sQry + "             ON a.U_ItemCode = b.ItemCode ";
//							sQry = sQry + " WHERE   a.U_jakName = '" + JakName + "'";

//							oRecordSet01.DoQuery(sQry);

//						}

//						//UPGRADE_WARNING: oMat01.Columns(SubNo1).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SubNo1").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_SubNo1").Value);
//						//UPGRADE_WARNING: oMat01.Columns(SubNo2).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SubNo2").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_SubNo2").Value);
//						//UPGRADE_WARNING: oMat01.Columns(RegNum).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("RegNum").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_RegNum").Value);
//						//UPGRADE_WARNING: oMat01.Columns(ItemCode).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_ItemCode").Value);
//						//UPGRADE_WARNING: oMat01.Columns(ItemName).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_ItemName").Value);
//						//UPGRADE_WARNING: oMat01.Columns(Material).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("Material").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_Material").Value);
//						//UPGRADE_WARNING: oMat01.Columns(Unit).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("Unit").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("SalUnitMsr").Value);
//						//UPGRADE_WARNING: oMat01.Columns(Size).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("Size").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_Size").Value);
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ItmBSort").Cells.Item(oRow).Specific.Select(Strings.Trim(oRecordSet01.Fields.Item("U_ItmBSort").Value), SAPbouiCOM.BoSearchKey.psk_ByValue);
//						//UPGRADE_WARNING: oMat01.Columns(SjDocNum).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SjDocNum").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_SjDocNum").Value);
//						//UPGRADE_WARNING: oMat01.Columns(SjLinNum).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SjLinNum").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_SjLinNum").Value);
//						//UPGRADE_WARNING: oMat01.Columns(CardCode).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("CardCode").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_CardCode").Value);
//						//UPGRADE_WARNING: oMat01.Columns(CardName).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("CardName").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_CardName").Value);
//						//UPGRADE_WARNING: oMat01.Columns(ShipCode).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ShipCode").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_ShipCode").Value);
//						//UPGRADE_WARNING: oMat01.Columns(ShipName).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ShipName").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_ShipName").Value);
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("InOutGbn").Cells.Item(oRow).Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//						//외주구분을 선택하도록 유도하기 위해 수정(2012.07.12 송명규)
//						//UPGRADE_WARNING: oMat01.Columns(JakDate).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("JakDate").Cells.Item(oRow).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
//						//UPGRADE_WARNING: oMat01.Columns(ProDate).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ProDate").Cells.Item(oRow).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_ProDate").Value), "YYYYMMDD");
//						//UPGRADE_WARNING: oMat01.Columns(ReDate).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ReDate").Cells.Item(oRow).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_SjDuDate").Value), "YYYYMMDD");
//						//UPGRADE_WARNING: oMat01.Columns(SjWeight).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SjWeight").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_SjWeight").Value);
//						//UPGRADE_WARNING: oMat01.Columns(SjDcDate).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SjDcDate").Cells.Item(oRow).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_SjDcDate").Value), "YYYYMMDD");
//						//UPGRADE_WARNING: oMat01.Columns(SjDuDate).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SjDuDate").Cells.Item(oRow).Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_SjDuDate").Value), "YYYYMMDD");
//						//UPGRADE_WARNING: oMat01.Columns(SlePrice).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("SlePrice").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_SlePrice").Value);
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("WorkGbn").Cells.Item(oRow).Specific.Select(Strings.Trim(oRecordSet01.Fields.Item("U_WorkGbn").Value), SAPbouiCOM.BoSearchKey.psk_ByValue);
//						//UPGRADE_WARNING: oMat01.Columns(PP010Doc).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("PP010Doc").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("DocEntry").Value);
//						//UPGRADE_WARNING: oMat01.Columns(ReqCod).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ReqCod").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_ReqCod").Value);
//						//UPGRADE_WARNING: oMat01.Columns(WrWeight).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("WrWeight").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("U_ReWeight").Value);
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("YearPdYN").Cells.Item(oRow).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
//						//연간품여부(2015.06.15 송명규 추가)

//						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + MDC_PS_Common.User_MSTCOD() + "'";
//						oRecordSet01.DoQuery(sQry);
//						//UPGRADE_WARNING: oMat01.Columns(ReqNam).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ReqNam").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);

//						oMat01.Columns.Item("Comments").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						oForm01.Freeze(false);
//					} else if (oCol == "ItemCode") {
//						oForm01.Freeze(true);
//						if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if ((oRow == oMat01.RowCount | oMat01.VisualRowCount == 0) & !string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.VALUE))) {
//								oMat01.FlushToDataSource();
//								Add_MatrixRow(oMat01.RowCount, false);
//							}
//						}

//						sQry = "        Select  a.ItemName, ";
//						sQry = sQry + "         a.U_Material, ";
//						sQry = sQry + "         a.SalUnitMsr, ";
//						sQry = sQry + "         a.U_Size, ";
//						sQry = sQry + "         a.U_ItmBSort, ";
//						sQry = sQry + "         b.U_CPNaming  ";
//						sQry = sQry + " From    OITM a ";
//						sQry = sQry + "         Left Join ";
//						sQry = sQry + "         [@PSH_ItmMSort] b";
//						sQry = sQry + "             ON a.U_ItmBSort = b.U_rCode ";
//						sQry = sQry + "             And a.U_ItmMsort = b.U_Code ";
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = sQry + " WHERE   a.ItemCode = '" + Strings.Trim(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.VALUE) + "'";
//						oRecordSet01.DoQuery(sQry);
//						//UPGRADE_WARNING: oMat01.Columns(ItemName).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);

//						oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						oForm01.Freeze(false);
//					} else if (oCol == "ShipCode") {
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = "Select CardName From OCRD Where CardCode = '" + Strings.Trim(oMat01.Columns.Item("ShipCode").Cells.Item(oRow).Specific.VALUE) + "'";
//						oRecordSet01.DoQuery(sQry);

//						//UPGRADE_WARNING: oMat01.Columns(ShipName).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ShipName").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);

//					} else if (oCol == "ReqCod") {

//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + Strings.Trim(oMat01.Columns.Item("ReqCod").Cells.Item(oRow).Specific.VALUE) + "'";
//						oRecordSet01.DoQuery(sQry);
//						//UPGRADE_WARNING: oMat01.Columns(ReqNam).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("ReqNam").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);

//					}
//					break;
//			}

//			return;
//			FlushToItemValue_Error:
//			oForm01.Freeze(false);
//			MDC_Com.MDC_GF_Message(ref "FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		private bool HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short ErrNum = 0;

//			ErrNum = 0;

//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE)):
//					ErrNum = 1;
//					goto HeaderSpaceLineDel_Error;
//					break;
//			}

//			functionReturnValue = true;
//			return functionReturnValue;
//			HeaderSpaceLineDel_Error:

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "사업장은 필수사항입니다. 확인하세요.", ref "E");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "HeaderSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private bool MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			short ErrNum = 0;
//			string sQry = null;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			ErrNum = 0;

//			oMat01.FlushToDataSource();

//			//라인
//			if (oMat01.VisualRowCount == 0) {
//				ErrNum = 1;
//				goto MatrixSpaceLineDel_Error;
//			}

//			for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
//				switch (true) {
//					//외주구분 필수 체크(2012.07.12 송명규 수정, 최수환이사 요청)
//					case Strings.Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", i - 1)) == "%":
//						ErrNum = 6;
//						goto MatrixSpaceLineDel_Error;
//						break;
//					//작번중복체크(서브작번 포함) 기능 필요
//				}
//			}
//			oMat01.LoadFromDataSource();

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			MatrixSpaceLineDel_Error:

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 6) {
//				MDC_Com.MDC_GF_Message(ref "" + i + "번 라인의 외주구분이 선택되지 않았습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 7) {
//				MDC_Com.MDC_GF_Message(ref "" + i + "번 라인의 작번은 이미 등록되었습니다. 확인하세요.", ref "E");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		public void Delete_EmptyRow()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;

//			//    oMat01.FlushToDataSource
//			//
//			//    For i = 0 To oMat01.VisualRowCount - 1
//			//        If Trim(oDS_PS_MM010L.GetValue("U_CGNo", i)) = "" Then
//			//            oDS_PS_MM010L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
//			//        End If
//			//    Next i
//			//
//			//    oMat01.LoadFromDataSource
//			return;
//			Delete_EmptyRow_Error:

//			MDC_Com.MDC_GF_Message(ref "Delete_EmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		private void LoadCaption()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				//UPGRADE_WARNING: oForm01.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm01.Items.Item("Btn01").Specific.Caption = "추가";
//			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				//UPGRADE_WARNING: oForm01.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm01.Items.Item("Btn01").Specific.Caption = "확인";
//			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//				//UPGRADE_WARNING: oForm01.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm01.Items.Item("Btn01").Specific.Caption = "갱신";
//			}

//			return;
//			LoadCaption_Error:

//			MDC_Com.MDC_GF_Message(ref "Delete_EmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public object PS_PP020_AddData()
//		{
//			object functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : PS_PP020_AddData()
//			//해당모듈    : PS_PP020
//			//기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 선행프로세스 일자 체크
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short loopCount = 0;
//			string ErrNum = null;
//			string ErrOrdNum = null;
//			//선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장

//			string Query01 = null;
//			string Query02 = null;

//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			SAPbobsCOM.Recordset RecordSet02 = null;
//			RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string DocEntry = null;
//			string BPLId = null;
//			string JakName = null;
//			string SubNo1 = null;
//			string SubNo2 = null;
//			string JakMyung = null;
//			string JakSize = null;
//			string JakUnit = null;
//			string RegNum = null;
//			string ItemCode = null;
//			string ItemName = null;
//			string Material = null;
//			string Unit = null;
//			string Size = null;
//			string ItmBsort = null;
//			string SjDocNum = null;
//			string SjLinNum = null;
//			string SjWeight = null;
//			string WrWeight = null;
//			string SjDcDate = null;
//			string SjDuDate = null;
//			string SlePrice = null;
//			string WorkGbn = null;
//			string CardCode = null;
//			string CardName = null;
//			string ShipCode = null;
//			string ShipName = null;
//			string InOutGbn = null;
//			string JakDate = null;
//			string ProDate = null;
//			string ReDate = null;
//			string Comments = null;
//			string PP010Doc = null;
//			string ReqCod = null;
//			string ReqNam = null;
//			string YearPdYN = null;
//			string OrderAmt = null;
//			string NegoAmt = null;
//			string TrgtAmt = null;
//			string DrawQty = null;
//			string ReWkRn = null;
//			string Status = null;

//			string BaseEntry = null;
//			string BaseLine = null;
//			string DocType = null;
//			string CurDocDate = null;

//			short MinusNum = 0;
//			//화면모드에 따라 VisualRowCount에서 빼줄 수

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, false);

//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			BPLId = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);

//			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				MinusNum = 2;
//			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//				MinusNum = 1;
//			}

//			oMat01.FlushToDataSource();
//			for (loopCount = 0; loopCount <= oMat01.VisualRowCount - MinusNum; loopCount++) {

//				DocEntry = Strings.Trim(oDS_PS_PP020H.GetValue("DocEntry", loopCount));
//				//DocEntry는 프로시저에서 처리
//				WorkGbn = Strings.Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", loopCount));
//				JakName = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakName", loopCount));
//				SubNo1 = Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo1", loopCount));
//				SubNo2 = Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", loopCount));
//				JakMyung = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakMyung", loopCount));
//				JakSize = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakSize", loopCount));
//				JakUnit = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakUnit", loopCount));
//				RegNum = Strings.Trim(oDS_PS_PP020H.GetValue("U_RegNum", loopCount));
//				ItemCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_ItemCode", loopCount));
//				ItemName = MDC_PS_Common.Make_ItemName(Strings.Trim(oDS_PS_PP020H.GetValue("U_ItemName", loopCount)));
//				Material = Strings.Trim(oDS_PS_PP020H.GetValue("U_Material", loopCount));
//				Unit = Strings.Trim(oDS_PS_PP020H.GetValue("U_Unit", loopCount));
//				Size = Strings.Trim(oDS_PS_PP020H.GetValue("U_Size", loopCount));
//				ReqCod = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReqCod", loopCount));
//				ReqNam = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReqNam", loopCount));
//				ItmBsort = Strings.Trim(oDS_PS_PP020H.GetValue("U_ItmBSort", loopCount));
//				SjDocNum = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDocNum", loopCount));
//				SjLinNum = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjLinNum", loopCount));
//				SjWeight = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjWeight", loopCount));
//				WrWeight = Strings.Trim(oDS_PS_PP020H.GetValue("U_WrWeight", loopCount));
//				SjDcDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDcDate", loopCount));
//				SjDuDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDuDate", loopCount));
//				SlePrice = Strings.Trim(oDS_PS_PP020H.GetValue("U_SlePrice", loopCount));
//				CardCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardCode", loopCount));
//				CardName = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardName", loopCount));
//				ShipCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipCode", loopCount));
//				ShipName = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipName", loopCount));
//				Comments = Strings.Trim(oDS_PS_PP020H.GetValue("U_Comments", loopCount));
//				InOutGbn = Strings.Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", loopCount));
//				JakDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakDate", loopCount));
//				ProDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_ProDate", loopCount));
//				ReDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReDate", loopCount));
//				PP010Doc = Strings.Trim(oDS_PS_PP020H.GetValue("U_PP010Doc", loopCount));
//				DrawQty = Strings.Trim(oDS_PS_PP020H.GetValue("U_DrawQty", loopCount));
//				YearPdYN = Strings.Trim(oDS_PS_PP020H.GetValue("U_YearPdYN", loopCount));
//				OrderAmt = Strings.Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", loopCount));
//				NegoAmt = Strings.Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", loopCount));
//				TrgtAmt = Strings.Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", loopCount));
//				ReWkRn = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReWkRn", loopCount));
//				Status = "O";

//				Query01 = "         EXEC PS_PP020_03 '";
//				Query01 = Query01 + DocEntry + "','";
//				Query01 = Query01 + BPLId + "','";
//				Query01 = Query01 + JakName + "','";
//				Query01 = Query01 + SubNo1 + "','";
//				Query01 = Query01 + SubNo2 + "','";
//				Query01 = Query01 + JakMyung + "','";
//				Query01 = Query01 + JakSize + "','";
//				Query01 = Query01 + JakUnit + "','";
//				Query01 = Query01 + RegNum + "','";
//				Query01 = Query01 + ItemCode + "','";
//				Query01 = Query01 + ItemName + "','";
//				Query01 = Query01 + Material + "','";
//				Query01 = Query01 + Unit + "','";
//				Query01 = Query01 + Size + "','";
//				Query01 = Query01 + ItmBsort + "','";
//				Query01 = Query01 + SjDocNum + "','";
//				Query01 = Query01 + SjLinNum + "','";
//				Query01 = Query01 + SjWeight + "','";
//				Query01 = Query01 + WrWeight + "','";
//				Query01 = Query01 + SjDcDate + "','";
//				Query01 = Query01 + SjDuDate + "','";
//				Query01 = Query01 + SlePrice + "','";
//				Query01 = Query01 + WorkGbn + "','";
//				Query01 = Query01 + CardCode + "','";
//				Query01 = Query01 + CardName + "','";
//				Query01 = Query01 + ShipCode + "','";
//				Query01 = Query01 + ShipName + "','";
//				Query01 = Query01 + InOutGbn + "','";
//				Query01 = Query01 + JakDate + "','";
//				Query01 = Query01 + ProDate + "','";
//				Query01 = Query01 + ReDate + "','";
//				Query01 = Query01 + Comments + "','";
//				Query01 = Query01 + PP010Doc + "','";
//				Query01 = Query01 + ReqCod + "','";
//				Query01 = Query01 + ReqNam + "','";
//				Query01 = Query01 + YearPdYN + "','";
//				Query01 = Query01 + OrderAmt + "','";
//				Query01 = Query01 + NegoAmt + "','";
//				Query01 = Query01 + TrgtAmt + "','";
//				Query01 = Query01 + DrawQty + "','";
//				Query01 = Query01 + ReWkRn + "','";
//				Query01 = Query01 + Status + "'";

//				//선행프로세스 대비 일자체크_S
//				//Entry = Split(SjDocLin, "-")
//				BaseEntry = PP010Doc;
//				BaseLine = "0";
//				DocType = "PS_PP020";
//				CurDocDate = JakDate;

//				Query02 = "         EXEC PS_Z_CHECK_DATE '";
//				Query02 = Query02 + BaseEntry + "','";
//				Query02 = Query02 + BaseLine + "','";
//				Query02 = Query02 + DocType + "','";
//				Query02 = Query02 + CurDocDate + "'";

//				RecordSet02.DoQuery(Query02);
//				//선행프로세스 대비 일자체크_E

//				if (RecordSet02.Fields.Item("ReturnValue").Value == "True") {

//					RecordSet01.DoQuery(Query01);
//					//등록

//				} else {

//					ErrOrdNum = ErrOrdNum + " [" + ItemCode + "]";

//				}

//			}

//			//하나라도 선행프로세스 일자가 빠른 작번이 있으면
//			if (!string.IsNullOrEmpty(ErrOrdNum)) {
//				ErrNum = "1";
//				goto PS_PP020_AddData_Error;
//			}

//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			MDC_Com.MDC_GF_Message(ref "전체 저장 완료!", ref "S");

//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet02 = null;

//			//UPGRADE_WARNING: PS_PP020_AddData 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = true;
//			return functionReturnValue;
//			PS_PP020_AddData_Error:



//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			//UPGRADE_WARNING: PS_PP020_AddData 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = false;
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet02 = null;

//			if (ErrNum == "1") {

//				SubMain.Sbo_Application.MessageBox("작번등록일은 생산의뢰접수일과 같거나 늦어야합니다. 확인하십시오." + Strings.Chr(13) + ErrOrdNum, 1);

//				//등록되지 않은 작번이 있어도 화면 Clear_S
//				oMat01.Clear();
//				oMat01.FlushToDataSource();
//				oMat01.LoadFromDataSource();
//				Add_MatrixRow(0, true);
//				//등록되지 않은 작번이 있어도 화면 Clear_E

//			} else {

//				MDC_Com.MDC_GF_Message(ref "PS_PP020_AddData_Error:" + Err().Number + " - " + Err().Description, ref "E");

//			}
//			return functionReturnValue;

//		}

//		public bool Add_JakNum(ref SAPbouiCOM.ItemEvent pval)
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : Add_JakNum()
//			//해당모듈    : PS_PP020
//			//기능        : 데이터 INSERT
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD020_AddData 로 대체 (2017.02.27 송명규)
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string SjLinNum = null;
//			string CpName = null;
//			string Size = null;
//			string Material = null;
//			string ItemCode = null;
//			string BPLId = null;
//			string DocEntry = null;
//			string RegNum = null;
//			string ItemName = null;
//			string Unit = null;
//			string ItmBsort = null;
//			string SjDocNum = null;
//			string PP010Doc = null;
//			int SjQty = 0;
//			decimal SjWeight = default(decimal);
//			decimal SlePrice = default(decimal);
//			string ShipCode = null;
//			string CardCode = null;
//			string SjDuDate = null;
//			string SjDcDate = null;
//			string WorkGbn = null;
//			string CardName = null;
//			string ShipName = null;
//			string JakSize = null;
//			string ReDate = null;
//			string JakDate = null;
//			string Temp01 = null;
//			string SubNo1 = null;
//			string Comments = null;
//			string JakName = null;
//			string SubNo2 = null;
//			string InOutGbn = null;
//			string ProDate = null;
//			string JakMyung = null;
//			string JakUnit = null;
//			decimal WrWeight = default(decimal);

//			double DrawQty = 0;
//			string YearPdYN = null;
//			//연간품여부(2015.06.15 송명규)
//			string ReqCod = null;
//			//요청자(2018.07.27 황영수)
//			string ReqNam = null;
//			//요청자(2018.07.27 황영수)

//			decimal OrderAmt = default(decimal);
//			//수주금액(2015.10.12 송명규)
//			decimal NegoAmt = default(decimal);
//			//Nego금액(2015.10.12 송명규)
//			decimal TrgtAmt = default(decimal);
//			//목표금액(생산)(2015.10.12 송명규)

//			oMat01.FlushToDataSource();

//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			BPLId = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);
//			for (i = 0; i <= oMat01.RowCount - 2; i++) {
//				WorkGbn = Strings.Trim(oDS_PS_PP020H.GetValue("U_WorkGbn", i));
//				JakName = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakName", i));
//				SubNo1 = Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo1", i));
//				SubNo2 = Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", i));
//				JakMyung = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakMyung", i));
//				JakSize = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakSize", i));
//				JakUnit = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakUnit", i));
//				RegNum = Strings.Trim(oDS_PS_PP020H.GetValue("U_RegNum", i));
//				ItemCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_ItemCode", i));
//				ItemName = MDC_PS_Common.Make_ItemName(Strings.Trim(oDS_PS_PP020H.GetValue("U_ItemName", i)));
//				Material = Strings.Trim(oDS_PS_PP020H.GetValue("U_Material", i));
//				Unit = Strings.Trim(oDS_PS_PP020H.GetValue("U_Unit", i));
//				Size = Strings.Trim(oDS_PS_PP020H.GetValue("U_Size", i));
//				ItmBsort = Strings.Trim(oDS_PS_PP020H.GetValue("U_ItmBSort", i));
//				SjDocNum = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDocNum", i));
//				SjLinNum = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjLinNum", i));
//				SjWeight = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_SjWeight", i)));
//				WrWeight = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_WrWeight", i)));
//				SjDcDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDcDate", i));
//				SjDuDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_SjDuDate", i));
//				SlePrice = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_SlePrice", i)));
//				CardCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardCode", i));
//				CardName = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardName", i));
//				ShipCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipCode", i));
//				ShipName = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipName", i));
//				Comments = Strings.Trim(oDS_PS_PP020H.GetValue("U_Comments", i));
//				InOutGbn = Strings.Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", i));
//				JakDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakDate", i));
//				ProDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_ProDate", i));
//				ReDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReDate", i));
//				PP010Doc = Strings.Trim(oDS_PS_PP020H.GetValue("U_PP010Doc", i));
//				ReqCod = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReqCod", i));
//				ReqNam = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReqNam", i));
//				DrawQty = Convert.ToDouble(Strings.Trim(oDS_PS_PP020H.GetValue("U_DrawQty", i)));
//				YearPdYN = Strings.Trim(oDS_PS_PP020H.GetValue("U_YearPdYN", i));
//				OrderAmt = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", i)));
//				NegoAmt = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", i)));
//				TrgtAmt = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", i)));

//				//DocEntry
//				sQry = "Select IsNULL(Max(DocEntry), 0) From [@PS_PP020H] ";
//				RecordSet01.DoQuery(sQry);
//				if (Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) == 0) {
//					DocEntry = Convert.ToString(1);
//				} else {
//					DocEntry = Convert.ToString(Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1);
//				}

//				sQry = "INSERT INTO [@PS_PP020H]";
//				sQry = sQry + " (";
//				sQry = sQry + " DocEntry,";
//				sQry = sQry + " DocNum,";
//				sQry = sQry + " U_BPLId,";
//				sQry = sQry + " U_JakName,";
//				sQry = sQry + " U_SubNo1,";
//				sQry = sQry + " U_SubNo2,";
//				sQry = sQry + " U_JakMyung,";
//				sQry = sQry + " U_JakSize,";
//				sQry = sQry + " U_JakUnit,";
//				sQry = sQry + " U_RegNum,";
//				sQry = sQry + " U_ItemCode,";
//				sQry = sQry + " U_ItemName,";
//				sQry = sQry + " U_Material,";
//				sQry = sQry + " U_Unit,";
//				sQry = sQry + " U_Size,";
//				sQry = sQry + " U_ItmBSort,";
//				sQry = sQry + " U_SjDocNum,";
//				sQry = sQry + " U_SjLinNum,";
//				sQry = sQry + " U_SjWeight,";
//				sQry = sQry + " U_WrWeight,";
//				sQry = sQry + " U_SjDcDate,";
//				sQry = sQry + " U_SjDuDate,";
//				sQry = sQry + " U_SlePrice,";
//				sQry = sQry + " U_WorkGbn,";
//				sQry = sQry + " U_CardCode,";
//				sQry = sQry + " U_CardName,";
//				sQry = sQry + " U_ShipCode,";
//				sQry = sQry + " U_ShipName,";
//				sQry = sQry + " U_InOutGbn,";
//				sQry = sQry + " U_JakDate,";
//				sQry = sQry + " U_ProDate,";
//				sQry = sQry + " U_ReDate,";
//				sQry = sQry + " U_Comments,";
//				sQry = sQry + " U_PP010Doc,";
//				sQry = sQry + " U_YearPdYN,";
//				sQry = sQry + " U_OrderAmt,";
//				//수주금액
//				sQry = sQry + " U_NegoAmt,";
//				//Nego금액
//				sQry = sQry + " U_TrgtAmt,";
//				//목표금액
//				sQry = sQry + " U_Status,";
//				sQry = sQry + " U_DrawQty,";
//				sQry = sQry + " CreateDate";
//				sQry = sQry + " U_ReqCod,";
//				sQry = sQry + " U_ReqNam,";
//				sQry = sQry + " ) ";
//				sQry = sQry + "VALUES(";
//				sQry = sQry + DocEntry + ",";
//				sQry = sQry + DocEntry + ",";
//				sQry = sQry + "'" + BPLId + "',";
//				sQry = sQry + "'" + JakName + "',";
//				sQry = sQry + "'" + SubNo1 + "',";
//				sQry = sQry + "'" + SubNo2 + "',";
//				sQry = sQry + "'" + JakMyung + "',";
//				sQry = sQry + "'" + JakSize + "',";
//				sQry = sQry + "'" + JakUnit + "',";
//				sQry = sQry + "'" + RegNum + "',";
//				sQry = sQry + "'" + ItemCode + "',";
//				sQry = sQry + "'" + ItemName + "',";
//				sQry = sQry + "'" + Material + "',";
//				sQry = sQry + "'" + Unit + "',";
//				sQry = sQry + "'" + Size + "',";
//				sQry = sQry + "'" + ItmBsort + "',";
//				sQry = sQry + "'" + SjDocNum + "',";
//				sQry = sQry + "'" + SjLinNum + "',";
//				sQry = sQry + "'" + SjWeight + "',";
//				sQry = sQry + "'" + WrWeight + "',";
//				sQry = sQry + "'" + SjDcDate + "',";
//				sQry = sQry + "'" + SjDuDate + "',";
//				sQry = sQry + "'" + SlePrice + "',";
//				sQry = sQry + "'" + WorkGbn + "',";
//				sQry = sQry + "'" + CardCode + "',";
//				sQry = sQry + "'" + CardName + "',";
//				sQry = sQry + "'" + ShipCode + "',";
//				sQry = sQry + "'" + ShipName + "',";
//				sQry = sQry + "'" + InOutGbn + "',";
//				sQry = sQry + "'" + JakDate + "',";
//				sQry = sQry + "'" + ProDate + "',";
//				sQry = sQry + "'" + ReDate + "',";
//				sQry = sQry + "'" + Comments + "',";
//				sQry = sQry + "'" + PP010Doc + "',";
//				sQry = sQry + "'" + YearPdYN + "',";
//				sQry = sQry + "'" + OrderAmt + "',";
//				//수주금액
//				sQry = sQry + "'" + NegoAmt + "',";
//				//Nego금액
//				sQry = sQry + "'" + TrgtAmt + "',";
//				//목표금액
//				sQry = sQry + "'O',";
//				sQry = sQry + "'" + DrawQty + "',";
//				sQry = sQry + " GETDATE()";
//				sQry = sQry + "'" + ReqCod + "',";
//				sQry = sQry + "'" + ReqNam + "',";
//				sQry = sQry + ")";
//				RecordSet01.DoQuery(sQry);

//				sQry = "Update [@PS_PP010H] Set U_Status = 'C' Where DocEntry = '" + PP010Doc + "'";
//				RecordSet01.DoQuery(sQry);
//			}

//			MDC_Com.MDC_GF_Message(ref "작번등록 완료!", ref "S");

//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			Add_JakNum_Error:

//			functionReturnValue = false;
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "Add_JakNum_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public bool Update_JakNum(ref SAPbouiCOM.ItemEvent pval)
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : Update_JakNum()
//			//해당모듈    : PS_PP020
//			//기능        : 데이터 UPDATE
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 2017.02.17 패치 이후 사용 안함. PS_SD020_AddData 로 대체 (2017.02.17 송명규)
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string SjDocNum = null;
//			string ItmBsort = null;
//			string Unit = null;
//			string ItemName = null;
//			string RegNum = null;
//			string ItemCode = null;
//			string Material = null;
//			string Size = null;
//			string CpName = null;
//			string SjLinNum = null;
//			int SjQty = 0;
//			decimal SjWeight = default(decimal);
//			decimal SlePrice = default(decimal);
//			string ShipCode = null;
//			string CardCode = null;
//			string SjDuDate = null;
//			string SjDcDate = null;
//			string WorkGbn = null;
//			string CardName = null;
//			string ShipName = null;
//			string JakSize = null;
//			string ReDate = null;
//			string Status = null;
//			string SubNo2 = null;
//			string JakName = null;
//			string ProDate = null;
//			string PuDate = null;
//			string Comments = null;
//			string SubNo1 = null;
//			string Temp01 = null;
//			string InOutGbn = null;
//			string JakMyung = null;
//			string JakUnit = null;
//			decimal WrWeight = default(decimal);

//			double DrawQty = 0;
//			string ReWkRn = null;
//			//재작업사유(2013.09.10 송명규 추가)
//			string YearPdYN = null;
//			//연간품여부(2015.06.15 송명규 추가)

//			decimal OrderAmt = default(decimal);
//			//수주금액(2015.10.12 송명규)
//			decimal NegoAmt = default(decimal);
//			//Nego금액(2015.10.12 송명규)
//			decimal TrgtAmt = default(decimal);
//			//목표금액(생산)(2015.10.12 송명규)

//			oMat01.FlushToDataSource();

//			for (i = 0; i <= oMat01.RowCount - 1; i++) {
//				JakMyung = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakMyung", i));
//				JakSize = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakSize", i));
//				JakUnit = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakUnit", i));
//				CardCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardCode", i));
//				CardName = Strings.Trim(oDS_PS_PP020H.GetValue("U_CardName", i));
//				ShipCode = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipCode", i));
//				ShipName = Strings.Trim(oDS_PS_PP020H.GetValue("U_ShipName", i));
//				InOutGbn = Strings.Trim(oDS_PS_PP020H.GetValue("U_InOutGbn", i));
//				ProDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_ProDate", i));
//				ReDate = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReDate", i));
//				Comments = Strings.Trim(oDS_PS_PP020H.GetValue("U_Comments", i));
//				JakName = Strings.Trim(oDS_PS_PP020H.GetValue("U_JakName", i));
//				SubNo1 = Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo1", i));
//				SubNo2 = Strings.Trim(oDS_PS_PP020H.GetValue("U_SubNo2", i));
//				Status = Strings.Trim(oDS_PS_PP020H.GetValue("U_Status", i));
//				WrWeight = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_WrWeight", i)));
//				DrawQty = Convert.ToDouble(Strings.Trim(oDS_PS_PP020H.GetValue("U_DrawQty", i)));
//				ReWkRn = Strings.Trim(oDS_PS_PP020H.GetValue("U_ReWkRn", i));
//				YearPdYN = Strings.Trim(oDS_PS_PP020H.GetValue("U_YearPdYN", i));

//				OrderAmt = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_OrderAmt", i)));
//				NegoAmt = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_NegoAmt", i)));
//				TrgtAmt = Convert.ToDecimal(Strings.Trim(oDS_PS_PP020H.GetValue("U_TrgtAmt", i)));

//				sQry = "UPDATE [@PS_PP020H] ";
//				sQry = sQry + "SET ";
//				sQry = sQry + "U_JakMyung = '" + JakMyung + "', ";
//				sQry = sQry + "U_JakSize = '" + JakSize + "', ";
//				sQry = sQry + "U_JakUnit = '" + JakUnit + "', ";
//				sQry = sQry + "U_CardCode = '" + CardCode + "', ";
//				sQry = sQry + "U_CardName = '" + CardName + "', ";
//				sQry = sQry + "U_ShipCode = '" + ShipCode + "', ";
//				sQry = sQry + "U_ShipName = '" + ShipName + "', ";
//				sQry = sQry + "U_InOutGbn = '" + InOutGbn + "', ";
//				sQry = sQry + "U_ProDate = '" + ProDate + "', ";
//				sQry = sQry + "U_ReDate = '" + ReDate + "', ";
//				sQry = sQry + "U_WrWeight = '" + WrWeight + "', ";
//				sQry = sQry + "U_Comments = '" + Comments + "', ";
//				sQry = sQry + "U_Status = '" + Status + "', ";
//				sQry = sQry + "U_DrawQty = '" + DrawQty + "', ";
//				sQry = sQry + "U_YearPdYN = '" + YearPdYN + "', ";
//				sQry = sQry + "U_OrderAmt = '" + OrderAmt + "', ";
//				//수주금액
//				sQry = sQry + "U_NegoAmt = '" + NegoAmt + "', ";
//				//Nego금액
//				sQry = sQry + "U_TrgtAmt = '" + TrgtAmt + "', ";
//				//목표금액
//				sQry = sQry + "U_ReWkRn = '" + ReWkRn + "' ";
//				sQry = sQry + "Where U_JakName = '" + JakName + "' And U_SubNo1 = '" + SubNo1 + "' And U_SubNo2 = '" + SubNo2 + "'";

//				RecordSet01.DoQuery(sQry);
//			}

//			MDC_Com.MDC_GF_Message(ref "작번등록 갱신 완료!", ref "S");

//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			Update_JakNum_Error:

//			functionReturnValue = false;
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "Update_JakNum_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public void LoadData()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string SjDocNum = null;
//			string ItmBsort = null;
//			string Unit = null;
//			string ItemName = null;
//			string RegNum = null;
//			string ItemCode = null;
//			string Material = null;
//			string Size = null;
//			string CpName = null;
//			string SjLinNum = null;
//			int SjQty = 0;
//			decimal SjWeight = default(decimal);
//			decimal SlePrice = default(decimal);
//			string ShipCode = null;
//			string CardCode = null;
//			string SjDuDate = null;
//			string SjDcDate = null;
//			string WorkGbn = null;
//			string CardName = null;
//			string ShipName = null;
//			string JakDateTo = null;
//			string BPLId = null;
//			string Temp01 = null;
//			string SubNo1 = null;
//			string Comments = null;
//			string PuDate = null;
//			string ProDate = null;
//			string JakName = null;
//			string SubNo2 = null;
//			string Status = null;
//			string JakDateFr = null;
//			string JakGbn = null;

//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			BPLId = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ItmBsort = Strings.Trim(oForm01.Items.Item("ItmBSort").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			WorkGbn = Strings.Trim(oForm01.Items.Item("WorkGbn").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CardCode = Strings.Trim(oForm01.Items.Item("CardCode").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ItemCode = Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JakDateFr = Strings.Trim(oForm01.Items.Item("JakDateFr").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JakDateTo = Strings.Trim(oForm01.Items.Item("JakDateTo").Specific.VALUE);
//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JakGbn = Strings.Trim(oForm01.Items.Item("JakGbn").Specific.VALUE);

//			if (string.IsNullOrEmpty(BPLId))
//				BPLId = "%";
//			if (string.IsNullOrEmpty(CardCode))
//				CardCode = "%";
//			if (string.IsNullOrEmpty(ItemCode))
//				ItemCode = "%";
//			if (string.IsNullOrEmpty(JakDateFr))
//				JakDateFr = "19000101";
//			if (string.IsNullOrEmpty(JakDateTo))
//				JakDateTo = "20991231";
//			if (string.IsNullOrEmpty(JakGbn))
//				JakGbn = "%";

//			sQry = "      EXEC [PS_PP020_01] '";
//			sQry = sQry + BPLId + "','";
//			sQry = sQry + ItmBsort + "','";
//			sQry = sQry + WorkGbn + "','";
//			sQry = sQry + CardCode + "','";
//			sQry = sQry + ItemCode + "','";
//			sQry = sQry + JakDateFr + "','";
//			sQry = sQry + JakDateTo + "','";
//			sQry = sQry + JakGbn + "'";

//			oRecordSet01.DoQuery(sQry);

//			oMat01.Clear();
//			oDS_PS_PP020H.Clear();

//			if ((oRecordSet01.RecordCount == 0)) {
//				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
//				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet01 = null;
//				return;
//			}

//			oForm01.Freeze(true);
//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
//				if (i + 1 > oDS_PS_PP020H.Size) {
//					oDS_PS_PP020H.InsertRecord((i));
//				}

//				oMat01.AddRow();
//				oDS_PS_PP020H.Offset = i;
//				oDS_PS_PP020H.SetValue("DocNum", i, Convert.ToString(i + 1));
//				oDS_PS_PP020H.SetValue("DocEntry", i, Strings.Trim(oRecordSet01.Fields.Item("DocEntry").Value));
//				oDS_PS_PP020H.SetValue("U_JakName", i, Strings.Trim(oRecordSet01.Fields.Item("U_JakName").Value));
//				oDS_PS_PP020H.SetValue("U_SubNo1", i, Strings.Trim(oRecordSet01.Fields.Item("U_SubNo1").Value));
//				oDS_PS_PP020H.SetValue("U_SubNo2", i, Strings.Trim(oRecordSet01.Fields.Item("U_SubNo2").Value));
//				oDS_PS_PP020H.SetValue("U_JakMyung", i, Strings.Trim(oRecordSet01.Fields.Item("U_JakMyung").Value));
//				oDS_PS_PP020H.SetValue("U_JakSize", i, Strings.Trim(oRecordSet01.Fields.Item("U_JakSize").Value));
//				oDS_PS_PP020H.SetValue("U_JakUnit", i, Strings.Trim(oRecordSet01.Fields.Item("U_JakUnit").Value));
//				oDS_PS_PP020H.SetValue("U_RegNum", i, Strings.Trim(oRecordSet01.Fields.Item("U_RegNum").Value));
//				oDS_PS_PP020H.SetValue("U_ItemCode", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItemCode").Value));
//				oDS_PS_PP020H.SetValue("U_ItemName", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItemName").Value));
//				oDS_PS_PP020H.SetValue("U_Material", i, Strings.Trim(oRecordSet01.Fields.Item("U_Material").Value));
//				oDS_PS_PP020H.SetValue("U_Unit", i, Strings.Trim(oRecordSet01.Fields.Item("U_Unit").Value));
//				oDS_PS_PP020H.SetValue("U_Size", i, Strings.Trim(oRecordSet01.Fields.Item("U_Size").Value));
//				oDS_PS_PP020H.SetValue("U_ItmBSort", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItmBSort").Value));
//				oDS_PS_PP020H.SetValue("U_SjDocNum", i, Strings.Trim(oRecordSet01.Fields.Item("U_SjDocNum").Value));
//				oDS_PS_PP020H.SetValue("U_SjLinNum", i, Strings.Trim(oRecordSet01.Fields.Item("U_SjLinNum").Value));
//				oDS_PS_PP020H.SetValue("U_SjWeight", i, Strings.Trim(oRecordSet01.Fields.Item("U_SjWeight").Value));
//				oDS_PS_PP020H.SetValue("U_WrWeight", i, Strings.Trim(oRecordSet01.Fields.Item("U_WrWeight").Value));
//				oDS_PS_PP020H.SetValue("U_SjDcDate", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_SjDcDate").Value), "YYYYMMDD"));
//				oDS_PS_PP020H.SetValue("U_SjDuDate", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_SjDuDate").Value), "YYYYMMDD"));
//				oDS_PS_PP020H.SetValue("U_SlePrice", i, Strings.Trim(oRecordSet01.Fields.Item("U_SlePrice").Value));
//				oDS_PS_PP020H.SetValue("U_WorkGbn", i, Strings.Trim(oRecordSet01.Fields.Item("U_WorkGbn").Value));
//				oDS_PS_PP020H.SetValue("U_CardCode", i, Strings.Trim(oRecordSet01.Fields.Item("U_CardCode").Value));
//				oDS_PS_PP020H.SetValue("U_CardName", i, Strings.Trim(oRecordSet01.Fields.Item("U_CardName").Value));
//				oDS_PS_PP020H.SetValue("U_ShipCode", i, Strings.Trim(oRecordSet01.Fields.Item("U_ShipCode").Value));
//				oDS_PS_PP020H.SetValue("U_ShipName", i, Strings.Trim(oRecordSet01.Fields.Item("U_ShipName").Value));
//				oDS_PS_PP020H.SetValue("U_JakDate", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_JakDate").Value), "YYYYMMDD"));
//				oDS_PS_PP020H.SetValue("U_ProDate", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_ProDate").Value), "YYYYMMDD"));
//				oDS_PS_PP020H.SetValue("U_ReDate", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_ReDate").Value), "YYYYMMDD"));
//				oDS_PS_PP020H.SetValue("U_Comments", i, Strings.Trim(oRecordSet01.Fields.Item("U_Comments").Value));
//				oDS_PS_PP020H.SetValue("U_InOutGbn", i, Strings.Trim(oRecordSet01.Fields.Item("U_InOutGbn").Value));
//				oDS_PS_PP020H.SetValue("U_Status", i, Strings.Trim(oRecordSet01.Fields.Item("U_Status").Value));
//				oDS_PS_PP020H.SetValue("U_ReqCod", i, Strings.Trim(oRecordSet01.Fields.Item("U_ReqCod").Value));
//				oDS_PS_PP020H.SetValue("U_ReqNam", i, Strings.Trim(oRecordSet01.Fields.Item("U_ReqNam").Value));
//				oDS_PS_PP020H.SetValue("U_PP010Doc", i, Strings.Trim(oRecordSet01.Fields.Item("U_PP010Doc").Value));
//				oDS_PS_PP020H.SetValue("U_DrawQty", i, Strings.Trim(oRecordSet01.Fields.Item("U_DrawQty").Value));
//				oDS_PS_PP020H.SetValue("U_ReWkRn", i, Strings.Trim(oRecordSet01.Fields.Item("U_ReWkRn").Value));
//				oDS_PS_PP020H.SetValue("U_YearPdYN", i, Strings.Trim(oRecordSet01.Fields.Item("U_YearPdYN").Value));

//				oDS_PS_PP020H.SetValue("U_OrderAmt", i, Strings.Trim(oRecordSet01.Fields.Item("U_OrderAmt").Value));
//				//수주금액
//				oDS_PS_PP020H.SetValue("U_NegoAmt", i, Strings.Trim(oRecordSet01.Fields.Item("U_NegoAmt").Value));
//				//Nego금액
//				oDS_PS_PP020H.SetValue("U_TrgtAmt", i, Strings.Trim(oRecordSet01.Fields.Item("U_TrgtAmt").Value));
//				//목표금액

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

//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "LoadData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public bool Add_SubJakNum(ref SAPbouiCOM.ItemEvent pval)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			int ErrNum = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string SjDocNum = null;
//			string ItmBsort = null;
//			string Unit = null;
//			string ItemName = null;
//			string RegNum = null;
//			string DocEntry = null;
//			string BPLId = null;
//			string ItemCode = null;
//			string Material = null;
//			string Size = null;
//			string CpName = null;
//			string SjLinNum = null;
//			int SjQty = 0;
//			decimal SjWeight = default(decimal);
//			decimal SlePrice = default(decimal);
//			string ShipCode = null;
//			string CardCode = null;
//			string SjDuDate = null;
//			string SjDcDate = null;
//			string WorkGbn = null;
//			string CardName = null;
//			string ShipName = null;
//			string JakSize = null;
//			string ReDate = null;
//			string JakDate = null;
//			string Temp01 = null;
//			string SubNo1 = null;
//			string Comments = null;
//			string JakName = null;
//			string SubNo2 = null;
//			string InOutGbn = null;
//			string ProDate = null;
//			string JakMyung = null;
//			string JakUnit = null;
//			decimal WrWeight = default(decimal);
//			string Status = null;

//			string ReWkRn = null;
//			//재작업 사유
//			string ReWork = null;
//			//재작업 구분

//			decimal OrderAmt = default(decimal);
//			//수주금액(2015.10.12 송명규)
//			decimal NegoAmt = default(decimal);
//			//Nego금액(2015.10.12 송명규)
//			decimal TrgtAmt = default(decimal);
//			//목표금액(생산)(2015.10.12 송명규)

//			string ErrOrdNum = null;
//			//선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
//			string BaseEntry = null;
//			string BaseLine = null;
//			string DocType = null;
//			string CurDocDate = null;

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, false);

//			oMat01.FlushToDataSource();

//			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			BPLId = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);

//			for (i = 0; i <= oMat02.RowCount - 1; i++) {
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				JakName = Strings.Trim(oMat02.Columns.Item("JakName").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SubNo1 = Strings.Trim(oMat02.Columns.Item("SubNo1").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SubNo2 = Strings.Trim(oMat02.Columns.Item("SubNo2").Cells.Item(i + 1).Specific.VALUE);

//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ReWkRn = Strings.Trim(oMat02.Columns.Item("ReWkRn").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ReWork = oForm01.Items.Item("ReWork").Specific.VALUE;

//				sQry = "Select COUNT(*) From [@PS_PP020H] Where U_JakName = '" + JakName + "' And U_SubNo1 = '" + SubNo1 + "' And U_SubNo2 = '" + SubNo2 + "'";
//				oRecordSet01.DoQuery(sQry);

//				if (oRecordSet01.Fields.Item(0).Value > 0) {
//					ErrNum = 1;
//					goto Add_SubJakNum_Error;
//				} else {

//					if (ReWork == "20" & string.IsNullOrEmpty(ReWkRn)) {
//						ErrNum = 2;
//						goto Add_SubJakNum_Error;
//					}

//				}
//			}

//			for (i = 0; i <= oMat02.RowCount - 1; i++) {

//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				WorkGbn = Strings.Trim(oMat02.Columns.Item("WorkGbn").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				JakName = Strings.Trim(oMat02.Columns.Item("JakName").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SubNo1 = Strings.Trim(oMat02.Columns.Item("SubNo1").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SubNo2 = Strings.Trim(oMat02.Columns.Item("SubNo2").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				JakMyung = Strings.Trim(oMat02.Columns.Item("JakMyung").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				JakSize = Strings.Trim(oMat02.Columns.Item("JakSize").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				JakUnit = Strings.Trim(oMat02.Columns.Item("JakUnit").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				RegNum = Strings.Trim(oMat02.Columns.Item("RegNum").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ItemCode = Strings.Trim(oMat02.Columns.Item("ItemCode").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ItemName = MDC_PS_Common.Make_ItemName(Strings.Trim(oMat02.Columns.Item("ItemName").Cells.Item(i + 1).Specific.VALUE));
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Material = Strings.Trim(oMat02.Columns.Item("Material").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Unit = Strings.Trim(oMat02.Columns.Item("Unit").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Size = Strings.Trim(oMat02.Columns.Item("Size").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ItmBsort = Strings.Trim(oMat02.Columns.Item("ItmBSort").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SjDocNum = Strings.Trim(oMat02.Columns.Item("SjDocNum").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SjLinNum = Strings.Trim(oMat02.Columns.Item("SjLinNum").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SjWeight = Convert.ToDecimal(Strings.Trim(oMat02.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.VALUE));
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				WrWeight = Convert.ToDecimal(Strings.Trim(oMat02.Columns.Item("WrWeight").Cells.Item(i + 1).Specific.VALUE));
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SjDcDate = Strings.Trim(oMat02.Columns.Item("SjDcDate").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SjDuDate = Strings.Trim(oMat02.Columns.Item("SjDuDate").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				SlePrice = Convert.ToDecimal(Strings.Trim(oMat02.Columns.Item("SlePrice").Cells.Item(i + 1).Specific.VALUE));
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CardCode = Strings.Trim(oMat02.Columns.Item("CardCode").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CardName = Strings.Trim(oMat02.Columns.Item("CardName").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ShipCode = Strings.Trim(oMat02.Columns.Item("ShipCode").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ShipName = Strings.Trim(oMat02.Columns.Item("ShipName").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Comments = Strings.Trim(oMat02.Columns.Item("Comments").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				InOutGbn = Strings.Trim(oMat02.Columns.Item("InOutGbn").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				JakDate = Strings.Trim(oMat02.Columns.Item("JakDate").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ProDate = Strings.Trim(oMat02.Columns.Item("ProDate").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ReDate = Strings.Trim(oMat02.Columns.Item("ReDate").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				ReWkRn = Strings.Trim(oMat02.Columns.Item("ReWkRn").Cells.Item(i + 1).Specific.VALUE);
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				OrderAmt = Convert.ToDecimal(Strings.Trim(oMat02.Columns.Item("OrderAmt").Cells.Item(i + 1).Specific.VALUE));
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				NegoAmt = Convert.ToDecimal(Strings.Trim(oMat02.Columns.Item("NegoAmt").Cells.Item(i + 1).Specific.VALUE));
//				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				TrgtAmt = Convert.ToDecimal(Strings.Trim(oMat02.Columns.Item("TrgtAmt").Cells.Item(i + 1).Specific.VALUE));
//				Status = "O";

//				sQry = "      EXEC PS_PP020_04 '";
//				sQry = sQry + DocEntry + "','";
//				sQry = sQry + BPLId + "','";
//				sQry = sQry + JakName + "','";
//				sQry = sQry + SubNo1 + "','";
//				sQry = sQry + SubNo2 + "','";
//				sQry = sQry + JakMyung + "','";
//				sQry = sQry + JakSize + "','";
//				sQry = sQry + JakUnit + "','";
//				sQry = sQry + RegNum + "','";
//				sQry = sQry + ItemCode + "','";
//				sQry = sQry + ItemName + "','";
//				sQry = sQry + Material + "','";
//				sQry = sQry + Unit + "','";
//				sQry = sQry + Size + "','";
//				sQry = sQry + ItmBsort + "','";
//				sQry = sQry + SjDocNum + "','";
//				sQry = sQry + SjLinNum + "','";
//				sQry = sQry + SjWeight + "','";
//				sQry = sQry + WrWeight + "','";
//				sQry = sQry + SjDcDate + "','";
//				sQry = sQry + SjDuDate + "','";
//				sQry = sQry + SlePrice + "','";
//				sQry = sQry + WorkGbn + "','";
//				sQry = sQry + CardCode + "','";
//				sQry = sQry + CardName + "','";
//				sQry = sQry + ShipCode + "','";
//				sQry = sQry + ShipName + "','";
//				sQry = sQry + InOutGbn + "','";
//				sQry = sQry + JakDate + "','";
//				sQry = sQry + ProDate + "','";
//				sQry = sQry + ReDate + "','";
//				sQry = sQry + Comments + "','";
//				sQry = sQry + OrderAmt + "','";
//				sQry = sQry + NegoAmt + "','";
//				sQry = sQry + TrgtAmt + "','";
//				sQry = sQry + Status + "','";
//				sQry = sQry + ReWkRn + "'";

//				oRecordSet01.DoQuery(sQry);
//				//등록

//			}

//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			MDC_Com.MDC_GF_Message(ref "Sub작번등록 완료!", ref "S");

//			oDS_PS_TEMPTABLE.Clear();
//			oMat02.Clear();

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			Add_SubJakNum_Error:

//			functionReturnValue = false;

//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번째 라인의 작번이 이미 등록되어 있습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 2) {
//				MDC_Com.MDC_GF_Message(ref "재작업인 경우는 재작업 사유를 필수로 입력해야합니다.  [" + i + 1 + "행]", ref "E");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "Add_SubJakNum_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//			return functionReturnValue;
//		}
//	}
//}
