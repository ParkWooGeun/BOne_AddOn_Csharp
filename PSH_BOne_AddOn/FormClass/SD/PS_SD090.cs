using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 이동처리등록
	/// </summary>
	internal class PS_SD090 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01; 
		private SAPbouiCOM.DBDataSource oDS_PS_SD090H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD090L; //등록라인
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oSeq;

		// 입고 DI를 위한 정보를 가지는 구조체
		public class StockInfos
		{
			public string CardCode; //고객코드
			public string ItemCode; //품목코드
			public string FromWarehouseCode; //창고코드
			public string ToWarehouseCode; //창고코드
			public double Weight; //중량
			public double UnWeight;
			public string BatchNum; //배치번호
			public double BatchWeight;//배치중량
			public int Qty; //수량
			public string TransNo; //재고이전문서번호
			public string Chk;
			public int MatrixRow;
			public string StockTransDocEntry; //재고이전문서번호
			public string StockTransLineNum; //재고이전라인번호
			public string Indate; //전기일
		}

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD090.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD090_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD090");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);

				//CreateItems();
				//ComboBox_Setting();
				//CF_ChooseFromList();
				//FormItemEnabled();
				//FormClear();
				//AddMatrixRow(0, oMat01.RowCount, true);

				oForm.EnableMenu("1293", true); //행삭제
				oForm.EnableMenu("1283", false); //제거
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1285", false); //복원
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}
		}

		
		//public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	string sQry = null;
		//	string sQry02 = null;
		//	SAPbobsCOM.Recordset oRecordSet01 = null;
		//	SAPbobsCOM.Recordset oRecordset02 = null;
		//	object TempForm01 = null;
		//	short ErrNum = 0;

		//	int SumQty = 0;
		//	decimal SumWeight = default(decimal);

		//	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	//// 객체 정의 및 데이터 할당
		//	oRecordset02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


		//	short i = 0;
		//	short j = 0;
		//	short DocEntry = 0;
		//	short LineId = 0;
		//	////BeforeAction = True
		//	if ((pVal.BeforeAction == true)) {
		//		switch (pVal.EventType) {

		//			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
		//				////1
		//				if (pVal.ItemUID == "1") {
		//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
		//						if (HeaderSpaceLineDel() == false) {
		//							BubbleEvent = false;
		//							//BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
		//							return;
		//						}

		//						if (MatrixSpaceLineDel() == false) {
		//							BubbleEvent = false;
		//							return;
		//						}

		//						//// 재고 이동 DI API
		//						if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
		//							if (StockTrans() == true) {
		//								UpdateUserField();
		//							} else {
		//								AddMatrixRow(1, oMat01.RowCount, ref true);
		//								BubbleEvent = false;
		//								return;
		//							}
		//						}
		//					}

		//				} else if (pVal.ItemUID == "ChulPrin") {
		//					PS_SD090_Print_Report01();
		//					//            ElseIf pVal.ItemUID = "GuraPrin" Then

		//				} else {
		//					if (pVal.ItemChanged == true) {
		//						if (pVal.ItemUID == "Mat01" & pVal.ColUID == "ItemCode") {
		//							FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
		//						}
		//					}
		//				}
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
		//				////2

		//				// 거래처코드
		//				//UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value)) {
		//					if (pVal.ItemUID == "CardCode" & pVal.CharPressed == 9) {
		//						////CharPressed: The character that was pressed to trigger this event.
		//						oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
		//						BubbleEvent = false;
		//					}
		//				}

		//				// 아이템코드
		//				if (pVal.ItemUID == "Mat01" & pVal.ColUID == "SD091HNo" & pVal.CharPressed == 9) {
		//					//UPGRADE_WARNING: oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String)) {
		//						oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
		//						BubbleEvent = false;
		//					}
		//				}

		//				// 담당자
		//				//UPGRADE_WARNING: oForm.Items(RepName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				if (string.IsNullOrEmpty(oForm.Items.Item("RepName").Specific.Value)) {
		//					if (pVal.ItemUID == "RepName" & pVal.CharPressed == 9) {
		//						oForm.Items.Item("RepName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
		//						BubbleEvent = false;
		//					}
		//				}

		//				// 납품처
		//				//UPGRADE_WARNING: oForm.Items(ShipTo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				if (string.IsNullOrEmpty(oForm.Items.Item("ShipTo").Specific.Value)) {
		//					if (pVal.ItemUID == "ShipTo" & pVal.CharPressed == 9) {
		//						oForm.Items.Item("ShipTo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
		//						BubbleEvent = false;
		//					}
		//				}

		//				// 운송업체

		//				// 차량번호

		//				// 도착장소

		//				// 질별

		//				// 출고창고
		//				if (pVal.ItemUID == "OutWhCd" & pVal.CharPressed == 9) {
		//					//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Specific.String)) {
		//						//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
		//						BubbleEvent = false;
		//					}
		//				}

		//				// 입고창고
		//				if (pVal.ItemUID == "InWhCd" & pVal.CharPressed == 9) {
		//					//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Specific.String)) {
		//						//UPGRADE_WARNING: oForm.Items().Cells 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item(pVal.ColUID).Cells(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
		//						BubbleEvent = false;
		//					}
		//				}
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
		//				////5
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_CLICK:
		//				////6
		//				break;


		//			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
		//				////7
		//				if (pVal.ItemChanged == true) {
		//					if (pVal.ItemUID == "Mat01" & pVal.ColUID == "ItemCode") {
		//						FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
		//					}
		//				}
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
		//				////8
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
		//				////10
		//				if (pVal.ItemChanged == true) {

		//					// 거래처 이름 Query
		//					if (pVal.ItemUID == "CardCode") {
		//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						sQry = "Select CardName From [OCRD] Where CardCode = '" + Strings.Trim(oForm.Items.Item("CardCode").Specific.Value) + "'";
		//						oRecordSet01.DoQuery(sQry);
		//						//UPGRADE_WARNING: oForm.Items(CardName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item("CardName").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
		//					}

		//					// 사원 이름 Query
		//					if (pVal.ItemUID == "RepName") {
		//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + Strings.Trim(oForm.Items.Item("RepName").Specific.Value) + "'";
		//						oRecordSet01.DoQuery(sQry);
		//						//UPGRADE_WARNING: oForm.Items(RepNm1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item("RepNm1").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
		//					}

		//					// 납품처 이름 Query
		//					if (pVal.ItemUID == "ShipTo") {
		//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						sQry = "Select CardName From [OCRD] Where CardCode = '" + Strings.Trim(oForm.Items.Item("ShipTo").Specific.Value) + "'";
		//						oRecordSet01.DoQuery(sQry);
		//						//UPGRADE_WARNING: oForm.Items(ShipNm).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item("ShipNm").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
		//					}

		//					// 아이템 코드
		//					if (pVal.ItemUID == "Mat01" & pVal.ColUID == "ItemCode") {
		//						FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
		//					}

		//					// 출고 창고
		//					if (pVal.ItemUID == "OutWhCd") {
		//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						sQry = "Select WhsName From [OWHS] Where WhsCode = '" + Strings.Trim(oForm.Items.Item("OutWhCd").Specific.Value) + "'";
		//						oRecordSet01.DoQuery(sQry);
		//						//UPGRADE_WARNING: oForm.Items(OutWhName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item("OutWhName").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
		//					}

		//					// 입고 창고
		//					if (pVal.ItemUID == "InWhCd") {
		//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						sQry = "Select WhsName From [OWHS] Where WhsCode = '" + Strings.Trim(oForm.Items.Item("InWhCd").Specific.Value) + "'";
		//						oRecordSet01.DoQuery(sQry);
		//						//UPGRADE_WARNING: oForm.Items(InWhCd).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item("InWhCd").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
		//					}

		//					// 작업요청
		//					if (pVal.ItemUID == "Mat01" & pVal.ColUID == "SD091HNo") {
		//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						j = Strings.InStr(Strings.Trim(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value), "-");
		//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						DocEntry = Convert.ToInt16(Strings.Left(Strings.Trim(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value), j - 1));
		//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						LineId = Convert.ToInt16(Strings.Mid(Strings.Trim(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value), j + 1));
		//						sQry = "Select U_ItemCode, U_ItemName, U_ItemGu, U_Qty, ";
		//						sQry = sQry + "U_Unweight, U_Weight, U_Comments, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, ";
		//						sQry = sQry + "U_Quality, U_Mark, U_CallSize, U_SbasUnit ";
		//						sQry = sQry + "From [@PS_SD091L] Where DocEntry = '" + DocEntry + "' And LineId = '" + LineId + "'";
		//						oRecordSet01.DoQuery(sQry);
		//						oMat01.FlushToDataSource();
		//						oDS_PS_SD090L.SetValue("U_ItemCode", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(0).Value));
		//						oDS_PS_SD090L.SetValue("U_ItemName", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(1).Value));
		//						oDS_PS_SD090L.SetValue("U_ItemGu", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(2).Value));
		//						oDS_PS_SD090L.SetValue("U_Qty", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(3).Value));
		//						oDS_PS_SD090L.SetValue("U_Unweight", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(4).Value));
		//						oDS_PS_SD090L.SetValue("U_Weight", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(5).Value));
		//						oDS_PS_SD090L.SetValue("U_Comments", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(6).Value));
		//						oDS_PS_SD090L.SetValue("U_ItmBsort", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(7).Value));
		//						oDS_PS_SD090L.SetValue("U_ItmMsort", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(8).Value));
		//						oDS_PS_SD090L.SetValue("U_Unit1", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(9).Value));
		//						oDS_PS_SD090L.SetValue("U_Size", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(10).Value));
		//						oDS_PS_SD090L.SetValue("U_ItemType", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(11).Value));
		//						oDS_PS_SD090L.SetValue("U_Quality", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(12).Value));
		//						oDS_PS_SD090L.SetValue("U_Mark", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(13).Value));
		//						oDS_PS_SD090L.SetValue("U_CallSize", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(14).Value));
		//						oDS_PS_SD090L.SetValue("U_SbasUnit", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(15).Value));
		//						//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//						oRecordSet01 = null;

		//						AddMatrixRow(1, oMat01.VisualRowCount, ref false);
		//						for (i = 1; i <= oMat01.VisualRowCount; i++) {
		//							if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", i - 1))) {
		//								oDS_PS_SD090L.RemoveRecord(i - 1);
		//								oMat01.LoadFromDataSource();
		//							}
		//						}

		//						oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//						sQry = "Select U_OutWhCd, U_InWhCd From [@PS_SD091H] Where DocEntry = '" + DocEntry + "'";
		//						oRecordSet01.DoQuery(sQry);
		//						oDS_PS_SD090H.SetValue("U_OutWhCd", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(0).Value));
		//						oDS_PS_SD090H.SetValue("U_InWhCd", pVal.Row - 1, Strings.Trim(oRecordSet01.Fields.Item(1).Value));
		//						//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//						oRecordSet01 = null;


		//						for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
		//							//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//							if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
		//								SumQty = SumQty;
		//							} else {
		//								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//								SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
		//							}
		//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//							SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

		//						}
		//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item("SumQty").Specific.Value = SumQty;
		//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oForm.Items.Item("SumWeight").Specific.Value = SumWeight;


		//						AddMatrixRow(1, oMat01.VisualRowCount, ref false);
		//					}

		//					//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//					oRecordSet01 = null;
		//				}
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
		//				////11
		//				break;
		//			//                AddMatrixRow 1, oMat01.VisualRowCount, False
		//			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
		//				////18
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
		//				////19
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
		//				////20
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
		//				////27
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
		//				////3
		//				oLast_Item_UID = pVal.ItemUID;
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
		//				////4
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
		//				////17
		//				break;
		//		}
		//	////BeforeAction = False
		//	} else if ((pVal.BeforeAction == false)) {
		//		switch (pVal.EventType) {
		//			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
		//				////1
		//				//저장 후 추가 가능처리
		//				if (pVal.ItemUID == "1") {
		//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true) {
		//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
		//						SubMain.Sbo_Application.ActivateMenuItem("1282");
		//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == false) {
		//						FormItemEnabled();
		//						AddMatrixRow(1, oMat01.RowCount, ref true);
		//					}
		//				}
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
		//				////2
		//				if (pVal.Action_Success == true) {
		//					oSeq = 1;
		//				}
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
		//				////5
		//				if (pVal.ItemChanged == true) {
		//					oForm.Freeze(true);
		//					if ((pVal.ItemUID == "BPLId")) {
		//						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						if (oForm.Items.Item("BPLId").Specific.Value == "1") {
		//							oDS_PS_SD090H.SetValue("U_OutWhCd", 0, "101");
		//							oDS_PS_SD090H.SetValue("U_InWhCd", 0, "104");
		//							//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						} else if (oForm.Items.Item("BPLId").Specific.Value == "4") {
		//							oDS_PS_SD090H.SetValue("U_OutWhCd", 0, "104");
		//							oDS_PS_SD090H.SetValue("U_InWhCd", 0, "101");
		//						}
		//					}
		//					oForm.Update();
		//					oForm.Freeze(false);
		//				}
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_CLICK:
		//				////6
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
		//				////7
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
		//				////8
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
		//				////10
		//				break;
		//			// 이동요청 문서 Query


		//			//            ' 품목 이름 Query
		//			//                If (pVal.ItemUID = "Mat01" And (pVal.ColUID = "ItemCode")) Then
		//			//                    sQry = "Select ItemName, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, U_Quality, U_Mark, U_CallSize, U_SbasUnit From [OITM] Where "
		//			//                    sQry = sQry & "ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(pVal.Row).Specific.Value) & "'"
		//			//                    oRecordSet01.DoQuery sQry
		//			//                    oMat01.Columns("ItemName").Cells(pVal.Row).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
		//			//            ' 품목 대분류
		//			//                    sQry02 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet01.Fields(1).Value) & "'"
		//			//                    oRecordSet02.DoQuery sQry02
		//			//                    Call oMat01.Columns("ItmBsort").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
		//			//            ' 품목 중분류
		//			//                    sQry02 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet01.Fields(2).Value) & "'"
		//			//                    oRecordSet02.DoQuery sQry02
		//			//                    Call oMat01.Columns("ItmBsort").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
		//			//            ' 형태타입
		//			//                    sQry02 = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" & Trim(oRecordSet01.Fields(5).Value) & "'"
		//			//                    oRecordSet02.DoQuery sQry02
		//			//                    Call oMat01.Columns("ItemType").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
		//			//            ' 질별
		//			//                    sQry02 = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" & Trim(oRecordSet01.Fields(6).Value) & "'"
		//			//                    oRecordSet02.DoQuery sQry02
		//			//                    Call oMat01.Columns("Quality").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
		//			//            ' 인증기호
		//			//                    sQry02 = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" & Trim(oRecordSet01.Fields(7).Value) & "'"
		//			//                    oRecordSet02.DoQuery sQry02
		//			//                    Call oMat01.Columns("Mark").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
		//			//            ' 판매기준단위
		//			//                    sQry02 = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" & Trim(oRecordSet01.Fields(9).Value) & "'"
		//			//                    oRecordSet02.DoQuery sQry02
		//			//                    Call oMat01.Columns("SbasUnit").Cells(pVal.Row).Specific.Select(oRecordSet02.Fields(0).Value, psk_ByValue)
		//			//                    oMat01.Columns("Unit1").Cells(pVal.Row).Specific.Value = Trim(oRecordSet01.Fields(3).Value)
		//			//                    oMat01.Columns("Size").Cells(pVal.Row).Specific.Value = Trim(oRecordSet01.Fields(4).Value)
		//			//                End If
		//			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
		//				////11

		//				for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
		//					//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
		//						SumQty = SumQty;
		//					} else {
		//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
		//					}
		//					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

		//				}

		//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				oForm.Items.Item("SumQty").Specific.Value = SumQty;
		//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

		//				AddMatrixRow(1, oMat01.VisualRowCount, ref true);
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
		//				////18
		//				if (oSeq == 1) {
		//					oSeq = 0;
		//				}
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
		//				////19
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
		//				////20
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
		//				////27
		//				break;

		//			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
		//				////3
		//				oLast_Item_UID = pVal.ItemUID;
		//				break;
		//			//                If oLast_Item_UID = "매트릭스" Then
		//			//                    If pVal.Row > 0 Then
		//			//                        oLast_Item_UID = pVal.ItemUID
		//			//                        oLast_Col_UID = pVal.ColUID
		//			//                        oLast_Col_Row = pVal.Row
		//			//                    End If
		//			//                Else
		//			//                    oLast_Item_UID = pVal.ItemUID
		//			//                    oLast_Col_UID = ""
		//			//                    oLast_Col_Row = 0
		//			//                End If
		//			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
		//				////4
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
		//				////17
		//				SubMain.RemoveForms(oFormUniqueID01);
		//				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//				oForm = null;
		//				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//				oMat01 = null;
		//				break;
		//		}
		//	}

		//	return;
		//	Raise_ItemEvent_Error:
		//	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}



		//public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;
		//	int SumQty = 0;
		//	decimal SumWeight = default(decimal);

		//	////BeforeAction = True
		//	if ((pVal.BeforeAction == true)) {
		//		switch (pVal.MenuUID) {
		//			case "1284":
		//				//취소
		//				break;
		//			case "1286":
		//				//닫기
		//				break;
		//			case "1293":
		//				//행닫기
		//				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
		//				break;
		//			case "1281":
		//				//찾기
		//				break;
		//			case "1282":
		//				//추가
		//				break;
		//			case "1288":
		//			case "1289":
		//			case "1290":
		//			case "1291":
		//				//레코드이동버튼
		//				break;

		//		}
		//	////BeforeAction = False
		//	} else if ((pVal.BeforeAction == false)) {
		//		switch (pVal.MenuUID) {
		//			case "1284":
		//				//취소
		//				break;
		//			case "1286":
		//				//닫기
		//				break;
		//			case "1281":
		//				//찾기
		//				FormItemEnabled();
		//				break;
		//			//                oForm.Items("ItemCode").Click ct_Regular
		//			case "1282":
		//				//추가
		//				FormItemEnabled();
		//				FormClear();
		//				AddMatrixRow(0, 0, ref true);
		//				oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
		//				break;

		//			case "1288":
		//			case "1289":
		//			case "1290":
		//			case "1291":
		//				//레코드이동버튼
		//				FormItemEnabled();
		//				if (oMat01.VisualRowCount > 0) {
		//					//UPGRADE_WARNING: oMat01.Columns(SD091HNo).Cells(oMat01.VisualRowCount).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					if (!string.IsNullOrEmpty(oMat01.Columns.Item("SD091HNo").Cells.Item(oMat01.VisualRowCount).Specific.Value)) {
		//						AddMatrixRow(1, oMat01.RowCount, ref true);
		//					}
		//				}
		//				break;
		//			case "1293":
		//				//행닫기
		//				if (oMat01.RowCount != oMat01.VisualRowCount) {
		//					for (i = 1; i <= oMat01.VisualRowCount; i++) {
		//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
		//					}
		//					oMat01.FlushToDataSource();
		//					// DBDataSource에 레코드가 한줄 더 생긴다.
		//					oDS_PS_SD090L.RemoveRecord(oDS_PS_SD090L.Size - 1);
		//					// 레코드 한 줄을 지운다.
		//					oMat01.LoadFromDataSource();
		//					// DBDataSource를 매트릭스에 올리고
		//					if (oMat01.RowCount == 0) {
		//						AddMatrixRow(1, 0, ref true);
		//					} else {
		//						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD090L.GetValue("U_ItemCode", oMat01.RowCount - 1)))) {
		//							AddMatrixRow(1, oMat01.RowCount, ref true);

		//						}
		//					}


		//					for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
		//						//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
		//							SumQty = SumQty;
		//						} else {
		//							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//							SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
		//						}
		//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;

		//					}
		//					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					oForm.Items.Item("SumQty").Specific.Value = SumQty;
		//					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
		//				}
		//				break;
		//		}
		//	}
		//	return;
		//	Raise_MenuEvent_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////BeforeAction = True
		//	if ((BusinessObjectInfo.BeforeAction == true)) {
		//		switch (BusinessObjectInfo.EventType) {
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
		//				////33
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
		//				////34 - 추가
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
		//				////35 - 업데이트
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
		//				////36
		//				break;
		//		}
		//	////BeforeAction = False
		//	} else if ((BusinessObjectInfo.BeforeAction == false)) {
		//		switch (BusinessObjectInfo.EventType) {
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
		//				////33
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
		//				////34
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
		//				////35
		//				break;
		//			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
		//				////36
		//				break;
		//		}
		//	}
		//	return;
		//	Raise_FormDataEvent_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	if ((eventInfo.BeforeAction == true)) {
		//		////작업
		//	} else if ((eventInfo.BeforeAction == false)) {
		//		////작업
		//	}
		//	return;
		//	Raise_RightClickEvent_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public bool CreateItems()
		//{
		//	bool functionReturnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	string sQry = null;
		//	SAPbobsCOM.Recordset oRecordSet01 = null;

		//	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oDS_PS_SD090H = oForm.DataSources.DBDataSources("@PS_SD090H");
		//	oDS_PS_SD090L = oForm.DataSources.DBDataSources("@PS_SD090L");
		//	oMat01 = oForm.Items.Item("Mat01").Specific;
		//	//// 매트릭스 데이터 셋

		//	oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
		//	oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);

		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");
		//	//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");

		//	oDS_PS_SD090H.SetValue("U_DocDate", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd"));

		//	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("BPLId").Specific), ref "SELECT BPLId, BPLName FROM OBPL  Where BPLId = '1' Or BPLId = '4' order by BPLId", ref "1", ref false, ref false);

		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//	return functionReturnValue;
		//	CreateItems_Error:
		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//	SubMain.Sbo_Application.SetStatusBarMessage("CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	return functionReturnValue;
		//}

		//public void ComboBox_Setting()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////콤보에 기본값설정
		//	// 반출등록
		//	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("OutWhCd").Specific), ref "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", ref "104", ref false, ref true);
		//	MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("InWhCd").Specific), ref "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", ref "101", ref false, ref true);

		//	// 품목대분류
		//	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code");
		//	// 품목중분류
		//	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code");
		//	// 형태타입
		//	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code");
		//	// 질별
		//	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] ORDER BY Code");
		//	// 인증기호
		//	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code");
		//	// 매입기준단위
		//	MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("SbasUnit"), "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code");
		//	return;
		//	ComboBox_Setting_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void CF_ChooseFromList()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////ChooseFromList 설정
		//	return;
		//	CF_ChooseFromList_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void Initial_Setting()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	// 사업장
		//	//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);
		//	// 인수자
		//	//UPGRADE_WARNING: oForm.Items(RepName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	oForm.Items.Item("RepName").Specific.Value = MDC_PS_Common.User_MSTCOD();
		//	return;
		//	Initial_Setting_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Initial_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void FormItemEnabled()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement


		//	SAPbouiCOM.ComboBox oCombo = null;
		//	//콤보박스
		//	SAPbouiCOM.ComboBox oCombo1 = null;
		//	//콤보박스
		//	SAPbouiCOM.Column oColumn = null;
		//	string lQuery = null;
		//	SAPbobsCOM.Recordset lRecordSet = null;

		//	lRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
		//		////각모드에따른 아이템설정
		//		oForm.Items.Item("DocNum").Enabled = false;
		//		oForm.Items.Item("CardCode").Enabled = true;
		//		oForm.Items.Item("BPLId").Enabled = true;
		//		oForm.Items.Item("RepName").Enabled = true;
		//		oForm.Items.Item("DocDate").Enabled = true;
		//		oForm.Items.Item("ShipTo").Enabled = true;
		//		oForm.Items.Item("CarCo").Enabled = true;
		//		oForm.Items.Item("CarNo").Enabled = true;
		//		oForm.Items.Item("ArrSite").Enabled = true;
		//		oForm.Items.Item("Fare").Enabled = true;
		//		oForm.Items.Item("Specific").Enabled = true;
		//		oForm.Items.Item("ChulPrin").Enabled = false;
		//		//        oForm.Items("OutWhCd").Enabled = True
		//		//        oForm.Items("InWhCd").Enabled = True
		//		oMat01.Columns.Item("ItemCode").Editable = true;
		//		oMat01.Columns.Item("ItemGu").Editable = true;
		//		//////////////////////////////////////////////////폼상태변경//////////////////////////////////
		//		//UPGRADE_WARNING: oForm.Items(DocDate).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oForm.Items.Item("DocDate").Specific.String = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd");
		//		oForm.Items.Item("RepName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

		//	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
		//		////각모드에따른 아이템설정
		//		oForm.Items.Item("DocNum").Enabled = false;
		//		oForm.Items.Item("CardCode").Enabled = true;
		//		oForm.Items.Item("BPLId").Enabled = true;
		//		oForm.Items.Item("RepName").Enabled = true;
		//		oForm.Items.Item("DocDate").Enabled = true;
		//		oForm.Items.Item("ShipTo").Enabled = true;
		//		oForm.Items.Item("CarCo").Enabled = true;
		//		oForm.Items.Item("CarNo").Enabled = true;
		//		oForm.Items.Item("ArrSite").Enabled = true;
		//		oForm.Items.Item("Fare").Enabled = false;
		//		oForm.Items.Item("Specific").Enabled = false;
		//		oForm.Items.Item("ChulPrin").Enabled = true;
		//		//False

		//		//        oForm.Items("OutWhCd").Enabled = False
		//		//        oForm.Items("InWhCd").Enabled = False
		//		oMat01.Columns.Item("ItemCode").Editable = false;
		//		oMat01.Columns.Item("ItemGu").Editable = false;

		//	} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
		//		//////////////////////////////////////////////////폼상태변경//////////////////////////////////
		//		//        lQuery = "Select Status From [@PS_SD090H] Where DocNum = '"
		//		//        lQuery = lQuery + oForm.Items("DocNum").Specific.Value
		//		//        lQuery = lQuery + "'"
		//		//
		//		//        lRecordSet.DoQuery lQuery
		//		//        If (lRecordSet.Fields(0).Value = "O") Then
		//		//            oForm.Items("DocNum").Enabled = False
		//		//            oForm.Items("BPLId").Enabled = False
		//		//            oForm.Items("RepName").Enabled = False
		//		//            oMat01.Columns("ItemCode").Editable = False
		//		//            oMat01.Columns("OutWhCd").Editable = False
		//		//            oMat01.Columns("InWhCd").Editable = False
		//		//            oMat01.Columns("BatchNum").Editable = False
		//		//        ElseIf (lRecordSet.Fields(0).Value = "C") Then
		//		//            oForm.Items("DocNum").Enabled = False
		//		//            oForm.Items("BPLId").Enabled = False
		//		//            oForm.Items("RepName").Enabled = False
		//		//            oMat01.Columns("ItemCode").Editable = False
		//		//            oMat01.Columns("OutWhCd").Editable = False
		//		//            oMat01.Columns("InWhCd").Editable = False
		//		//            oMat01.Columns("BatchNum").Editable = False
		//		//        End If
		//		////Status 설정
		//		oForm.Items.Item("DocNum").Enabled = false;
		//		oForm.Items.Item("CardCode").Enabled = false;
		//		oForm.Items.Item("BPLId").Enabled = false;
		//		oForm.Items.Item("RepName").Enabled = false;
		//		oForm.Items.Item("DocDate").Enabled = false;
		//		oForm.Items.Item("ShipTo").Enabled = false;
		//		oForm.Items.Item("CarCo").Enabled = false;
		//		oForm.Items.Item("CarNo").Enabled = false;
		//		oForm.Items.Item("ArrSite").Enabled = false;
		//		oForm.Items.Item("Fare").Enabled = false;
		//		oForm.Items.Item("Specific").Enabled = false;
		//		oForm.Items.Item("ChulPrin").Enabled = true;
		//		//        oForm.Items("OutWhCd").Enabled = False
		//		//        oForm.Items("InWhCd").Enabled = False
		//		oMat01.Columns.Item("ItemCode").Editable = false;
		//		oMat01.Columns.Item("ItemGu").Editable = false;

		//		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		lQuery = "SELECT Status,Canceled FROM [@PS_SD090H] WHERE DocEntry = '" + oForm.Items.Item("DocNum").Specific.Value + "'";
		//		lRecordSet.DoQuery(lQuery);
		//		if ((lRecordSet.Fields.Item(0).Value == "O")) {
		//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			oForm.Items.Item("Status").Specific.Value = "미결";
		//		} else if ((lRecordSet.Fields.Item(0).Value == "C")) {
		//			if ((lRecordSet.Fields.Item(1).Value == "Y")) {
		//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				oForm.Items.Item("Status").Specific.Value = "취소";
		//			} else {
		//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				oForm.Items.Item("Status").Specific.Value = "종료";
		//			}
		//		}
		//		//////////////////////////////////////////////////폼상태변경//////////////////////////////////
		//	}
		//	oMat01.AutoResizeColumns();

		//	return;
		//	FormItemEnabled_Error:

		//	SubMain.Sbo_Application.SetStatusBarMessage("FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public void AddMatrixRow(short oSeq, int oRow, ref bool RowIserted = false)
		//{
		//	//On Error GoTo AddMatrixRow_Error

		//	switch (oSeq) {
		//		case 0:
		//			oMat01.AddRow();
		//			// 매트릭스에 새로운 로를 추가한다.
		//			oDS_PS_SD090L.SetValue("U_LIneNum", oRow, Convert.ToString(oRow + 1));
		//			oMat01.LoadFromDataSource();
		//			break;
		//		case 1:
		//			oDS_PS_SD090L.InsertRecord(oRow);
		//			oDS_PS_SD090L.SetValue("U_LIneNum", oRow, Convert.ToString(oRow + 1));
		//			oMat01.LoadFromDataSource();
		//			break;
		//	}
		//	//AddMatrixRow_Error:
		//	//    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
		//}

		//public void FormClear()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	string DocNum = null;
		//	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	DocNum = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_SD090'", ref "");
		//	if (Convert.ToDouble(DocNum) == 0) {
		//		//        oForm.Items("DocEntry").Specific.String = 1
		//		oDS_PS_SD090H.SetValue("DocNum", 0, "1");
		//	} else {
		//		//        oForm.Items("DocEntry").Specific.String = DocNum
		//		oDS_PS_SD090H.SetValue("DocNum", 0, DocNum);
		//		// 화면에 적용이 안되기 때문
		//	}
		//	return;
		//	FormClear_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//public bool DataValidCheck()
		//{
		//	bool functionReturnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	return functionReturnValue;
		//	DataValidCheck_Error:
		//	////유효성검사
		//	SubMain.Sbo_Application.SetStatusBarMessage("DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	return functionReturnValue;
		//}

		//private void FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
		//{
		//	////데이터변화에 따른처리
		//	int i = 0;
		//	SAPbobsCOM.Recordset oRecordSet01 = null;
		//	string sQry = null;

		//	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	switch (oUID) {

		//		case "Mat01":
		//			//rowcount: 로 카운트를 반환, VisualRowCount: 삭제된 로를 제외하고 현재 보이는 로 카운트를 반환
		//			//            If oCol = "ItemCode" Then
		//			//                sQry = "SELECT U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, U_Quality, U_Mark, U_CallSize, U_SbasUnit"
		//			//                sQry = sQry & "FROM [OITM] WHERE ItemCode ='" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific) & "'"
		//			//                oRecordSet01.DoQuery (sQry)
		//			//
		//			//                oRecordSet01.MoveFirst
		//			//
		//			//                For i = 0 To oRecordSet01.EOF
		//			//                oMat01.Columns("ItmBsort").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
		//			//                oMat01.Columns("ItmMsort").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
		//			//                oMat01.Columns("Unit1").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(2).Value)
		//			//                oMat01.Columns("Size").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(3).Value)
		//			//                oMat01.Columns("ItemType").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(4).Value)
		//			//                oMat01.Columns("Quality").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(5).Value)
		//			//                oMat01.Columns("Mark").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(6).Value)
		//			//                oMat01.Columns("CallSize").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(7).Value)
		//			//                oMat01.Columns("SbasUnit").Cells(oRecordSet01.RecordCount).Specific.Value = Trim(oRecordSet01.Fields(8).Value)
		//			//                oRecordSet01.MoveNext
		//			//                Next i
		//			//            End If

		//			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			if ((oRow == oMat01.RowCount | oMat01.VisualRowCount == 0) & !string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value))) {
		//				oMat01.FlushToDataSource();
		//				//데이터 소스를 지우고 매트릭스로부터 데이터 소스 레코드로 각 로를 복사한다.
		//				AddMatrixRow(1, oMat01.RowCount, ref true);
		//				oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
		//				//Column: 칼럼 오브젝트의 collection을 반환한다.
		//			}
		//			break;
		//	}
		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//}

		//private void MTX01()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	////메트릭스에 데이터 로드
		//	return;
		//	MTX01_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private bool HeaderSpaceLineDel()
		//{
		//	bool functionReturnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	short ErrNum = 0;
		//	string DocNum = null;

		//	ErrNum = 0;

		//	//// Check
		//	switch (true) {
		//		case string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_DocDate", 0)):
		//			ErrNum = 1;
		//			goto HeaderSpaceLineDel_Error;
		//			break;
		//		case string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_BPLId", 0)):
		//			ErrNum = 2;
		//			goto HeaderSpaceLineDel_Error;
		//			break;
		//		case string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_OutWhCd", 0)):
		//			ErrNum = 3;
		//			goto HeaderSpaceLineDel_Error;
		//			break;
		//		case string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_InWhCd", 0)):
		//			ErrNum = 4;
		//			goto HeaderSpaceLineDel_Error;
		//			break;
		//	}

		//	functionReturnValue = true;
		//	return functionReturnValue;
		//	HeaderSpaceLineDel_Error:
		//	///////////////////////////////////////////////////////////////////////////////////////////////////////////
		//	if (ErrNum == 1) {
		//		MDC_Com.MDC_GF_Message(ref "거래일자는 필수입력 사항입니다. 확인하세요.", ref "E");
		//	} else if (ErrNum == 2) {
		//		MDC_Com.MDC_GF_Message(ref "사업장은 필수입력 사항입니다. 확인하세요.", ref "E");
		//	} else if (ErrNum == 3) {
		//		MDC_Com.MDC_GF_Message(ref "출고창고는 필수입력 사항입니다. 확인하세요.", ref "E");
		//	} else if (ErrNum == 4) {
		//		MDC_Com.MDC_GF_Message(ref "입고창고는 필수입력 사항입니다. 확인하세요.", ref "E");
		//	} else {
		//		MDC_Com.MDC_GF_Message(ref "HeaderSpaceLineDel_Error:" + Err().Description, ref "E");
		//	}
		//	functionReturnValue = false;
		//	return functionReturnValue;
		//}

		//private bool MatrixSpaceLineDel()
		//{
		//	bool functionReturnValue = false;
		//	//------------------------------------------------------------------------------
		//	// 저장할 데이터의 유효성을 점검한다
		//	//------------------------------------------------------------------------------
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;
		//	int K = 0;
		//	short ErrNum = 0;
		//	string Chk_Data = null;
		//	short oRow = 0;
		//	SAPbobsCOM.Recordset oRecordSet01 = null;

		//	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	ErrNum = 0;

		//	//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
		//	//// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
		//	//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
		//	// Flushes current data from the user interface to the bounded data source, as follows:
		//	// 1. Cleans the data source.
		//	// 2. Copies each row from the matrix to the corresponding data source record.
		//	oMat01.FlushToDataSource();

		//	//// 라인
		//	if (oMat01.VisualRowCount <= 1) {
		//		ErrNum = 1;
		//		goto MatrixSpaceLineDel_Error;
		//	}

		//	//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
		//	//// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
		//	//// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
		//	//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
		//	if (oMat01.VisualRowCount > 0) {

		//		for (i = 0; i <= oMat01.VisualRowCount - 2; i++) {
		//			oDS_PS_SD090L.Offset = i;
		//			switch (true) {
		//				case string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_ItemCode", i)):
		//					ErrNum = 2;
		//					goto MatrixSpaceLineDel_Error;
		//					break;

		//				case string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_Qty", i)):
		//					ErrNum = 5;
		//					goto MatrixSpaceLineDel_Error;
		//					break;

		//				case string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", i)):
		//					ErrNum = 6;
		//					goto MatrixSpaceLineDel_Error;
		//					break;
		//			}
		//		}

		//		if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", oMat01.VisualRowCount - 1))) {
		//			oDS_PS_SD090L.RemoveRecord(oMat01.VisualRowCount - 1);
		//		}
		//	}
		//	//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
		//	//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
		//	//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
		//	oMat01.LoadFromDataSource();

		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//	functionReturnValue = true;
		//	return functionReturnValue;
		//	MatrixSpaceLineDel_Error:
		//	///////////////////////////////////////////////////////////////////////////////////////////////////
		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//	if (ErrNum == 1) {
		//		MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
		//	} else if (ErrNum == 2) {
		//		MDC_Com.MDC_GF_Message(ref "아이템 데이터는 필수입니다. 확인하세요.", ref "E");
		//	} else if (ErrNum == 5) {
		//		MDC_Com.MDC_GF_Message(ref "수량은 필수입니다. 확인하세요.", ref "E");
		//	} else if (ErrNum == 6) {
		//		MDC_Com.MDC_GF_Message(ref "이동요청 문서는 필수입니다. 확인하세요.", ref "E");
		//	} else {
		//		MDC_Com.MDC_GF_Message(ref "MatrixSpaceLineDel_Error:" + Err().Description, ref "E");
		//	}
		//	functionReturnValue = false;
		//	return functionReturnValue;
		//}

		//private string Exist_YN(ref string DocNum)
		//{
		//	string functionReturnValue = null;

		//	SAPbobsCOM.Recordset oRecordSet01 = null;
		//	string sQry = null;

		//	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	sQry = "SELECT Top 1 T1.DocNum FROM [@PS_SD090H] T1 ";
		//	sQry = sQry + " WHERE T1.DocNum  = '" + DocNum + "'";
		//	oRecordSet01.DoQuery(sQry);

		//	while (!(oRecordSet01.EoF)) {
		//		//UPGRADE_WARNING: oRecordSet01().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		functionReturnValue = oRecordSet01.Fields.Item(0).Value;
		//		oRecordSet01.MoveNext();
		//	}

		//	if (string.IsNullOrEmpty(Strings.Trim(Exist_YN()))) {
		//		functionReturnValue = "";
		//		return functionReturnValue;
		//	}

		//	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet01 = null;
		//	return functionReturnValue;
		//}

		//private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;
		//	if (pVal.BeforeAction == true) {
		//		if (SubMain.Sbo_Application.MessageBox("정말 삭제 하시겠습니까?", 1, "OK", "NO") != 1) {
		//			BubbleEvent = false;
		//		}
		//		////행삭제전 행삭제가능여부검사
		//	} else if (pVal.BeforeAction == false) {
		//		for (i = 1; i <= oMat01.VisualRowCount; i++) {
		//			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
		//		}
		//		//        oMat01.Clear
		//		oMat01.FlushToDataSource();
		//		oDS_PS_SD090L.RemoveRecord(oDS_PS_SD090L.Size - 1);
		//		oMat01.LoadFromDataSource();
		//		if (oMat01.RowCount == 0) {
		//			AddMatrixRow(0, oMat01.RowCount, ref true);
		//		} else {
		//			if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD090L.GetValue("U_ItemCode", oMat01.RowCount - 1)))) {
		//				AddMatrixRow(1, oMat01.RowCount, ref true);
		//			}
		//		}
		//	}
		//	return;
		//	Raise_EVENT_ROW_DELETE_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}

		//private bool StockTrans()
		//{
		//	bool functionReturnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	int RetVal = 0;
		//	int ErrCode = 0;
		//	string ErrMsg = null;
		//	string lQuery = null;
		//	SAPbobsCOM.Recordset lRecordSet = null;
		//	SAPbobsCOM.Recordset oRecordSet = null;
		//	int lMaxBatchNum = 0;
		//	//해당 품목의 최대 배치번호
		//	double lBatchWeight = 0;
		//	//배치별 중량
		//	short lTypeCount = 0;
		//	//전체 StockInfo 구조체배열의 RowCount
		//	object Q = null;
		//	object j = null;
		//	object i = null;
		//	object K = null;
		//	object z = null;
		//	int r = 0;
		//	int DocCnt = 0;
		//	string Chk1_Val = null;
		//	string sCur_ItemCode = null;
		//	string sNxt_ItemCode = null;
		//	string sCur_TrCardCode = null;
		//	string sCur_TrOutWhs = null;
		//	string sNxt_TrOutWhs = null;
		//	string sCur_TrInWhs = null;
		//	string sNxt_TrInWhs = null;
		//	string RtnDocNum = null;
		//	SAPbobsCOM.StockTransfer oStockTrans = null;
		//	SAPbouiCOM.ProgressBar oPrgBar = null;
		//	int StockTransLineCounter = 0;

		//	lRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	string BatchNum = null;

		//	functionReturnValue = true;

		//	for (i = 0; i <= oMat01.RowCount - 1; i++) {
		//		Array.Resize(ref StockInfo, lTypeCount + 1);
		//		//DI API
		//		StockInfo[lTypeCount].CardCode = Strings.Trim(oDS_PS_SD090H.GetValue("U_CardCode", 0));
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].ItemCode = Strings.Trim(oDS_PS_SD090L.GetValue("U_ItemCode", i));
		//		StockInfo[lTypeCount].FromWarehouseCode = Strings.Trim(oDS_PS_SD090H.GetValue("U_OutWhCd", 0));
		//		StockInfo[lTypeCount].ToWarehouseCode = Strings.Trim(oDS_PS_SD090H.GetValue("U_InWhCd", 0));
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].BatchNum = Strings.Trim(oDS_PS_SD090L.GetValue("U_BatchNum", i));

		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].Weight = System.Math.Round(Convert.ToDouble(Strings.Trim(oDS_PS_SD090L.GetValue("U_Weight", i))), 2);
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].UnWeight = System.Math.Round(Convert.ToDouble(Strings.Trim(oDS_PS_SD090L.GetValue("U_Unweight", i))), 2);
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].BatchWeight = System.Math.Round(Convert.ToDouble(Strings.Trim(oDS_PS_SD090L.GetValue("U_Qty", i))), 2);
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].Qty = Conversion.Val(Strings.Trim(oDS_PS_SD090L.GetValue("U_Qty", i)));

		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].TransNo = oForm.Items.Item("DocNum").Specific.Value + (i + 1);
		//		StockInfo[lTypeCount].Chk = "N";
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].MatrixRow = (i + 1);
		//		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[lTypeCount].Indate = oForm.Items.Item("DocDate").Specific.Value;
		//		lTypeCount = lTypeCount + 1;
		//	}

		//	for (i = 0; i <= (Information.UBound(StockInfo)); i++) {
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[i].StockTransDocEntry = "";
		//	}

		//	SubMain.Sbo_Company.StartTransaction();
		//	for (i = 0; i <= (Information.UBound(StockInfo)); i++) {

		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		Chk1_Val = StockInfo[i].Chk;

		//		if (Chk1_Val != "N")
		//			goto Continue_First;

		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		sCur_TrOutWhs = StockInfo[i].FromWarehouseCode;

		//		oStockTrans = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oStockTrans.CardCode = StockInfo[i].CardCode;
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oStockTrans.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(StockInfo[i].Indate, "&&&&-&&-&&"));
		//		oStockTrans.FromWarehouse = sCur_TrOutWhs;
		//		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oStockTrans.Comments = "재고이전" + oForm.Items.Item("DocNum").Specific.Value + ".";

		//		StockTransLineCounter = -1;
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		for (K = i; K <= (Information.UBound(StockInfo)); K++) {
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			Chk1_Val = StockInfo[K].Chk;

		//			if (Chk1_Val != "N")
		//				goto Continue_Second;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			sCur_TrCardCode = StockInfo[K].CardCode;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			sNxt_TrOutWhs = StockInfo[K].FromWarehouseCode;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			sCur_ItemCode = StockInfo[K].ItemCode;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			sCur_TrInWhs = StockInfo[K].ToWarehouseCode;

		//			if ((sCur_TrOutWhs != sNxt_TrOutWhs)) {
		//				goto Continue_Second;
		//			}

		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			if ((i != K)) {
		//				oStockTrans.Lines.Add();
		//			}
		//			StockTransLineCounter = StockTransLineCounter + 1;
		//			//---------------------------------------------------------------------------< Line >----------
		//			var _with1 = oStockTrans.Lines;

		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.ItemCode = StockInfo[K].ItemCode;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.UserFields.Fields.Item("U_Qty").Value = Strings.Trim(Convert.ToString(StockInfo[K].Qty));
		//			//// 수량
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.UserFields.Fields.Item("U_UnWeight").Value = Strings.Trim(Convert.ToString(StockInfo[K].UnWeight));
		//			//// 단중
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.Quantity = System.Math.Round(StockInfo[K].Qty, 2);
		//			//// 수량
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.WarehouseCode = StockInfo[K].ToWarehouseCode;
		//			////ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.UserFields.Fields.Item("U_BatchNum").Value = StockInfo[K].BatchNum;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.BatchNumbers.BatchNumber = StockInfo[K].BatchNum;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.BatchNumbers.Quantity = System.Math.Round(StockInfo[K].BatchWeight, 2);

		//			_with1.BatchNumbers.Notes = "재고이전(Addon)";
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[K].Chk = "Y";
		//			/// 적용한 라인에 대한 표시
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[K].StockTransDocEntry = "Checked";
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[K].StockTransLineNum = Convert.ToString(StockTransLineCounter);

		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			for (Q = K + 1; Q <= (Information.UBound(StockInfo)); Q++) {
		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				Chk1_Val = StockInfo[Q].Chk;

		//				if (Chk1_Val != "N")
		//					goto Continue_Sixth;
		//				/// 체크2 에 않된건 Skip

		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				sNxt_TrOutWhs = StockInfo[Q].FromWarehouseCode;
		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				sNxt_ItemCode = StockInfo[Q].ItemCode;
		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				sNxt_TrInWhs = StockInfo[Q].ToWarehouseCode;

		//				if (sNxt_TrOutWhs == sCur_TrOutWhs & sCur_ItemCode == sNxt_ItemCode & sCur_TrInWhs == sNxt_TrInWhs) {
		//					////ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
		//					//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT ManBatchNum FROM OITM WHERE ITEMCODE = '', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//					if (MDC_PS_Common.GetValue("SELECT ManBatchNum FROM OITM WHERE ITEMCODE = ''", 0, 1) == "Y") {
		//						_with1.BatchNumbers.Add();
		//						//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						_with1.BatchNumbers.BatchNumber = StockInfo[Q].BatchNum;
		//						//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						_with1.BatchNumbers.Quantity = System.Math.Round(StockInfo[Q].BatchWeight, 2);
		//						//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						_with1.UserFields.Fields.Item("Quantity").Value = _with1.UserFields.Fields.Item("Quantity").Value + Strings.Trim(Convert.ToString(StockInfo[Q].Qty));
		//						////수량
		//						//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						_with1.Quantity = _with1.Quantity + System.Math.Round(StockInfo[Q].Weight, 2);
		//						////중량을 합함
		//						_with1.BatchNumbers.Notes = "재고이전(Addon)";
		//						//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						StockInfo[Q].Chk = "Y";
		//						//// 적용한 라인에 대한 표시
		//						//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						StockInfo[Q].StockTransDocEntry = "Checked";
		//						//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//						StockInfo[Q].StockTransLineNum = Convert.ToString(StockTransLineCounter);
		//					}
		//				}
		//				Continue_Sixth:
		//			}
		//			Continue_Second:
		//		}
		//		//---------------------------------------------------------------------------------------------

		//		RetVal = oStockTrans.Add();
		//		if (RetVal == 0) {
		//			DocCnt = DocCnt + 1;
		//			SubMain.Sbo_Company.GetNewObjectCode(out RtnDocNum);
		//			////재고이전문서번호
		//			for (r = 0; r <= Information.UBound(StockInfo); r++) {
		//				if ((StockInfo[r].StockTransDocEntry == "Checked")) {
		//					StockInfo[r].StockTransDocEntry = RtnDocNum;
		//				}
		//			}
		//			//// 데이터 업데이트
		//		} else {
		//			goto StockTrans_Error;
		//		}
		//		Continue_First:
		//	}
		//	//-----------------------------------------------------------------------------------------------< First For End

		//	if ((SubMain.Sbo_Company.InTransaction)) {
		//		SubMain.Sbo_Company.EndTransaction((SAPbobsCOM.BoWfTransOpt.wf_Commit));
		//	}
		//	SubMain.Sbo_Application.SetStatusBarMessage(DocCnt + " 개의 재고이전 문서가 발행되었습니다 !", SAPbouiCOM.BoMessageTime.bmt_Short, false);
		//	//UPGRADE_NOTE: oStockTrans 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oStockTrans = null;
		//	//UPGRADE_NOTE: lRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	lRecordSet = null;
		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	return functionReturnValue;
		//	StockTrans_Error:
		//	//************Error Process************
		//	if ((SubMain.Sbo_Company.InTransaction)) {
		//		SubMain.Sbo_Company.EndTransaction((SAPbobsCOM.BoWfTransOpt.wf_RollBack));
		//	}
		//	SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
		//	SubMain.Sbo_Application.SetStatusBarMessage(ErrCode + " : " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	functionReturnValue = false;
		//	//UPGRADE_NOTE: oStockTrans 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oStockTrans = null;
		//	//UPGRADE_NOTE: lRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	lRecordSet = null;
		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	return functionReturnValue;
		//	//************Error Process************

		//}

		//private bool UpdateUserField()
		//{
		//	bool functionReturnValue = false;
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;
		//	string lQuery = null;
		//	SAPbobsCOM.Recordset lRecordSet = null;
		//	lRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	SAPbobsCOM.Recordset RecordSet01 = null;
		//	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oDS_PS_SD090H.SetValue("U_StoTrDoc", 0, (StockInfo[i].StockTransDocEntry));
		//	//    oForm.Items("StoTrDoc").Specific.Value = StockInfo(i).StockTransDocEntry

		//	functionReturnValue = true;
		//	return functionReturnValue;
		//	UpdateUserField_Error:
		//	functionReturnValue = false;
		//	return functionReturnValue;
		//}

		//private void PS_SD090_Print_Report01()
		//{
		//	 // ERROR: Not supported in C#: OnErrorStatement

		//	string DocNum = null;
		//	string WinTitle = null;
		//	string ReportName = null;
		//	string sQry01 = null;

		//	MDC_PS_Common.ConnectODBC();
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	DocNum = Strings.Trim(oForm.Items.Item("DocNum").Specific.Value);
		//	WinTitle = "[PS_SD090] 출고원부/반출증";
		//	ReportName = "PS_SD090_10.rpt";
		//	sQry01 = "EXEC PS_SD090_10 '" + DocNum + "'";
		//	MDC_Globals.gRpt_Formula = new string[2];
		//	MDC_Globals.gRpt_Formula_Value = new string[2];
		//	MDC_Globals.gRpt_SRptSqry = new string[2];
		//	MDC_Globals.gRpt_SRptName = new string[2];
		//	MDC_Globals.gRpt_SFormula = new string[2, 2];
		//	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

		//	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry01, "1", "Y", "V") == false) {
		//		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	}
		//	return;
		//	PS_SD090_Print_Report01_Error:
		//	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD090_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//}
	}
}
