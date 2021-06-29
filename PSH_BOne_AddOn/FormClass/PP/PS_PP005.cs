using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
namespace MDC_PS_Addon
{
	internal class PS_PP005
	{
//****************************************************************************************************************
////  File           : PS_PP005.cls
////  Module         : PP
////  Description    : 제품 원재료관계등록
////  FormType       : PS_PP005H
////  Create Date    : 2010.10.20
////  Creator        : Lee Byong Gak
////  Company        : Poongsan Holdings
//****************************************************************************************************************

		public string oFormUniqueID01;
		public SAPbouiCOM.Form oForm01;
		public SAPbouiCOM.Matrix oMat01;
			//등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP005H;

			//클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Item_UID;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private string oLast_Col_UID;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Col_Row;

		private int oLast_Mode;

//****************************************************************************************************************
// .srf 파일로부터 폼을 로드한다.
//****************************************************************************************************************
		public void LoadForm()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string oInnerXml01 = null;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_PP005.srf");
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

			//매트릭스의 타이틀높이와 셀높이를 고정
			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
			}

			oFormUniqueID01 = "PS_PP005_" + GetTotalFormsCount();
			SubMain.AddForms(this, oFormUniqueID01);
			////폼추가
			SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));

			//폼 할당
			oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

			oForm01.SupportedModes = -1;
			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

			//////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//************************************************************************************************************
			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
			//   oForm01.DataBrowser.BrowseBy = "DocNum"
			//************************************************************************************************************
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////

			oForm01.Freeze(true);
			CreateItems();
			ComboBox_Setting();
			//    Call Initialization
			//    Call PS_PP005_CF_ChooseFromList
			//    Call PS_PP005_FormItemEnabled
			//    Call FormClear
			Add_MatrixRow(0, ref true);

			FormItemEnabled();

			oForm01.EnableMenu(("1283"), false);
			//// 삭제
			oForm01.EnableMenu(("1287"), false);
			//// 복제
			oForm01.EnableMenu(("1286"), true);
			//// 닫기
			oForm01.EnableMenu(("1284"), true);
			//// 취소
			oForm01.EnableMenu(("1293"), true);
			//// 행삭제
			oForm01.EnableMenu(("1282"), true);

			oForm01.Update();
			oForm01.Freeze(false);
			oForm01.Visible = true;

			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			return;
			LoadForm_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			oForm01.Update();
			oForm01.Freeze(false);
			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			if ((oForm01 == null) == false) {
				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oForm01 = null;
			}
			MDC_Com.MDC_GF_Message(ref "LoadForm_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

//// 데이타 insert
		public bool Add_PurchaseDemand(ref SAPbouiCOM.ItemEvent pval)
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement

			short i = 0;
			string sQry = null;
			string sQry1 = null;
			short ReturnValue = 0;
			SAPbobsCOM.Recordset RecordSet01 = null;
			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string BPLId = null;
			string DocEntry = null;
			string DocNum = null;
			string LineNum = null;
			decimal WorkTime = default(decimal);
			double UnWeight = 0;
			string ItemNam2 = null;
			string ItemNam1 = null;
			string ItemCod1 = null;
			string ItemCod2 = null;
			string Indate = null;
			string baseChk = null;
			string convChk = null;


			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemCod1 = Strings.Trim(oForm01.Items.Item("ItemCod1").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemNam1 = Strings.Trim(oForm01.Items.Item("ItemNam1").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemCod2 = Strings.Trim(oForm01.Items.Item("ItemCod2").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemNam2 = Strings.Trim(oForm01.Items.Item("ItemNam2").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			UnWeight = Convert.ToDouble(Strings.Trim(oForm01.Items.Item("UnWeight").Specific.VALUE));

			//UPGRADE_WARNING: oForm01.Items(BaseChk).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (oForm01.Items.Item("BaseChk").Specific.Checked == true) {
				baseChk = "Y";
			} else {
				baseChk = "N";
			}

			//UPGRADE_WARNING: oForm01.Items(ConvChk).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (oForm01.Items.Item("ConvChk").Specific.Checked == true) {
				convChk = "Y";
			} else {
				convChk = "N";
			}


			Indate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "yymmdd");
			//오늘날짜를 넣어주기 위한 선언부분

			sQry1 = "Select U_ItemCod1, U_ItemCod2 From [@PS_PP005H] Where U_ItemCod1 ='" + ItemCod1 + "' AND U_ItemCod2 = '" + ItemCod2 + "'";

			RecordSet01.DoQuery(sQry1);
			if ((RecordSet01.RecordCount > 0)) {
				MDC_Com.MDC_GF_Message(ref "기존자료가 존재합니다.:" + Err().Number + " - " + Err().Description, ref "E");
				//        ReturnValue = Sbo_Application.MessageBox("기존자료가 존재합니다.", 1, "&확인", "&취소")   '오류시 팝업창에 Error표시
				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				RecordSet01 = null;
				return functionReturnValue;
			}

			if (UnWeight <= 0) {
				MDC_Com.MDC_GF_Message(ref "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.:" + Err().Number + " - " + Err().Description, ref "E");
				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				RecordSet01 = null;
				return functionReturnValue;
			}

			sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_PP005H]";
			RecordSet01.DoQuery(sQry);
			//        If Trim(RecordSet01.Fields(0).Value) = 0 Then
			//            DocEntry = Left(DocDate, 6) + "0001"
			//        Else
			DocEntry = Convert.ToString(Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1);
			//        End If

			//        ItemCod1 = Trim(oDS_PS_PP005H.GetValue("U_ItemCod1", i))
			//        CntcName = Trim(oDS_PS_PP060H.GetValue("U_CntcName", i))
			//        ItmBSort = Trim(oDS_PS_PP060H.GetValue("U_ItmBsort", i))
			//        WorkNote = Trim(oDS_PS_PP060H.GetValue("U_WorkNote", i))
			//        WorkTime = Trim(oDS_PS_PP060H.GetValue("U_WorkTime", i))

			sQry = "INSERT INTO [@PS_PP005H]";
			sQry = sQry + " (";
			sQry = sQry + " DocEntry,";
			sQry = sQry + " DocNum,";
			sQry = sQry + " U_ItemCod1,";
			sQry = sQry + " U_ItemNam1,";
			sQry = sQry + " U_ItemCod2,";
			sQry = sQry + " U_ItemNam2,";
			sQry = sQry + " U_UnWeight,";
			sQry = sQry + " U_InDate,";
			sQry = sQry + " U_BaseChk,";
			sQry = sQry + " U_ConvChk";
			sQry = sQry + " ) ";
			sQry = sQry + "VALUES(";
			sQry = sQry + DocEntry + ",";
			sQry = sQry + DocEntry + ",";
			sQry = sQry + "'" + ItemCod1 + "',";
			sQry = sQry + "'" + ItemNam1 + "',";
			sQry = sQry + "'" + ItemCod2 + "',";
			sQry = sQry + "'" + ItemNam2 + "',";
			sQry = sQry + UnWeight + ",";
			sQry = sQry + Indate + ",";
			sQry = sQry + "'" + baseChk + "',";
			sQry = sQry + "'" + convChk + "'";
			sQry = sQry + ")";
			RecordSet01.DoQuery(sQry);
			//        Next

			MDC_Com.MDC_GF_Message(ref "제품코드 및 원자재코드 정상등록!", ref "S");

			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			functionReturnValue = true;
			return functionReturnValue;
			Add_PurchaseDemand_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			functionReturnValue = false;
			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			MDC_Com.MDC_GF_Message(ref "Add_PurchaseDemand_Error:" + Err().Number + " - " + Err().Description, ref "E");
			return functionReturnValue;
		}

//// 데이타 Update
		public bool UpdateData()
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement

			short i = 0;
			string sQry = null;
			SAPbobsCOM.Recordset RecordSet01 = null;
			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string ItemCode = null;
			string DocEntry = null;
			string ItemName = null;
			int Qty = 0;
			decimal Price = default(decimal);
			decimal Weight = default(decimal);
			decimal LinTotal = default(decimal);
			string ItemCod2 = null;
			object ItemNam2 = null;
			string Chk = null;
			string MoDate = null;
			double UnWeight = 0;
			string baseChk = null;
			string convChk = null;

			MoDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "yymmdd");
			//오늘날짜를 불러오기 위한 변수선언


			oMat01.FlushToDataSource();

			for (i = 1; i <= oMat01.RowCount; i++) {
				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

				if (Convert.ToBoolean(Chk) == true) {
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					UnWeight = oMat01.Columns.Item("UnWeight").Cells.Item(i).Specific.VALUE;

					if (UnWeight <= 0) {
						MDC_Com.MDC_GF_Message(ref "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.:" + Err().Number + " - " + Err().Description, ref "E");
						return functionReturnValue;
					}
				}
			}

			for (i = 1; i <= oMat01.RowCount; i++) {
				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

				if (Convert.ToBoolean(Chk) == true) {
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.VALUE;
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ItemCod2 = oMat01.Columns.Item("ItemCod2").Cells.Item(i).Specific.VALUE;
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					//UPGRADE_WARNING: ItemNam2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ItemNam2 = oMat01.Columns.Item("ItemNam2").Cells.Item(i).Specific.VALUE;
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					UnWeight = oMat01.Columns.Item("UnWeight").Cells.Item(i).Specific.VALUE;

					//UPGRADE_WARNING: oMat01.Columns(BaseChk).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					if (oMat01.Columns.Item("BaseChk").Cells.Item(i).Specific.Checked == true) {
						baseChk = "Y";
					} else {
						baseChk = "N";
					}

					//UPGRADE_WARNING: oMat01.Columns(ConvChk).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					if (oMat01.Columns.Item("ConvChk").Cells.Item(i).Specific.Checked == true) {
						convChk = "Y";
					} else {
						convChk = "N";
					}


					sQry = "Update [@PS_PP005H] set ";
					sQry = sQry + " U_ItemCod2   = '" + ItemCod2 + "',";
					//UPGRADE_WARNING: ItemNam2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sQry = sQry + " U_ItemNam2   = '" + ItemNam2 + "',";
					sQry = sQry + " U_UnWeight   = " + UnWeight + ",";
					sQry = sQry + " U_BaseChk   = '" + baseChk + "',";
					sQry = sQry + " U_ConvChk   = '" + convChk + "',";
					sQry = sQry + " U_MoDate     = '" + MoDate + "'";
					//오늘날짜 가져오기

					sQry = sQry + " Where DocEntry = '" + DocEntry + "'";
					RecordSet01.DoQuery(sQry);
				}
			}

			MDC_Com.MDC_GF_Message(ref "원자재코드 수정완료!", ref "S");

			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			functionReturnValue = true;
			return functionReturnValue;
			UpdateData_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			functionReturnValue = false;
			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			RecordSet01 = null;
			MDC_Com.MDC_GF_Message(ref "UpdateData_Error:" + Err().Number + " - " + Err().Description, ref "E");
			return functionReturnValue;
		}

		public void DeleteData()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			SAPbouiCOM.ProgressBar ProgBar01 = null;
			short i = 0;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;
			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string ItemCode = null;
			string DocEntry = null;
			string ItemName = null;
			int Qty = 0;
			decimal Price = default(decimal);
			decimal Weight = default(decimal);
			decimal LinTotal = default(decimal);
			string ItemCod2 = null;
			object ItemNam2 = null;
			string Chk = null;
			string MoDate = null;

			oMat01.FlushToDataSource();

			for (i = 1; i <= oMat01.RowCount; i++) {
				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

				if (Convert.ToBoolean(Chk) == true) {
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.VALUE;
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ItemCod2 = oMat01.Columns.Item("ItemCod2").Cells.Item(i).Specific.VALUE;
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					//UPGRADE_WARNING: ItemNam2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ItemNam2 = oMat01.Columns.Item("ItemNam2").Cells.Item(i).Specific.VALUE;

					sQry = "Delete From [@PS_PP005H] where DocEntry = '" + DocEntry + "'";
					oRecordSet01.DoQuery(sQry);
				}

			}

			oMat01.Clear();
			oMat01.FlushToDataSource();
			oMat01.LoadFromDataSource();
			Add_MatrixRow(0, ref true);
			LoadData();
			return;
			DeleteData_Error:

			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			ProgBar01.Stop();
			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgBar01 = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			MDC_Com.MDC_GF_Message(ref "DeleteData_Error:" + Err().Number + " - " + Err().Description, ref "E");

		}

////조회데이타 가져오기
		public void LoadData()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			short i = 0;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;
			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			string ItmBsort = null;
			string ItemCod1 = null;
			string ItemCod2 = null;
			string ItmMsort = null;

			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemCod1 = Strings.Trim(oForm01.Items.Item("ItemCod1").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemCod2 = Strings.Trim(oForm01.Items.Item("ItemCod2").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItmBsort = Strings.Trim(oForm01.Items.Item("ItmBSort").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItmMsort = Strings.Trim(oForm01.Items.Item("ItmMSort").Specific.VALUE);

			if (string.IsNullOrEmpty(ItemCod1))
				ItemCod1 = "%";
			if (string.IsNullOrEmpty(ItemCod2))
				ItemCod2 = "%";
			if (string.IsNullOrEmpty(ItmMsort))
				ItmMsort = "%";
			if (string.IsNullOrEmpty(ItmBsort))
				ItmBsort = "%";

			sQry = "EXEC [PS_PP005_01] '" + ItmBsort + "','" + ItmMsort + "','" + ItemCod1 + "','" + ItemCod2 + "'";

			oRecordSet01.DoQuery(sQry);

			oMat01.Clear();
			oDS_PS_PP005H.Clear();
			oMat01.LoadFromDataSource();

			if ((oRecordSet01.RecordCount == 0)) {
				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oRecordSet01 = null;
				return;
			}

			oForm01.Freeze(true);
			SAPbouiCOM.ProgressBar ProgBar01 = null;
			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
				if (i + 1 > oDS_PS_PP005H.Size) {
					oDS_PS_PP005H.InsertRecord((i));
				}

				oMat01.AddRow();
				oDS_PS_PP005H.Offset = i;

				oDS_PS_PP005H.SetValue("DocEntry", i, Strings.Trim(oRecordSet01.Fields.Item("DocEntry").Value));
				oDS_PS_PP005H.SetValue("DocNum", i, Convert.ToString(i + 1));
				oDS_PS_PP005H.SetValue("U_ItemCod1", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItemCod1").Value));
				oDS_PS_PP005H.SetValue("U_ItemNam1", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItemNam1").Value));
				oDS_PS_PP005H.SetValue("U_ItmMSort", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItmMSort").Value));
				oDS_PS_PP005H.SetValue("U_BaseChk", i, Strings.Trim(oRecordSet01.Fields.Item("U_BaseChk").Value));
				oDS_PS_PP005H.SetValue("U_ConvChk", i, Strings.Trim(oRecordSet01.Fields.Item("U_ConvChk").Value));
				//        oMat01.Columns("ItmMsort").Cells(i + 1).Specific.String = Trim(oRecordSet01.Fields("U_ItmMsort").Value)
				oDS_PS_PP005H.SetValue("U_ItemCod2", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItemCod2").Value));
				oDS_PS_PP005H.SetValue("U_ItemNam2", i, Strings.Trim(oRecordSet01.Fields.Item("U_ItemNam2").Value));
				oDS_PS_PP005H.SetValue("U_UnWeight", i, Strings.Trim(oRecordSet01.Fields.Item("U_UnWeight").Value));
				oDS_PS_PP005H.SetValue("U_InDate", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_InDate").Value), "YYYYMMDD"));
				oDS_PS_PP005H.SetValue("U_MoDate", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("U_MoDate").Value), "YYYYMMDD"));

				oRecordSet01.MoveNext();
				ProgBar01.Value = ProgBar01.Value + 1;
				ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
			}
			oMat01.LoadFromDataSource();
			oMat01.AutoResizeColumns();
			ProgBar01.Stop();

			oForm01.Freeze(false);

			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgBar01 = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;

			return;
			LoadData_Error:

			MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");

		}

//****************************************************************************************************************
//// ItemEventHander
//****************************************************************************************************************
		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{

			 // ERROR: Not supported in C#: OnErrorStatement


			short i = 0;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;

			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			int sCount = 0;
			int sSeq = 0;
			////BeforeAction = True
			if ((pval.BeforeAction == true)) {
				switch (pval.EventType) {
					//et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
						////1                               '버튼 클릭시 발생하는 Event
						if (pval.ItemUID == "Btn01") {
							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
								if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
									//저장 버튼클릭
									if (pval.ItemUID == "Btn01") {
										if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
											if (HeaderSpaceLineDel() == false) {
												BubbleEvent = false;
												return;
											}

											if (Add_PurchaseDemand(ref pval) == false) {
												BubbleEvent = false;
												return;
											}

											oMat01.Clear();
											oMat01.FlushToDataSource();
											oMat01.LoadFromDataSource();
											Add_MatrixRow(0, ref true);

											//                        Call Delete_EmptyRow
											oLast_Mode = oForm01.Mode;
											//                                ElseIf oForm01.Mode = fm_UPDATE_MODE Then
											//                                    If Updatedata(pval) = False Then
											//                                        BubbleEvent = False
											//                                        Exit Sub
											//                                    End If
											//                                    Call LoadData
										}
									}
								}
								oLast_Mode = oForm01.Mode;
							}
						//내역조회
						} else if (pval.ItemUID == "Btn02") {
							LoadData();
						} else if (pval.ItemUID == "Btn03") {
							oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							///fm_VIEW_MODE
							DeleteData();
						} else if (pval.ItemUID == "Btn04") {
							oForm01.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							UpdateData();
						}
						return;

						break;
					//et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
						////2                         '탭키를 눌렀을 때 발생하는 Event

						// 질의 관리자 창을 사용할 때... 선언부분은 FlushTo ItemValue에서 선언

						if (pval.CharPressed == 9) {
							if (pval.ItemUID == "ItemCod1") {
								//UPGRADE_WARNING: oForm01.Items(ItemCod1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								if (string.IsNullOrEmpty(oForm01.Items.Item("ItemCod1").Specific.VALUE)) {
									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
									//질의 관리자 사용시(탭키를 눌렀을 때)
									BubbleEvent = false;
								}
							} else if (pval.ItemUID == "ItemCod2") {
								//UPGRADE_WARNING: oForm01.Items(ItemCod2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								if (string.IsNullOrEmpty(oForm01.Items.Item("ItemCod2").Specific.VALUE)) {
									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
									BubbleEvent = false;
								}
							// Matrix에 질의 관리자 사용시 선언
							} else if (pval.ColUID == "ItemCod2") {
								//UPGRADE_WARNING: oMat01.Columns(ItemCod2).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCod2").Cells.Item(pval.Row).Specific.VALUE)) {
									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
									BubbleEvent = false;
								}
							}
						}
						break;

					//et_COMBO_SELECT ////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
						////5
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:
						////6
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
						////7
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
						////8
						break;
					//et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
						////10
						break;


					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
						////11
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
						////18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
						////19
						break;
					//et_FORM_RESIZE//////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
						////20
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
						////27
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
						////3
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
						////4
						break;
					//et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
						////17
						break;
				}
			////BeforeAction = False
			} else if ((pval.BeforeAction == false)) {
				switch (pval.EventType) {
					//et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
						////1
						if (pval.ItemUID == "1") {
							Add_MatrixRow(oMat01.RowCount, ref false);
							oLast_Mode = 0;
							//                    If oForm01.Mode = fm_OK_MODE And oLast_Mode = fm_UPDATE_MODE Then
							//
							//                    End If
						}
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
						////2
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
						////5

						if (pval.ItemUID == "ItmBSort") {
							//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sCount = Convert.ToInt32(Strings.Trim(oForm01.Items.Item("ItmMSort").Specific.ValidValues.Count));
							sSeq = sCount;
							for (i = 1; i <= sCount; i++) {
								//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								oForm01.Items.Item("ItmMSort").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
								sSeq = sSeq - 1;
							}

							////중분류
							//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sQry = "SELECT U_Code, U_CodeName From [@PSH_ITMMSORT] Where U_rCode = '" + Strings.Trim(oForm01.Items.Item("ItmBSort").Specific.VALUE) + "' Order by Code";
							oRecordSet01.DoQuery(sQry);
							//                    oForm01.Items("ItmMSort").Specific.ValidValues.Add "", ""
							while (!(oRecordSet01.EoF)) {
								//UPGRADE_WARNING: oForm01.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								oForm01.Items.Item("ItmMSort").Specific.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
								oRecordSet01.MoveNext();
							}
							//                    If oRecordSet01.RecordCount <> 0 Then
							//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							oForm01.Items.Item("ItmMSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
							//                    Else
							//                        oForm01.Items("ItmMSort").Specific.Select "", psk_ByValue
							//                    End If
						} else if (pval.ItemUID == "Mat01") {
							///                    If oForm01.Mode = fm_ADD_MODE Then
							///                    Else
							///                        oForm01.Mode = fm_UPDATE_MODE
							///                        Call LoadCaption
							///                    End If
						}
						break;

					case SAPbouiCOM.BoEventTypes.et_CLICK:
						////6
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
						////7
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
						////8
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
						////10       ' 필드의 값이 바뀌었을 때 동작하는 부분
						if (pval.ItemChanged == true) {
							if (pval.ItemUID == "ItemCod1") {
								FlushToItemValue(pval.ItemUID);
							}
							if (pval.ItemUID == "ItemCod2") {
								FlushToItemValue(pval.ItemUID);
							}
							if (pval.ColUID == "ItemCod2") {
								FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
							}
						}
						break;

					//                If pval.ItemChanged = True Then
					//                    If pval.ItemUID = "CardCode" Then
					//                        FlushToItemValue pval.ItemUID
					//                    ElseIf pval.ItemUID = "CntcCode" Then
					//                        FlushToItemValue pval.ItemUID
					//                    ElseIf pval.ItemUID = "Mat01" Then
					//                        If pval.ColUID = "GADocLin" Then
					//                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
					//                        End If
					//                    End If
					//                End If

					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
						////11
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
						////18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
						////19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
						////20
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
						////27
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
						////3
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
						////4
						break;
					//et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
						////17
						SubMain.RemoveForms(oFormUniqueID01);
						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						oForm01 = null;
						//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						oMat01 = null;
						break;
				}
			}
			return;
			Raise_ItemEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;

			////BeforeAction = True
			if ((pval.BeforeAction == true)) {
				switch (pval.MenuUID) {
					case "1284":
						//취소
						break;


					case "1286":
						//닫기
						break;


					case "1293":
						//행삭제
						break;
					case "1281":
						//찾기
						break;
					case "1282":
						//추가
						break;
					case "1288":
					case "1289":
					case "1290":
					case "1291":
						//레코드이동버튼
						break;
				}
			////BeforeAction = False
			} else if ((pval.BeforeAction == false)) {
				switch (pval.MenuUID) {
					//            Case "1284": '취소
					//                FormItemEnabled
					//                oForm01.Items("DocNum").Click ct_Regular
					case "1286":
						//닫기
						break;
					case "1293":
						//행삭제

						if (oMat01.RowCount != oMat01.VisualRowCount) {
							for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.VALUE = i + 1;
							}

							oMat01.FlushToDataSource();
							oDS_PS_PP005H.RemoveRecord(oDS_PS_PP005H.Size - 1);
							//// Mat01에 마지막라인(빈라인) 삭제
							oMat01.Clear();
							oMat01.LoadFromDataSource();

							//UPGRADE_WARNING: oMat01.Columns(PQDocNum).Cells(oMat01.RowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							if (!string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(oMat01.RowCount).Specific.VALUE)) {
								Add_MatrixRow(oMat01.RowCount, ref false);
							}
						}
						break;

					case "1281":
						//찾기
						FormItemEnabled();
						break;
					//oForm01.Items("DocNum").Click ct_Regular

					case "1282":
						//추가
						//                oForm01.Items("ItemCod1").Specific.Value = ""
						//                oForm01.Items("ItemCod2").Specific.Value = ""
						oMat01.Clear();
						oDS_PS_PP005H.Clear();
						break;
					//                oDS_PS_PP005H.GetValue("U_ItemCod1",0)

					case "1288":
					case "1289":
					case "1290":
					case "1291":
						//레코드이동버튼
						FormItemEnabled();
						if (oMat01.VisualRowCount > 0) {
							//UPGRADE_WARNING: oMat01.Columns(PQDocNum).Cells(oMat01.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							if (!string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(oMat01.VisualRowCount).Specific.VALUE)) {
								if (oDS_PS_PP005H.GetValue("Status", 0) == "O") {
									Add_MatrixRow(oMat01.RowCount, ref false);
								}
							}
						}
						break;

				}
			}
			return;
			Raise_MenuEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////BeforeAction = True
			if ((BusinessObjectInfo.BeforeAction == true)) {
				switch (BusinessObjectInfo.EventType) {
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
						////33
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
						////34
						break;
					//                If Add_oPurchaseOrders(2) = False Then
					//                    BubbleEvent = False
					//                    Exit Sub
					//                Else
					//                    Call Delete_EmptyRow
					//                End If

					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
						////35
						if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
							//                    If Update_oPurchaseOrders(2) = False Then
							//                        oLast_Mode = 0
							//                        BubbleEvent = False
							//                        Exit Sub
							//                    Else
							//                        oLast_Mode = 0
							//                        Call Delete_EmptyRow
							//                    End If
						////취소
						} else if (oLast_Mode == 101) {
							//                    If Cancel_oPurchaseOrders(2) = False Then
							//                        oLast_Mode = 0
							//                        BubbleEvent = False
							//                        Exit Sub
							//                    Else
							//                        oLast_Mode = 0
							//                    End If
						////닫기
						} else if (oLast_Mode == 102) {
							//                    If Close_oPurchaseOrders(2) = False Then
							//                        oLast_Mode = 0
							//                        BubbleEvent = False
							//                        Exit Sub
							//                    Else
							//                        oLast_Mode = 0
							//                    End If
						}
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
						////36
						break;
				}
			////BeforeAction = False
			} else if ((BusinessObjectInfo.BeforeAction == false)) {
				switch (BusinessObjectInfo.EventType) {
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
						////33
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
						////34
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
						////35
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
						////36
						break;
				}
			}
			return;
			Raise_FormDataEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if ((eventInfo.BeforeAction == true)) {
				////작업
			} else if ((eventInfo.BeforeAction == false)) {
				////작업
			}
			return;
			Raise_RightClickEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		private void CreateItems()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////디비데이터 소스 개체 할당
			//    oForm01.DataSources.DataTables.Add ("PS_PP005")

			oDS_PS_PP005H = oForm01.DataSources.DBDataSources("@PS_PP005H");

			//// 메트릭스 개체 할당
			oMat01 = oForm01.Items.Item("Mat01").Specific;


			// 화면 호출시 등록일자에 오늘날짜 뿌려주기
			//
			oForm01.DataSources.UserDataSources.Add("UnWeight", SAPbouiCOM.BoDataType.dt_PERCENT, 10);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("UnWeight").Specific.DataBind.SetBound(true, "", "UnWeight");


			oForm01.DataSources.UserDataSources.Add("BaseChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("BaseChk").Specific.DataBind.SetBound(true, "", "BaseChk");
			//UPGRADE_WARNING: oForm01.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("BaseChk").Specific.Checked = false;

			oForm01.DataSources.UserDataSources.Add("ConvChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("ConvChk").Specific.DataBind.SetBound(true, "", "ConvChk");
			//UPGRADE_WARNING: oForm01.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("ConvChk").Specific.Checked = false;

			//   Call oForm01.DataSources.UserDataSources.Add("InDateFr", dt_DATE, 8)
			//   Call oForm01.DataSources.UserDataSources.Add("InDateTo", dt_DATE, 8)
			//   Call oForm01.Items("InDateFr").Specific.DataBind.SetBound(True, "", "InDateFr")
			//   Call oForm01.Items("InDateTo").Specific.DataBind.SetBound(True, "", "InDateTo")
			//
			//   oForm01.Items("InDateFr").Specific.Value = Format(Now, "yyyymmdd")
			//   oForm01.Items("InDateTo").Specific.Value = Format(Now, "yyyymmdd")


			//    oDS_PS_PP005H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")

			return;
			CreateItems_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void ComboBox_Setting()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////콤보에 기본값설정
			SAPbouiCOM.ComboBox oCombo = null;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;

			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			//// 대분류
			oCombo = oForm01.Items.Item("ItmBSort").Specific;
			sQry = "SELECT Code,Name From [@PSH_ITMBSORT] Order by Code";
			oRecordSet01.DoQuery(sQry);
			while (!(oRecordSet01.EoF)) {
				oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
				oRecordSet01.MoveNext();
			}

			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			return;
			ComboBox_Setting_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			MDC_Com.MDC_GF_Message(ref "ComboBox_Setting_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void CF_ChooseFromList()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////ChooseFromList 설정
			return;
			PS_PP005_CF_ChooseFromList_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "CF_ChooseFromList_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void FormItemEnabled()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			//    If (oForm01.Mode = fm_ADD_MODE) Then
			//        oForm01.Items("DocNum").Enabled = False
			//        oForm01.Items("DocDate").Enabled = True
			//    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
			//        oForm01.Items("DocNum").Enabled = True
			//        oForm01.Items("DocDate").Enabled = True
			//    ElseIf (oForm01.Mode = fm_OK_MODE) Then
			//        oForm01.Items("DocNum").Enabled = False
			//        oForm01.Items("DocDate").Enabled = False
			//    End If
			return;
			FormItemEnabled_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "FormItemEnabled_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void FormClear()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			string DocNum = null;
			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DocNum = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PS_PP005'", ref "");
			if (Convert.ToDouble(DocNum) == 0) {
				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oForm01.Items.Item("DocNum").Specific.VALUE = 1;
			} else {
				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oForm01.Items.Item("DocNum").Specific.VALUE = DocNum;
			}
			return;
			FormClear_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "FormClear_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}


		public void Add_MatrixRow(int oRow, ref bool RowIserted = false)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			//    If RowIserted = False Then '//행추가여부
			//        oDS_PS_PP005H.InsertRecord (oRow)
			//    End If
			//    oMat01.AddRow
			//    oDS_PS_PP005H.Offset = oRow
			//    oDS_PS_PP005H.setValue "U_DocNum", oRow, oRow + 1
			//    oDS_PS_PP005H.setValue "U_Chk", oRow, oRow + 1
			//    oDS_PS_PP005H.setValue "U_ItemCod1", oRow, oRow + 1
			//    oDS_PS_PP005H.setValue "U_ItemNam1", oRow, oRow + 1
			//    oDS_PS_PP005H.setValue "U_ItemCod2", oRow, oRow + 1
			//'    oDS_PS_PP005H.setValue "U_ItmMsort", oRow, oRow + 1
			///    oDS_OITM.setValue "U_ItmMsort", oRow, oRow + 1
			//    oDS_PS_PP005H.setValue "U_ItemNam2", oRow, oRow + 1
			//    oDS_PS_PP005H.setValue "U_InDate", oRow, oRow + 1
			//    oDS_PS_PP005H.setValue "U_MoDate", oRow, oRow + 1
			oMat01.LoadFromDataSource();
			return;
			Add_MatrixRow_Error:
			//'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Add_MatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

// 코드에 대한 품명을 뿌려주기 위한 선언부분

		private void FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			short ErrNum = 0;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;
			int sRow = 0;
			string sSeq = null;

			SAPbouiCOM.ComboBox oCombo = null;

			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			sRow = oRow;

			switch (oUID) {
				case "ItemCod1":
					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sQry = "Select ItemName, ItmMSort = U_ItmMsort  From OITM Where ItemCode = '" + Strings.Trim(oForm01.Items.Item("ItemCod1").Specific.VALUE) + "'";
					oRecordSet01.DoQuery(sQry);

					//UPGRADE_WARNING: oForm01.Items(ItemNam1).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oForm01.Items.Item("ItemNam1").Specific.String = Strings.Trim(oRecordSet01.Fields.Item("ItemName").Value);

					oCombo = oForm01.Items.Item("ItmMSort").Specific;
					oCombo.Select(Strings.Trim(oRecordSet01.Fields.Item("ItmMSort").Value), SAPbouiCOM.BoSearchKey.psk_ByValue);
					break;
				//oDS_PS_PP005H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")

				case "ItemCod2":
					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sQry = "Select ItemName  From OITM Where ItemCode = '" + Strings.Trim(oForm01.Items.Item("ItemCod2").Specific.VALUE) + "'";
					oRecordSet01.DoQuery(sQry);

					//UPGRADE_WARNING: oForm01.Items(ItemNam2).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oForm01.Items.Item("ItemNam2").Specific.String = Strings.Trim(oRecordSet01.Fields.Item("ItemName").Value);
					break;

			}

			// Matrix 필드에 질의 응답 창 띄워주기
			switch (oCol) {
				case "ItemCod2":
					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sQry = "Select ItemName  From OITM Where ItemCode = '" + Strings.Trim(oMat01.Columns.Item("ItemCod2").Cells.Item(oRow).Specific.VALUE) + "'";
					oRecordSet01.DoQuery(sQry);

					//UPGRADE_WARNING: oMat01.Columns(ItemNam2).Cells(oRow).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oMat01.Columns.Item("ItemNam2").Cells.Item(oRow).Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item("ItemName").Value);
					break;
			}

			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			return;
			FlushToItemValue_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			if (ErrNum == 1) {
				MDC_Com.MDC_GF_Message(ref "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
			} else {
				MDC_Com.MDC_GF_Message(ref "FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
			}
		}

		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement

			short ErrNum = 0;
			string DocNum = null;

			ErrNum = 0;

			// 저장버튼 클릭시 필수입력 필드에 값이 있는지를 Check 한다.

			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			switch (true) {
				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("ItemCod1").Specific.VALUE)):
					ErrNum = 1;
					goto HeaderSpaceLineDel_Error;
					break;
				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("ItemNam1").Specific.VALUE)):
					ErrNum = 2;
					goto HeaderSpaceLineDel_Error;
					break;
				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("ItemCod2").Specific.VALUE)):
					ErrNum = 3;
					goto HeaderSpaceLineDel_Error;
					break;
				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("ItemNam2").Specific.VALUE)):
					ErrNum = 4;
					goto HeaderSpaceLineDel_Error;
					break;
			}

			functionReturnValue = true;
			return functionReturnValue;
			HeaderSpaceLineDel_Error:

			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

			if (ErrNum == 1) {
				MDC_Com.MDC_GF_Message(ref "제품코드는 필수입력사항입니다. 확인하세요.", ref "E");
			} else if (ErrNum == 2) {
				MDC_Com.MDC_GF_Message(ref "제품코드명은 필수입력사항입니다. 확인하세요.", ref "E");
			} else if (ErrNum == 3) {
				MDC_Com.MDC_GF_Message(ref "원자재코드는 필수입력사항입니다. 확인하세요.", ref "E");
			} else if (ErrNum == 4) {
				MDC_Com.MDC_GF_Message(ref "원자재명은 필수입력사항입니다. 확인하세요.", ref "E");
			} else {
				MDC_Com.MDC_GF_Message(ref "HeaderSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
			}
			functionReturnValue = false;
			return functionReturnValue;
		}

		private bool MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			short ErrNum = 0;
			SAPbobsCOM.Recordset oRecordSet = null;
			string sQry = null;

			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			ErrNum = 0;

			oMat01.FlushToDataSource();

			//// 라인
			if (oMat01.VisualRowCount == 0) {
				ErrNum = 1;
				goto MatrixSpaceLineDel_Error;
			}

			oMat01.LoadFromDataSource();

			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			functionReturnValue = true;
			return functionReturnValue;
			MatrixSpaceLineDel_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			if (ErrNum == 1) {
				MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
			} else {
				MDC_Com.MDC_GF_Message(ref "MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
			}
			functionReturnValue = false;
			return functionReturnValue;
		}

//Sub Delete_EmptyRow()
//On Error GoTo Delete_EmptyRow_Error
//    Dim i&
//
//    oMat01.FlushToDataSource
//
//    For i = 0 To oMat01.VisualRowCount - 1
//        If Trim(oDS_PS_PP005L.GetValue("U_ItemCode", i)) = "" Then
//            oDS_PS_PP005L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
//        End If
//    Next i
//
//    oMat01.LoadFromDataSource
//    Exit Sub
//'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Delete_EmptyRow_Error:
//    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
//End Sub
	}
}
