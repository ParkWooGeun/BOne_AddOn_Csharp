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
//	[System.Runtime.InteropServices.ProgId("ZPY505_NET.ZPY505")]
//	public class ZPY505
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY505.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 기부금명세 등록
//////  FormType       : 2000060505
//////  Create Date    : 2006.01.15
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//			//시스템코드 헤더
//		private SAPbouiCOM.DBDataSource oDS_ZPY505H;
//			//시스템코드 라인
//		private SAPbouiCOM.DBDataSource oDS_ZPY505L;
//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string Col_Last_Uid;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int Col_Last_Row;
//		private string oOLDCHK;

//		private void FormItemEnabled()
//		{
//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = true;
//				oForm.Items.Item("MSTCOD").Enabled = true;
//				oForm.Items.Item("MSTNAM").Enabled = true;
//				oForm.Items.Item("DocNum").Enabled = true;
//				oForm.Items.Item("ENDCHK").Enabled = true;
//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = true;
//				oForm.Items.Item("MSTCOD").Enabled = true;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = false;
//				oForm.Items.Item("ENDCHK").Enabled = true;
//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = false;
//				oForm.Items.Item("MSTCOD").Enabled = false;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = false;
//				//// 년마감된것은 비활성화
//				oOLDCHK = oDS_ZPY505H.GetValue("U_ENDCHK", 0);
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(U_ENDCHK, U_JOBYER, [ZPY509L], ' & oDS_ZPY505H.GetValue(U_JSNYER, 0) & ',  AND Code = ' & oDS_ZPY505H.GetValue(U_CLTCOD, 0) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (MDC_SetMod.Get_ReData(ref "U_ENDCHK", ref "U_JOBYER", ref "[@ZPY509L]", ref "'" + oDS_ZPY505H.GetValue("U_JSNYER", 0) + "'", ref " AND Code = '" + oDS_ZPY505H.GetValue("U_CLTCOD", 0) + "'") == "Y") {
//					oForm.Items.Item("ENDCHK").Enabled = false;
//				} else {
//					oForm.Items.Item("ENDCHK").Enabled = true;
//				}

//			}
//		}

//		private void FormClear()
//		{
//			int DocNum = 0;

//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocNum = MDC_SetMod.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'ZPY505'", ref "");

//			if (DocNum == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.String = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.String = DocNum;
//			}
//			FlushToItemValue("JSNYER");
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{
//			int i = 0;
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);
//			double TOTCNT = 0;
//			double TOTAMT = 0;

//			switch (oUID) {
//				case "JSNYER":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oUID).Specific.String))) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MDC_Globals.ZPAY_GBL_JSNYER.Value = oForm.Items.Item(oUID).Specific.String;
//					} else {
//						oDS_ZPY505H.SetValue("U_JSNYER", 0, MDC_Globals.ZPAY_GBL_JSNYER.Value);
//					}
//					oForm.Items.Item(oUID).Update();
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oDS_ZPY505H.SetValue("U_MSTCOD", 0, "");
//						oDS_ZPY505H.SetValue("U_MSTNAM", 0, "");
//						oDS_ZPY505H.SetValue("U_EmpID", 0, "");
//						oDS_ZPY505H.SetValue("U_CLTCOD", 0, "");
//					} else {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oDS_ZPY505H.SetValue("U_MSTCOD", 0, Strings.UCase(oForm.Items.Item(oUID).Specific.String));
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMast = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oDS_ZPY505H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//						oDS_ZPY505H.SetValue("U_EmpID", 0, oMast.EmpID);
//						oDS_ZPY505H.SetValue("U_CLTCOD", 0, oMast.CLTCOD);
//					}
//					oForm.Items.Item("MSTNAM").Update();
//					oForm.Items.Item("EmpID").Update();
//					oForm.Items.Item("CLTCOD").Update();
//					oForm.Items.Item(oUID).Update();
//					break;
//			}

//			////ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ/
//			switch (oUID) {
//				case "Col5":
//				case "Col6":
//					oMat1.FlushToDataSource();
//					for (i = 1; i <= oMat1.VisualRowCount; i++) {
//						TOTCNT = TOTCNT + Conversion.Val(oDS_ZPY505L.GetValue("U_GBUCNT", i - 1));
//						TOTAMT = TOTAMT + Conversion.Val(oDS_ZPY505L.GetValue("U_GBUAMT", i - 1));
//					}

//					oDS_ZPY505H.SetValue("U_TOTCNT", 0, Convert.ToString(TOTCNT));
//					oDS_ZPY505H.SetValue("U_TOTAMT", 0, Convert.ToString(TOTAMT));
//					oForm.Items.Item("TOTCNT").Update();
//					oForm.Items.Item("TOTAMT").Update();

//					oDS_ZPY505L.Offset = oRow - 1;
//					break;
//				//            oMat1.SetLineData oRow
//				case "Col8":
//					oDS_ZPY505L.Offset = oRow - 1;
//					//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oDS_ZPY505L.SetValue("U_FAMNAM", oRow - 1, oMat1.Columns.Item(oUID).Cells.Item(oRow).Specific.Value);
//					Display_GibuMan(ref oRow - 1);
//					oMat1.SetLineData(oRow);
//					break;
//				case "Col1":
//					oMat1.FlushToDataSource();
//					oDS_ZPY505L.Offset = oRow - 1;

//					if (oRow == oMat1.RowCount & !string.IsNullOrEmpty(Strings.Trim(oDS_ZPY505L.GetValue("U_GBUYMM", oRow - 1)))) {
//						Matrix_AddRow(oRow);
//						oMat1.Columns.Item("Col1").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					}
//					break;
//			}
//		}

//		private void Display_GibuMan(ref int sRow)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string MSTCOD = null;
//			string JSNYER = null;
//			string FAMNAM = null;


//			JSNYER = oDS_ZPY505H.GetValue("U_JSNYER", 0);
//			MSTCOD = oDS_ZPY505H.GetValue("U_MSTCOD", 0);
//			FAMNAM = oDS_ZPY505L.GetValue("U_FAMNAM", sRow);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//// 기부자명입력시 소득자료등록화면의 부양가족명세중 본인에 해당하는 나머지 정보 표시해줌.
//			sQry = "SELECT  T0.U_FAMNAM AS FAMNAM, T0.U_CHKINT AS INTGBN, T0.U_FAMPER AS PERNBR, ";
//			sQry = sQry + " CASE T0.U_CHKCOD WHEN '0' THEN '1' WHEN '3' THEN '2' WHEN '4' THEN '3' ELSE '' END AS GWANGE ";
//			sQry = sQry + " FROM [@ZPY501L] T0 INNER JOIN [@ZPY501H] T1 ON T0.DocEntry = T1.DocEntry";
//			sQry = sQry + " WHERE T1.U_JSNYER = '" + Strings.Trim(JSNYER) + "'";
//			sQry = sQry + " AND   T1.U_MSTCOD = '" + Strings.Trim(MSTCOD) + "'";
//			sQry = sQry + " AND   T0.U_FAMNAM = '" + Strings.Trim(FAMNAM) + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				oDS_ZPY505L.SetValue("U_INTGBN", sRow, oRecordSet.Fields.Item("INTGBN").Value);
//				oDS_ZPY505L.SetValue("U_PERNBR", sRow, oRecordSet.Fields.Item("PERNBR").Value);
//				oDS_ZPY505L.SetValue("U_GWANGE", sRow, oRecordSet.Fields.Item("GWANGE").Value);
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Display_GibuMan Error:" + Err().Number + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private bool MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//저장할 데이터의 유효성을 점검한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int iRow = 0;
//			int kRow = 0;
//			short ErrNum = 0;
//			string Chk_Data = null;
//			//    Dim CHAAMT      As Double
//			string GovidChk = null;
//			ErrNum = 0;
//			/// 헤더부분 체크
//			switch (true) {
//				case Strings.Len(Strings.Trim(oDS_ZPY505H.GetValue("U_JSNYER", 0))) != 4:
//					ErrNum = 4;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY505H.GetValue("U_MSTCOD", 0)):
//					ErrNum = 5;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY505H.GetValue("U_CLTCOD", 0)):
//					ErrNum = 17;
//					goto Error_Message;
//					break;

//			}
//			/// 주민번호체크유무
//			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			GovidChk = MDC_SetMod.Get_ReData(ref "ISNULL(T0.U_GovIDChk,'N')", ref "T1.U_MSTCOD", ref "[@PH_PY005A] T0 INNER JOIN [@PH_PY001A] T1 ON T0.CODE = T1.U_CLTCOD", ref "'" + Strings.Trim(oDS_ZPY505H.GetValue("U_MSTCOD", 0)) + "'", ref "");

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oMat1.FlushToDataSource();

//			//// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
//			if (oMat1.RowCount == 1) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
//			////oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
//			for (iRow = 0; iRow <= oMat1.VisualRowCount - 2; iRow++) {
//				oDS_ZPY505L.Offset = iRow;
//				//        CHAAMT = Val(oDS_ZPY505L.GetValue("U_BEFAMT", irow)) + Val(oDS_ZPY505L.GetValue("U_CURAMT", irow)) + Val(oDS_ZPY505L.GetValue("U_CHAAMT", irow))
//				if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY505L.GetValue("U_GBUYMM", iRow)))) {
//					ErrNum = 2;
//					oMat1.Columns.Item("Col1").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (MDC_SetMod.ChkYearMonth(ref oDS_ZPY505L.GetValue("U_GBUYMM", iRow)) == false) {
//					ErrNum = 7;
//					oMat1.Columns.Item("Col1").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY505L.GetValue("U_GBUNAM", iRow)))) {
//					ErrNum = 6;
//					oMat1.Columns.Item("Col3").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (Conversion.Val(oDS_ZPY505L.GetValue("U_GBUCNT", iRow)) == 0) {
//					ErrNum = 8;
//					oMat1.Columns.Item("Col5").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (Conversion.Val(oDS_ZPY505L.GetValue("U_GBUAMT", iRow)) == 0) {
//					ErrNum = 9;
//					oMat1.Columns.Item("Col6").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY505L.GetValue("U_FAMNAM", iRow)))) {
//					ErrNum = 10;
//					oMat1.Columns.Item("Col8").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY505L.GetValue("U_GWANGE", iRow)))) {
//					ErrNum = 11;
//					oMat1.Columns.Item("Col9").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY505L.GetValue("U_INTGBN", iRow)))) {
//					ErrNum = 12;
//					oMat1.Columns.Item("Col10").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY505L.GetValue("U_PERNBR", iRow)))) {
//					ErrNum = 13;
//					oMat1.Columns.Item("Col11").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//					//        ElseIf Trim$(oDS_ZPY505L.GetValue("U_GBUCOD", irow)) = "31" And CHAAMT <> Val(oDS_ZPY505L.GetValue("U_GBUAMT", irow)) Then
//					//            ErrNum = 14
//					//            oMat1.Columns("Col14").Cells(irow + 1).CLICK ct_Regular
//					//            GoTo Error_Message
//					//        ElseIf Trim$(oDS_ZPY505L.GetValue("U_GBUCOD", irow)) <> "31" And CHAAMT <> 0 Then
//					//            ErrNum = 15
//					//            oMat1.Columns("Col14").Cells(irow + 1).CLICK ct_Regular
//					//            GoTo Error_Message
//				} else {
//					//// 6.주민번호 오류 체크
//					if (Strings.Trim(GovidChk) == "Y" & Strings.Len(oDS_ZPY505L.GetValue("U_PERNBR", iRow)) > 0) {
//						if (MDC_Com.GovIDCheck(ref oDS_ZPY505L.GetValue("U_PERNBR", iRow)) == false) {
//							ErrNum = 16;
//							oMat1.Columns.Item("Col11").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							goto Error_Message;
//						}
//					}
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//중복체크작업
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					Chk_Data = Strings.Trim(oDS_ZPY505L.GetValue("U_GBUYMM", iRow)) + Strings.Trim(oDS_ZPY505L.GetValue("U_GBUCOD", iRow)) + Strings.Trim(oDS_ZPY505L.GetValue("U_GBUNBR", iRow)) + Strings.Trim(oDS_ZPY505L.GetValue("U_PERNBR", iRow));
//					for (kRow = iRow + 1; kRow <= oMat1.VisualRowCount - 2; kRow++) {
//						oDS_ZPY505L.Offset = kRow;
//						if (Strings.Trim(Chk_Data) == Strings.Trim(oDS_ZPY505L.GetValue("U_GBUYMM", kRow)) + Strings.Trim(oDS_ZPY505L.GetValue("U_GBUCOD", kRow)) + Strings.Trim(oDS_ZPY505L.GetValue("U_GBUNBR", kRow)) + Strings.Trim(oDS_ZPY505L.GetValue("U_PERNBR", kRow))) {
//							ErrNum = 3;
//							oMat1.Columns.Item("Col1").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							goto Error_Message;
//						}
//					}
//				}
//			}

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//			////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oDS_ZPY505L.RemoveRecord(oDS_ZPY505L.Size - 1);
//			//// Mat1에 마지막라인(빈라인) 삭제

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부연월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부연월의 기부코드, 기부처 사업자번호가 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부처 상호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부연월을 확인하세요. Ex)2006년1월->200601", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 8) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부금건수가 0입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 9) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부금금액가 0입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 10) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부자 성명은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 11) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부자 관계코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 12) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부자 내외국인구분은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 13) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기부자 주민등록번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				//    ElseIf ErrNum = 14 Then
//				//        Sbo_Application.StatusBar.SetText "31-공익법인기부금신탁일 경우 기부금액과 (이월액잔액+해당과세기간공제액+이월액)이 일치하지 않습니다.", bmt_Short, smt_Error
//				//    ElseIf ErrNum = 15 Then
//				//        Sbo_Application.StatusBar.SetText "31-공익법인기부금신탁가 아닐경우 (이월액잔액+해당과세기간공제액+이월액)는 입력하지 않습니다.", bmt_Short, smt_Error
//			} else if (ErrNum == 16) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("주민등록번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 17) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("MatrixSpaceLineDel Error:" + Err().Number + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
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
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//									//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (MDC_SetMod.Value_ChkYn(ref "[@ZPY505H]", ref "U_JSNYER", ref "'" + oForm.Items.Item("JSNYER").Specific.String + "'", ref " AND U_MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.String + "'") == false) {
//										MDC_Globals.Sbo_Application.StatusBar.SetText("이미 저장되어져 있는 헤더의 내용과 일치합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//										BubbleEvent = false;
//										return;
//									}
//								}

//								if (Strings.Trim(oDS_ZPY505H.GetValue("U_ENDCHK", 0)) == "Y" & Strings.Trim(oOLDCHK) == "Y") {
//									MDC_Globals.Sbo_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//									BubbleEvent = false;
//									return;
//								} else if (MatrixSpaceLineDel() == false) {
//									BubbleEvent = false;
//								}
//							}
//						/// ChooseBtn사원리스트
//						} else if (pval.ItemUID == "CBtn1" & oForm.Items.Item("MSTCOD").Enabled == true) {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						} else if (pval.ItemUID == "Btn1" & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)) {
//							BeforeBalance();
//						}
//					} else {
//						if (pval.ItemUID == "1" & pval.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							MDC_Globals.Sbo_Application.ActivateMenuItem("1282");
//						}
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
//					if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == true & Last_Item == "Mat1" & Col_Last_Uid == "Col1" & Col_Last_Row > 0 & (Col_Last_Uid != pval.ColUID | Col_Last_Row != pval.Row) & pval.ItemUID != "1000001" & pval.ItemUID != "2") {
//						if (Col_Last_Row > oMat1.VisualRowCount) {
//							return;
//						}
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUID == "MSTCOD" | pval.ItemUID == "JSNYER")) {
//						FlushToItemValue(pval.ItemUID);
//					} else if (pval.BeforeAction == false & pval.ItemChanged == true & pval.ItemUID == "Mat1" & (pval.ColUID == "Col5" | pval.ColUID == "Col6" | pval.ColUID == "Col1" | pval.ColUID == "Col8")) {
//						FlushToItemValue(pval.ColUID, ref pval.Row);
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
//							if (oMat1.RowCount > 0) {
//								oMat1.Columns.Item("Col1").Cells.Item(oMat1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								BubbleEvent = false;
//							}
//						}
//					} else if (pval.BeforeAction == true & pval.ColUID == "Col1" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(Strings.Trim(oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.String))) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("기부연월은 필수입니다. 입력하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					} else if (pval.BeforeAction == true & pval.ColUID == "Col3" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(Strings.Trim(oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.String))) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("기부처상호는 필수입니다. 입력하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					} else if (pval.BeforeAction == true & pval.ColUID == "Col4" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(Strings.Trim(oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.String))) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("기부처의 사업자(주민)번호는 필수입니다. 입력하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						} else {
//							/// 사업자번호 체크
//							//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (Strings.Len(oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.String) <= 12) {
//								//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (MDC_SetMod.TaxNoCheck(oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.String) == false) {
//									MDC_Globals.Sbo_Application.StatusBar.SetText("사업자번호가 틀립니다. 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//									BubbleEvent = false;
//								}
//							}
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
//						//UPGRADE_NOTE: oDS_ZPY505H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY505H = null;
//						//UPGRADE_NOTE: oDS_ZPY505L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY505L = null;
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

//			}

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Raise_FormItemEvent_Error:", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
////*******************************************************************
////// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
////*******************************************************************
//		private void Matrix_AddRow(int oRow, ref bool Insert_YN = false)
//		{
//			if (Insert_YN == false) {
//				oDS_ZPY505L.InsertRecord((oRow));
//			}
//			oDS_ZPY505L.Offset = oRow;
//			oDS_ZPY505L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//			oDS_ZPY505L.SetValue("U_GBUYMM", oRow, "");
//			oDS_ZPY505L.SetValue("U_GBUCOD", oRow, "");
//			oDS_ZPY505L.SetValue("U_GBUNAM", oRow, "");
//			oDS_ZPY505L.SetValue("U_GBUNBR", oRow, "");
//			oDS_ZPY505L.SetValue("U_GBUSEQ", oRow, "");
//			oDS_ZPY505L.SetValue("U_GBUCNT", oRow, "");
//			oDS_ZPY505L.SetValue("U_GBUAMT", oRow, "");
//			oDS_ZPY505L.SetValue("U_GWANGE", oRow, "");
//			oDS_ZPY505L.SetValue("U_FAMNAM", oRow, "");
//			oDS_ZPY505L.SetValue("U_INTGBN", oRow, "");
//			oDS_ZPY505L.SetValue("U_PERNBR", oRow, "");
//			oDS_ZPY505L.SetValue("U_BEFAMT", oRow, "");
//			oDS_ZPY505L.SetValue("U_CURAMT", oRow, "");
//			oDS_ZPY505L.SetValue("U_CHAAMT", oRow, "");
//			oMat1.LoadFromDataSource();
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
//						/// 제거
//						if (Strings.Trim(oDS_ZPY505H.GetValue("U_ENDCHK", 0)) == "Y") {
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
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@ZPY505H", ref "DocNum");
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
//						break;
//					case "1293":
//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//							////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							for (i = 0; i <= oMat1.VisualRowCount - 1; i++) {
//								//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat1.Columns.Item("Col0").Cells.Item(i + 1).Specific.Value = i + 1;
//							}

//							oMat1.FlushToDataSource();
//							oDS_ZPY505L.RemoveRecord(oDS_ZPY505L.Size - 1);
//							//// Mat1에 마지막라인(빈라인) 삭제
//							oMat1.Clear();
//							oMat1.LoadFromDataSource();
//						}
//						FlushToItemValue("Col5", ref 1);
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
//		public void LoadForm(ref string JSNYER = "", ref string MSTCOD = "", ref string CLTCOD = "")
//		{
//			//Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY505.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY505_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY505");
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

//			oForm.EnableMenu(("1293"), true);
//			/// 행삭제
//			oForm.EnableMenu(("1283"), true);
//			/// 제거
//			oForm.EnableMenu(("1284"), false);
//			/// 취소

//			oForm.Freeze(false);
//			oForm.Update();
//			//oForm.Visible = True

//			if (!string.IsNullOrEmpty(JSNYER)) {
//				ShowSource(ref JSNYER, ref MSTCOD, ref CLTCOD);
//			}

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

//		private void BeforeBalance()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string JSNYER = null;
//			string MSTCOD = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			int ErrNum = 0;
//			int iRow = 0;
//			bool DupChk = false;

//			JSNYER = Strings.Trim(oDS_ZPY505H.GetValue("U_JSNYER", 0));
//			MSTCOD = Strings.Trim(oDS_ZPY505H.GetValue("U_MSTCOD", 0));

//			if (string.IsNullOrEmpty(JSNYER) | string.IsNullOrEmpty(MSTCOD)) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			oMat1.FlushToDataSource();
//			DupChk = false;
//			for (iRow = 0; iRow <= oDS_ZPY505L.Size - 1; iRow++) {
//				if (Conversion.Val(oDS_ZPY505L.GetValue("U_BEFDOC", iRow)) != 0) {
//					DupChk = true;
//					break; // TODO: might not be correct. Was : Exit For
//				}
//			}
//			if (DupChk == true) {
//				ErrNum = 3;
//				goto Error_Message;
//			}

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			sQry = "EXEC ZPY505_1 '" + JSNYER + "', '" + MSTCOD + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 2;
//				goto Error_Message;
//			}

//			iRow = oDS_ZPY505L.Size - 1;
//			while (!(oRecordSet.EoF)) {
//				if (iRow == oDS_ZPY505L.Size) {
//					oDS_ZPY505L.InsertRecord((iRow));
//				}
//				oDS_ZPY505L.Offset = iRow;
//				oDS_ZPY505L.SetValue("U_LINENUM", iRow, Convert.ToString(iRow + 1));
//				oDS_ZPY505L.SetValue("U_GBUYMM", iRow, oRecordSet.Fields.Item("U_GBUYMM").Value);
//				oDS_ZPY505L.SetValue("U_GBUCOD", iRow, oRecordSet.Fields.Item("U_GBUCOD").Value);
//				oDS_ZPY505L.SetValue("U_GBUNAM", iRow, oRecordSet.Fields.Item("U_GBUNAM").Value);
//				oDS_ZPY505L.SetValue("U_GBUCNT", iRow, oRecordSet.Fields.Item("U_GBUCNT").Value);
//				oDS_ZPY505L.SetValue("U_GBUAMT", iRow, oRecordSet.Fields.Item("U_GBUAMT").Value);
//				oDS_ZPY505L.SetValue("U_FAMNAM", iRow, oRecordSet.Fields.Item("U_FAMNAM").Value);
//				oDS_ZPY505L.SetValue("U_GWANGE", iRow, oRecordSet.Fields.Item("U_GWANGE").Value);
//				oDS_ZPY505L.SetValue("U_INTGBN", iRow, oRecordSet.Fields.Item("U_INTGBN").Value);
//				oDS_ZPY505L.SetValue("U_PERNBR", iRow, oRecordSet.Fields.Item("U_PERNBR").Value);
//				oDS_ZPY505L.SetValue("U_BEFAMT", iRow, oRecordSet.Fields.Item("U_CHAAMT").Value);

//				oRecordSet.MoveNext();
//				iRow = iRow + 1;
//			}
//			Matrix_AddRow(iRow);

//			MDC_Globals.Sbo_Application.StatusBar.SetText("이전년도 기부금 이월금액 가져오기를 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Error_Message:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도와 사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("이전년도에서 이월된 기부금 내역이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("이미 이월금액 가져오기를 실행한 상태입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("BeforeBalance Error : " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}

//		}

//		private void ShowSource(ref string JSNYER, ref string MSTCOD, ref string CLTCOD)
//		{
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string DocEntry = null;
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			sQry = "SELECT DocNum FROM [@ZPY505H]";
//			sQry = sQry + "   WHERE U_JSNYER = N'" + JSNYER + "'";
//			sQry = sQry + "   AND   U_MSTCOD = N'" + MSTCOD + "'";
//			sQry = sQry + "   AND   U_CLTCOD = N'" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				while (!(oRecordSet.EoF)) {
//					//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					DocEntry = oRecordSet.Fields.Item(0).Value;
//					oRecordSet.MoveNext();
//				}
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("JSNYER").Specific.Value = JSNYER;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("MSTCOD").Specific.String = MSTCOD;
//				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("CLTCOD").Specific.Select(CLTCOD, SAPbouiCOM.BoSearchKey.psk_ByValue);
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.Value = DocEntry;

//				oForm.Items.Item("DocNum").Update();
//				oMat1.LoadFromDataSource();
//				oForm.Update();
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//				MDC_Globals.Sbo_Application.ActivateMenuItem("1282");

//				oDS_ZPY505H.SetValue("U_JSNYER", 0, JSNYER);
//				oDS_ZPY505H.SetValue("U_MSTCOD", 0, MSTCOD);
//				oDS_ZPY505H.SetValue("U_CLTCOD", 0, CLTCOD);
//				//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oMast = MDC_SetMod.Get_EmpID_InFo(ref MSTCOD);
//				oDS_ZPY505H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//				oDS_ZPY505H.SetValue("U_EmpID", 0, oMast.EmpID);

//				oForm.Update();

//				MDC_Globals.Sbo_Application.SendKeys("{TAB}");
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//		}
////*******************************************************************
////
////*******************************************************************
//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.CheckBox oCheck = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////디비데이터 소스 개체 할당
//			oDS_ZPY505H = oForm.DataSources.DBDataSources("@ZPY505H");
//			oDS_ZPY505L = oForm.DataSources.DBDataSources("@ZPY505L");

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			//// 사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}

//			//// 기부코드
//			oColumn = oMat1.Columns.Item("Col2");
//			oColumn.ValidValues.Add("10", "법정기부금");
//			oColumn.ValidValues.Add("20", "정치자금");
//			//oColumn.ValidValues.Add "21", "문화예술진흥기금" '2011년
//			oColumn.ValidValues.Add("30", "특례기부금");
//			oColumn.ValidValues.Add("31", "공익법인기부금신탁");
//			oColumn.ValidValues.Add("40", "지정기부금(종교단체외)");
//			oColumn.ValidValues.Add("41", "지정기부금(종교단체)");
//			oColumn.ValidValues.Add("42", "우리사주조합 기부금");
//			oColumn.ValidValues.Add("50", "공제제외기부금");

//			//// 관계코드
//			oColumn = oMat1.Columns.Item("Col9");
//			oColumn.ValidValues.Add("1", "거주자(본인)");
//			oColumn.ValidValues.Add("2", "배우자");
//			oColumn.ValidValues.Add("3", "직계비속");
//			oColumn.ValidValues.Add("4", "직계존속");
//			oColumn.ValidValues.Add("5", "형제자매");
//			oColumn.ValidValues.Add("6", "그 외");

//			//// 내외국인
//			oColumn = oMat1.Columns.Item("Col10");
//			oColumn.ValidValues.Add("1", "내국인");
//			oColumn.ValidValues.Add("9", "외국인");

//			oColumn = oMat1.Columns.Item("Col19");
//			oColumn.ValOff = "N";
//			oColumn.ValOn = "Y";

//			//// 영수증일련번호(2008년 제외)
//			oMat1.Columns.Item("Col7").Visible = false;

//			/// Check 버튼
//			oCheck = oForm.Items.Item("ENDCHK").Specific;
//			oCheck.ValOff = "N";
//			oCheck.ValOn = "Y";

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
//	}
//}
