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
//	[System.Runtime.InteropServices.ProgId("ZPY502_NET.ZPY502")]
//	public class ZPY502
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY502.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 종(전)근무지 등록
//////  FormType       : 2010110502
//////  Create Date    : 2006.01.15
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//			//시스템코드 헤더
//		private SAPbouiCOM.DBDataSource oDS_ZPY502H;
//			//시스템코드 라인
//		private SAPbouiCOM.DBDataSource oDS_ZPY502L;
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
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("DocNum").Enabled = true;
//				oForm.Items.Item("ENDCHK").Enabled = true;
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = true;
//				oForm.Items.Item("MSTCOD").Enabled = true;
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = true;
//				oForm.Items.Item("ENDCHK").Enabled = true;
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//				oForm.Items.Item("JSNYER").Enabled = false;
//				oForm.Items.Item("MSTCOD").Enabled = false;
//				oForm.Items.Item("MSTNAM").Enabled = false;
//				oForm.Items.Item("CLTCOD").Enabled = false;
//				oForm.Items.Item("DocNum").Enabled = false;
//				//// 급여월마감된것은 비활성화
//				oOLDCHK = oDS_ZPY502H.GetValue("U_ENDCHK", 0);
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(U_ENDCHK, U_JOBYER, [ZPY509L], ' & oDS_ZPY502H.GetValue(U_JSNYER, 0) & ',  AND Code = ' & oDS_ZPY502H.GetValue(U_CLTCOD, 0) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (MDC_SetMod.Get_ReData(ref "U_ENDCHK", ref "U_JOBYER", ref "[@ZPY509L]", ref "'" + oDS_ZPY502H.GetValue("U_JSNYER", 0) + "'", ref " AND Code = '" + oDS_ZPY502H.GetValue("U_CLTCOD", 0) + "'") == "Y") {
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
//			DocNum = MDC_SetMod.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'ZPY502'", ref "");

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
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);
//			double TOTAMT = 0;
//			string JSNYER = null;
//			switch (oUID) {
//				case "JSNYER":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oUID).Specific.String))) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MDC_Globals.ZPAY_GBL_JSNYER.Value = oForm.Items.Item(oUID).Specific.String;
//					} else {
//						oDS_ZPY502H.SetValue("U_JSNYER", 0, MDC_Globals.ZPAY_GBL_JSNYER.Value);
//					}
//					oForm.Items.Item(oUID).Update();
//					Matrix_TitleSetting();
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						oDS_ZPY502H.SetValue("U_MSTCOD", 0, "");
//						oDS_ZPY502H.SetValue("U_MSTNAM", 0, "");
//						oDS_ZPY502H.SetValue("U_EmpID", 0, "");
//						oDS_ZPY502H.SetValue("U_CLTCOD", 0, "");
//					} else {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oDS_ZPY502H.SetValue("U_MSTCOD", 0, Strings.UCase(oForm.Items.Item(oUID).Specific.String));
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMast = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oDS_ZPY502H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//						oDS_ZPY502H.SetValue("U_EmpID", 0, oMast.EmpID);
//						oDS_ZPY502H.SetValue("U_CLTCOD", 0, oMast.CLTCOD);
//					}
//					oForm.Items.Item("MSTNAM").Update();
//					oForm.Items.Item("EmpID").Update();
//					oForm.Items.Item("CLTCOD").Update();
//					oForm.Items.Item(oUID).Update();
//					break;
//			}
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			if (Strings.Left(oUID, 3) == "Col") {
//				oMat1.FlushToDataSource();

//				oDS_ZPY502L.Offset = oRow - 1;

//				switch (oUID) {
//					case "Col1":
//					case "Col2":
//					case "Col3":
//						//            oMat1.SetLineData oRow
//						//
//						if (oRow == oMat1.RowCount & !string.IsNullOrEmpty(Strings.Trim(oDS_ZPY502L.GetValue("U_JONNAM", oRow - 1)))) {
//							Matrix_AddRow(oRow);
//							oMat1.Columns.Item("Col1").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						}
//						break;
//					case "Col5":
//					case "Col6":
//					case "Col7":
//					case "Col13":
//					case "Col17":
//					case "Col20":
//					case "Col23":
//					case "Col24":
//					case "Col25":
//					case "Col26":
//					case "Col27":
//					case "Col28":
//					case "Col29":
//					case "Col30":
//					case "Col31":
//					case "Col32":
//					case "Col33":
//					case "Col34":
//					case "Col35":
//					case "Col36":
//					case "Col37":
//					case "Col38":
//					case "Col39":
//					case "Col40":
//					case "Col41":
//					case "Col42":
//					case "Col43":
//					case "Col44":
//					case "Col45":
//					case "Col46":
//					case "Col47":
//					case "Col48":
//					case "Col50":

//						JSNYER = Strings.Trim(oDS_ZPY502H.GetValue("U_JSNYER", 0));
//						TOTAMT = 0;
//						if (JSNYER <= "2008") {
//							TOTAMT = Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT1", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT2", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT3", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBU3", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT4", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT5", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT6", oRow - 1));

//						} else if (JSNYER == "2009") {
//							TOTAMT = Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT1", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT2", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT3", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBU3", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT4", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT5", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT6", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTG01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH05", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH06", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH07", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH08", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH09", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH10", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH11", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH12", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH13", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTI01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTK01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTM01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTM02", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTM03", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTO01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTQ01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTS01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTT01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTX01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY02", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY03", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY20", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTZ01", oRow - 1));
//						} else if (JSNYER >= "2010") {
//							TOTAMT = Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT1", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT2", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT3", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBU3", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT4", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT5", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JONBT6", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTG01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH05", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH06", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH07", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH08", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH09", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH10", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH11", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH12", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTH13", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTI01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTK01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTM01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTM02", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTM03", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTO01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTQ01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTS01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTT01", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY02", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY03", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY21", oRow - 1)) + Conversion.Val(oDS_ZPY502L.GetValue("U_JBTZ01", oRow - 1));
//						}

//						oDS_ZPY502L.SetValue("U_JBTTOT", oRow - 1, Convert.ToString(TOTAMT));
//						oMat1.SetLineData(oRow);
//						break;
//				}
//			}
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
//			string Chk_Data1 = null;
//			string JSNYER = null;
//			ErrNum = 0;
//			/// 헤더부분 체크
//			switch (true) {
//				case string.IsNullOrEmpty(oDS_ZPY502H.GetValue("U_JSNYER", 0)):
//					ErrNum = 4;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY502H.GetValue("U_MSTCOD", 0)):
//					ErrNum = 5;
//					goto Error_Message;
//					break;
//			}
//			JSNYER = oDS_ZPY502H.GetValue("U_JSNYER", 0);
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
//				oDS_ZPY502L.Offset = iRow;
//				if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY502L.GetValue("U_JONNAM", iRow)))) {
//					ErrNum = 2;
//					oMat1.Columns.Item("Col1").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY502L.GetValue("U_JONNBR", iRow)))) {
//					ErrNum = 6;
//					oMat1.Columns.Item("Col2").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY502L.GetValue("U_JONSTR", iRow)))) {
//					ErrNum = 7;
//					oMat1.Columns.Item("Col18").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (string.IsNullOrEmpty(Strings.Trim(oDS_ZPY502L.GetValue("U_JONEND", iRow)))) {
//					ErrNum = 8;
//					oMat1.Columns.Item("Col19").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				} else if (Strings.Left(oDS_ZPY502L.GetValue("U_JONSTR", iRow), 4) != Strings.Trim(JSNYER) | Strings.Left(oDS_ZPY502L.GetValue("U_JONEND", iRow), 4) != Strings.Trim(JSNYER)) {
//					ErrNum = 9;
//					oMat1.Columns.Item("Col18").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					goto Error_Message;
//				//// 2010년 폐지 비과세 항목에 금액이 입력된 경우 금액을 초기화 시킴
//				} else if (JSNYER >= "2010") {
//					if (Conversion.Val(oDS_ZPY502L.GetValue("U_JBTX01", iRow)) != 0) {
//						oDS_ZPY502L.Offset = iRow;
//						oDS_ZPY502L.SetValue("U_JBTX01", iRow, Convert.ToString(0));
//						oMat1.SetLineData((iRow + 1));
//					}
//					if (Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY01", iRow)) != 0) {
//						oDS_ZPY502L.Offset = iRow;
//						oDS_ZPY502L.SetValue("U_JBTY01", iRow, Convert.ToString(0));
//						oMat1.SetLineData((iRow + 1));
//					}
//					if (Conversion.Val(oDS_ZPY502L.GetValue("U_JBTY20", iRow)) != 0) {
//						oDS_ZPY502L.Offset = iRow;
//						oDS_ZPY502L.SetValue("U_JBTY20", iRow, Convert.ToString(0));
//						oMat1.SetLineData((iRow + 1));
//					}
//				} else {
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//중복체크작업
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					Chk_Data = Strings.Trim(oDS_ZPY502L.GetValue("U_JONNBR", iRow));
//					Chk_Data1 = Strings.Trim(oDS_ZPY502L.GetValue("U_JONSTR", iRow));
//					for (kRow = iRow + 1; kRow <= oMat1.VisualRowCount - 2; kRow++) {
//						oDS_ZPY502L.Offset = kRow;
//						if (Strings.Trim(Chk_Data) == Strings.Trim(oDS_ZPY502L.GetValue("U_JONNBR", kRow)) & Strings.Trim(Chk_Data1) == Strings.Trim(oDS_ZPY502L.GetValue("U_JONSTR", kRow))) {
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
//			oDS_ZPY502L.RemoveRecord(oDS_ZPY502L.Size - 1);
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("근무처명은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("근무처 사업자번호/귀속시작일이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("사업자번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 7) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속시작일은 필수입니다. (귀속연도+0101 또는 종전근무지의 당해입사일)을 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 8) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속종료일은 필수입니다. (종전근무지의 퇴사일)을 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 9) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속연도와 귀속시작일 또는 귀속종료일의 연도가 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 10) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("2010년부터 폐지된 비과세 항목에 금액이 입력되어 있습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}

//			functionReturnValue = false;
//			return functionReturnValue;
//		}


//		private bool HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string DocNum = null;

//			ErrNum = 0;
//			/// Check
//			switch (true) {
//				case string.IsNullOrEmpty(oDS_ZPY502H.GetValue("U_JSNYER", 0)):
//					ErrNum = 1;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY502H.GetValue("U_MSTCOD", 0)):
//					ErrNum = 2;
//					goto Error_Message;
//					break;
//				case string.IsNullOrEmpty(oDS_ZPY502H.GetValue("U_CLTCOD", 0)):
//					ErrNum = 3;
//					goto Error_Message;
//					break;
//			}

//			if (Strings.Trim(oDS_ZPY502H.GetValue("U_ENDCHK", 0)) == "Y" & Strings.Trim(oOLDCHK) == "Y") {
//				ErrNum = 5;
//				goto Error_Message;
//			}

//			DocNum = Exist_YN(ref oDS_ZPY502H.GetValue("U_JSNYER", 0), ref oDS_ZPY502H.GetValue("U_MSTCOD", 0), ref oDS_ZPY502H.GetValue("U_CLTCOD", 0));
//			if (!string.IsNullOrEmpty(Strings.Trim(DocNum)) & Strings.Trim(oDS_ZPY502H.GetValue("DocNum", 0)) != Strings.Trim(DocNum)) {
//				//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//				//같은데이터가 존재하는데 자기 자신이 현재 자기자신이 아니라면(같은월에는 취소한거 아니면 하나만 존재해야함)
//				//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//				ErrNum = 4;
//				goto Error_Message;
//			}



//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 연도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("문서번호" + DocNum + " 와(과) 데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("잠금 자료입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//							////추가및 업데이트시에
//							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//								////추가및 업데이트시에
//								//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//								if (HeaderSpaceLineDel() == false) {
//									BubbleEvent = false;
//									return;
//								} else {
//									if (MatrixSpaceLineDel() == false) {
//										BubbleEvent = false;
//									}
//								}
//							}
//						/// ChooseBtn사원리스트
//						} else if (pval.ItemUID == "CBtn1" & oForm.Items.Item("MSTCOD").Enabled == true) {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
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
//					} else if (pval.BeforeAction == false & pval.ItemChanged == true & pval.ItemUID == "Mat1" & Strings.Left(pval.ColUID, 3) == "Col") {
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
//							MDC_Globals.Sbo_Application.StatusBar.SetText("근무처명은 필수입니다. 입력하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					} else if (pval.BeforeAction == true & pval.ColUID == "Col2" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(Strings.Trim(oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.String))) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("근무처 사업자번호는 필수입니다. 입력하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//						//UPGRADE_NOTE: oDS_ZPY502H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY502H = null;
//						//UPGRADE_NOTE: oDS_ZPY502L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_ZPY502L = null;
//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//					}
//					break;
//				//et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					if (pval.BeforeAction == false) {
//						Matrix_TitleSetting();
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


//		private string Exist_YN(ref string JOBYER, ref string MSTCOD, ref string CLTCOD)
//		{
//			string functionReturnValue = null;
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//저장할 데이터의 기존데이터가 있는지 확인한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "SELECT Top 1 T1.DocNum FROM [@ZPY502H] T1 ";
//			sQry = sQry + " WHERE T1.U_JSNYER = N'" + Strings.Trim(JOBYER) + "'";
//			sQry = sQry + "   AND T1.U_MSTCOD = N'" + Strings.Trim(MSTCOD) + "'";
//			sQry = sQry + "   AND T1.U_CLTCOD = N'" + Strings.Trim(CLTCOD) + "'";
//			oRecordSet.DoQuery(sQry);

//			while (!(oRecordSet.EoF)) {
//				//UPGRADE_WARNING: oRecordSet().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
////*******************************************************************
////// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
////*******************************************************************
//		private void Matrix_AddRow(int oRow, ref bool Insert_YN = false)
//		{
//			if (Insert_YN == false) {
//				oDS_ZPY502L.InsertRecord((oRow));
//			}
//			oDS_ZPY502L.Offset = oRow;
//			oDS_ZPY502L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//			oDS_ZPY502L.SetValue("U_JONNAM", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONNBR", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONSTR", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONEND", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONPAY", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBNS", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBT1", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBT2", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBT3", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBU3", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBT4", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBT5", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONBT6", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONMED", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONGBH", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONKUK", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONKUE", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONRET", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONGAB", oRow, "");
//			oDS_ZPY502L.SetValue("U_JONJUM", oRow, "");
//			/// 2009년추가
//			oDS_ZPY502L.SetValue("U_JBTG01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH05", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH06", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH07", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH08", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH09", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH10", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH11", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH12", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTH13", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTI01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTK01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTM01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTM02", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTM03", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTO01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTQ01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTS01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTT01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTX01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTY01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTY02", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTY03", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTY20", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTY21", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTZ01", oRow, "");
//			oDS_ZPY502L.SetValue("U_JBTTOT", oRow, "");

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
//						if (Strings.Trim(oDS_ZPY502H.GetValue("U_ENDCHK", 0)) == "Y") {
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
//					// oForm.Items("Btn1").Visible = True
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
//								//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oMat1.Columns.Item("Col0").Cells.Item(i + 1).Specific.VALUE = i + 1;
//							}

//							oMat1.FlushToDataSource();
//							oDS_ZPY502L.RemoveRecord(oDS_ZPY502L.Size - 1);
//							//// Mat1에 마지막라인(빈라인) 삭제
//							oMat1.Clear();
//							oMat1.LoadFromDataSource();
//						}
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


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY502.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY502_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY502");
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
//			oForm.EnableMenu(("1284"), false);
//			/// 취소

//			if (!string.IsNullOrEmpty(Strings.Trim(JSNYER))) {
//				ShowSource(ref JSNYER, ref MSTCOD, ref CLTCOD);
//			}

//			oForm.Freeze(false);
//			oForm.Update();
//			//oForm.Visible = True

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
//			oDS_ZPY502H = oForm.DataSources.DBDataSources("@ZPY502H");
//			oDS_ZPY502L = oForm.DataSources.DBDataSources("@ZPY502L");

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			//// 사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT Code, Name FROM [@PH_PY005A]";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oCombo.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;
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

//		private void ShowSource(ref string JSNYER, ref string MSTCOD, ref string CLTCOD)
//		{
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string DocEntry = null;
//			ZPAY_g_EmpID oMast = default(ZPAY_g_EmpID);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "SELECT DocEntry FROM [@ZPY502H]";
//			sQry = sQry + "   WHERE U_JSNYER = N'" + JSNYER + "'";
//			sQry = sQry + "   AND   U_MSTCOD = N'" + MSTCOD + "'";
//			sQry = sQry + "   AND   U_CLTCOD = N'" + CLTCOD + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				while (!(oRecordSet.EoF)) {
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					DocEntry = oRecordSet.Fields.Item(0).Value;
//					oRecordSet.MoveNext();
//				}
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("JSNYER").Specific.VALUE = JSNYER;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("MSTCOD").Specific.String = MSTCOD;
//				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("CLTCOD").Specific.Select(CLTCOD, SAPbouiCOM.BoSearchKey.psk_ByValue);
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocNum").Specific.VALUE = DocEntry;

//				oForm.Items.Item("DocNum").Update();
//				oMat1.LoadFromDataSource();
//				oForm.Update();
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//				MDC_Globals.Sbo_Application.ActivateMenuItem("1282");

//				oDS_ZPY502H.SetValue("U_JSNYER", 0, JSNYER);
//				oDS_ZPY502H.SetValue("U_MSTCOD", 0, MSTCOD);
//				oDS_ZPY502H.SetValue("U_CLTCOD", 0, CLTCOD);
//				//UPGRADE_WARNING: oMast 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oMast = MDC_SetMod.Get_EmpID_InFo(ref MSTCOD);
//				oDS_ZPY502H.SetValue("U_MSTNAM", 0, oMast.MSTNAM);
//				oDS_ZPY502H.SetValue("U_EmpID", 0, oMast.EmpID);

//				oForm.Update();

//				MDC_Globals.Sbo_Application.SendKeys("{TAB}");
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

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

//			JSNYER = Strings.Trim(oDS_ZPY502H.GetValue("U_JSNYER", 0));

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
//				oMat1.Columns.Item("Col5").TitleObject.Caption = "비과세-생산";
//				oMat1.Columns.Item("Col6").TitleObject.Caption = "비과세-보육";
//				oMat1.Columns.Item("Col7").TitleObject.Caption = "비과세-국외";
//				oMat1.Columns.Item("Col13").TitleObject.Caption = "비과세-연구";
//				oMat1.Columns.Item("Col20").TitleObject.Caption = "비과세-외국인";

//				oMat1.Columns.Item("Col23").TitleObject.Caption = "비과세(G01)";
//				oMat1.Columns.Item("Col24").TitleObject.Caption = "비과세(H01)";
//				oMat1.Columns.Item("Col25").TitleObject.Caption = "비과세(H05)";
//				oMat1.Columns.Item("Col26").TitleObject.Caption = "비과세(H06)";
//				oMat1.Columns.Item("Col27").TitleObject.Caption = "비과세(H07)";
//				oMat1.Columns.Item("Col28").TitleObject.Caption = "비과세(H08)";
//				oMat1.Columns.Item("Col29").TitleObject.Caption = "비과세(H09)";
//				oMat1.Columns.Item("Col30").TitleObject.Caption = "비과세(H10)";
//				oMat1.Columns.Item("Col31").TitleObject.Caption = "비과세(H11)";
//				oMat1.Columns.Item("Col32").TitleObject.Caption = "비과세(H12)";
//				oMat1.Columns.Item("Col33").TitleObject.Caption = "비과세(H13)";
//				oMat1.Columns.Item("Col34").TitleObject.Caption = "비과세(I01)";
//				oMat1.Columns.Item("Col35").TitleObject.Caption = "비과세(K01)";
//				oMat1.Columns.Item("Col36").TitleObject.Caption = "비과세(M01)";
//				oMat1.Columns.Item("Col37").TitleObject.Caption = "비과세(M02)";
//				oMat1.Columns.Item("Col38").TitleObject.Caption = "비과세(M03)";
//				oMat1.Columns.Item("Col39").TitleObject.Caption = "비과세(O01)";
//				oMat1.Columns.Item("Col40").TitleObject.Caption = "비과세(Q01)";
//				oMat1.Columns.Item("Col41").TitleObject.Caption = "비과세(S01)";
//				oMat1.Columns.Item("Col42").TitleObject.Caption = "비과세(T01)";
//				oMat1.Columns.Item("Col43").TitleObject.Caption = "비과세(X01)";
//				oMat1.Columns.Item("Col44").TitleObject.Caption = "비과세(Y01)";
//				oMat1.Columns.Item("Col45").TitleObject.Caption = "비과세(Y02)";
//				oMat1.Columns.Item("Col46").TitleObject.Caption = "비과세(Y03)";
//				oMat1.Columns.Item("Col47").TitleObject.Caption = "비과세(Y20)";
//				oMat1.Columns.Item("Col50").TitleObject.Caption = "비과세(Y21)";
//				oMat1.Columns.Item("Col48").TitleObject.Caption = "비과세(Z01)";

//				//// 비과세 컬럼 전체 표시
//				for (iCol = 5; iCol <= 50; iCol++) {
//					if ((iCol >= 5 & iCol <= 7) | iCol == 13 | iCol == 20 | (iCol >= 23 & iCol <= 50)) {
//						oMat1.Columns.Item("Col" + Convert.ToString(iCol)).Visible = true;
//					}
//				}

//			} else if (Conversion.Val(JSNYER) <= 2008) {
//				while (!(oRecordSet.EoF)) {
//					//// 비과세 코드에 따라 컬럼UID 확인
//					switch (oRecordSet.Fields.Item("U_BTXCOD").Value) {
//						case "01":
//							COLNAM = "Col5";
//							/// 비과세-생산
//							break;
//						case "07":
//							COLNAM = "Col6";
//							/// 비과세-보육
//							break;
//						case "03":
//							COLNAM = "Col7";
//							/// 비과세-국외
//							break;
//						case "06":
//							COLNAM = "Col13";
//							/// 비과세-연구
//							break;
//						case "04":
//							COLNAM = "Col20";
//							/// 비과세-외국인
//							break;
//						default:
//							COLNAM = "";
//							break;
//					}

//					//// 컬럼명, 화면표시 여부 적용
//					if (!string.IsNullOrEmpty(COLNAM)) {
//						var _with2 = oMat1.Columns.Item(COLNAM);
//						//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						_with2.TitleObject.Caption = oRecordSet.Fields.Item("U_BTXNAM").Value;
//						_with2.Editable = false;
//					}

//					oRecordSet.MoveNext();
//				}

//			} else {
//				while (!(oRecordSet.EoF)) {
//					//// 비과세 코드에 따라 컬럼UID 확인
//					switch (oRecordSet.Fields.Item("U_BTXCOD").Value) {
//						case "G01":
//							COLNAM = "Col23";
//							//// 비과세(G01)
//							break;
//						case "H01":
//							COLNAM = "Col24";
//							//// 비과세(H01)
//							break;
//						case "H05":
//							COLNAM = "Col25";
//							//// 비과세(H05)
//							break;
//						case "H06":
//							COLNAM = "Col26";
//							//// 비과세(H06)
//							break;
//						case "H07":
//							COLNAM = "Col27";
//							//// 비과세(H07)
//							break;
//						case "H08":
//							COLNAM = "Col28";
//							//// 비과세(H08)
//							break;
//						case "H09":
//							COLNAM = "Col29";
//							//// 비과세(H09)
//							break;
//						case "H10":
//							COLNAM = "Col30";
//							//// 비과세(H10)
//							break;
//						case "H11":
//							COLNAM = "Col31";
//							//// 비과세(H11)
//							break;
//						case "H12":
//							COLNAM = "Col32";
//							//// 비과세(H12)
//							break;
//						case "H13":
//							COLNAM = "Col33";
//							//// 비과세(H13)
//							break;
//						case "I01":
//							COLNAM = "Col34";
//							//// 비과세(I01)
//							break;
//						case "K01":
//							COLNAM = "Col35";
//							//// 비과세(K01)
//							break;
//						case "M01":
//							COLNAM = "Col36";
//							//// 비과세(M01)
//							break;
//						case "M02":
//							COLNAM = "Col37";
//							//// 비과세(M02)
//							break;
//						case "M03":
//							COLNAM = "Col38";
//							//// 비과세(M03)
//							break;
//						case "O01":
//							COLNAM = "Col39";
//							//// 비과세(O01)
//							break;
//						case "Q01":
//							COLNAM = "Col40";
//							//// 비과세(Q01)
//							break;
//						case "S01":
//							COLNAM = "Col41";
//							//// 비과세(S01)
//							break;
//						case "T01":
//							COLNAM = "Col42";
//							//// 비과세(T01)
//							break;
//						case "X01":
//							COLNAM = "Col43";
//							//// 비과세(X01)
//							break;
//						case "Y01":
//						case "R10":
//							COLNAM = "Col44";
//							//// 비과세(Y01)
//							break;
//						case "Y02":
//							COLNAM = "Col45";
//							//// 비과세(Y02)
//							break;
//						case "Y03":
//							COLNAM = "Col46";
//							//// 비과세(Y03)
//							break;
//						case "Y20":
//						case "Y22":
//							COLNAM = "Col47";
//							//// 비과세(Y20)
//							break;
//						case "Y21":
//							COLNAM = "Col50";
//							//// 비과세(Y21)
//							break;
//						case "Z01":
//							COLNAM = "Col48";
//							//// 비과세(Z01)
//							break;
//					}

//					//// 컬럼명, 화면표시 여부 적용
//					var _with1 = oMat1.Columns.Item(COLNAM);
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					_with1.TitleObject.Caption = oRecordSet.Fields.Item("U_BTXNAM").Value;

//					oRecordSet.MoveNext();
//				}
//				if (JSNYER >= "2010") {
//					oMat1.Columns.Item("Col43").Visible = false;
//				}

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
