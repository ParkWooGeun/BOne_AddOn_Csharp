using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작번별 진행금액 일괄조회
	/// </summary>
	internal class PS_PP985 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Matrix oMat03;
		private SAPbouiCOM.Matrix oMat04;
		private SAPbouiCOM.Matrix oMat05;
		private SAPbouiCOM.Matrix oMat06;

		//등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP985A; //자재비		
		private SAPbouiCOM.DBDataSource oDS_PS_PP985B; //자체가공비	
		private SAPbouiCOM.DBDataSource oDS_PS_PP985C; //외주가공비	
		private SAPbouiCOM.DBDataSource oDS_PS_PP985D; //외주제작비	
		private SAPbouiCOM.DBDataSource oDS_PS_PP985E; //설계비	
		private SAPbouiCOM.DBDataSource oDS_PS_PP985F; //기성매출

		/// <summary>
		/// 화면 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP985.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP985_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP985");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP985_CreateItems();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oMat01.Columns.Item("PurchaseNm").Visible = false;  //품의구분은 굳이 필요 없을 것 같아 Hidden 처리
				oMat01.AutoResizeColumns();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP985_CreateItems
		/// </summary>
		private void PS_PP985_CreateItems()
		{
			try
			{
				oDS_PS_PP985A = oForm.DataSources.DBDataSources.Item("@PS_USERDS06");
				oDS_PS_PP985B = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
				oDS_PS_PP985C = oForm.DataSources.DBDataSources.Item("@PS_USERDS03");
				oDS_PS_PP985D = oForm.DataSources.DBDataSources.Item("@PS_USERDS04");
				oDS_PS_PP985E = oForm.DataSources.DBDataSources.Item("@PS_USERDS05");
				oDS_PS_PP985F = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				// 매트릭스 개체 할당
				//자재비 매트릭스
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat01.AutoResizeColumns();
				//자체가공비 매트릭스
				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat02.AutoResizeColumns();
				//외주가공비 매트릭스
				oMat03 = oForm.Items.Item("Mat03").Specific;
				oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat03.AutoResizeColumns();
				//외주제작비 매트릭스
				oMat04 = oForm.Items.Item("Mat04").Specific;
				oMat04.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat04.AutoResizeColumns();
				//설계비 매트릭스
				oMat05 = oForm.Items.Item("Mat05").Specific;
				oMat05.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat05.AutoResizeColumns();
				//기성매출 매트릭스
				oMat06 = oForm.Items.Item("Mat06").Specific;
				oMat06.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat06.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

				oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

				oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");

				oForm.DataSources.UserDataSources.Add("Opt04", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt04").Specific.DataBind.SetBound(true, "", "Opt04");

				oForm.DataSources.UserDataSources.Add("Opt05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt05").Specific.DataBind.SetBound(true, "", "Opt05");

				oForm.DataSources.UserDataSources.Add("Opt06", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt06").Specific.DataBind.SetBound(true, "", "Opt06");

				oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
				oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");
				oForm.Items.Item("Opt01").Specific.GroupWith("Opt04");
				oForm.Items.Item("Opt01").Specific.GroupWith("Opt05");
				oForm.Items.Item("Opt01").Specific.GroupWith("Opt06");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("SubNo01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("SubNo1").Specific.DataBind.SetBound(true, "", "SubNo01");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("SubNo02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("SubNo2").Specific.DataBind.SetBound(true, "", "SubNo02");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//규격
				oForm.DataSources.UserDataSources.Add("Spec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Spec").Specific.DataBind.SetBound(true, "", "Spec");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//수주금액
				oForm.DataSources.UserDataSources.Add("SjAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("SjAmt").Specific.DataBind.SetBound(true, "", "SjAmt");

				//자재비
				oForm.DataSources.UserDataSources.Add("MatAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("MatAmt").Specific.DataBind.SetBound(true, "", "MatAmt");

				//자체가공비
				oForm.DataSources.UserDataSources.Add("GagongAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("GagongAmt").Specific.DataBind.SetBound(true, "", "GagongAmt");

				//외주가공비
				oForm.DataSources.UserDataSources.Add("OutgAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("OutgAmt").Specific.DataBind.SetBound(true, "", "OutgAmt");

				//외주제작비
				oForm.DataSources.UserDataSources.Add("OutmAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("OutmAmt").Specific.DataBind.SetBound(true, "", "OutmAmt");

				//설계비
				oForm.DataSources.UserDataSources.Add("DrawAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("DrawAmt").Specific.DataBind.SetBound(true, "", "DrawAmt");

				//계
				oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");

				//생산완료여부
				oForm.DataSources.UserDataSources.Add("PCYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("PCYN").Specific.DataBind.SetBound(true, "", "PCYN");

				//생산완료일자
				oForm.DataSources.UserDataSources.Add("EndDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("EndDate").Specific.DataBind.SetBound(true, "", "EndDate");

				//수주일
				oForm.DataSources.UserDataSources.Add("SjDocDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SjDocDate").Specific.DataBind.SetBound(true, "", "SjDocDate");

				//작업지시일
				oForm.DataSources.UserDataSources.Add("WODate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("WODate").Specific.DataBind.SetBound(true, "", "WODate");

				//납기일
				oForm.DataSources.UserDataSources.Add("SjDueDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SjDueDate").Specific.DataBind.SetBound(true, "", "SjDueDate");

				//협력업체(외주제작)
				oForm.DataSources.UserDataSources.Add("OtCardNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("OtCardNm").Specific.DataBind.SetBound(true, "", "OtCardNm");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP985_AddMatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP985_AddMatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP985A.InsertRecord(oRow);
				}

				oMat01.AddRow();
				oDS_PS_PP985A.Offset = oRow;
				oDS_PS_PP985A.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat01.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP985_AddMatrixRow02
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP985_AddMatrixRow02(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP985B.InsertRecord(oRow);
				}

				oMat02.AddRow();
				oDS_PS_PP985B.Offset = oRow;
				oDS_PS_PP985B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat02.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP985_AddMatrixRow03
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP985_AddMatrixRow03(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP985C.InsertRecord(oRow);
				}

				oMat03.AddRow();
				oDS_PS_PP985C.Offset = oRow;
				oDS_PS_PP985C.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat03.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP985_AddMatrixRow04
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP985_AddMatrixRow04(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP985D.InsertRecord(oRow);
				}

				oMat04.AddRow();
				oDS_PS_PP985D.Offset = oRow;
				oDS_PS_PP985D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat04.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP985_AddMatrixRow05
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP985_AddMatrixRow05(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP985E.InsertRecord(oRow);
				}

				oMat05.AddRow();
				oDS_PS_PP985E.Offset = oRow;
				oDS_PS_PP985E.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat05.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP985_AddMatrixRow06
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP985_AddMatrixRow06(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP985F.InsertRecord(oRow);
				}

				oMat06.AddRow();
				oDS_PS_PP985F.Offset = oRow;
				oDS_PS_PP985F.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat06.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP985_Accumulate 데이터 집계 및 테이블 저장
		/// </summary>
		private void PS_PP985_Accumulate()
		{
			string sQry;
			string CntcCode; //조회자 사번
			string OrdNum;	 //메인작번
			string SubNo1;	 //서브작번1
			string SubNo2;	 //서브작번2
			string FrDt;	 //기간(시작)
			string ToDt;     //기간(종료)
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();
				FrDt   = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt   = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "비용 집계중...!";

				//메인정보 집계
				sQry = "EXEC [PS_PP985_91] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";
				oRecordSet.DoQuery(sQry);

				//자재비 집계
				sQry = "EXEC [PS_PP985_92] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";
				oRecordSet.DoQuery(sQry);

				//자체가공비 집계
				sQry = "EXEC [PS_PP985_93] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";
				oRecordSet.DoQuery(sQry);

				//외주가공비 집계
				sQry = "EXEC [PS_PP985_94] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";
				oRecordSet.DoQuery(sQry);

				//외주제작비 집계
				sQry = "EXEC [PS_PP985_95] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";
				oRecordSet.DoQuery(sQry);

				//설계비 집계
				sQry = "EXEC [PS_PP985_96] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";
				oRecordSet.DoQuery(sQry);

				//기성매출 집계
				sQry = "EXEC [PS_PP985_97] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.MessageBox("집계가 완료되었습니다. 조회를 시작하십시오.");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP985_Select0 주 데이터 조회
		/// </summary>
		private void PS_PP985_Select0()
		{
			string sQry;
			string errMessage = string.Empty;
			string CntcCode; //조회자 사번
			string OrdNum;	 //메인작번
			string SubNo1;	 //서브작번1
			string SubNo2;   //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP985_01] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "데이터가 존재하지 않습니다.확인하세요.";
					throw new Exception();
				}

				oForm.Items.Item("SjAmt").Specific.VALUE = oRecordSet.Fields.Item("SjAmt").Value.ToString().Trim();         //수주금액
				oForm.Items.Item("MatAmt").Specific.VALUE = oRecordSet.Fields.Item("MatAmt").Value.ToString().Trim();       //자재비
				oForm.Items.Item("GagongAmt").Specific.VALUE = oRecordSet.Fields.Item("GagongAmt").Value.ToString().Trim(); //자체가공비
				oForm.Items.Item("OutgAmt").Specific.VALUE = oRecordSet.Fields.Item("OutgAmt").Value.ToString().Trim();     //외주가공비
				oForm.Items.Item("OutmAmt").Specific.VALUE = oRecordSet.Fields.Item("OutmAmt").Value.ToString().Trim();     //외주제작비
				oForm.Items.Item("DrawAmt").Specific.VALUE = oRecordSet.Fields.Item("DrawAmt").Value.ToString().Trim();     //설계비
				oForm.Items.Item("Total").Specific.VALUE = oRecordSet.Fields.Item("Total").Value.ToString().Trim();         //계
				oForm.Items.Item("PCYN").Specific.VALUE = oRecordSet.Fields.Item("PCYN").Value.ToString().Trim();           //생산완료여부
				oForm.Items.Item("EndDate").Specific.VALUE = oRecordSet.Fields.Item("EndDate").Value.ToString().Trim();     //생산완료일자
				oForm.Items.Item("SjDocDate").Specific.VALUE = oRecordSet.Fields.Item("SjDocDate").Value.ToString().Trim(); //수주일
				oForm.Items.Item("WODate").Specific.VALUE = oRecordSet.Fields.Item("WODate").Value.ToString().Trim();       //작업지시일
				oForm.Items.Item("SjDueDate").Specific.VALUE = oRecordSet.Fields.Item("SjDueDate").Value.ToString().Trim(); //납기일
				oForm.Items.Item("OtCardNm").Specific.VALUE = oRecordSet.Fields.Item("OtCardNm").Value.ToString().Trim();   //협력업체명(외주제작)
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP985_Select1  자재비 조회
		/// </summary>
		private void PS_PP985_Select1()
		{
			int i;
			string sQry;
			string CntcCode; //조회자 사번
			string OrdNum;   //메인작번
			string SubNo1;   //서브작번1
			string SubNo2;   //서브작번2

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP985_02] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "'";

				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_PP985A.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP985A.Size)
					{
						oDS_PS_PP985A.InsertRecord(i);
					}

					oMat01.AddRow();
					oDS_PS_PP985A.Offset = i;
					oDS_PS_PP985A.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP985A.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());		//작번
					oDS_PS_PP985A.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("PurchaseNm").Value.ToString().Trim());	//품의구분
					oDS_PS_PP985A.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());	//자재품목코드
					oDS_PS_PP985A.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());	//자재품명
					oDS_PS_PP985A.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("MatAmt").Value.ToString().Trim());		//자재비

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP985_Select2  자체가공비 조회
		/// </summary>
		private void PS_PP985_Select2()
		{
			int i;
			string sQry;
			string CntcCode; //조회자 사번
			string OrdNum;   //메인작번
			string SubNo1;   //서브작번1
			string SubNo2;   //서브작번2

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP985_03] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "'";
				oRecordSet.DoQuery(sQry);

				oMat02.Clear();
				oDS_PS_PP985B.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP985B.Size)
					{
						oDS_PS_PP985B.InsertRecord(i);
					}

					oMat02.AddRow();
					oDS_PS_PP985B.Offset = i;
					oDS_PS_PP985B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP985B.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());    //작번
					oDS_PS_PP985B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());    //공정명
					oDS_PS_PP985B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("StdTime").Value.ToString().Trim());   //표준공수
					oDS_PS_PP985B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("WkTime").Value.ToString().Trim());    //실동공수
					oDS_PS_PP985B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amt").Value.ToString().Trim());       //가공비(실동)
					oDS_PS_PP985B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("CompltDt").Value.ToString().Trim());  //완료요구일
					oDS_PS_PP985B.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("FirstWkDt").Value.ToString().Trim()); //최초작업일
					oDS_PS_PP985B.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("LastWkDt").Value.ToString().Trim());  //최종작업일

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP985_Select3  외주가공비 조회
		/// </summary>
		private void PS_PP985_Select3()
		{
			int i;
			string sQry;
			string CntcCode; //조회자 사번
			string OrdNum;   //메인작번
			string SubNo1;   //서브작번1
			string SubNo2;   //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP985_04] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "'";
				oRecordSet.DoQuery(sQry);

				oMat03.Clear();
				oDS_PS_PP985C.Clear();
				oMat03.FlushToDataSource();
				oMat03.LoadFromDataSource();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP985C.Size)
					{
						oDS_PS_PP985C.InsertRecord(i);
					}

					oMat03.AddRow();
					oDS_PS_PP985C.Offset = i;
					oDS_PS_PP985C.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP985C.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());	 //작번
					oDS_PS_PP985C.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	 //공정코드
					oDS_PS_PP985C.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	 //공정명
					oDS_PS_PP985C.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
					oDS_PS_PP985C.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amt").Value.ToString().Trim());      //금액

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat03.LoadFromDataSource();
				oMat03.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP985_Select4  외주제작비 조회
		/// </summary>
		private void PS_PP985_Select4()
		{
			int i;
			string sQry;
			string CntcCode; //조회자 사번
			string OrdNum;   //메인작번
			string SubNo1;   //서브작번1
			string SubNo2;   //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP985_05] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "'";
				oRecordSet.DoQuery(sQry);

				oMat04.Clear();
				oDS_PS_PP985D.Clear();
				oMat04.FlushToDataSource();
				oMat04.LoadFromDataSource();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP985D.Size)
					{
						oDS_PS_PP985D.InsertRecord(i);
					}

					oMat04.AddRow();
					oDS_PS_PP985D.Offset = i;
					oDS_PS_PP985D.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP985D.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());	 //작번
					oDS_PS_PP985D.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	 //공정코드
					oDS_PS_PP985D.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	 //공정명
					oDS_PS_PP985D.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
					oDS_PS_PP985D.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amt").Value.ToString().Trim());      //금액

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat04.LoadFromDataSource();
				oMat04.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP985_Select5  설계비 조회
		/// </summary>
		private void PS_PP985_Select5()
		{
			int i;
			string sQry;
			string CntcCode; //조회자 사번
			string OrdNum;   //메인작번
			string SubNo1;   //서브작번1
			string SubNo2;   //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP985_06] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "'";
				oRecordSet.DoQuery(sQry);

				oMat05.Clear();
				oDS_PS_PP985E.Clear();
				oMat05.FlushToDataSource();
				oMat05.LoadFromDataSource();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP985E.Size)
					{
						oDS_PS_PP985E.InsertRecord(i);
					}

					oMat05.AddRow();
					oDS_PS_PP985E.Offset = i;
					oDS_PS_PP985E.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP985E.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());	 //작번
					oDS_PS_PP985E.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("DocDate").Value.ToString().Trim());	 //일자
					oDS_PS_PP985E.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("WorkCode").Value.ToString().Trim()); //작업자사번
					oDS_PS_PP985E.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("WorkName").Value.ToString().Trim()); //작업자명
					oDS_PS_PP985E.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("PQty").Value.ToString().Trim());	 //도면매수
					oDS_PS_PP985E.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amt").Value.ToString().Trim());      //설계비

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat05.LoadFromDataSource();
				oMat05.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP985_Select6  기성매출 조회
		/// </summary>
		private void PS_PP985_Select6()
		{
			int i;
			string sQry;
			string CntcCode; //조회자 사번
			string OrdNum;   //메인작번
			string SubNo1;   //서브작번1
			string SubNo2;   //서브작번2
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP985_07] '";
				sQry += CntcCode + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "'";
				oRecordSet.DoQuery(sQry);

				oMat06.Clear();
				oDS_PS_PP985F.Clear();
				oMat06.FlushToDataSource();
				oMat06.LoadFromDataSource();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{

					if (i + 1 > oDS_PS_PP985F.Size)
					{
						oDS_PS_PP985F.InsertRecord(i);
					}

					oMat06.AddRow();
					oDS_PS_PP985F.Offset = i;
					oDS_PS_PP985F.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP985F.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("StdYear").Value.ToString().Trim()); //년도
					oDS_PS_PP985F.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Month01").Value.ToString().Trim()); //1월
					oDS_PS_PP985F.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("Month02").Value.ToString().Trim()); //2월
					oDS_PS_PP985F.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("Month03").Value.ToString().Trim()); //3월
					oDS_PS_PP985F.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("Month04").Value.ToString().Trim()); //4월
					oDS_PS_PP985F.SetValue("U_ColSum05", i, oRecordSet.Fields.Item("Month05").Value.ToString().Trim()); //5월
					oDS_PS_PP985F.SetValue("U_ColSum06", i, oRecordSet.Fields.Item("Month06").Value.ToString().Trim()); //6월
					oDS_PS_PP985F.SetValue("U_ColSum07", i, oRecordSet.Fields.Item("Month07").Value.ToString().Trim()); //7월
					oDS_PS_PP985F.SetValue("U_ColSum08", i, oRecordSet.Fields.Item("Month08").Value.ToString().Trim()); //8월
					oDS_PS_PP985F.SetValue("U_ColSum09", i, oRecordSet.Fields.Item("Month09").Value.ToString().Trim()); //9월
					oDS_PS_PP985F.SetValue("U_ColSum10", i, oRecordSet.Fields.Item("Month10").Value.ToString().Trim()); //10월
					oDS_PS_PP985F.SetValue("U_ColSum11", i, oRecordSet.Fields.Item("Month11").Value.ToString().Trim()); //11월
					oDS_PS_PP985F.SetValue("U_ColSum12", i, oRecordSet.Fields.Item("Month12").Value.ToString().Trim()); //12월
					oDS_PS_PP985F.SetValue("U_ColSum13", i, oRecordSet.Fields.Item("Total").Value.ToString().Trim());   //계

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat06.LoadFromDataSource();
				oMat06.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP985_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP985_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim()))
				{
					errMessage = "메인 작번은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP985_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP985_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			string OrdNum;   //메인작번
			string SubNo1;   //서브작번1
			string SubNo2;   //서브작번2
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oUID == "OrdNum" || oUID == "SubNo1" || oUID == "SubNo2")
				{
					OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
					SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
					SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

					sQry = "  SELECT   CASE";
					sQry += "          WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "          ELSE T0.U_JakMyung";
					sQry += "          END AS [ItemName],";
					sQry += "          CASE";
					sQry += "          WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "          ELSE T0.U_JakSize";
					sQry += "          END AS [SPEC]";
					sQry += " FROM     [@PS_PP020H] AS T0";
					sQry += " WHERE    T0.U_JakName = '" + OrdNum + "'";
					sQry += "          AND T0.U_SubNo1 = CASE WHEN '" + SubNo1 + "' = '' THEN '00' ELSE '" + SubNo1 + "' END";
					sQry += "          AND T0.U_SubNo2 = CASE WHEN '" + SubNo2 + "' = '' THEN '000' ELSE '" + SubNo2 + "' END";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("ItemName").Specific.VALUE = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
					oForm.Items.Item("Spec").Specific.VALUE = oRecordSet.Fields.Item("Spec").Value.ToString().Trim();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP985_PrintReport
		/// </summary>
		[STAThread]
		private void PS_PP985_PrintReport()
		{
			string WinTitle;
			string ReportName;

			string FrDt;
			string ToDt;
			string CntcCode;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CntcCode = dataHelpClass.User_MSTCOD();

				WinTitle = "[PS_PP985] 작번별 진행금액 일괄조회";
				ReportName = "PS_PP985_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt));

				//SubReport Parameter
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode, "PS_PP985_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode, "PS_PP985_SUB_02"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode, "PS_PP985_SUB_03"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode, "PS_PP985_SUB_04"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode, "PS_PP985_SUB_05"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode, "PS_PP985_SUB_06"));

				formHelpClass.CrystalReportOpen(dataPackParameter, dataPackSubReportParameter, WinTitle, ReportName);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Form Item Event
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">pVal</param>
		/// <param name="BubbleEvent">Bubble Event</param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			switch (pVal.EventType)
			{
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                   // Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                   // Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
            }
		}

		/// <summary>
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				if (PVal.BeforeAction == true)
				{
					if (PVal.ItemUID == "BtnAccm")
					{
						if (PS_PP985_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							PS_PP985_Accumulate();
						}
					}

					if (PVal.ItemUID == "BtnSearch")
					{
						if (PS_PP985_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							PS_PP985_Select0(); //메인 조회
							PS_PP985_Select1(); //자재비 조회
							PS_PP985_Select2(); //자체가공비 조회
							PS_PP985_Select3(); //외주가공비 조회
							PS_PP985_Select4(); //외주제작비 조회
							PS_PP985_Select5(); //설계비 조회
							PS_PP985_Select6(); //기성매출 조회
						}
					}

					if (PVal.ItemUID == "BtnPrint")
					{
						if (PS_PP985_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP985_PrintReport);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}

				}
				else if (PVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				if (PVal.BeforeAction == true)
				{
					if (PVal.ItemUID == "Opt01")
					{
						oForm.Freeze(true);
						oForm.Settings.MatrixUID = "Mat01";
						oForm.Settings.EnableRowFormat = true;
						oForm.Settings.Enabled = true;
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
						oMat03.AutoResizeColumns();
						oMat04.AutoResizeColumns();
						oMat05.AutoResizeColumns();
						oMat06.AutoResizeColumns();
						oForm.Freeze(false);
					}
					if (PVal.ItemUID == "Opt02")
					{
						oForm.Freeze(true);
						oForm.Settings.MatrixUID = "Mat02";
						oForm.Settings.EnableRowFormat = true;
						oForm.Settings.Enabled = true;
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
						oMat03.AutoResizeColumns();
						oMat04.AutoResizeColumns();
						oMat05.AutoResizeColumns();
						oMat06.AutoResizeColumns();
						oForm.Freeze(false);
					}
					if (PVal.ItemUID == "Opt03")
					{
						oForm.Freeze(true);
						oForm.Settings.MatrixUID = "Mat03";
						oForm.Settings.EnableRowFormat = true;
						oForm.Settings.Enabled = true;
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
						oMat03.AutoResizeColumns();
						oMat04.AutoResizeColumns();
						oMat05.AutoResizeColumns();
						oMat06.AutoResizeColumns();
						oForm.Freeze(false);
					}
					if (PVal.ItemUID == "Opt04")
					{
						oForm.Freeze(true);
						oForm.Settings.MatrixUID = "Mat04";
						oForm.Settings.EnableRowFormat = true;
						oForm.Settings.Enabled = true;
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
						oMat03.AutoResizeColumns();
						oMat04.AutoResizeColumns();
						oMat05.AutoResizeColumns();
						oMat06.AutoResizeColumns();
						oForm.Freeze(false);
					}
					if (PVal.ItemUID == "Opt05")
					{
						oForm.Freeze(true);
						oForm.Settings.MatrixUID = "Mat05";
						oForm.Settings.EnableRowFormat = true;
						oForm.Settings.Enabled = true;
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
						oMat03.AutoResizeColumns();
						oMat04.AutoResizeColumns();
						oMat05.AutoResizeColumns();
						oMat06.AutoResizeColumns();
						oForm.Freeze(false);
					}
					if (PVal.ItemUID == "Opt06")
					{
						oForm.Freeze(true);
						oForm.Settings.MatrixUID = "Mat06";
						oForm.Settings.EnableRowFormat = true;
						oForm.Settings.Enabled = true;
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
						oMat03.AutoResizeColumns();
						oMat04.AutoResizeColumns();
						oMat05.AutoResizeColumns();
						oMat06.AutoResizeColumns();
						oForm.Freeze(false);
					}

					if (PVal.ItemUID == "Mat01")
					{
						if (PVal.Row > 0)
						{
							oMat01.SelectRow(PVal.Row, true, false);
						}
					}
					else if (PVal.ItemUID == "Mat02")
					{
						if (PVal.Row > 0)
						{
							oMat02.SelectRow(PVal.Row, true, false);
						}
					}
					else if (PVal.ItemUID == "Mat03")
					{
						if (PVal.Row > 0)
						{
							oMat03.SelectRow(PVal.Row, true, false);
						}
					}
					else if (PVal.ItemUID == "Mat04")
					{
						if (PVal.Row > 0)
						{
							oMat04.SelectRow(PVal.Row, true, false);
						}
					}
					else if (PVal.ItemUID == "Mat05")
					{
						if (PVal.Row > 0)
						{
							oMat05.SelectRow(PVal.Row, true, false);
						}
					}
					else if (PVal.ItemUID == "Mat06")
					{
						if (PVal.Row > 0)
						{
							oMat06.SelectRow(PVal.Row, true, false);
						}
					}
				}
				else if (PVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				if (PVal.Before_Action == true)
				{
				}
				else if (PVal.Before_Action == false)
				{
					if (PVal.ItemChanged == true)
					{
						PS_PP985_FlushToItemValue(PVal.ItemUID, 0, "");
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				if (PVal.BeforeAction == true)
				{
					if (PVal.ItemUID == "Mat01")
					{
						if (PVal.Row == 0)
						{
							oMat01.Columns.Item(PVal.ColUID).TitleObject.Sortable = true;
							oMat01.FlushToDataSource();
						}
					}
				}
				else if (PVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (PVal.BeforeAction == true)
				{
					if (PVal.ItemChanged == true)
					{
						PS_PP985_FlushToItemValue(PVal.ItemUID, 0, "");
					}
				}
				else if (PVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
            {
				oForm.Freeze(false);
			}
		}

		/// Raise_EVENT_FORM_UNLOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			if (pVal.Before_Action == true)
			{
			}
			else if (pVal.Before_Action == false)
			{
				SubMain.Remove_Forms(oFormUniqueID);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat04);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat05);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat06);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP985A);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP985B);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP985C);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP985D);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP985E);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP985F);
			}
		}

		/// <summary>
		/// FormMenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1285": //복원
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							oForm.Freeze(true);
							if (oForm.Settings.MatrixUID == "Mat01")
							{
								PS_PP985_AddMatrixRow01(oMat01.VisualRowCount, false);
							}
							oForm.Freeze(false);
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1285": //복원
							break;
						case "1293": //행삭제
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							if (oForm.Settings.MatrixUID == "Mat01")
							{
								oDS_PS_PP985A.RemoveRecord(oDS_PS_PP985A.Size - 1);
								oMat01.LoadFromDataSource();
							}
							oForm.Freeze(false);
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
