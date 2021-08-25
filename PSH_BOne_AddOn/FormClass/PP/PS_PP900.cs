using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품목별 공정 진행현황
	/// </summary>
	internal class PS_PP900 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.DBDataSource oDS_PS_PP900H;  //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP900L;  //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP900M;  //세부사항라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP900.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP900_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP900");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP900_CreateItems();
				PS_PP900_SetComboBox();
				PS_PP900_Initialize();
				PS_PP900_ResizeForm();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
			}
		}

		/// <summary>
		/// PS_PP900_CreateItems
		/// </summary>
		private void PS_PP900_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_PP900H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_PP900L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_PP900M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				//매트릭스 초기화
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//품목대분류
				oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemCls").Specific.DataBind.SetBound(true, "", "ItemCls");

				//작지등록일시작
				oForm.DataSources.UserDataSources.Add("RgFrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("RgFrDt").Specific.DataBind.SetBound(true, "", "RgFrDt");

				//작지등록일종료
				oForm.DataSources.UserDataSources.Add("RgToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("RgToDt").Specific.DataBind.SetBound(true, "", "RgToDt");

				//거래처코드
				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

				//거래처명
				oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

				//품목코드
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//품목명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//생산완료여부
				oForm.DataSources.UserDataSources.Add("PrdYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("PrdYN").Specific.DataBind.SetBound(true, "", "PrdYN");

				//기준일자
				oForm.DataSources.UserDataSources.Add("StdDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("StdDt").Specific.DataBind.SetBound(true, "", "StdDt");
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

		/// <summary>
		/// PS_PP900_SetComboBox
		/// </summary>
		private void PS_PP900_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				//사업장
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//품목대분류
				oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN ('105','106') order by Code", "", false, false);
				oForm.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//품목구분
				oForm.Items.Item("ItemCls").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemCls").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", false, false);
				oForm.Items.Item("ItemCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//생산완료여부
				oForm.Items.Item("PrdYN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("PrdYN").Specific.ValidValues.Add("Y", "생산완료");
				oForm.Items.Item("PrdYN").Specific.ValidValues.Add("N", "생산미완료");
				oForm.Items.Item("PrdYN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//품목대분류(매트릭스)
				dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");

				//품목구분
				dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemCls"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", "");

				//작업일보여부
				dataHelpClass.Combo_ValidValues_Insert("PS_PP900", "Mat02", "ReportYN", "Y", "예");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP900", "Mat02", "ReportYN", "N", "아니오");
				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat02.Columns.Item("ReportYN"), "PS_PP900", "Mat02", "ReportYN", false);
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

		/// <summary>
		/// PS_PP900_Initialize
		/// </summary>
		private void PS_PP900_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //사업장 사용자의 소속 사업장 선택
				oForm.Items.Item("StdDt").Visible = false; //기준일자
				oForm.Items.Item("Static08").Visible = false;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// PS_PP900_ResizeForm
		/// </summary>
		private void PS_PP900_ResizeForm()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Mat01").Height + 20;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Mat01").Width + 25;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// PS_PP900_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP900_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_PP900L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_PP900L.Offset = oRow;
				oMat01.LoadFromDataSource();
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

		/// <summary>
		/// PS_PP900_MTX01
		/// </summary>
		private void PS_PP900_MTX01()
		{
			int ErrNum = 0;
			int loopCount;
			string sQry;

			string BPLID;           //사업장
			string ItmBsort;        //품목대분류
			string ItemCls;         //품목구분
			string RgFrDt;          //작지등록일시작
			string RgToDt;          //작지등록일종료
			string CardCode;        //거래처
			string ItemCode;        //품목코드
			string PrdYN;           //생산완료여부

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim();
				ItemCls = oForm.Items.Item("ItemCls").Specific.Selected.Value.ToString().Trim();
				RgFrDt = oForm.Items.Item("RgFrDt").Specific.Value.ToString().Trim();
				RgToDt = oForm.Items.Item("RgToDt").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				PrdYN = oForm.Items.Item("PrdYN").Specific.Selected.Value.ToString().Trim();

				oForm.Freeze(true);

				sQry = "EXEC PS_PP900_01 '" + BPLID + "','" + ItmBsort + "','" + ItemCls + "','" + RgFrDt + "','" + RgToDt + "','" + CardCode + "','" + ItemCode + "','" + PrdYN + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat01.Clear();
					ErrNum = 1;
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP900L.InsertRecord(loopCount);
					}
					oDS_PS_PP900L.Offset = loopCount;

					oDS_PS_PP900L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
					oDS_PS_PP900L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim()); //거래처코드
					oDS_PS_PP900L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("CardName").Value.ToString().Trim()); //거래처명
					oDS_PS_PP900L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim()); //작번
					oDS_PS_PP900L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim()); //품목코드
					oDS_PS_PP900L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
					oDS_PS_PP900L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("Spec").Value.ToString().Trim()); //규격
					oDS_PS_PP900L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ItmBsort").Value.ToString().Trim()); //품목대분류
					oDS_PS_PP900L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("ItemCls").Value.ToString().Trim()); //품목구분
					oDS_PS_PP900L.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("FirPrcNm").Value.ToString().Trim()); //최초공정
					oDS_PS_PP900L.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("FirDt").Value.ToString().Trim()); //완료요구일(최초)
					oDS_PS_PP900L.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("CurPrcNm").Value.ToString().Trim()); //현재공정
					oDS_PS_PP900L.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("CurDt").Value.ToString().Trim()); //완료요구일(현재)
					oDS_PS_PP900L.SetValue("U_ColReg13", loopCount, oRecordSet.Fields.Item("WPRgDt").Value.ToString().Trim()); //작업일보등록일
					oDS_PS_PP900L.SetValue("U_ColReg14", loopCount, oRecordSet.Fields.Item("WODate").Value.ToString().Trim()); //작업지시완료일
					oDS_PS_PP900L.SetValue("U_ColReg16", loopCount, oRecordSet.Fields.Item("EndPrcNm").Value.ToString().Trim()); //최종공정
					oDS_PS_PP900L.SetValue("U_ColReg17", loopCount, oRecordSet.Fields.Item("EndDt").Value.ToString().Trim()); //완료요구일(최종)
					oDS_PS_PP900L.SetValue("U_ColReg18", loopCount, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim()); //작업지시문서번호

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP900_MTX02
		/// </summary>
		/// <param name="prmDocEntry"></param>
		private void PS_PP900_MTX02(int prmDocEntry)
		{
			int loopCount;
			int ErrNum = 0;
			string sQry;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 0, false);

			try
			{
				oForm.Freeze(true);

				sQry = "EXEC PS_PP900_02 " + prmDocEntry;
				oRecordSet.DoQuery(sQry);

				oMat02.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat02.Clear();
					ErrNum = 1;
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP900M.InsertRecord(loopCount);
					}
					oDS_PS_PP900M.Offset = loopCount;

					oDS_PS_PP900M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
					oDS_PS_PP900M.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("PrcCd").Value.ToString().Trim()); //공정코드
					oDS_PS_PP900M.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("PrcNm").Value.ToString().Trim()); //공정명
					oDS_PS_PP900M.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("StdWT").Value.ToString().Trim()); //표준공수
					oDS_PS_PP900M.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("StdDt").Value.ToString().Trim()); //완료요구일
					oDS_PS_PP900M.SetValue("U_ColQty02", loopCount, oRecordSet.Fields.Item("WorkWT").Value.ToString().Trim()); //실동공수
					oDS_PS_PP900M.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("WorkDt").Value.ToString().Trim()); //등록일
					oDS_PS_PP900M.SetValue("U_ColQty03", loopCount, oRecordSet.Fields.Item("PrcPnt").Value.ToString().Trim()); //공정진행율
					oDS_PS_PP900M.SetValue("U_ColQty04", loopCount, oRecordSet.Fields.Item("TotPnt").Value.ToString().Trim()); //전체진행율
					oDS_PS_PP900M.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ReportYN").Value.ToString().Trim()); //작업일보여부

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
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
			switch (pVal.EventType) {
				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:				//1
					Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:					//2
					Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:				//5
					//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK:					    //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:				//7
					Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:		//8
					Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE:					//10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:				//11
					//Raise_EVENT_MATRIX_LOAD(rFormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:				//18
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:			//19
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:				//20
					Raise_EVENT_RESIZE(FormUID, pVal, BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:			//27
					//Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:					//3
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:					//4
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:				//17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
			}
		}

		/// <summary>
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							oForm.Items.Item("Folder01").Specific.Select();		//공정현황조회 선택
							oForm.Items.Item("ItemCode2").Specific.Value = "";	//품목코드
							oForm.Items.Item("ItemName2").Specific.Value = "";	//품목명
							oForm.Items.Item("Spec").Specific.Value = "";		//규격
							oForm.Items.Item("OrdNum").Specific.Value = "";		//작번
							oForm.Items.Item("WODocNum").Specific.Value = "";	//작업지시문서번호
							oForm.Items.Item("WODate").Specific.Value = "";		//작업지시완료일

							oMat02.Clear();		//세부현황 매트릭스 초기화
							PS_PP900_MTX01();	//매트릭스에 데이터 로드
						}
					}
					else if (pVal.ItemUID == "Link04")
					{
						PS_PP030 oTempClass = new PS_PP030();
						oTempClass.LoadForm(oForm.Items.Item("WODocNum").Specific.Value.ToString().Trim());
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Folder01")  //Folder01이 선택되었을 때
					{
						oForm.PaneLevel = 1;
					}
					else if (pVal.ItemUID == "Folder02")  //Folder02가 선택되었을 때
					{
						oForm.PaneLevel = 2;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //거래처코드 포맷서치 활성
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); //품목코드(작번) 포맷서치 활성
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat01.SelectRow(pVal.Row, true, false);
							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;
						}
					}
					else if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row > 0)
						{
							oMat02.SelectRow(pVal.Row, true, false);
							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;
						}
					}
					else
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = "";
						oLastColRow01 = 0;
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row == 0)
						{
							oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true; //정렬
							oMat01.FlushToDataSource();
						}
						else
						{
							//품목별 공정 세부현황 탭 선택
							oForm.Items.Item("Folder02").Specific.Select();
							oForm.Items.Item("ItemCode2").Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(); //품목코드
							oForm.Items.Item("ItemName2").Specific.Value = oMat01.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(); //품목명
							oForm.Items.Item("Spec").Specific.Value = oMat01.Columns.Item("Spec").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(); //규격
							oForm.Items.Item("OrdNum").Specific.Value = oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(); //작번
							oForm.Items.Item("WODocNum").Specific.Value = oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(); //작업지시문서번호
							oForm.Items.Item("WODate").Specific.Value = oMat01.Columns.Item("WODate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(); //작업지시완료일

							PS_PP900_MTX02(Convert.ToInt32(oForm.Items.Item("WODocNum").Specific.Value));  //공정세부현황조회
						}
					}
				}
				else if (pVal.BeforeAction == false)
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

		/// <summary>
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					PS_PP030 oTempClass = new PS_PP030();
					oTempClass.LoadForm(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							sQry = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "CntcCode")
						{
							sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RESIZE(string FormUID, SAPbouiCOM.ItemEvent pVal, bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP900_ResizeForm();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = pVal.ColUID;
						oLastColRow01 = pVal.Row;
					}
				}
				else
				{
					oLastItemUID01 = pVal.ItemUID;
					oLastColUID01 = "";
					oLastColRow01 = 0;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_UNLOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP900H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP900L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP900M);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							PS_PP900_AddMatrixRow(oMat01.VisualRowCount, false); //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
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
							oDS_PS_PP900L.RemoveRecord(oDS_PS_PP900L.Size - 1);
							oMat01.LoadFromDataSource();
							oForm.Freeze(false);
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}
	}
}
