using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 실패비용 관리(기계)
	/// </summary>
	internal class PS_PP250 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Matrix oMat03;
		private SAPbouiCOM.DBDataSource oDS_PS_PP250L; //라인(작업일보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP250M; //라인(구매요청)
		private SAPbouiCOM.DBDataSource oDS_PS_PP250N; //라인(공용)
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP250.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}
				oFormUniqueID = "PS_PP250_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP250");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP250_CreateItems();
				PS_PP250_SetComboBox();
				PS_PP250_ResizeForm();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Items.Item("Folder01").Specific.Select();	//폼이 로드 될 때 Folder01이 선택됨
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP250_CreateItems
		/// </summary>
		private void PS_PP250_CreateItems()
		{
			try
			{
				oDS_PS_PP250L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_PP250M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
				oDS_PS_PP250N = oForm.DataSources.DBDataSources.Item("@PS_USERDS03");

				//매트릭스 초기화
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				oMat03 = oForm.Items.Item("Mat03").Specific;
				oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat03.AutoResizeColumns();

				//작업일보
				//작업일보기간(Fr)
				oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");
				oForm.DataSources.UserDataSources.Item("FrDt01").Value = DateTime.Now.ToString("yyyyMM01");

				//작업일보기간(To)
				oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");
				oForm.DataSources.UserDataSources.Item("ToDt01").Value = DateTime.Now.ToString("yyyyMMdd");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 11);
				oForm.Items.Item("OrdNum01").Specific.DataBind.SetBound(true, "", "OrdNum01");

				//작명
				oForm.DataSources.UserDataSources.Add("OrdName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("OrdName01").Specific.DataBind.SetBound(true, "", "OrdName01");

				//규격
				oForm.DataSources.UserDataSources.Add("OrdSpec01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("OrdSpec01").Specific.DataBind.SetBound(true, "", "OrdSpec01");

				//공정
				oForm.DataSources.UserDataSources.Add("CpCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode01").Specific.DataBind.SetBound(true, "", "CpCode01");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName01").Specific.DataBind.SetBound(true, "", "CpName01");

				//작업자
				oForm.DataSources.UserDataSources.Add("WorkCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("WorkCode01").Specific.DataBind.SetBound(true, "", "WorkCode01");

				//작업자성명
				oForm.DataSources.UserDataSources.Add("WorkName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("WorkName01").Specific.DataBind.SetBound(true, "", "WorkName01");

				//작업상태
				oForm.DataSources.UserDataSources.Add("WorkCls01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("WorkCls01").Specific.DataBind.SetBound(true, "", "WorkCls01");

				//연간품제외
				oForm.DataSources.UserDataSources.Add("YearPDYN01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("YearPDYN01").Specific.DataBind.SetBound(true, "", "YearPDYN01");
				oForm.Items.Item("YearPDYN01").Specific.Checked = false;

				//구매요청
				//청구기간(Fr)
				oForm.DataSources.UserDataSources.Add("FrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt02").Specific.DataBind.SetBound(true, "", "FrDt02");
				oForm.DataSources.UserDataSources.Item("FrDt02").Value = DateTime.Now.ToString("yyyyMM01");

				//청구기간(To)
				oForm.DataSources.UserDataSources.Add("ToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt02").Specific.DataBind.SetBound(true, "", "ToDt02");
				oForm.DataSources.UserDataSources.Item("ToDt02").Value = DateTime.Now.ToString("yyyyMMdd");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 11);
				oForm.Items.Item("OrdNum02").Specific.DataBind.SetBound(true, "", "OrdNum02");

				//작명
				oForm.DataSources.UserDataSources.Add("OrdName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("OrdName02").Specific.DataBind.SetBound(true, "", "OrdName02");

				//규격
				oForm.DataSources.UserDataSources.Add("OrdSpec02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("OrdSpec02").Specific.DataBind.SetBound(true, "", "OrdSpec02");

				//청구자
				oForm.DataSources.UserDataSources.Add("CntcCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode02").Specific.DataBind.SetBound(true, "", "CntcCode02");

				//청구자성명
				oForm.DataSources.UserDataSources.Add("CntcName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName02").Specific.DataBind.SetBound(true, "", "CntcName02");

				//품의구분
				oForm.DataSources.UserDataSources.Add("POType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("POType02").Specific.DataBind.SetBound(true, "", "POType02");

				//청구사유
				oForm.DataSources.UserDataSources.Add("OrdCls02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OrdCls02").Specific.DataBind.SetBound(true, "", "OrdCls02");

				//연간품제외
				oForm.DataSources.UserDataSources.Add("YearPDYN02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("YearPDYN02").Specific.DataBind.SetBound(true, "", "YearPDYN02");
				oForm.Items.Item("YearPDYN02").Specific.Checked = false;

				//공용
				//공용기간(Fr)
				oForm.DataSources.UserDataSources.Add("FrDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt03").Specific.DataBind.SetBound(true, "", "FrDt03");
				oForm.DataSources.UserDataSources.Item("FrDt03").Value = DateTime.Now.ToString("yyyyMM01");

				//공용기간(To)
				oForm.DataSources.UserDataSources.Add("ToDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt03").Specific.DataBind.SetBound(true, "", "ToDt03");
				oForm.DataSources.UserDataSources.Item("ToDt03").Value = DateTime.Now.ToString("yyyyMMdd");

				//사번
				oForm.DataSources.UserDataSources.Add("CntcCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode03").Specific.DataBind.SetBound(true, "", "CntcCode03");

				//성명
				oForm.DataSources.UserDataSources.Add("CntcName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName03").Specific.DataBind.SetBound(true, "", "CntcName03");

				//목적구분
				oForm.DataSources.UserDataSources.Add("ObjCls03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ObjCls03").Specific.DataBind.SetBound(true, "", "ObjCls03");

				//목적
				oForm.DataSources.UserDataSources.Add("Object03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Object03").Specific.DataBind.SetBound(true, "", "Object03");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum03").Specific.DataBind.SetBound(true, "", "OrdNum03");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP250_SetComboBox
		/// </summary>
		private void PS_PP250_SetComboBox()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//작업일보
				//작업상태(헤더)
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'P203'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("WorkCls01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkCls01").Specific, sQry, "%", false, false);

				//작업상태(라인)
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'P203'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCls"), sQry, "", "");

				//구매요청
				//품의구분(헤더)
				sQry = " SELECT      Code, ";
				sQry += "             Name ";
				sQry += " FROM        [@PSH_ORDTYP] ";
				sQry += " ORDER BY    Code";
				oForm.Items.Item("POType02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("POType02").Specific, sQry, "%", false, false);

				//청구사유(헤더)
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'P203'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += "             AND U_Minor <> 'A'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("OrdCls02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OrdCls02").Specific, sQry, "%", false, false);

				//청구사유(라인)
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'P203'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += "             AND U_Minor <> 'A'";
				sQry += " ORDER BY    U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat02.Columns.Item("MM005RCode"), sQry, "", "");

				//공용
				//목적구분(헤더)
				sQry = " SELECT      U_Code,";
				sQry += "             U_CodeNm";
				sQry += " FROM        [@PS_HR200L]";
				sQry += " WHERE       Code = 'P224'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("ObjCls03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ObjCls03").Specific, sQry, "%", false, false);

				//목적구분(라인)
				sQry = " SELECT      U_Code,";
				sQry += "             U_CodeNm";
				sQry += " FROM        [@PS_HR200L]";
				sQry += " WHERE       Code = 'P224'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat03.Columns.Item("ObjCls"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP250_ResizeForm
		/// </summary>
		private void PS_PP250_ResizeForm()
		{
			try
			{
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Mat01").Height + 125;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Mat01").Width + 20;

				oMat01.AutoResizeColumns();
				oMat02.AutoResizeColumns();
				oMat03.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP250_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="prmMat"></param>
		/// <param name="prmDataSource"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP250_AddMatrixRow(int oRow, SAPbouiCOM.Matrix prmMat, SAPbouiCOM.DBDataSource prmDataSource, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				
				if (RowIserted == false) //행추가여부
				{
					prmDataSource.InsertRecord(oRow);
				}
				prmMat.AddRow();
				prmDataSource.Offset = oRow;
				prmMat.LoadFromDataSource();
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
		/// PS_PP250_MTX01 작업일보조회
		/// </summary>
		private void PS_PP250_MTX01()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;

			string FrDt;	 //작업일보일자(시작)
			string ToDt;	 //작업일보일자(종료)
			string OrdNum;	 //작번
			string CpCode;	 //공정
			string WorkCode; //작업자사번
			string WorkCls;	 //작업상태
			string YearPdYN; //연간품제외

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				FrDt     = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();
				OrdNum   = oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim();
				CpCode   = oForm.Items.Item("CpCode01").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode01").Specific.Value.ToString().Trim();
				WorkCls  = oForm.Items.Item("WorkCls01").Specific.Value.ToString().Trim();
				
				if (oForm.DataSources.UserDataSources.Item("YearPDYN01").Value == "Y") //연간품여부
				{
					YearPdYN = "Y";
				}
				else
				{
					YearPdYN = "N";
				}

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = "EXEC PS_PP250_01 '";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += OrdNum + "','";
				sQry += CpCode + "','";
				sQry += WorkCode + "','";
				sQry += WorkCls + "','";
				sQry += YearPdYN + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat01.Clear();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP250L.InsertRecord(loopCount);
					}
					oDS_PS_PP250L.Offset = loopCount;

					oDS_PS_PP250L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));					            //라인번호
					oDS_PS_PP250L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Select").Value.ToString().Trim());	 	//선택
					oDS_PS_PP250L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("PP040Entry").Value.ToString().Trim());	//일보문서번호
					oDS_PS_PP250L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("PP040Line").Value.ToString().Trim());	//일보라인번호
					oDS_PS_PP250L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());		//작번
					oDS_PS_PP250L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());		//서브작번1
					oDS_PS_PP250L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());		//서브작번2
					oDS_PS_PP250L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("OrdName").Value.ToString().Trim());		//작명
					oDS_PS_PP250L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("OrdSpec").Value.ToString().Trim());		//규격
					oDS_PS_PP250L.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("WorkCode").Value.ToString().Trim());	//작업자사번
					oDS_PS_PP250L.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("WorkName").Value.ToString().Trim());	//작업자성명
					oDS_PS_PP250L.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());		//공정코드
					oDS_PS_PP250L.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());		//공정명
					oDS_PS_PP250L.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("LaboTime").Value.ToString().Trim());	//실동공수
					oDS_PS_PP250L.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("Expense").Value.ToString().Trim());		//비용
					oDS_PS_PP250L.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("Cost").Value.ToString().Trim());		//원가
					oDS_PS_PP250L.SetValue("U_ColReg15", loopCount, oRecordSet.Fields.Item("WorkCls").Value.ToString().Trim());		//작업상태
					oDS_PS_PP250L.SetValue("U_ColReg16", loopCount, oRecordSet.Fields.Item("PP030Date").Value.ToString().Trim());	//작지등록일
					oDS_PS_PP250L.SetValue("U_ColReg17", loopCount, oRecordSet.Fields.Item("PP040Date").Value.ToString().Trim());	//일보등록일
					oDS_PS_PP250L.SetValue("U_ColReg18", loopCount, oRecordSet.Fields.Item("PP080Date").Value.ToString().Trim());	//완료등록일
					oDS_PS_PP250L.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("PP080Diff").Value.ToString().Trim());	//일보일-완료일
					oDS_PS_PP250L.SetValue("U_ColReg20", loopCount, oRecordSet.Fields.Item("PP030Entry").Value.ToString().Trim());	//작지문서번호
					oDS_PS_PP250L.SetValue("U_ColReg21", loopCount, oRecordSet.Fields.Item("PP030Line").Value.ToString().Trim());	//작지라인번호
					oDS_PS_PP250L.SetValue("U_ColReg22", loopCount, oRecordSet.Fields.Item("PP030YM").Value.ToString().Trim());		//일보년월
					oDS_PS_PP250L.SetValue("U_ColReg23", loopCount, oRecordSet.Fields.Item("YearPDYN").Value.ToString().Trim());    //연간품여부

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
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
		/// PS_PP250_MTX02 구매요청
		/// </summary>
		private void PS_PP250_MTX02()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;

			string FrDt;     //작업일보일자(시작)
			string ToDt;     //작업일보일자(종료)
			string OrdNum;   //작번
			string CntcCode;   //청구자사번
			string POType; //품의구분
			string OrdCls;  //청구사유
			string YearPdYN; //연간품제외

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				FrDt = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum02").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim();
				POType = oForm.Items.Item("POType02").Specific.Value.ToString().Trim();
				OrdCls = oForm.Items.Item("OrdCls02").Specific.Value.ToString().Trim();

				//연간품여부
				if (oForm.DataSources.UserDataSources.Item("YearPDYN02").Value == "Y")
				{
					YearPdYN = "Y";
				}
				else
				{
					YearPdYN = "N";
				}

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = "EXEC PS_PP250_11 '";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += OrdNum + "','";
				sQry += CntcCode + "','";
				sQry += POType + "','";
				sQry += OrdCls + "','";
				sQry += YearPdYN + "'";
				oRecordSet.DoQuery(sQry);

				oMat02.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat02.Clear();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP250M.InsertRecord(loopCount);
					}
					oDS_PS_PP250M.Offset = loopCount;

					oDS_PS_PP250M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                                //라인번호
					oDS_PS_PP250M.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Select").Value.ToString().Trim());      //선택
					oDS_PS_PP250M.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("MM005Entry").Value.ToString().Trim());  //청구번호
					oDS_PS_PP250M.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("MM005RCode").Value.ToString().Trim());  //청구사유
					oDS_PS_PP250M.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("MM005Reasn").Value.ToString().Trim());  //세부사항
					oDS_PS_PP250M.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("PP030Entry").Value.ToString().Trim());  //작지문서번호
					oDS_PS_PP250M.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("PP030Date").Value.ToString().Trim());   //작지등록일
					oDS_PS_PP250M.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("MM005Date").Value.ToString().Trim());   //구매요청일
					oDS_PS_PP250M.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("POTypeNM").Value.ToString().Trim());    //품의구분
					oDS_PS_PP250M.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("PP080Date").Value.ToString().Trim());   //생산완료일
					oDS_PS_PP250M.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("CreateDate").Value.ToString().Trim());  //등록일(시스템)
					oDS_PS_PP250M.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("MM005SDiff").Value.ToString().Trim());  //시스템일-청구일
					oDS_PS_PP250M.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("PP030Diff").Value.ToString().Trim());   //작지등록일-청구일
					oDS_PS_PP250M.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("PP080Diff").Value.ToString().Trim());   //생산완료일-청구일
					oDS_PS_PP250M.SetValue("U_ColReg14", loopCount, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());    //품목코드
					oDS_PS_PP250M.SetValue("U_ColReg15", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());    //품명
					oDS_PS_PP250M.SetValue("U_ColReg16", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());       //작번
					oDS_PS_PP250M.SetValue("U_ColReg17", loopCount, oRecordSet.Fields.Item("OrdName").Value.ToString().Trim());      //작명
					oDS_PS_PP250M.SetValue("U_ColReg18", loopCount, oRecordSet.Fields.Item("OrdSpec").Value.ToString().Trim());      //규격
					oDS_PS_PP250M.SetValue("U_ColReg19", loopCount, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim());     //청구자사번
					oDS_PS_PP250M.SetValue("U_ColReg20", loopCount, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim());     //청구자성명
					oDS_PS_PP250M.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("MM030Amt").Value.ToString().Trim());     //품의금액
					oDS_PS_PP250M.SetValue("U_ColReg22", loopCount, oRecordSet.Fields.Item("MM030Date").Value.ToString().Trim());    //품의일자
					oDS_PS_PP250M.SetValue("U_ColReg23", loopCount, oRecordSet.Fields.Item("YearPDYN").Value.ToString().Trim());     //연간품여부

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
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
		/// PS_PP250_MTX03 공용조회
		/// </summary>
		private void PS_PP250_MTX03()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;

			string FrDt;     //작업일보일자(시작)
			string ToDt;     //작업일보일자(종료)
			string OrdNum;   //작번
			string CntcCode;   //청구자사번
			string ObjCls; //목적구분
			string Object_Renamed;  //목적

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				FrDt     = oForm.Items.Item("FrDt03").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt03").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim();
				ObjCls   = oForm.Items.Item("ObjCls03").Specific.Value.ToString().Trim();
				Object_Renamed = oForm.Items.Item("Object03").Specific.Value.ToString().Trim();
				OrdNum   = oForm.Items.Item("OrdNum03").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = " EXEC PS_PP250_21 '";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CntcCode + "','";
				sQry += ObjCls + "','";
				sQry += Object_Renamed + "','";
				sQry += OrdNum + "'";

				oRecordSet.DoQuery(sQry);

				oMat03.Clear();
				oMat03.FlushToDataSource();
				oMat03.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat03.Clear();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP250N.InsertRecord(loopCount);
					}
					oDS_PS_PP250N.Offset = loopCount;

					oDS_PS_PP250N.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));				             //라인번호
					oDS_PS_PP250N.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Select").Value.ToString().Trim());	 //선택
					oDS_PS_PP250N.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim()); //관리번호
					oDS_PS_PP250N.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("MSTCOD").Value.ToString().Trim());	 //사원번호
					oDS_PS_PP250N.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("MSTNAM").Value.ToString().Trim());	 //사원성명
					oDS_PS_PP250N.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("FrDate").Value.ToString().Trim());	 //출발일자
					oDS_PS_PP250N.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("FrHour").Value.ToString().Trim());	 //출발시각
					oDS_PS_PP250N.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ToDate").Value.ToString().Trim());	 //도착일자
					oDS_PS_PP250N.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("ToHour").Value.ToString().Trim());	 //도착시각
					oDS_PS_PP250N.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());	 //작번
					oDS_PS_PP250N.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("ObjCls").Value.ToString().Trim());	 //목적구분
					oDS_PS_PP250N.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("Object").Value.ToString().Trim());	 //목적내용
					oDS_PS_PP250N.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("Dest2").Value.ToString().Trim());	 //목적지
					oDS_PS_PP250N.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("TransExp").Value.ToString().Trim()); //교통비
					oDS_PS_PP250N.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("DayExp").Value.ToString().Trim());	 //일비
					oDS_PS_PP250N.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("FoodExp").Value.ToString().Trim());	 //식비
					oDS_PS_PP250N.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("ParkExp").Value.ToString().Trim());	 //주차비
					oDS_PS_PP250N.SetValue("U_ColSum05", loopCount, oRecordSet.Fields.Item("TollExp").Value.ToString().Trim());	 //도로비
					oDS_PS_PP250N.SetValue("U_ColSum06", loopCount, oRecordSet.Fields.Item("TotalExp").Value.ToString().Trim()); //계

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat03.LoadFromDataSource();
				oMat03.AutoResizeColumns();
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
		/// PS_PP250_SaveData01 작업일보의 작업상태 업데이트
		/// </summary>
		private void PS_PP250_SaveData01()
		{
			int loopCount;
			string sQry;

			string PP040Entry;	//일보문서번호
			string PP040Line;	//일보라인번호
			string WorkCls;     //작업상태

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oMat01.FlushToDataSource();

				ProgressBar01.Text = "저장 중...";

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP250L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						PP040Entry = oDS_PS_PP250L.GetValue("U_ColReg02", loopCount).ToString().Trim();
						PP040Line  = oDS_PS_PP250L.GetValue("U_ColReg03", loopCount).ToString().Trim();
						WorkCls    = oDS_PS_PP250L.GetValue("U_ColReg15", loopCount).ToString().Trim();

						sQry = "EXEC [PS_PP250_02] '";
						sQry += PP040Entry + "','";
						sQry += PP040Line + "','";
						sQry += WorkCls + "'";
						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat01.VisualRowCount - 1) + "건 저장중...";
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_PP250_SaveData02 구매요청의 청구사유, 세부사항 업데이트
		/// </summary>
		private void PS_PP250_SaveData02()
		{
			int loopCount;
			string sQry;

			string MM005Entry;			//청구번호
			string MM005RCode;			//청구사유
			string MM005Reasn;          //세부사항

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oMat02.FlushToDataSource();

				ProgressBar01.Text = "저장 중...";

				for (loopCount = 0; loopCount <= oMat02.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP250M.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						MM005Entry = oDS_PS_PP250M.GetValue("U_ColReg02", loopCount).ToString().Trim();
						MM005RCode = oDS_PS_PP250M.GetValue("U_ColReg03", loopCount).ToString().Trim();
						MM005Reasn = oDS_PS_PP250M.GetValue("U_ColReg04", loopCount).ToString().Trim();

						sQry = "EXEC [PS_PP250_12] '";
						sQry += MM005Entry + "','";
						sQry += MM005RCode + "','";
						sQry += MM005Reasn + "'";
						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat02.VisualRowCount - 1) + "건 저장중...";
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_PP250_SaveData03
		/// </summary>
		private void PS_PP250_SaveData03()
		{
			int loopCount;
			string sQry;

			string DocEntry; //공용관리번호
			string OrdNum;	 //작번
			string ObjCls;   //공용목적구분

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oMat03.FlushToDataSource();

				ProgressBar01.Text = "저장 중...";

				for (loopCount = 0; loopCount <= oMat03.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP250N.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						DocEntry = oDS_PS_PP250N.GetValue("U_ColReg02", loopCount).ToString().Trim();
						OrdNum = oDS_PS_PP250N.GetValue("U_ColReg09", loopCount).ToString().Trim();
						ObjCls = oDS_PS_PP250N.GetValue("U_ColReg10", loopCount).ToString().Trim();

						sQry = "EXEC [PS_PP250_22] '";
						sQry += DocEntry + "','";
						sQry += OrdNum + "','";
						sQry += ObjCls + "'";
						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat03.VisualRowCount - 1) + "건 저장중...";
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_PP250_DeleteData02 청구사유삭제
		/// </summary>
		private void PS_PP250_DeleteData02()
		{
			int loopCount;
			string sQry;

			string MM005Entry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat02.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat02.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP250M.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						MM005Entry = oDS_PS_PP250M.GetValue("U_ColReg02", loopCount).ToString().Trim();

						sQry = "EXEC [PS_PP250_13] '";
						sQry += MM005Entry + "'";	//청구번호
						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "BtnSrch01")
					{
						PS_PP250_MTX01(); //매트릭스에 데이터 로드
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						PS_PP250_MTX02();
					}
					else if (pVal.ItemUID == "BtnSrch03")
					{
						PS_PP250_MTX03();
					}
					else if (pVal.ItemUID == "BtnSave01")
					{
						PS_PP250_SaveData01();
					}
					else if (pVal.ItemUID == "BtnSave02")
					{
						PS_PP250_SaveData02();
					}
					else if (pVal.ItemUID == "BtnSave03")
					{
						PS_PP250_SaveData03();
					}
					else if (pVal.ItemUID == "BtnDel02")
					{
						PS_PP250_DeleteData02();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스
					if (pVal.ItemUID == "Folder01")
					{
						oForm.Freeze(true);
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
						oForm.Settings.MatrixUID = "Mat01";
						oMat01.AutoResizeColumns();
						oForm.Freeze(false);
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.Freeze(true);
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
						oForm.Settings.MatrixUID = "Mat02";
						oMat02.AutoResizeColumns();
						oForm.Freeze(false);
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.Freeze(true);
						oForm.PaneLevel = 3;
						oForm.DefButton = "BtnSrch03";
						oForm.Settings.MatrixUID = "Mat03";
						oMat03.AutoResizeColumns();
						oForm.Freeze(false);
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum01", "");   //작업일보-작번
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode01", "");   //작업일보-공정
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "WorkCode01", ""); //작업일보-작업자

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum02", "");   //구매요청-작번
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode02", ""); //구매요청-청구자
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode03", ""); //구매요청-사번
				}
				else if (pVal.BeforeAction == false)
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
					else if (pVal.ItemUID == "Mat03")
					{
						if (pVal.Row > 0)
						{
							oMat03.SelectRow(pVal.Row, true, false);
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
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						//작업일보-작번
						if (pVal.ItemUID == "OrdNum01")
						{
							oForm.Items.Item("OrdName01").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("OrdSpec01").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim() + "'", "");
						}
						else if (pVal.ItemUID == "CpCode01")
						{
							oForm.Items.Item("CpName01").Specific.Value = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item("CpCode01").Specific.Value.ToString().Trim() + "'", "");
						}
						else if(pVal.ItemUID == "WorkCode01")
						{
							oForm.Items.Item("WorkName01").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("WorkCode01").Specific.Value.ToString().Trim() + "'", "");
						}
						else if (pVal.ItemUID == "OrdNum02")
						{
							oForm.Items.Item("OrdName02").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum02").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("OrdSpec02").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum02").Specific.Value.ToString().Trim() + "'", "");
						}
						else if (pVal.ItemUID == "CntcCode02")
						{
							oForm.Items.Item("CntcName02").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim() + "'", "");
						}
						else if (pVal.ItemUID == "CntcCode03")
						{
							oForm.Items.Item("CntcName03").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim() + "'", "");
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP250_ResizeForm();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				else if (pVal.ItemUID == "Mat02")
				{
					if (pVal.Row > 0)
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = pVal.ColUID;
						oLastColRow01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat03")
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
		}

		/// <summary>
		/// FORM_UNLOAD 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP250L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP250M);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP250N);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
				}

				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = pVal.ColUID;
						oLastColRow01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat02")
				{
					if (pVal.Row > 0)
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = pVal.ColUID;
						oLastColRow01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat03")
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
				oForm.Freeze(true);

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
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							PS_PP250_AddMatrixRow(oMat01.VisualRowCount, oMat01, oDS_PS_PP250L, false); //작업일보
							PS_PP250_AddMatrixRow(oMat02.VisualRowCount, oMat02, oDS_PS_PP250M, false); //구매요청
							PS_PP250_AddMatrixRow(oMat03.VisualRowCount, oMat03, oDS_PS_PP250N, false); //공용
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
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1287": //복제
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							oDS_PS_PP250L.RemoveRecord(oDS_PS_PP250L.Size - 1); //작업일보
							oMat01.LoadFromDataSource();
							oDS_PS_PP250M.RemoveRecord(oDS_PS_PP250M.Size - 1); //구매요청
							oMat02.LoadFromDataSource();
							oDS_PS_PP250N.RemoveRecord(oDS_PS_PP250N.Size - 1); //공용
							oMat03.LoadFromDataSource();
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
				oForm.Freeze(false);
			}
		}
	}
}
