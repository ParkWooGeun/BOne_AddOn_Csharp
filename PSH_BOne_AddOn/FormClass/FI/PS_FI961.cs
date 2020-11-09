using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 월별계정별비용현황
	/// </summary>
	internal class PS_FI961 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.DBDataSource oDS_PS_FI961L;  //라인1
		private SAPbouiCOM.DBDataSource oDS_PS_FI961M;  //라인2
		private string oLastItemUID01;          //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;           //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;              //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI961.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI961_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI961");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_FI961_CreateItems();
				PS_FI961_ComboBox_Setting();
				PS_FI961_Initial_Setting();
				PS_FI961_SetDocument(oFormDocEntry01);

				oForm.Items.Item("Folder01").Specific.Select();				//폼이 로드 될 때 Folder01이 선택됨
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc01); //메모리 해제
			}
		}

		/// <summary>
		/// PS_FI961_CreateItems
		/// </summary>
		private void PS_FI961_CreateItems()
		{
			try
			{
				oForm.Freeze(true);
				oDS_PS_FI961L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_FI961M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				//매트릭스 초기화
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				//사업장_S
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//기간 시작_S
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//기간 종료_S
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//계정과목코드_S
				oForm.DataSources.UserDataSources.Add("AcctCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("AcctCode").Specific.DataBind.SetBound(true, "", "AcctCode");

				//계정과목명_S
				oForm.DataSources.UserDataSources.Add("AcctName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("AcctName").Specific.DataBind.SetBound(true, "", "AcctName");
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
		/// PS_FI961_ComboBox_Setting
		/// </summary>
		private void PS_FI961_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
		/// PS_FI961_Initial_Setting
		/// </summary>
		private void PS_FI961_Initial_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장 사용자의 소속 사업장 선택
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//날짜 설정
				oForm.Items.Item("ToDt").Specific.Value = DateTime.Now.ToString("yyyy") + "1231";
				oForm.Items.Item("FrDt").Specific.Value = DateTime.Now.ToString("yyyy") + "0101";
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
		/// PS_FI961_FormItemEnabled
		/// </summary>
		private void PS_FI961_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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
		/// PS_FI961_AddMatrixRow1
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_FI961_AddMatrixRow1(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				//행추가여부
				if (RowIserted == false)
				{
					oDS_PS_FI961L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_FI961L.Offset = oRow;
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
		/// PS_FI961_AddMatrixRow2
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_FI961_AddMatrixRow2(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				//행추가여부
				if (RowIserted == false)
				{
					oDS_PS_FI961M.InsertRecord(oRow);
				}
				oMat02.AddRow();
				oDS_PS_FI961M.Offset = oRow;
				oMat02.LoadFromDataSource();
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
		/// PS_FI961_FormClear
		/// </summary>
		private void PS_FI961_FormClear()
		{
			string DocEntry = String.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FI961'", "");
				if (Convert.ToDouble(DocEntry) == 0)
				{
					oForm.Items.Item("DocEntry").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
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
		/// PS_FI961_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		private void PS_FI961_SetDocument(string oFormDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry01))
				{
					PS_FI961_FormItemEnabled();
				}
				else
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
		/// 배부처별 조회
		/// </summary>
		private void PS_FI961_MTX01()
		{
			int loopCount = 0;
			int ErrNum = 0;
			string sQry = String.Empty;

			string BPLID = string.Empty;            //사업장
			string FrDt = string.Empty;         //기간시작
			string ToDt = string.Empty;         //기간종료
			string AcctCode = string.Empty;         //계정과목코드

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				AcctCode = oForm.Items.Item("AcctCode").Specific.Value.ToString().Trim();

				oForm.Freeze(true);

				sQry = "EXEC PS_FI961_01 '" + BPLID + "','" + FrDt + "','" + ToDt + "','" + AcctCode + "'";
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
						oDS_PS_FI961L.InsertRecord(loopCount);
					}
					oDS_PS_FI961L.Offset = loopCount;

					oDS_PS_FI961L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                    //라인번호
					oDS_PS_FI961L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("AcctCode01").Value);        //계정코드
					oDS_PS_FI961L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("AcctName01").Value);        //계정명
					oDS_PS_FI961L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("AcctCode02").Value);        //계정과목코드
					oDS_PS_FI961L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("AcctName02").Value);        //계정과목명
					oDS_PS_FI961L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("AcctCode03").Value);        //세부계정과목코드
					oDS_PS_FI961L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("AcctName03").Value);        //세부계정과목명
					oDS_PS_FI961L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("TeamName").Value);          //배부처(팀)
					oDS_PS_FI961L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("RspName").Value);           //배부처(담당)
					oDS_PS_FI961L.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("Month01").Value);           //1월
					oDS_PS_FI961L.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("Month02").Value);           //2월
					oDS_PS_FI961L.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("Month03").Value);           //3월
					oDS_PS_FI961L.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("Month04").Value);           //4월
					oDS_PS_FI961L.SetValue("U_ColSum05", loopCount, oRecordSet.Fields.Item("Month05").Value);           //5월
					oDS_PS_FI961L.SetValue("U_ColSum06", loopCount, oRecordSet.Fields.Item("Month06").Value);           //6월
					oDS_PS_FI961L.SetValue("U_ColSum07", loopCount, oRecordSet.Fields.Item("Month07").Value);           //7월
					oDS_PS_FI961L.SetValue("U_ColSum08", loopCount, oRecordSet.Fields.Item("Month08").Value);           //8월
					oDS_PS_FI961L.SetValue("U_ColSum09", loopCount, oRecordSet.Fields.Item("Month09").Value);           //9월
					oDS_PS_FI961L.SetValue("U_ColSum10", loopCount, oRecordSet.Fields.Item("Month10").Value);           //10월
					oDS_PS_FI961L.SetValue("U_ColSum11", loopCount, oRecordSet.Fields.Item("Month11").Value);           //11월
					oDS_PS_FI961L.SetValue("U_ColSum12", loopCount, oRecordSet.Fields.Item("Month12").Value);           //12월
					oDS_PS_FI961L.SetValue("U_ColSum13", loopCount, oRecordSet.Fields.Item("Total").Value);             //계

					oRecordSet.MoveNext();
					ProgressBar01.Value = ProgressBar01.Value + 1;
					ProgressBar01.Text = "배부처별 집계 " + ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();

			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				oForm.Freeze(false);
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// 계정과목별 조회
		/// </summary>
		private void PS_FI961_MTX02()
		{
			int loopCount = 0;
			int ErrNum = 0;
			string sQry = String.Empty;

			string BPLID = string.Empty;            //사업장
			string FrDt = string.Empty;         //기간시작
			string ToDt = string.Empty;         //기간종료
			string AcctCode = string.Empty;         //계정과목코드

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				AcctCode = oForm.Items.Item("AcctCode").Specific.Value.ToString().Trim();

				oForm.Freeze(true);

				sQry = "EXEC PS_FI961_02 '" + BPLID + "','" + FrDt + "','" + ToDt + "','" + AcctCode + "'";
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
						oDS_PS_FI961M.InsertRecord(loopCount);
					}
					oDS_PS_FI961M.Offset = loopCount;

					oDS_PS_FI961M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                    //라인번호
					oDS_PS_FI961M.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("AcctCode01").Value);                    //계정코드
					oDS_PS_FI961M.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("AcctName01").Value);                    //계정명
					oDS_PS_FI961M.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("AcctCode02").Value);                    //계정과목코드
					oDS_PS_FI961M.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("AcctName02").Value);                    //계정과목명
					oDS_PS_FI961M.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("AcctCode03").Value);                    //세부계정과목코드
					oDS_PS_FI961M.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("AcctName03").Value);                    //세부계정과목명
					oDS_PS_FI961M.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("Month01").Value);                   //1월
					oDS_PS_FI961M.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("Month02").Value);                   //2월
					oDS_PS_FI961M.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("Month03").Value);                   //3월
					oDS_PS_FI961M.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("Month04").Value);                   //4월
					oDS_PS_FI961M.SetValue("U_ColSum05", loopCount, oRecordSet.Fields.Item("Month05").Value);                   //5월
					oDS_PS_FI961M.SetValue("U_ColSum06", loopCount, oRecordSet.Fields.Item("Month06").Value);                   //6월
					oDS_PS_FI961M.SetValue("U_ColSum07", loopCount, oRecordSet.Fields.Item("Month07").Value);                   //7월
					oDS_PS_FI961M.SetValue("U_ColSum08", loopCount, oRecordSet.Fields.Item("Month08").Value);                   //8월
					oDS_PS_FI961M.SetValue("U_ColSum09", loopCount, oRecordSet.Fields.Item("Month09").Value);                   //9월
					oDS_PS_FI961M.SetValue("U_ColSum10", loopCount, oRecordSet.Fields.Item("Month10").Value);                   //10월
					oDS_PS_FI961M.SetValue("U_ColSum11", loopCount, oRecordSet.Fields.Item("Month11").Value);                   //11월
					oDS_PS_FI961M.SetValue("U_ColSum12", loopCount, oRecordSet.Fields.Item("Month12").Value);                   //12월
					oDS_PS_FI961M.SetValue("U_ColSum13", loopCount, oRecordSet.Fields.Item("Total").Value);                 //계

					oRecordSet.MoveNext();
					ProgressBar01.Value = ProgressBar01.Value + 1;
					ProgressBar01.Text = "계정과목별 집계 " + ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				oForm.Freeze(false);
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_FI961_GetDetail
		/// </summary>
		private void PS_FI961_GetDetail()
		{
			string BPLID = string.Empty;
			string FrDt = string.Empty;
			string ToDt = string.Empty;
			string AcctCode03 = string.Empty;           //세부계정과목코드

			PS_FI962 oTempClass = new PS_FI962();

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();

				if (oForm.PaneLevel == 1)
				{
					AcctCode03 = oMat02.Columns.Item("AcctCode03").Cells.Item(oLastColRow01).Specific.Value.ToString().Trim();
				}
				else
				{
					AcctCode03 = oMat01.Columns.Item("AcctCode03").Cells.Item(oLastColRow01).Specific.Value.ToString().Trim();
				}

				oTempClass.LoadForm(BPLID, FrDt, ToDt, AcctCode03);
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
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				switch (pVal.EventType)
				{
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
						Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						//2
						Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					//5
						//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:						    //6
						Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:					//7
						Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:			//8
						//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:						//10
						Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:					//11
						Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:					//18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:				//19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:					//20
						//Raise_EVENT_RESIZE(FormUID, pVal, BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				//27
						//Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:						//3
						Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:						//4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:					//17
						Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
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
							PS_FI961_MTX02();                           //계정과목별 집계
							PS_FI961_MTX01();                           //배부처별 집계
						}
					}
					else if (pVal.ItemUID == "BtnPrint01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							// PS_FI961_Print_Report01();
						}
					}
					else if (pVal.ItemUID == "BtnPrint02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							// PS_FI961_Print_Report02();
						}
					}
					else if (pVal.ItemUID == "BtnDetail")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_FI961_GetDetail();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스_S
					//Folder01이 선택되었을 때
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
					}
					//Folder02가 선택되었을 때
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
					}
					//폴더를 사용할 때는 필수 소스_E
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "AcctCode", "");                  //계정코드 포맷서치 활성
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
						//            oLastColUID01 = ""
						//            oLastColRow01 = 0
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
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row == 0)
						{
							oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat01.FlushToDataSource();
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
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "AcctCode")
						{
							sQry = "SELECT AcctName FROM [OACT] WHERE AcctCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("AcctName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_FI961_FormItemEnabled();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI961L); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI961M); //메모리 해제
					SubMain.Remove_Forms(oFormUniqueID01);
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
		/// Raise_FormMenuEvent
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
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행삭제
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							break;

						case "7169":							//엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							if (oForm.PaneLevel == 1)
							{
								PS_FI961_AddMatrixRow1(oMat01.VisualRowCount, false);
							}
							else
							{
								PS_FI961_AddMatrixRow2(oMat02.VisualRowCount, false);
							}
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행삭제
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							break;
						case "7169":							//엑셀 내보내기
							//엑셀 내보내기 이후 처리
							if (oForm.PaneLevel == 1)
							{
								oForm.Freeze(true);
								oDS_PS_FI961L.RemoveRecord(oDS_PS_FI961L.Size - 1);
								oMat01.LoadFromDataSource();
								oForm.Freeze(false);
							}
							else
							{
								oForm.Freeze(true);
								oDS_PS_FI961M.RemoveRecord(oDS_PS_FI961M.Size - 1);
								oMat02.LoadFromDataSource();
								oForm.Freeze(false);
							}
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

		/// <summary>
		/// Raise_FormDataEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="BusinessObjectInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			try
			{
				if (BusinessObjectInfo.BeforeAction == true)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:							//33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:							//34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:						//35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:						//36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:							//33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:							//34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:						//35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:						//36
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
	}
}
